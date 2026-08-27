using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace HybridSurvey
{
    internal static partial class HybridCommands
    {
        private const string kOwnerRecord = "Owner";

        internal static bool TryRequireUtmMode(string actionName)
        {
            var document = AcadApp.DocumentManager.MdiActiveDocument;
            if (document == null)
            {
                MessageBox.Show(
                    "No active AutoCAD drawing is available.",
                    "Hybrid Vertex Editor",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return false;
            }

            if (JtbUtmModeGuard.TryRequireUtmMode(document.Database, actionName, out var detail))
                return true;

            document.Editor.WriteMessage("\n" + detail);
            MessageBox.Show(
                detail,
                "UTM mode required",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return false;
        }

        private sealed class LayerUnlockScope : IDisposable
        {
            private readonly LayerTableRecord _layer;
            private readonly bool _relock;

            internal LayerUnlockScope(Transaction tr, ObjectId layerId)
            {
                _layer = tr.GetObject(layerId, OpenMode.ForWrite, false) as LayerTableRecord;
                if (_layer == null) return;

                _relock = _layer.IsLocked;
                if (_relock) _layer.IsLocked = false;
            }

            public void Dispose()
            {
                if (_relock && _layer != null && !_layer.IsDisposed && !_layer.IsLocked)
                    _layer.IsLocked = true;
            }
        }

        private static LayerUnlockScope UnlockLayer(Transaction tr, ObjectId layerId)
        {
            return new LayerUnlockScope(tr, layerId);
        }

        private static LayerUnlockScope UnlockLayer(Transaction tr, Database db, string layerName)
        {
            EnsureLayer(tr, db, layerName);
            var layers = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            return UnlockLayer(tr, layers[layerName]);
        }

        private static List<VertexInfo> NormalizeVertices(IEnumerable<VertexInfo> vertices)
        {
            var result = new List<VertexInfo>();
            if (vertices == null) return result;
            var seenCoordinates = new HashSet<string>(StringComparer.Ordinal);

            foreach (var source in vertices)
            {
                if (source == null) continue;
                string coordinateKey = HybridValuePolicy.StoredCoordinateKey(source.E, source.N);
                if (!seenCoordinates.Add(coordinateKey)) continue;

                var point = new Point3d(source.E, source.N, source.Pt.Z);
                result.Add(VertexAt(point, source));
            }

            var usedIds = new HashSet<int>();
            int nextId = Math.Max(1, result.Where(v => v.ID > 0).Select(v => v.ID).DefaultIfEmpty(0).Max() + 1);
            foreach (var vertex in result)
            {
                if (vertex.ID <= 0 || !usedIds.Add(vertex.ID))
                {
                    while (!usedIds.Add(nextId)) nextId++;
                    vertex.ID = nextId++;
                }
            }

            return result;
        }

        private static void ApplyPolylineGeometry(Polyline polyline, IList<VertexInfo> vertices)
        {
            int existing = polyline.NumberOfVertices;
            int target = vertices?.Count ?? 0;
            int common = Math.Min(existing, target);

            if (target < 2 && polyline.Closed)
                polyline.Closed = false;

            for (int i = 0; i < common; i++)
            {
                polyline.SetPointAt(
                    i,
                    new Point2d(
                        StoreCoordinate(vertices[i].E),
                        StoreCoordinate(vertices[i].N)));
            }

            for (int i = existing; i < target; i++)
            {
                polyline.AddVertexAt(
                    i,
                    new Point2d(
                        StoreCoordinate(vertices[i].E),
                        StoreCoordinate(vertices[i].N)),
                    0,
                    0,
                    0);
            }

            for (int i = existing - 1; i >= target; i--)
                polyline.RemoveVertexAt(i);
        }

        private static DBDictionary GetOrCreateHybridDataDictionary(
            Transaction tr,
            DBObject owner)
        {
            if (!owner.IsWriteEnabled) owner.UpgradeOpen();
            if (owner.ExtensionDictionary.IsNull) owner.CreateExtensionDictionary();

            var extension = (DBDictionary)tr.GetObject(owner.ExtensionDictionary, OpenMode.ForWrite);
            if (!extension.Contains("HybridData"))
            {
                var data = new DBDictionary();
                extension.SetAt("HybridData", data);
                tr.AddNewlyCreatedDBObject(data, true);
            }

            return (DBDictionary)tr.GetObject(extension.GetAt("HybridData"), OpenMode.ForWrite);
        }

        private static void WriteOwnerMetadata(
            Transaction tr,
            DBObject owner,
            string ownerKey)
        {
            var data = GetOrCreateHybridDataDictionary(tr, owner);
            Xrecord record;
            if (data.Contains(kOwnerRecord))
                record = (Xrecord)tr.GetObject(data.GetAt(kOwnerRecord), OpenMode.ForWrite);
            else
            {
                record = new Xrecord();
                data.SetAt(kOwnerRecord, record);
                tr.AddNewlyCreatedDBObject(record, true);
            }

            record.Data = new ResultBuffer(
                new TypedValue((int)DxfCode.Text, ownerKey ?? string.Empty));
        }

        private static string ReadOwnerMetadata(Transaction tr, DBObject owner)
        {
            try
            {
                if (owner == null || owner.ExtensionDictionary.IsNull) return string.Empty;
                var extension = tr.GetObject(owner.ExtensionDictionary, OpenMode.ForRead, false) as DBDictionary;
                if (extension == null || !extension.Contains("HybridData")) return string.Empty;
                var data = tr.GetObject(extension.GetAt("HybridData"), OpenMode.ForRead, false) as DBDictionary;
                if (data == null || !data.Contains(kOwnerRecord)) return string.Empty;
                var record = tr.GetObject(data.GetAt(kOwnerRecord), OpenMode.ForRead, false) as Xrecord;
                if (record?.Data == null) return string.Empty;

                foreach (TypedValue value in record.Data)
                    if (value.TypeCode == (int)DxfCode.Text)
                        return Convert.ToString(value.Value) ?? string.Empty;
            }
            catch (Autodesk.AutoCAD.Runtime.Exception) { }

            return string.Empty;
        }

        private static ObjectId FindOwnedTableId(
            Transaction tr,
            Database db,
            ObjectId styleId,
            string ownerKey,
            IList<VertexInfo> currentVertices,
            IList<VertexInfo> legacyMatchVertices)
        {
            var blocks = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var model = (BlockTableRecord)tr.GetObject(
                blocks[BlockTableRecord.ModelSpace],
                OpenMode.ForRead);
            var legacy = new List<ObjectId>();

            foreach (ObjectId id in model)
            {
                var table = tr.GetObject(id, OpenMode.ForRead, false) as Table;
                if (table == null || table.TableStyle != styleId ||
                    !string.Equals(table.Layer, kTableLayer, StringComparison.OrdinalIgnoreCase))
                    continue;

                string recordedOwner = ReadOwnerMetadata(tr, table);
                if (string.Equals(recordedOwner, ownerKey, StringComparison.OrdinalIgnoreCase))
                    return id;
                if (string.IsNullOrWhiteSpace(recordedOwner) &&
                    (TableMetadataMatches(tr, table, currentVertices) ||
                     TableMetadataMatches(tr, table, legacyMatchVertices)))
                    legacy.Add(id);
            }

            // Adopt only a single legacy table whose stored coordinates match
            // this polyline. Never rebind an arbitrary sole table.
            return legacy.Count == 1 ? legacy[0] : ObjectId.Null;
        }

        private static bool TableMetadataMatches(
            Transaction tr,
            Table table,
            IList<VertexInfo> vertices)
        {
            if (table == null || vertices == null || table.ExtensionDictionary.IsNull)
                return false;

            try
            {
                var extension = tr.GetObject(table.ExtensionDictionary, OpenMode.ForRead, false) as DBDictionary;
                if (extension == null || !extension.Contains("HybridData")) return false;
                var data = tr.GetObject(extension.GetAt("HybridData"), OpenMode.ForRead, false) as DBDictionary;
                if (data == null || !data.Contains("Data")) return false;
                var record = tr.GetObject(data.GetAt("Data"), OpenMode.ForRead, false) as Xrecord;
                if (record?.Data == null) return false;

                foreach (TypedValue value in record.Data)
                {
                    if (value.TypeCode != (int)DxfCode.Text || !(value.Value is string json)) continue;
                    var stored = JsonConvert.DeserializeObject<List<SimpleVertex>>(json);
                    if (stored == null || stored.Count != vertices.Count) return false;

                    for (int i = 0; i < stored.Count; i++)
                        if (Math.Abs(stored[i].X - vertices[i].E) > kMatchTol ||
                            Math.Abs(stored[i].Y - vertices[i].N) > kMatchTol)
                            return false;
                    return true;
                }
            }
            catch (System.Exception) { }

            return false;
        }

        private static void FillTable(
            Transaction tr,
            Table table,
            IList<VertexInfo> vertices,
            string ownerKey)
        {
            // AutoCAD tables begin with a style-driven title row. Create one
            // temporary extra row and remove that title so every remaining row
            // is a normal five-column Hybrid data row.
            table.SetSize(vertices.Count + 1, 5);
            table.DeleteRows(0, 1);
            ApplyHybridTableLayout(table);

            for (int i = 0; i < vertices.Count; i++)
            {
                var vertex = vertices[i];
                Write(table, i, 0, (i + 1).ToString());
                Write(table, i, 1, DisplayCoordinate(vertex.N));
                Write(table, i, 2, DisplayCoordinate(vertex.E));
                Write(table, i, 3, vertex.Type);
                Write(table, i, 4, vertex.Desc);
            }

            WriteTableMetadata(table, vertices, tr);
            WriteOwnerMetadata(tr, table, ownerKey);
            table.GenerateLayout();
        }

        private static bool UpdateOwnedTable(
            Transaction tr,
            Database db,
            Editor editor,
            IList<VertexInfo> vertices,
            string ownerKey,
            bool createIfMissing,
            IList<VertexInfo> legacyMatchVertices = null)
        {
            ObjectId styleId = EnsureTableStyle(tr, db);
            EnsureLayer(tr, db, kTableLayer);
            ObjectId tableId = FindOwnedTableId(
                tr,
                db,
                styleId,
                ownerKey,
                vertices,
                legacyMatchVertices);

            if (vertices.Count == 0)
            {
                if (!tableId.IsNull)
                {
                    var emptyTable = tr.GetObject(tableId, OpenMode.ForRead, false) as Table;
                    if (emptyTable != null)
                    {
                        using (UnlockLayer(tr, emptyTable.LayerId))
                        {
                            emptyTable.UpgradeOpen();
                            emptyTable.Erase();
                        }
                    }
                }
                return true;
            }

            Table table;
            if (!tableId.IsNull)
            {
                table = (Table)tr.GetObject(tableId, OpenMode.ForRead, false);
            }
            else
            {
                if (!createIfMissing) return true;

                var pointResult = editor.GetPoint("\nPick table insertion point (Model Space):");
                if (pointResult.Status != PromptStatus.OK) return false;

                var blocks = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var model = (BlockTableRecord)tr.GetObject(
                    blocks[BlockTableRecord.ModelSpace],
                    OpenMode.ForWrite);

                using (UnlockLayer(tr, db, kTableLayer))
                {
                    table = new Table();
                    table.SetDatabaseDefaults(db);
                    table.TableStyle = styleId;
                    table.Layer = kTableLayer;
                    table.Position = pointResult.Value;
                    model.AppendEntity(table);
                    tr.AddNewlyCreatedDBObject(table, true);
                }
            }

            using (UnlockLayer(tr, table.LayerId))
            {
                if (!table.IsWriteEnabled) table.UpgradeOpen();
                FillTable(tr, table, vertices, ownerKey);
            }
            return true;
        }

        internal static bool SaveAndUpdate(
            ObjectId polylineId,
            IList<VertexInfo> vertices,
            bool updateMarkerBlocks,
            out List<VertexInfo> savedVertices,
            out string error)
        {
            savedVertices = NormalizeVertices(vertices);
            error = string.Empty;

            var document = AcadApp.DocumentManager.MdiActiveDocument;
            if (document == null || polylineId.IsNull || !polylineId.IsValid || polylineId.IsErased)
            {
                error = "Select a valid polyline first.";
                return false;
            }

            try
            {
                using (HybridGuard.Suspend())
                using (document.LockDocument())
                using (var tr = document.Database.TransactionManager.StartTransaction())
                {
                    var polyline = tr.GetObject(polylineId, OpenMode.ForRead, false) as Polyline;
                    if (polyline == null)
                    {
                        error = "The selected object is no longer an editable polyline.";
                        return false;
                    }
                    if (!IsModelSpaceEntity(tr, document.Database, polyline))
                    {
                        error = "Hybrid polylines must remain in Model Space.";
                        return false;
                    }

                    var previousVertices = ReadVertexData(tr, polyline);
                    string ownerKey = polyline.Handle.ToString();

                    bool erasePolyline = savedVertices.Count == 0;
                    if (!erasePolyline)
                    {
                        using (UnlockLayer(tr, polyline.LayerId))
                        {
                            polyline.UpgradeOpen();
                            ApplyPolylineGeometry(polyline, savedVertices);
                            WriteVertexData(tr, polyline, savedVertices);
                        }
                    }

                    if (!UpdateOwnedTable(
                            tr,
                            document.Database,
                            document.Editor,
                            savedVertices,
                            ownerKey,
                            createIfMissing: true,
                            legacyMatchVertices: previousVertices))
                        return false;

                    SyncHybridMarkerBlocks(
                        tr,
                        document.Database,
                        ownerKey,
                        previousVertices,
                        savedVertices,
                        updateMarkerBlocks);

                    SyncExistingNumberBubbles(
                        tr,
                        document.Database,
                        ownerKey,
                        previousVertices,
                        savedVertices);

                    if (erasePolyline)
                    {
                        using (UnlockLayer(tr, polyline.LayerId))
                        {
                            polyline.UpgradeOpen();
                            polyline.Erase();
                        }
                    }

                    tr.Commit();
                }

                document.Editor.Regen();
                return true;
            }
            catch (System.Exception ex)
            {
                error = "Hybrid data was not saved. No partial drawing changes were committed. " + ex.Message;
                return false;
            }
        }

        private static string EffectiveBlockName(Transaction tr, BlockReference block)
        {
            if (block == null) return string.Empty;
            if (!block.IsDynamicBlock) return (block.Name ?? string.Empty).Trim();

            var definition = tr.GetObject(
                block.DynamicBlockTableRecord,
                OpenMode.ForRead,
                false) as BlockTableRecord;
            return (definition?.Name ?? block.Name ?? string.Empty).Trim();
        }

        private static bool TryGetMarkerType(
            Transaction tr,
            BlockReference block,
            out string markerType)
        {
            markerType = string.Empty;
            string name = EffectiveBlockName(tr, block);
            if (name.Equals("Hybrid_XC", StringComparison.OrdinalIgnoreCase)) markerType = "XC";
            else if (name.Equals("Hybrid_RC", StringComparison.OrdinalIgnoreCase)) markerType = "RC";
            else if (name.Equals("Hybrid_EC", StringComparison.OrdinalIgnoreCase)) markerType = "EC";
            return markerType.Length > 0;
        }

        private static void EraseBlockSafely(Transaction tr, BlockReference block)
        {
            if (block == null || block.IsErased) return;
            using (UnlockLayer(tr, block.LayerId))
            {
                if (!block.IsWriteEnabled) block.UpgradeOpen();
                block.Erase();
            }
        }

        private static void SyncHybridMarkerBlocks(
            Transaction tr,
            Database db,
            string ownerKey,
            IList<VertexInfo> previousVertices,
            IList<VertexInfo> currentVertices,
            bool insertOrUpdate)
        {
            var previousPoints = new Dictionary<Point3d, VertexInfo>(new Point3dEquality(kMatchTol));
            foreach (var vertex in previousVertices ?? new List<VertexInfo>())
                previousPoints[vertex.Pt] = vertex;

            var currentPoints = new Dictionary<Point3d, VertexInfo>(new Point3dEquality(kMatchTol));
            foreach (var vertex in currentVertices ?? new List<VertexInfo>())
                currentPoints[vertex.Pt] = vertex;

            var desired = new Dictionary<Point3d, string>(new Point3dEquality(kMatchTol));
            if (insertOrUpdate)
            {
                foreach (var vertex in currentVertices)
                {
                    string type = HybridMarkerPolicy.DesiredMarkerType(vertex.Type, true);
                    if (type.Length > 0) desired[vertex.Pt] = type;
                }
            }

            var blocks = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var model = (BlockTableRecord)tr.GetObject(
                blocks[BlockTableRecord.ModelSpace],
                OpenMode.ForWrite);
            var toErase = new List<BlockReference>();

            foreach (ObjectId id in model)
            {
                if (id.ObjectClass != RXObject.GetClass(typeof(BlockReference))) continue;
                var block = tr.GetObject(id, OpenMode.ForRead, false) as BlockReference;
                if (block == null || !TryGetMarkerType(tr, block, out var existingType)) continue;

                string recordedOwner = ReadOwnerMetadata(tr, block);
                bool isOwned = string.Equals(
                    recordedOwner,
                    ownerKey,
                    StringComparison.OrdinalIgnoreCase);

                if (!isOwned && !string.IsNullOrWhiteSpace(recordedOwner))
                    continue;

                previousPoints.TryGetValue(block.Position, out var previousVertex);
                currentPoints.TryGetValue(block.Position, out var currentVertex);

                if (!isOwned)
                {
                    bool isMatchingLegacyMarker =
                        MarkerMatchesVertex(existingType, previousVertex) ||
                        MarkerMatchesVertex(existingType, currentVertex);
                    if (!isMatchingLegacyMarker ||
                        MarkerIsClaimedByOtherPolyline(
                            tr,
                            model,
                            ownerKey,
                            block.Position,
                            existingType))
                        continue;

                    using (UnlockLayer(tr, block.LayerId))
                        WriteOwnerMetadata(tr, block, ownerKey);
                    isOwned = true;
                }

                bool isCurrentPoint = currentVertex != null;

                // RC is a data-only type: never retain or insert an RC marker.
                if (existingType == "RC")
                {
                    toErase.Add(block);
                    continue;
                }

                // Always remove artifacts belonging to vertices that disappeared.
                if (!isCurrentPoint)
                {
                    toErase.Add(block);
                    continue;
                }

                // When marker updates are enabled, remove stale/wrong XC/EC
                // references before inserting the desired marker below.
                if (insertOrUpdate &&
                    (!desired.TryGetValue(block.Position, out var wantedType) ||
                     !string.Equals(existingType, wantedType, StringComparison.OrdinalIgnoreCase)))
                    toErase.Add(block);
            }

            foreach (var block in toErase.Distinct()) EraseBlockSafely(tr, block);
            if (!insertOrUpdate || desired.Count == 0) return;

            EnsureHybridBlock(tr, db, "Hybrid_XC", Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 1));
            EnsureHybridBlock(tr, db, "Hybrid_EC", Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 3));

            using (UnlockLayer(tr, db, kBlockLayer))
            {
                foreach (var pair in desired)
                {
                    bool present = false;
                    foreach (ObjectId id in model)
                    {
                        if (id.ObjectClass != RXObject.GetClass(typeof(BlockReference))) continue;
                        if (id.IsErased) continue;
                        var block = tr.GetObject(id, OpenMode.ForRead, false) as BlockReference;
                        if (block == null || block.IsErased ||
                            !TryGetMarkerType(tr, block, out var markerType)) continue;

                        if (string.Equals(markerType, pair.Value, StringComparison.OrdinalIgnoreCase) &&
                            block.Position.DistanceTo(pair.Key) <= kMatchTol &&
                            string.Equals(
                                ReadOwnerMetadata(tr, block),
                                ownerKey,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            present = true;
                            break;
                        }
                    }

                    if (present) continue;
                    var inserted = new BlockReference(pair.Key, blocks["Hybrid_" + pair.Value])
                    {
                        Layer = kBlockLayer,
                        ScaleFactors = new Scale3d(5, 5, 1)
                    };
                    model.AppendEntity(inserted);
                    tr.AddNewlyCreatedDBObject(inserted, true);
                    WriteOwnerMetadata(tr, inserted, ownerKey);
                }
            }
        }

        private static bool MarkerMatchesVertex(
            string markerType,
            VertexInfo vertex)
        {
            if (vertex == null) return false;

            if (string.Equals(markerType, "RC", StringComparison.OrdinalIgnoreCase))
                return string.Equals(
                    (vertex.Type ?? string.Empty).Trim(),
                    "RC",
                    StringComparison.OrdinalIgnoreCase);

            return string.Equals(
                markerType,
                HybridMarkerPolicy.DesiredMarkerType(vertex.Type, true),
                StringComparison.OrdinalIgnoreCase);
        }

        private static bool MarkerIsClaimedByOtherPolyline(
            Transaction tr,
            BlockTableRecord model,
            string ownerKey,
            Point3d markerPoint,
            string markerType)
        {
            foreach (ObjectId id in model)
            {
                if (id.ObjectClass != RXObject.GetClass(typeof(Polyline))) continue;
                var polyline = tr.GetObject(id, OpenMode.ForRead, false) as Polyline;
                if (polyline == null ||
                    string.Equals(
                        polyline.Handle.ToString(),
                        ownerKey,
                        StringComparison.OrdinalIgnoreCase))
                    continue;

                try
                {
                    foreach (var vertex in ReadVertexData(tr, polyline))
                        if (vertex.Pt.DistanceTo(markerPoint) <= kMatchTol &&
                            MarkerMatchesVertex(markerType, vertex))
                            return true;
                }
                catch (System.Exception)
                {
                    // Malformed metadata on another entity must not cause this
                    // workflow to claim an otherwise ambiguous legacy marker.
                    return true;
                }
            }

            return false;
        }

        private static string DetectTypeNear(
            Transaction tr,
            Database db,
            Point3d point)
        {
            var blocks = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var model = (BlockTableRecord)tr.GetObject(
                blocks[BlockTableRecord.ModelSpace],
                OpenMode.ForRead);

            foreach (ObjectId id in model)
            {
                if (id.ObjectClass != RXObject.GetClass(typeof(BlockReference))) continue;
                var block = tr.GetObject(id, OpenMode.ForRead, false) as BlockReference;
                if (block == null || block.Position.DistanceTo(point) > kMatchTol) continue;

                string name = EffectiveBlockName(tr, block).ToUpperInvariant();
                if (name == "HYBRID_XC") return "XC";
                if (name == "HYBRID_RC") return "RC";
                if (name == "HYBRID_EC") return "EC";
                if (name.StartsWith("FDI") || name.StartsWith("FDSPIKE")) return "OC";
            }

            return string.Empty;
        }

        private static List<VertexInfo> BuildCurrentVertexList(
            Transaction tr,
            Database db,
            Polyline polyline,
            IList<VertexInfo> oldVertices)
        {
            var oldByPoint = new Dictionary<Point3d, VertexInfo>(new Point3dEquality(kMatchTol));
            foreach (var oldVertex in oldVertices ?? new List<VertexInfo>())
                if (oldVertex != null && !oldByPoint.ContainsKey(oldVertex.Pt))
                    oldByPoint.Add(oldVertex.Pt, oldVertex);

            var current = new List<VertexInfo>();
            for (int i = 0; i < polyline.NumberOfVertices; i++)
            {
                Point3d point = StorePoint(polyline.GetPoint3dAt(i));
                if (oldByPoint.TryGetValue(point, out var metadata))
                    current.Add(VertexAt(point, metadata));
                else
                {
                    var added = VertexAt(point);
                    added.Type = DetectTypeNear(tr, db, point);
                    current.Add(added);
                }
            }

            return NormalizeVertices(current);
        }

        private static bool IsModelSpaceEntity(Transaction tr, Database db, Entity entity)
        {
            var blocks = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            return entity != null && entity.OwnerId == blocks[BlockTableRecord.ModelSpace];
        }

        internal static bool LoadPolylineVertices(
            ObjectId polylineId,
            out List<VertexInfo> vertices,
            out bool hadStoredData,
            out string error)
        {
            vertices = new List<VertexInfo>();
            hadStoredData = false;
            error = string.Empty;
            var document = AcadApp.DocumentManager.MdiActiveDocument;
            if (document == null || polylineId.IsNull || !polylineId.IsValid || polylineId.IsErased)
            {
                error = "The selected polyline is no longer available.";
                return false;
            }

            try
            {
                using (var tr = document.Database.TransactionManager.StartOpenCloseTransaction())
                {
                    var polyline = tr.GetObject(polylineId, OpenMode.ForRead, false) as Polyline;
                    if (polyline == null)
                    {
                        error = "Select a 2D polyline.";
                        return false;
                    }
                    if (!IsModelSpaceEntity(tr, document.Database, polyline))
                    {
                        error = "Hybrid polylines must be selected from Model Space.";
                        return false;
                    }

                    var stored = ReadVertexData(tr, polyline);
                    hadStoredData = stored.Count > 0;
                    vertices = BuildCurrentVertexList(tr, document.Database, polyline, stored);
                    tr.Commit();
                }

                return true;
            }
            catch (System.Exception ex)
            {
                error = "The selected polyline could not be loaded. " + ex.Message;
                return false;
            }
        }

        private static bool IsNumberBubble(Transaction tr, BlockReference block)
        {
            string name = EffectiveBlockName(tr, block);
            return name.Equals("Hybrd Num", StringComparison.OrdinalIgnoreCase) ||
                   name.Equals("Hybrid Num", StringComparison.OrdinalIgnoreCase);
        }

        private static AttributeReference FindAttribute(
            Transaction tr,
            BlockReference block,
            string tag,
            OpenMode mode)
        {
            foreach (ObjectId attributeId in block.AttributeCollection)
            {
                var attribute = tr.GetObject(attributeId, mode, false) as AttributeReference;
                if (attribute != null &&
                    string.Equals(attribute.Tag, tag, StringComparison.OrdinalIgnoreCase))
                    return attribute;
            }

            return null;
        }

        private static string ReadAttributeText(
            Transaction tr,
            BlockReference block,
            string tag)
        {
            return (FindAttribute(tr, block, tag, OpenMode.ForRead)?.TextString ?? string.Empty).Trim();
        }

        private static AttributeReference SetBubbleAttribute(
            Transaction tr,
            BlockReference block,
            BlockTableRecord definition,
            string tag,
            string value)
        {
            using (UnlockLayer(tr, block.LayerId))
                return SetBubbleAttributeUnlocked(tr, block, definition, tag, value);
        }

        private static AttributeReference SetBubbleAttributeUnlocked(
            Transaction tr,
            BlockReference block,
            BlockTableRecord definition,
            string tag,
            string value)
        {
            var attribute = FindAttribute(tr, block, tag, OpenMode.ForWrite);
            if (attribute == null)
            {
                AttributeDefinition attributeDefinition = null;
                foreach (ObjectId definitionId in definition)
                {
                    var candidate = tr.GetObject(definitionId, OpenMode.ForRead, false) as AttributeDefinition;
                    if (candidate != null &&
                        string.Equals(candidate.Tag, tag, StringComparison.OrdinalIgnoreCase))
                    {
                        attributeDefinition = candidate;
                        break;
                    }
                }

                if (attributeDefinition == null) return null;
                if (!block.IsWriteEnabled) block.UpgradeOpen();

                attribute = new AttributeReference();
                attribute.SetAttributeFromBlock(attributeDefinition, block.BlockTransform);
                attribute.Invisible = attributeDefinition.Invisible;
                block.AttributeCollection.AppendAttribute(attribute);
                tr.AddNewlyCreatedDBObject(attribute, true);
            }

            string normalized = value ?? string.Empty;
            if (!string.Equals(attribute.TextString, normalized, StringComparison.Ordinal))
            {
                attribute.TextString = normalized;
                attribute.AdjustAlignment(block.Database);
            }

            return attribute;
        }

        private sealed class BubbleRecord
        {
            internal BlockReference Block;
            internal int Id;
            internal int OriginalId;
            internal string Number;
            internal string Owner;
            internal double SelectedDistance;
            internal int SuppressionId;
            internal bool HasAuthoritativeEvidence;
            internal bool WasLegacyAdopted;
        }

        private sealed class NumberingPolylineSnapshot
        {
            internal List<VertexInfo> Vertices;
        }

        private static List<NumberingPolylineSnapshot> ReadOtherNumberingPolylines(
            Transaction tr,
            BlockTableRecord model,
            string ownerKey,
            out bool metadataUncertain)
        {
            metadataUncertain = false;
            var snapshots = new List<NumberingPolylineSnapshot>();

            foreach (ObjectId id in model)
            {
                if (id.ObjectClass != RXObject.GetClass(typeof(Polyline))) continue;
                var polyline = tr.GetObject(id, OpenMode.ForRead, false) as Polyline;
                if (polyline == null ||
                    string.Equals(
                        polyline.Handle.ToString(),
                        ownerKey,
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                try
                {
                    var vertices = ReadVertexData(tr, polyline)
                        .Where(vertex => vertex != null)
                        .ToList();
                    if (vertices.Count > 0)
                        snapshots.Add(new NumberingPolylineSnapshot { Vertices = vertices });
                }
                catch (System.Exception)
                {
                    // A malformed Hybrid payload makes legacy ownership
                    // uncertain. The caller will suppress replacement bubbles.
                    metadataUncertain = true;
                }
            }

            return snapshots;
        }

        private static bool TryResolveLegacyBubbleCandidate(
            BubbleRecord record,
            IList<VertexInfo> previousVertices,
            IList<VertexInfo> currentVertices,
            out int sequence,
            out bool hasStrongEvidence,
            out bool hasDirectIdEvidence)
        {
            sequence = 0;
            hasStrongEvidence = false;
            hasDirectIdEvidence = false;
            int resolvedId = record.Id;
            var currentPointIds = currentVertices
                .Where(vertex =>
                    vertex != null &&
                    vertex.ID > 0 &&
                    vertex.Pt.DistanceTo(record.Block.Position) <= kMatchTol)
                .Select(vertex => vertex.ID)
                .Distinct()
                .ToList();
            var previousPointIds = previousVertices
                .Where(vertex =>
                    vertex != null &&
                    vertex.ID > 0 &&
                    vertex.Pt.DistanceTo(record.Block.Position) <= kMatchTol)
                .Select(vertex => vertex.ID)
                .Distinct()
                .ToList();

            // The insertion point is the bubble's original vertex anchor. It is
            // stronger migration evidence than a stale ID that now happens to
            // exist at a different vertex after a rebuild or transfer.
            if (currentPointIds.Count == 1)
            {
                resolvedId = currentPointIds[0];
                hasStrongEvidence = true;
            }
            else if (currentPointIds.Count == 0 && previousPointIds.Count == 1)
            {
                resolvedId = previousPointIds[0];
                hasStrongEvidence = true;
            }

            bool idBelongsToSelected = resolvedId > 0 &&
                (currentVertices.Any(vertex => vertex != null && vertex.ID == resolvedId) ||
                 previousVertices.Any(vertex => vertex != null && vertex.ID == resolvedId));

            int.TryParse(record.Number, out sequence);
            if (sequence < 1) sequence = 0;

            if (!hasStrongEvidence && !idBelongsToSelected)
            {
                // Old constant-ID definitions were upgraded with generated IDs
                // unrelated to vertex identity. A point match is the strongest
                // recovery signal; otherwise NUMBER maps to the pre-edit order.
                var pointIds = previousVertices
                    .Concat(currentVertices)
                    .Where(vertex =>
                        vertex != null &&
                        vertex.ID > 0 &&
                        vertex.Pt.DistanceTo(record.Block.Position) <= kMatchTol)
                    .Select(vertex => vertex.ID)
                    .Distinct()
                    .ToList();

                if (pointIds.Count == 1)
                {
                    resolvedId = pointIds[0];
                    hasStrongEvidence = true;
                }
                else if (sequence <= previousVertices.Count &&
                         previousVertices[sequence - 1] != null &&
                         previousVertices[sequence - 1].ID > 0)
                {
                    resolvedId = previousVertices[sequence - 1].ID;
                }
                else if (sequence <= currentVertices.Count &&
                         currentVertices[sequence - 1] != null &&
                         currentVertices[sequence - 1].ID > 0)
                {
                    resolvedId = currentVertices[sequence - 1].ID;
                }
                else
                {
                    var nearest = previousVertices
                        .Concat(currentVertices)
                        .Where(vertex => vertex != null && vertex.ID > 0)
                        .Select(vertex => new
                        {
                            Vertex = vertex,
                            Distance = vertex.Pt.DistanceTo(record.Block.Position)
                        })
                        .OrderBy(candidate => candidate.Distance)
                        .FirstOrDefault();
                    if (nearest == null) return false;
                    resolvedId = nearest.Vertex.ID;
                }
            }

            hasDirectIdEvidence = !hasStrongEvidence &&
                record.OriginalId > 0 &&
                resolvedId == record.OriginalId &&
                (currentVertices.Any(vertex => vertex != null && vertex.ID == record.OriginalId) ||
                 previousVertices.Any(vertex => vertex != null && vertex.ID == record.OriginalId));

            var selectedPoints = previousVertices
                .Concat(currentVertices)
                .Where(vertex => vertex != null && vertex.ID == resolvedId)
                .Select(vertex => vertex.Pt)
                .ToList();
            if (selectedPoints.Count == 0) return false;

            record.Id = resolvedId;
            record.SelectedDistance = selectedPoints
                .Min(point => point.DistanceTo(record.Block.Position));
            if (record.SelectedDistance <= kMatchTol)
                hasStrongEvidence = true;
            if (currentVertices.Any(vertex => vertex != null && vertex.ID == resolvedId))
            {
                record.SuppressionId = resolvedId;
            }
            else if (sequence >= 1 && sequence <= currentVertices.Count &&
                     currentVertices[sequence - 1] != null)
            {
                record.SuppressionId = currentVertices[sequence - 1].ID;
            }
            else
            {
                var nearestCurrent = currentVertices
                    .Where(vertex => vertex != null && vertex.ID > 0)
                    .OrderBy(vertex => vertex.Pt.DistanceTo(record.Block.Position))
                    .FirstOrDefault();
                record.SuppressionId = nearestCurrent?.ID ?? 0;
            }
            return true;
        }

        private static double ClosestOtherLegacyCandidateDistance(
            BubbleRecord record,
            IList<NumberingPolylineSnapshot> otherPolylines,
            out bool hasCandidate)
        {
            hasCandidate = false;
            double closest = double.PositiveInfinity;

            foreach (var snapshot in otherPolylines)
            {
                var vertices = snapshot.Vertices;
                foreach (var vertex in vertices)
                {
                    if (vertex == null ||
                        (vertex.ID != record.Id && vertex.ID != record.OriginalId))
                    {
                        continue;
                    }

                    hasCandidate = true;
                    closest = Math.Min(
                        closest,
                        vertex.Pt.DistanceTo(record.Block.Position));
                }

            }

            return closest;
        }

        private static void ReconcileVertexIdsFromLegacyAnchors(
            Transaction tr,
            BlockTableRecord model,
            IList<VertexInfo> currentVertices)
        {
            var candidateIdsByVertex = new Dictionary<int, List<int>>();
            foreach (ObjectId id in model)
            {
                if (id.ObjectClass != RXObject.GetClass(typeof(BlockReference))) continue;
                var block = tr.GetObject(id, OpenMode.ForRead, false) as BlockReference;
                if (block == null || !IsNumberBubble(tr, block)) continue;
                if (!string.IsNullOrWhiteSpace(ReadAttributeText(tr, block, "OWNER"))) continue;
                if (!int.TryParse(ReadAttributeText(tr, block, "ID"), out var legacyId) || legacyId <= 0)
                    continue;

                var matchingVertices = Enumerable.Range(0, currentVertices.Count)
                    .Where(index =>
                        currentVertices[index] != null &&
                        currentVertices[index].Pt.DistanceTo(block.Position) <= kMatchTol)
                    .ToList();
                if (matchingVertices.Count != 1) continue;

                int vertexIndex = matchingVertices[0];
                if (!candidateIdsByVertex.TryGetValue(vertexIndex, out var candidates))
                    candidateIdsByVertex[vertexIndex] = candidates = new List<int>();
                candidates.Add(legacyId);
            }

            var currentIds = currentVertices
                .Select(vertex => vertex?.ID ?? 0)
                .ToList();
            var anchorIds = new int[currentVertices.Count];
            foreach (var pair in candidateIdsByVertex)
            {
                var distinctIds = pair.Value.Distinct().ToList();
                if (distinctIds.Count == 1)
                    anchorIds[pair.Key] = distinctIds[0];
            }

            var resolvedIds = HybridNumberingOwnershipPolicy
                .ReconcileVertexIdsWithLegacyAnchors(currentIds, anchorIds);
            for (int i = 0; i < currentVertices.Count; i++)
                if (currentVertices[i] != null)
                    currentVertices[i].ID = resolvedIds[i];
        }

        private static int BubbleCustomizationScore(BubbleRecord record, Point3d expectedPoint)
        {
            int score = record.Block.Position.DistanceTo(expectedPoint) > kMatchTol ? 2 : 0;
            var scale = record.Block.ScaleFactors;
            if (Math.Abs(scale.X - 1.0) > 1e-9 ||
                Math.Abs(scale.Y - 1.0) > 1e-9 ||
                Math.Abs(scale.Z - 1.0) > 1e-9 ||
                Math.Abs(record.Block.Rotation) > 1e-9)
            {
                score++;
            }

            return score;
        }

        private static int SyncOwnedNumberBubbles(
            Transaction tr,
            Database db,
            BlockTableRecord model,
            BlockTableRecord definition,
            string ownerKey,
            IList<VertexInfo> previousVertices,
            IList<VertexInfo> currentVertices,
            bool createMissing)
        {
            if (createMissing)
                ReconcileVertexIdsFromLegacyAnchors(tr, model, currentVertices);

            var expectedById = new Dictionary<int, Tuple<string, Point3d>>();
            var previousById = new Dictionary<int, VertexInfo>();
            foreach (var vertex in previousVertices ?? new List<VertexInfo>())
                if (vertex != null && vertex.ID > 0 && !previousById.ContainsKey(vertex.ID))
                    previousById.Add(vertex.ID, vertex);

            for (int i = 0; i < currentVertices.Count; i++)
            {
                var vertex = currentVertices[i];
                if (vertex.ID > 0 && !expectedById.ContainsKey(vertex.ID))
                    expectedById.Add(vertex.ID, Tuple.Create((i + 1).ToString(), vertex.Pt));
            }

            var previous = (previousVertices ?? new List<VertexInfo>())
                .Where(vertex => vertex != null)
                .ToList();
            var current = (currentVertices ?? new List<VertexInfo>())
                .Where(vertex => vertex != null)
                .ToList();
            var otherPolylines = ReadOtherNumberingPolylines(
                tr,
                model,
                ownerKey,
                out var otherMetadataUncertain);

            var allRecords = new List<BubbleRecord>();
            foreach (ObjectId id in model)
            {
                if (id.ObjectClass != RXObject.GetClass(typeof(BlockReference))) continue;
                var block = tr.GetObject(id, OpenMode.ForRead, false) as BlockReference;
                if (block == null || !IsNumberBubble(tr, block)) continue;

                string idText = ReadAttributeText(tr, block, "ID");
                int.TryParse(idText, out var parsedId);
                allRecords.Add(new BubbleRecord
                {
                    Block = block,
                    Id = parsedId,
                    OriginalId = parsedId,
                    Number = ReadAttributeText(tr, block, "NUMBER"),
                    Owner = ReadAttributeText(tr, block, "OWNER"),
                    SelectedDistance = double.PositiveInfinity
                });
            }

            var ownedById = new Dictionary<int, List<BubbleRecord>>();
            var adoptableById = new Dictionary<int, List<BubbleRecord>>();
            var suppressCreationForId = new HashSet<int>();
            var legacyCoverageForId = new HashSet<int>();
            var invalidOwned = new List<BlockReference>();

            foreach (var record in allRecords)
            {
                bool isOwned = string.Equals(record.Owner, ownerKey, StringComparison.OrdinalIgnoreCase);
                bool isUnowned = string.IsNullOrWhiteSpace(record.Owner);

                if (!isOwned && isUnowned)
                {
                    if (TryResolveLegacyBubbleCandidate(
                            record,
                            previous,
                            current,
                            out var sequence,
                            out var hasStrongEvidence,
                            out var hasDirectIdEvidence))
                    {
                        double otherDistance = ClosestOtherLegacyCandidateDistance(
                            record,
                            otherPolylines,
                            out var hasOtherCandidate);
                        var match = HybridNumberingOwnershipPolicy.ClassifyLegacyBubble(
                            hasSelectedCandidate: true,
                            hasStrongSelectedEvidence: hasStrongEvidence,
                            hasDirectSelectedIdEvidence: hasDirectIdEvidence,
                            selectedDistance: record.SelectedDistance,
                            hasOtherCandidate: hasOtherCandidate,
                            closestOtherDistance: otherDistance,
                            otherMetadataUncertain: otherMetadataUncertain,
                            tolerance: kMatchTol);

                        if (match == LegacyNumberBubbleMatch.AdoptForSelectedPolyline)
                        {
                            record.HasAuthoritativeEvidence = hasStrongEvidence || hasDirectIdEvidence;
                            if (!adoptableById.TryGetValue(record.Id, out var candidates))
                                adoptableById[record.Id] = candidates = new List<BubbleRecord>();
                            candidates.Add(record);
                        }
                        else if (match == LegacyNumberBubbleMatch.Ambiguous &&
                                 expectedById.ContainsKey(record.SuppressionId))
                        {
                            suppressCreationForId.Add(record.SuppressionId);
                            if (hasStrongEvidence)
                                legacyCoverageForId.Add(record.SuppressionId);
                        }
                    }
                }

                if (!isOwned) continue;
                if (record.Id <= 0)
                {
                    invalidOwned.Add(record.Block);
                    continue;
                }

                if (!ownedById.TryGetValue(record.Id, out var list))
                    ownedById[record.Id] = list = new List<BubbleRecord>();
                list.Add(record);
            }

            foreach (var pair in adoptableById)
            {
                var candidates = pair.Value
                    .OrderByDescending(record => record.HasAuthoritativeEvidence)
                    .ThenBy(record => record.SelectedDistance)
                    .ToList();
                bool hasConfirmedOwner = ownedById.TryGetValue(pair.Key, out var confirmed) &&
                    confirmed.Any(record => !record.Block.IsErased);
                if (!HybridNumberingOwnershipPolicy.CanAdoptResolvedLegacyCandidate(
                        hasConfirmedOwner,
                        candidates.Count,
                        candidates.Any(record => record.HasAuthoritativeEvidence)))
                {
                    int suppressionId = candidates
                        .Select(record => record.SuppressionId)
                        .FirstOrDefault(id => expectedById.ContainsKey(id));
                    if (suppressionId > 0) suppressCreationForId.Add(suppressionId);
                    continue;
                }

                var adopted = candidates[0];
                foreach (var duplicateLegacy in candidates.Skip(1))
                    EraseBlockSafely(tr, duplicateLegacy.Block);
                adopted.WasLegacyAdopted = true;
                adopted.Owner = ownerKey;
                SetBubbleAttribute(tr, adopted.Block, definition, "OWNER", ownerKey);
                SetBubbleAttribute(tr, adopted.Block, definition, "ID", adopted.Id.ToString());

                if (!ownedById.TryGetValue(adopted.Id, out var records))
                    ownedById[adopted.Id] = records = new List<BubbleRecord>();
                records.Add(adopted);
            }

            foreach (var block in invalidOwned) EraseBlockSafely(tr, block);

            // If a strong legacy anchor is shared ambiguously with another
            // polyline, keep the original ownerless block and remove only the
            // replacement this owner inserted at the same anchor. Suppression
            // below prevents it from being recreated.
            foreach (int coveredId in legacyCoverageForId)
            {
                if (!expectedById.TryGetValue(coveredId, out var expected) ||
                    !ownedById.TryGetValue(coveredId, out var ownedRecords))
                    continue;

                foreach (var record in ownedRecords)
                    if (!record.Block.IsErased &&
                        record.Block.Position.DistanceTo(expected.Item2) <= kMatchTol)
                        EraseBlockSafely(tr, record.Block);
            }

            foreach (var pair in ownedById)
            {
                if (expectedById.ContainsKey(pair.Key)) continue;
                foreach (var record in pair.Value) EraseBlockSafely(tr, record.Block);
            }

            int unresolvedLegacyCount = suppressCreationForId.Count;
            var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            foreach (var expected in expectedById)
            {
                ownedById.TryGetValue(expected.Key, out var records);
                BubbleRecord keeper = records?
                    .Where(record => !record.Block.IsErased)
                    .OrderByDescending(record => record.WasLegacyAdopted)
                    .ThenByDescending(record => BubbleCustomizationScore(record, expected.Value.Item2))
                    .FirstOrDefault();

                if (keeper == null)
                {
                    if (!HybridNumberingOwnershipPolicy.ShouldCreateMissingBubble(
                            createMissing,
                            hasKeeper: false,
                            insertionSuppressed: suppressCreationForId.Contains(expected.Key)))
                        continue;
                    var created = CreateBubble(
                        tr,
                        model,
                        blockTable,
                        definition,
                        expected.Value.Item2,
                        expected.Key,
                        ownerKey);
                    SetBubbleAttribute(tr, created, definition, "NUMBER", expected.Value.Item1);
                    continue;
                }

                foreach (var duplicate in records)
                    if (duplicate.Block.ObjectId != keeper.Block.ObjectId)
                        EraseBlockSafely(tr, duplicate.Block);

                // Bubbles inserted at the original vertex remain anchored when
                // that vertex moves. A bubble the user has deliberately offset
                // is left at its custom position.
                if (previousById.TryGetValue(expected.Key, out var previousVertex) &&
                    keeper.Block.Position.DistanceTo(previousVertex.Pt) <= kMatchTol &&
                    keeper.Block.Position.DistanceTo(expected.Value.Item2) > kMatchTol)
                {
                    using (UnlockLayer(tr, keeper.Block.LayerId))
                    {
                        if (!keeper.Block.IsWriteEnabled) keeper.Block.UpgradeOpen();
                        keeper.Block.Position = expected.Value.Item2;
                    }
                }

                SetBubbleAttribute(tr, keeper.Block, definition, "OWNER", ownerKey);
                SetBubbleAttribute(tr, keeper.Block, definition, "ID", expected.Key.ToString());
                SetBubbleAttribute(tr, keeper.Block, definition, "NUMBER", expected.Value.Item1);
            }

            return unresolvedLegacyCount;
        }

        private static void SyncExistingNumberBubbles(
            Transaction tr,
            Database db,
            string ownerKey,
            IList<VertexInfo> previousVertices,
            IList<VertexInfo> currentVertices)
        {
            var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            if (!blockTable.Has("Hybrd Num") && !blockTable.Has("Hybrid Num"))
                return;

            EnsureLayer(tr, db, kBlockLayer);
            EnsureHybridNumBlock(tr, db);
            var model = (BlockTableRecord)tr.GetObject(
                blockTable[BlockTableRecord.ModelSpace],
                OpenMode.ForWrite);
            var definition = (BlockTableRecord)tr.GetObject(
                blockTable["Hybrd Num"],
                OpenMode.ForRead);

            using (UnlockLayer(tr, db, kBlockLayer))
                SyncOwnedNumberBubbles(
                    tr,
                    db,
                    model,
                    definition,
                    ownerKey,
                    previousVertices,
                    currentVertices,
                    createMissing: false);
        }

        internal static bool AddOrUpdateNumbering(
            ObjectId polylineId,
            out List<VertexInfo> currentVertices,
            out string error)
        {
            currentVertices = new List<VertexInfo>();
            error = string.Empty;
            int unresolvedLegacyCount = 0;
            var document = AcadApp.DocumentManager.MdiActiveDocument;
            if (document == null || polylineId.IsNull || !polylineId.IsValid || polylineId.IsErased)
            {
                error = "Select a valid polyline first.";
                return false;
            }

            try
            {
                using (HybridGuard.Suspend())
                using (document.LockDocument())
                using (var tr = document.Database.TransactionManager.StartTransaction())
                {
                    var polyline = tr.GetObject(polylineId, OpenMode.ForRead, false) as Polyline;
                    if (polyline == null)
                    {
                        error = "The selected object is no longer an editable polyline.";
                        return false;
                    }
                    if (!IsModelSpaceEntity(tr, document.Database, polyline))
                    {
                        error = "Hybrid polylines must remain in Model Space.";
                        return false;
                    }

                    var previous = ReadVertexData(tr, polyline);
                    currentVertices = BuildCurrentVertexList(tr, document.Database, polyline, previous);
                    string ownerKey = polyline.Handle.ToString();

                    EnsureLayer(tr, document.Database, kBlockLayer);
                    EnsureHybridNumBlock(tr, document.Database);
                    var blockTable = (BlockTable)tr.GetObject(document.Database.BlockTableId, OpenMode.ForRead);
                    var model = (BlockTableRecord)tr.GetObject(
                        blockTable[BlockTableRecord.ModelSpace],
                        OpenMode.ForWrite);
                    var definition = (BlockTableRecord)tr.GetObject(
                        blockTable["Hybrd Num"],
                        OpenMode.ForRead);

                    using (UnlockLayer(tr, document.Database, kBlockLayer))
                        unresolvedLegacyCount = SyncOwnedNumberBubbles(
                            tr,
                            document.Database,
                            model,
                            definition,
                            ownerKey,
                            previous,
                            currentVertices,
                            createMissing: true);

                    using (UnlockLayer(tr, polyline.LayerId))
                        WriteVertexData(tr, polyline, currentVertices);

                    UpdateOwnedTable(
                        tr,
                        document.Database,
                        document.Editor,
                        currentVertices,
                        ownerKey,
                        createIfMissing: false,
                        legacyMatchVertices: previous);

                    tr.Commit();
                }

                document.Editor.Regen();
                if (unresolvedLegacyCount > 0)
                {
                    document.Editor.WriteMessage(
                        "\n" + unresolvedLegacyCount +
                        " legacy numbering match" + (unresolvedLegacyCount == 1 ? " was" : "es were") +
                        " ambiguous. Existing blocks were left unchanged and no new replacement was inserted for those vertices; review any existing duplicates manually.");
                }
                return true;
            }
            catch (System.Exception ex)
            {
                error = "Numbering was not changed. " + ex.Message;
                return false;
            }
        }

        [CommandMethod("SYNCNUMBERING")]
        public static void SyncNumberingCommand()
        {
            if (!TryRequireUtmMode("Add / Update Numbering")) return;

            var document = AcadApp.DocumentManager.MdiActiveDocument;
            var options = new PromptEntityOptions("\nSelect Hybrid polyline to add or update numbering: ");
            options.SetRejectMessage("\nThat is not a polyline.");
            options.AddAllowedClass(typeof(Polyline), false);
            var selection = document.Editor.GetEntity(options);
            if (selection.Status != PromptStatus.OK) return;

            if (AddOrUpdateNumbering(selection.ObjectId, out var vertices, out var error))
                document.Editor.WriteMessage("\nNumbering synchronized for " + vertices.Count + " vertices.");
            else
                MessageBox.Show(error, "Hybrid Vertex Editor", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        internal static bool DeletePointsAndUpdate(
            ObjectId polylineId,
            IList<VertexInfo> remainingVertices,
            bool updateMarkerBlocks,
            out List<VertexInfo> savedVertices,
            out string error)
        {
            savedVertices = NormalizeVertices(remainingVertices);
            error = string.Empty;
            var document = AcadApp.DocumentManager.MdiActiveDocument;
            if (document == null || polylineId.IsNull || !polylineId.IsValid || polylineId.IsErased)
            {
                error = "Select a valid polyline first.";
                return false;
            }

            try
            {
                using (HybridGuard.Suspend())
                using (document.LockDocument())
                using (var tr = document.Database.TransactionManager.StartTransaction())
                {
                    var polyline = tr.GetObject(polylineId, OpenMode.ForRead, false) as Polyline;
                    if (polyline == null)
                    {
                        error = "The selected object is no longer an editable polyline.";
                        return false;
                    }
                    if (!IsModelSpaceEntity(tr, document.Database, polyline))
                    {
                        error = "Hybrid polylines must remain in Model Space.";
                        return false;
                    }

                    var previous = ReadVertexData(tr, polyline);
                    string ownerKey = polyline.Handle.ToString();

                    bool erasePolyline = savedVertices.Count == 0;
                    if (!erasePolyline)
                    {
                        using (UnlockLayer(tr, polyline.LayerId))
                        {
                            polyline.UpgradeOpen();
                            ApplyPolylineGeometry(polyline, savedVertices);
                            WriteVertexData(tr, polyline, savedVertices);
                        }
                    }

                    if (!UpdateOwnedTable(
                            tr,
                            document.Database,
                            document.Editor,
                            savedVertices,
                            ownerKey,
                            createIfMissing: false,
                            legacyMatchVertices: previous))
                        return false;

                    SyncHybridMarkerBlocks(
                        tr,
                        document.Database,
                        ownerKey,
                        previous,
                        savedVertices,
                        updateMarkerBlocks);

                    SyncExistingNumberBubbles(
                        tr,
                        document.Database,
                        ownerKey,
                        previous,
                        savedVertices);

                    if (erasePolyline)
                    {
                        using (UnlockLayer(tr, polyline.LayerId))
                        {
                            polyline.UpgradeOpen();
                            polyline.Erase();
                        }
                    }

                    tr.Commit();
                }

                document.Editor.Regen();
                return true;
            }
            catch (System.Exception ex)
            {
                error = "The point was not deleted. No partial drawing changes were committed. " + ex.Message;
                return false;
            }
        }
    }
}
