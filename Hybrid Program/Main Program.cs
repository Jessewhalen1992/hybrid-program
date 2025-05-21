// HybridSurvey_MainProgram.cs
// Hybrid Survey – AutoCAD 2013–2025 plug-in

#region usings
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using Autodesk.AutoCAD.ApplicationServices;
using System.Security.Cryptography;
#endregion

[assembly: CommandClass(typeof(HybridSurvey.HybridCommands))]

namespace HybridSurvey
{
    internal static class HybridCommands
    {
        private const string kTableStyle = "Induction Bend";
        private const string kTableLayer = "Hybrid_Points_TBL";
        private const string kBlockLayer = "L-MON";
        private const double kTxtH = 2.5;
        private static readonly double[] kColW = { 40, 60, 60, 40, 120 };
        private const double kRowH = 4.0;

        // inside HybridCommands, replace your old SimpleVertex with this:
        private class SimpleVertex
        {
            public double X, Y;
            public string Type, Desc;
            public int ID;         // ← new
        }

        public static void WriteVertexData(ObjectId plId, Database db, List<VertexInfo> verts)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var pl = (Polyline)tr.GetObject(plId, OpenMode.ForWrite);
                if (pl.ExtensionDictionary.IsNull)
                    pl.CreateExtensionDictionary();
                var ext = (DBDictionary)tr.GetObject(pl.ExtensionDictionary, OpenMode.ForWrite);

                DBDictionary dataDict;
                if (ext.Contains("HybridData"))
                    dataDict = (DBDictionary)tr.GetObject(ext.GetAt("HybridData"), OpenMode.ForWrite);
                else
                {
                    dataDict = new DBDictionary();
                    ext.SetAt("HybridData", dataDict);
                    tr.AddNewlyCreatedDBObject(dataDict, true);
                }

                Xrecord xr;
                if (dataDict.Contains("Data"))
                    xr = (Xrecord)tr.GetObject(dataDict.GetAt("Data"), OpenMode.ForWrite);
                else
                {
                    xr = new Xrecord();
                    dataDict.SetAt("Data", xr);
                    tr.AddNewlyCreatedDBObject(xr, true);
                }

                // serialize including ID
                var list = verts.Select(v => new SimpleVertex
                {
                    X = v.Pt.X,
                    Y = v.Pt.Y,
                    Type = v.Type,
                    Desc = v.Desc,
                    ID = v.ID      // ← new
                }).ToList();

                xr.Data = new ResultBuffer(
                    new TypedValue((int)DxfCode.Text, JsonConvert.SerializeObject(list))
                );

                tr.Commit();
            }
        }
        public static List<VertexInfo> ReadVertexData(ObjectId plId, Database db)
        {
            var result = new List<VertexInfo>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var pl = (Polyline)tr.GetObject(plId, OpenMode.ForRead);
                if (pl.ExtensionDictionary.IsNull) { tr.Commit(); return result; }
                var ext = (DBDictionary)tr.GetObject(pl.ExtensionDictionary, OpenMode.ForRead);
                if (!ext.Contains("HybridData")) { tr.Commit(); return result; }
                var dataDict = (DBDictionary)tr.GetObject(ext.GetAt("HybridData"), OpenMode.ForRead);
                if (!dataDict.Contains("Data")) { tr.Commit(); return result; }
                var xr = (Xrecord)tr.GetObject(dataDict.GetAt("Data"), OpenMode.ForRead);

                foreach (TypedValue tv in xr.Data)
                {
                    if (tv.TypeCode == (int)DxfCode.Text)
                    {
                        var list = JsonConvert
                           .DeserializeObject<List<SimpleVertex>>((string)tv.Value);
                        if (list != null)
                        {
                            result.AddRange(list.Select(sv => new VertexInfo
                            {
                                Pt = new Point3d(sv.X, sv.Y, 0),
                                N = sv.Y,
                                E = sv.X,
                                Type = sv.Type,
                                Desc = sv.Desc,
                                ID = sv.ID    // ← new
                            }));
                        }
                    }
                }

                tr.Commit();
            }
            return result;
        }

        [CommandMethod("HYBRIDPLAN")]
        public static void HybridPlan()
        {
            var form = new VertexForm();
            form.Show();
        }

        public static void InsertOrUpdate(IList<VertexInfo> verts, bool insertHybrid)
        {
            if (verts == null || verts.Count == 0) return;

            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            // acquire document lock to avoid eLock errors on table creation/modification
            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var styleId = EnsureTableStyle(tr, db);
                EnsureLayer(tr, db, kTableLayer);

                ObjectId tblId = FindExistingTableId(tr, db, styleId);
                Table tbl;
                Point3d insertPt;

                if (tblId != ObjectId.Null)
                {
                    tbl = (Table)tr.GetObject(tblId, OpenMode.ForRead);
                    tbl.UpgradeOpen();
                    insertPt = tbl.Position;
                }
                else
                {
                    var ppr = ed.GetPoint("\nPick table insertion point:");
                    if (ppr.Status != PromptStatus.OK) return;
                    insertPt = ppr.Value;

                    tbl = new Table
                    {
                        TableStyle = styleId,
                        Layer = kTableLayer,
                        Position = insertPt
                    };
                    tbl.SetDatabaseDefaults();

                    var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    btr.AppendEntity(tbl);
                    tr.AddNewlyCreatedDBObject(tbl, true);
                }

                tbl.Position = insertPt;

                int dataRows = verts.Count;
                int cols = 5;

                tbl.SetSize(dataRows + 1, cols);
                tbl.DeleteRows(0, 1);

                for (int c = 0; c < cols; c++)
                    tbl.Columns[c].Width = kColW[c];
                for (int r = 0; r < dataRows; r++)
                    tbl.Rows[r].Height = kRowH;

                for (int i = 0; i < dataRows; i++)
                {
                    var v = verts[i];
                    Write(tbl, i, 0, (i + 1).ToString());
                    Write(tbl, i, 1, v.N.ToString("F2"));
                    Write(tbl, i, 2, v.E.ToString("F2"));
                    Write(tbl, i, 3, v.Type);
                    Write(tbl, i, 4, v.Desc);
                }

                if (insertHybrid)
                {
                    EnsureAllHybridBlocks(tr, db);
                    PlaceHybridBlocks(tr, db, verts);
                }

                tr.Commit();
            }
        }
        private static void Write(Table tbl, int row, int col, string txt)
        {
            var cell = tbl.Cells[row, col];
            cell.TextHeight = kTxtH;
            cell.TextString = txt ?? "";
        }
        /// <summary>
        /// Re-sorts all existing "Hybrd Num" blocks along the picked polyline
        /// and reassigns only the NUMBER attribute (ID stays fixed).
        /// </summary>

        private static ObjectId EnsureTableStyle(Transaction tr, Database db)
        {
            var dict = (DBDictionary)tr.GetObject(db.TableStyleDictionaryId, OpenMode.ForRead);
            if (dict.Contains(kTableStyle)) return dict.GetAt(kTableStyle);
            dict.UpgradeOpen();
            var ts = new TableStyle { Name = kTableStyle };
            var id = dict.SetAt(kTableStyle, ts);
            tr.AddNewlyCreatedDBObject(ts, true);
            return id;
        }

        private static void EnsureLayer(Transaction tr, Database db, string name)
        {
            var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

            if (!lt.Has(name))
            {
                lt.UpgradeOpen();
                var ltr = new LayerTableRecord { Name = name };
                lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
            }
            else
            {
                var ltr = (LayerTableRecord)tr.GetObject(lt[name], OpenMode.ForWrite);
                ltr.IsFrozen = false;
                ltr.IsLocked = false;
            }
        }

        private static ObjectId FindExistingTableId(Transaction tr, Database db, ObjectId styleId)
        {
            var ms = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
            foreach (ObjectId entId in ms)
            {
                var tbl = tr.GetObject(entId, OpenMode.ForRead) as Table;
                if (tbl != null && tbl.TableStyle == styleId)
                    return entId;
            }
            return ObjectId.Null;
        }

        private static void EnsureAllHybridBlocks(Transaction tr, Database db)
        {
            EnsureHybridBlock(tr, db, "Hybrid_XC", Color.FromColorIndex(ColorMethod.ByAci, 1));
            EnsureHybridBlock(tr, db, "Hybrid_RC", Color.FromColorIndex(ColorMethod.ByAci, 2));
            EnsureHybridBlock(tr, db, "Hybrid_EC", Color.FromColorIndex(ColorMethod.ByAci, 3));
        }

        private static void EnsureHybridBlock(Transaction tr, Database db, string name, Color col)
        {
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            if (bt.Has(name)) return;
            bt.UpgradeOpen();
            var btr = new BlockTableRecord { Name = name, Origin = Point3d.Origin };
            bt.Add(btr);
            tr.AddNewlyCreatedDBObject(btr, true);

            double half = 2.5;
            var l1 = new Line(new Point3d(-half, 0, 0), new Point3d(half, 0, 0)) { Color = col };
            var l2 = new Line(new Point3d(0, -half, 0), new Point3d(0, half, 0)) { Color = col };
            btr.AppendEntity(l1); tr.AddNewlyCreatedDBObject(l1, true);
            btr.AppendEntity(l2); tr.AddNewlyCreatedDBObject(l2, true);
        }

        private static void PlaceHybridBlocks(Transaction tr, Database db, IList<VertexInfo> verts)
        {
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            var tol = new Tolerance(1e-4, 1e-4);

            foreach (var v in verts)
            {
                if (v.Type != "XC" && v.Type != "RC" && v.Type != "EC") continue;

                const double near = 0.004;
                var same = (BlockReference)null;
                var others = new List<BlockReference>();

                foreach (ObjectId entId in space)
                {
                    if (entId.ObjectClass != RXObject.GetClass(typeof(BlockReference)))
                        continue;
                    var br = (BlockReference)tr.GetObject(entId, OpenMode.ForRead);
                    if (!br.Name.StartsWith("Hybrid_", StringComparison.OrdinalIgnoreCase))
                        continue;
                    if (br.Position.DistanceTo(v.Pt) > near) continue;

                    if (br.Name.Equals($"Hybrid_{v.Type}", StringComparison.OrdinalIgnoreCase))
                        same = br;
                    else
                        others.Add(br);
                }

                foreach (var br in others)
                {
                    br.UpgradeOpen();
                    br.Erase();
                }

                if (same != null)
                    continue;

                var insPt = others.Count > 0 ? others[0].Position : v.Pt;
                var nb = new BlockReference(insPt, bt[$"Hybrid_{v.Type}"])
                {
                    Layer = kBlockLayer,
                    ScaleFactors = new Scale3d(5, 5, 1)
                };
                space.AppendEntity(nb);
                tr.AddNewlyCreatedDBObject(nb, true);
            }
        }

        internal static string GuessType(Point3d pt, Transaction tr)
        {
            var db = AcadApp.DocumentManager.MdiActiveDocument.Database;
            var ms = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
            var tol = new Tolerance(1e-4, 1e-4);
            foreach (ObjectId id in ms)
            {
                if (id.ObjectClass != RXObject.GetClass(typeof(BlockReference))) continue;
                var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                if (!br.Position.IsEqualTo(pt, tol)) continue;
                var nm = br.Name.ToUpperInvariant();
                if (nm.Contains("HYBRID_XC")) return "XC";
                if (nm.Contains("HYBRID_RC")) return "RC";
                if (nm.Contains("HYBRID_EC")) return "EC";
                if (nm.StartsWith("FDI") || nm.StartsWith("FDSPIKE")) return "OC";
            }
            return "";
        }

        /// <summary>
        /// Ensures a block definition called "Hybrd Num" exists,
        /// with two attributes: NUMBER (visible, holds the vertex index)
        /// and ID (constant, holds a unique tag tied to the original N/E).
        /// </summary>
        /// <summary>
        /// Ensures a block definition called "Hybrd Num" exists,
        /// with two attributes:
        ///   • NUMBER – visible, holds the vertex index
        ///   • ID     – hidden, holds the immutable tag
        /// </summary>
        /// <summary>
        /// Ensures a block definition called **“Hybrd Num”** exists,
        /// with two attributes:
        ///   • NUMBER – visible, variable, shows the vertex order
        ///   • ID     – hidden, variable, forever stores the unique tag
        /// </summary>
        /// <summary>
        /// Makes sure the block definition “Hybrd Num” exists **and** that the
        /// ID attribute is VARIABLE (not constant). If the definition already
        /// exists but the ID attribute is still constant, the method converts
        /// it in-place so future inserts can carry unique IDs.
        /// </summary>
        private static void EnsureHybridNumBlock(Transaction tr, Database db)
        {
            const string blkName = "Hybrd Num";
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

            if (!bt.Has(blkName))
            {
                // ── create fresh definition ──────────────────────────────────────
                bt.UpgradeOpen();
                var btr = new BlockTableRecord { Name = blkName, Origin = Point3d.Origin };
                bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);

                // NUMBER (visible)
                var numDef = new AttributeDefinition
                {
                    Tag = "NUMBER",
                    Prompt = "Number",
                    Height = kTxtH,
                    Position = Point3d.Origin,
                    Invisible = false,
                    Constant = false
                };
                numDef.SetDatabaseDefaults();
                btr.AppendEntity(numDef); tr.AddNewlyCreatedDBObject(numDef, true);

                // ID (hidden, VARIABLE!)
                var idDef = new AttributeDefinition
                {
                    Tag = "ID",
                    Prompt = "Tag",
                    Height = kTxtH,
                    Position = Point3d.Origin,
                    Invisible = true,
                    Constant = false
                };
                idDef.SetDatabaseDefaults();
                btr.AppendEntity(idDef); tr.AddNewlyCreatedDBObject(idDef, true);
            }
            else
            {
                // ── definition exists – make sure ID is NOT constant ────────────
                var btr = (BlockTableRecord)tr.GetObject(bt[blkName], OpenMode.ForWrite);
                foreach (ObjectId entId in btr)
                {
                    if (tr.GetObject(entId, OpenMode.ForRead) is AttributeDefinition ad
                        && ad.Tag.Equals("ID", StringComparison.OrdinalIgnoreCase)
                        && ad.Constant)
                    {
                        ad.UpgradeOpen();
                        ad.Constant = false;   // flip to variable
                        ad.Invisible = true;    // keep hidden
                    }
                }
            }
        }
        /// <summary>
        /// Re-sorts all existing "Hybrd Num" blocks along the picked polyline
        /// and reassigns only the NUMBER attribute (ID stays fixed).
        /// </summary>

        // -----------------------------------------------------------------------------
        // UPDATE NUMBERING – never duplicates existing bubbles
        // -----------------------------------------------------------------------------
        // -----------------------------------------------------------------------------
        // UPDATE NUMBERING – uses ID↔X/Y binding, never looks at bubble position
        // -----------------------------------------------------------------------------
        // -----------------------------------------------------------------------------
        // UPDATE NUMBERING  – uses ID↔original-XY binding, no duplicates
        // -----------------------------------------------------------------------------
        // -----------------------------------------------------------------------------
        // UPDATE NUMBERING – keeps ID↔original-XY binding, no duplicates
        // -----------------------------------------------------------------------------
        [CommandMethod("UpdateNumbering")]
        public static void UpdateNumbering()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            // 1) pick polyline
            var peo = new PromptEntityOptions("\nSelect polyline to renumber: ");
            peo.SetRejectMessage("\nThat’s not a polyline.");
            peo.AddAllowedClass(typeof(Polyline), false);
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;
            var plId = per.ObjectId;

            // 2) load stored vertex metadata (ID + original XY)
            var oldData = ReadVertexData(plId, db);
            var oldMap = oldData.ToDictionary(v => v.Pt, v => v, new Point3dEquality());

            // 3) build current vertex list
            var verts = new List<VertexInfo>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var pl = (Polyline)tr.GetObject(plId, OpenMode.ForRead);
                for (int i = 0; i < pl.NumberOfVertices; i++)
                {
                    var pt = pl.GetPoint3dAt(i);
                    verts.Add(oldMap.TryGetValue(pt, out var vi)
                              ? vi
                              : new VertexInfo { Pt = pt, N = pt.Y, E = pt.X, ID = 0 });
                }
                tr.Commit();
            }

            // 4) work in model space
            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                EnsureLayer(tr, db, kBlockLayer);
                EnsureHybridNumBlock(tr, db);

                // -- upgrade legacy bubbles (constant ID) once --
                UpgradeExistingHybrdNumBubbles(
                    tr,
                    db,
                    (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite),
                    (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead));

                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                var defRec = (BlockTableRecord)tr.GetObject(bt["Hybrd Num"], OpenMode.ForRead);

                // build ID → bubble map
                var idToBr = new Dictionary<int, BlockReference>();
                int maxId = 0;
                foreach (ObjectId entId in space)
                {
                    if (entId.ObjectClass != RXObject.GetClass(typeof(BlockReference))) continue;
                    var br = (BlockReference)tr.GetObject(entId, OpenMode.ForRead);
                    if (!br.Name.Equals("Hybrd Num", StringComparison.OrdinalIgnoreCase)) continue;

                    foreach (ObjectId attId in br.AttributeCollection)
                    {
                        var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                        if (ar.Tag == "ID" && int.TryParse(ar.TextString, out int v))
                        {
                            idToBr[v] = br;
                            maxId = Math.Max(maxId, v);
                        }
                    }
                }

                // 5) ensure one bubble per vertex
                int nextId = maxId + 1;
                foreach (var v in verts)
                {
                    if (v.ID != 0)            // vertex already knows its ID
                    {
                        if (idToBr.ContainsKey(v.ID)) continue;  // bubble still present

                        // bubble was deleted → recreate with same ID
                        var brNew = CreateBubble(tr, space, bt, defRec, v.Pt, v.ID);
                        idToBr[v.ID] = brNew;
                    }
                    else                      // brand-new vertex
                    {
                        v.ID = nextId++;
                        var brNew = CreateBubble(tr, space, bt, defRec, v.Pt, v.ID);
                        idToBr[v.ID] = brNew;
                    }
                }

                // 6) renumber NUMBER attribute to match vertex order
                for (int i = 0; i < verts.Count; i++)
                {
                    if (!idToBr.TryGetValue(verts[i].ID, out var br)) continue;

                    foreach (ObjectId attId in br.AttributeCollection)
                    {
                        var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForWrite);
                        if (ar.Tag == "NUMBER")
                        {
                            ar.TextString = (i + 1).ToString();
                            ar.AdjustAlignment(db);
                        }
                    }
                }

                // 7) save metadata back on the polyline
                WriteVertexData(plId, db, verts);
                tr.Commit();
            }

            doc.Editor.Regen();
        }


        private static BlockReference CreateBubble(
            Transaction tr,
            BlockTableRecord space,
            BlockTable bt,
            BlockTableRecord defRec,
            Point3d pt,
            int id)
        {
            var nb = new BlockReference(pt, bt["Hybrd Num"])
            {
                Layer = kBlockLayer,
                ScaleFactors = new Scale3d(1, 1, 1)
            };
            space.AppendEntity(nb);
            tr.AddNewlyCreatedDBObject(nb, true);

            foreach (ObjectId defId in defRec)
            {
                if (tr.GetObject(defId, OpenMode.ForRead) is AttributeDefinition def)
                {
                    var ar = new AttributeReference();
                    ar.SetAttributeFromBlock(def, Matrix3d.Identity);
                    ar.Position = pt;
                    ar.Invisible = def.Invisible;
                    ar.TextString = def.Tag == "ID" ? id.ToString() : "";
                    nb.AttributeCollection.AppendAttribute(ar);
                    tr.AddNewlyCreatedDBObject(ar, true);
                }
            }
            return nb;
        }


        /// <summary>
        /// Converts any pre-existing Hybrd Num bubbles that *lack* an ID
        /// attribute (because they were inserted with the old constant-ID
        /// definition).  Each missing bubble is given a brand-new, unique,
        /// hidden ID attribute so the rest of the renumbering logic can work.
        /// </summary>
        private static void UpgradeExistingHybrdNumBubbles(
            Transaction tr, Database db, BlockTableRecord space, BlockTable bt)
        {
            if (!bt.Has("Hybrd Num")) return;

            // grab the *variable* ID definition we just ensured exists
            var defRec = (BlockTableRecord)tr.GetObject(bt["Hybrd Num"], OpenMode.ForRead);
            AttributeDefinition idDef = defRec
                .Cast<ObjectId>()
                .Select(id => tr.GetObject(id, OpenMode.ForRead) as AttributeDefinition)
                .FirstOrDefault(ad => ad != null && ad.Tag == "ID");

            if (idDef == null) return;   // should never happen

            // find max existing ID so we keep them unique
            int maxId = 0;
            foreach (ObjectId entId in space)
            {
                if (entId.ObjectClass != RXObject.GetClass(typeof(BlockReference))) continue;
                var br = (BlockReference)tr.GetObject(entId, OpenMode.ForRead);
                if (!br.Name.Equals("Hybrd Num", StringComparison.OrdinalIgnoreCase)) continue;

                foreach (ObjectId attId in br.AttributeCollection)
                    if (tr.GetObject(attId, OpenMode.ForRead) is AttributeReference ar
                        && ar.Tag == "ID"
                        && int.TryParse(ar.TextString, out int v))
                        maxId = Math.Max(maxId, v);
            }
            int nextId = maxId + 1;

            // add missing ID attributes
            foreach (ObjectId entId in space)
            {
                if (entId.ObjectClass != RXObject.GetClass(typeof(BlockReference))) continue;
                var br = (BlockReference)tr.GetObject(entId, OpenMode.ForRead);
                if (!br.Name.Equals("Hybrd Num", StringComparison.OrdinalIgnoreCase)) continue;

                bool hasId = br.AttributeCollection.Cast<ObjectId>().Any(attId =>
                {
                    var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                    return ar.Tag == "ID";
                });
                if (hasId) continue;             // already good

                // insert a fresh ID attribute
                br.UpgradeOpen();
                var arNew = new AttributeReference();
                arNew.SetAttributeFromBlock(idDef, Matrix3d.Identity);
                arNew.Position = br.Position;
                arNew.Invisible = true;
                arNew.TextString = nextId.ToString();
                br.AttributeCollection.AppendAttribute(arNew);
                tr.AddNewlyCreatedDBObject(arNew, true);
                nextId++;
            }
        }




        [CommandMethod("TransferVertexData")]
        public static void TransferVertexData()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            // 1) Pick the source polyline (must already have HybridData)
            var peoSrc = new PromptEntityOptions("\nSelect source polyline:");
            peoSrc.SetRejectMessage("\nThat’s not a polyline.");
            peoSrc.AddAllowedClass(typeof(Polyline), false);
            var perSrc = ed.GetEntity(peoSrc);
            if (perSrc.Status != PromptStatus.OK) return;
            ObjectId srcId = perSrc.ObjectId;

            // 2) Pick the target polyline
            var peoDst = new PromptEntityOptions("\nSelect target polyline:");
            peoDst.SetRejectMessage("\nThat’s not a polyline.");
            peoDst.AddAllowedClass(typeof(Polyline), false);
            var perDst = ed.GetEntity(peoDst);
            if (perDst.Status != PromptStatus.OK) return;
            ObjectId dstId = perDst.ObjectId;

            // 3) Read all metadata from source into a map by position
            var srcList = ReadVertexData(srcId, db);
            var srcMap = srcList.ToDictionary(v => v.Pt, v => v, new Point3dEquality());

            // 4) Build a new list for the target: preserve matches, blank others
            var newList = new List<VertexInfo>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var plDst = (Polyline)tr.GetObject(dstId, OpenMode.ForRead);
                for (int i = 0; i < plDst.NumberOfVertices; i++)
                {
                    var pt = plDst.GetPoint3dAt(i);
                    if (srcMap.TryGetValue(pt, out var vi))
                        newList.Add(new VertexInfo
                        {
                            Pt = vi.Pt,
                            N = vi.N,
                            E = vi.E,
                            Type = vi.Type,
                            Desc = vi.Desc
                        });
                    else
                        newList.Add(new VertexInfo
                        {
                            Pt = pt,
                            N = pt.Y,
                            E = pt.X,
                            Type = "",
                            Desc = ""
                        });
                }
                tr.Commit();
            }

            // 5) Write that metadata back onto the target polyline
            WriteVertexData(dstId, db, newList);
            ed.WriteMessage($"\nTransferred metadata to {newList.Count} vertices.");
        }

        public static void AddNumbering()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                EnsureLayer(tr, db, kBlockLayer);
                EnsureHybridNumBlock(tr, db);

                // pick polyline
                var peo = new PromptEntityOptions("\nSelect polyline to place numbers on: ");
                peo.SetRejectMessage("\nThat’s not a polyline.");
                peo.AddAllowedClass(typeof(Polyline), false);
                var per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK) return;
                var pl = (Polyline)tr.GetObject(per.ObjectId, OpenMode.ForRead);

                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                var defRec = (BlockTableRecord)tr.GetObject(bt["Hybrd Num"], OpenMode.ForRead);
                var tol = new Tolerance(1e-4, 1e-4);

                // collect existing IDs
                var existingIds = new List<int>();
                foreach (ObjectId entId in space)
                {
                    if (entId.ObjectClass != RXObject.GetClass(typeof(BlockReference))) continue;
                    var br = (BlockReference)tr.GetObject(entId, OpenMode.ForRead);
                    if (!br.Name.Equals("Hybrd Num", StringComparison.OrdinalIgnoreCase)) continue;

                    foreach (ObjectId attId in br.AttributeCollection)
                    {
                        var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                        if (ar.Tag == "ID" && int.TryParse(ar.TextString, out int v))
                            existingIds.Add(v);
                    }
                }
                int nextId = existingIds.Any() ? existingIds.Max() + 1 : 1;

                // final metadata list
                var verts = new List<VertexInfo>();

                // walk vertices
                for (int i = 0; i < pl.NumberOfVertices; i++)
                {
                    var pt = pl.GetPoint3dAt(i);
                    int id = 0;

                    // look for a bubble exactly on this vertex
                    BlockReference found = null;
                    foreach (ObjectId entId in space)
                    {
                        if (entId.ObjectClass != RXObject.GetClass(typeof(BlockReference))) continue;
                        var br = (BlockReference)tr.GetObject(entId, OpenMode.ForRead);
                        if (br.Name.Equals("Hybrd Num", StringComparison.OrdinalIgnoreCase) &&
                            br.Position.IsEqualTo(pt, tol))
                        {
                            found = br; break;
                        }
                    }

                    if (found != null)
                    {
                        // reuse bubble, grab its ID
                        foreach (ObjectId attId in found.AttributeCollection)
                        {
                            var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                            if (ar.Tag == "ID" && int.TryParse(ar.TextString, out int v)) { id = v; break; }
                        }
                        // update NUMBER
                        foreach (ObjectId attId in found.AttributeCollection)
                        {
                            var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForWrite);
                            if (ar.Tag == "NUMBER")
                            {
                                ar.TextString = (i + 1).ToString();
                                ar.AdjustAlignment(db);
                            }
                        }
                    }
                    else
                    {
                        // fresh bubble
                        id = nextId++;
                        CreateBubble(tr, space, bt, defRec, pt, id);   // ← reuse helper above
                    }

                    // record metadata row
                    verts.Add(new VertexInfo
                    {
                        Pt = pt,
                        N = pt.Y,
                        E = pt.X,
                        Type = "",
                        Desc = "",
                        ID = id
                    });
                }

                // save metadata to the polyline
                WriteVertexData(pl.ObjectId, db, verts);
                tr.Commit();
            }

            doc.Editor.Regen();
        }

        [CommandMethod("REBUILDPLINEFROMTABLE")]
        public static void RebuildPolylineFromTable()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            var peo = new PromptEntityOptions("\nSelect Hybrid-Points table: ");
            peo.SetRejectMessage("\nEntity must be a Table");
            peo.AddAllowedClass(typeof(Table), false);
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var tbl = (Table)tr.GetObject(per.ObjectId, OpenMode.ForRead);
                int rowCount = tbl.Rows.Count;
                var verts = new List<VertexInfo>();
                for (int r = 1; r < rowCount; r++)
                {
                    if (double.TryParse(tbl.Cells[r, 1].TextString, out double N) &&
                        double.TryParse(tbl.Cells[r, 2].TextString, out double E))
                    {
                        verts.Add(new VertexInfo
                        {
                            Pt = new Point3d(E, N, 0),
                            N = N,
                            E = E,
                            Type = tbl.Cells[r, 3].TextString,
                            Desc = tbl.Cells[r, 4].TextString
                        });
                    }
                }

                var btrOwner = (BlockTableRecord)tr.GetObject(tbl.OwnerId, OpenMode.ForWrite);
                var pl = new Polyline();
                for (int i = 0; i < verts.Count; i++)
                {
                    pl.AddVertexAt(i,
                        new Point2d(verts[i].Pt.X, verts[i].Pt.Y),
                        0, 0, 0);
                }
                btrOwner.AppendEntity(pl);
                tr.AddNewlyCreatedDBObject(pl, true);

                WriteVertexData(pl.ObjectId, db, verts);

                tr.Commit();
                ed.WriteMessage($"\nRebuilt polyline with {verts.Count} vertices.");
            }
        }
    }

    internal sealed class VertexForm : Form // test comit from vs
    {
        private readonly DataGridView _grid;
        private readonly CheckBox _chkHybrid;
        private List<VertexInfo> _verts;
        private ObjectId _currentPlId = ObjectId.Null;

        public VertexForm()
        {
            Text = "Hybrid Vertex Editor";
            ClientSize = new System.Drawing.Size(950, 580);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            var pnl = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 45,
                Padding = new Padding(5),
                WrapContents = false,
                AutoScroll = false
            };

            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };

            _grid.Columns.Add("#", "#");
            _grid.Columns["#"].ReadOnly = true;
            _grid.Columns.Add("Northing", "Northing");
            _grid.Columns.Add("Easting", "Easting");

            var typeCol = new DataGridViewComboBoxColumn
            {
                Name = "Type",
                HeaderText = "Type",
                FlatStyle = FlatStyle.Flat,
                DropDownWidth = 60
            };
            typeCol.Items.AddRange("", "XC", "RC", "EC", "OC", "Other");
            _grid.Columns.Add(typeCol);
            _grid.Columns.Add("Desc", "Description");

            Controls.Add(_grid);
            Controls.Add(pnl);

            // existing buttons
            var btnGet = new Button { Text = "Get Polyline", Width = 120 };
            var btnUpd = new Button { Text = "Update Polyline", Width = 120 };
            var btnAddNum = new Button { Text = "Add Numbering", Width = 120 };
            var btnUpdNum = new Button { Text = "Update Numbering", Width = 120 };
            var btnRebuild = new Button { Text = "Rebuild Polyline", Width = 140 };
            var btnTransfer = new Button { Text = "Transfer Data", Width = 120 };  // ← new
            var btnOK = new Button { Text = "Insert/Update", Width = 90 };
            _chkHybrid = new CheckBox { Text = "Hybrid Blocks", Checked = true };

            pnl.Controls.AddRange(new Control[] {
        btnGet, btnUpd, btnAddNum, btnUpdNum, btnRebuild, btnTransfer, _chkHybrid, btnOK
    });

            // wire up events
            btnGet.Click += (s, e) => PickAndPopulate();
            btnUpd.Click += (s, e) => PickAndPopulate();
            btnAddNum.Click += (s, e) => HybridCommands.AddNumbering();
            btnUpdNum.Click += (s, e) => HybridCommands.UpdateNumbering();
            btnRebuild.Click += (s, e) => HybridCommands.RebuildPolylineFromTable();
            btnTransfer.Click += (s, e) => HybridCommands.TransferVertexData();  // ← new
            btnOK.Click += InsertUpdate_Click;

            AcceptButton = btnOK;

            _grid.EditingControlShowing += (s, e) =>
            {
                if (_grid.CurrentCell.OwningColumn.Name == "Type" && e.Control is ComboBox cb)
                {
                    cb.DropDownStyle = ComboBoxStyle.DropDown;
                    cb.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                    cb.AutoCompleteSource = AutoCompleteSource.ListItems;
                    cb.Validating -= Combo_Validating;
                    cb.Validating += Combo_Validating;
                }
            };

            _grid.CellValidating += Grid_CellValidating;
            _grid.DataError += (s, e) =>
            {
                if (e.ColumnIndex == _grid.Columns["Type"].Index)
                    e.ThrowException = false;
            };

            _verts = new List<VertexInfo>();
        }
        /// <summary>
        /// Re-sorts all existing "Hybrd Num" blocks along the picked polyline
        /// and reassigns only the NUMBER attribute (ID stays fixed).
        /// </summary>
        private void PickAndPopulate()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var opts = new PromptEntityOptions("\nSelect polyline");
            opts.SetRejectMessage("\nMust be a polyline");
            opts.AddAllowedClass(typeof(Polyline), false);
            var res = ed.GetEntity(opts);
            if (res.Status != PromptStatus.OK) return;

            _currentPlId = res.ObjectId;
            var oldList = HybridCommands.ReadVertexData(_currentPlId, doc.Database);
            var oldMap = oldList.ToDictionary(v => v.Pt, v => v, new Point3dEquality());
            var newList = new List<VertexInfo>();

            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                var pl = (Polyline)tr.GetObject(res.ObjectId, OpenMode.ForRead);
                for (int i = 0; i < pl.NumberOfVertices; i++)
                {
                    var pt = pl.GetPoint3dAt(i);
                    if (oldMap.TryGetValue(pt, out var vi))
                        newList.Add(vi);
                    else
                        newList.Add(new VertexInfo
                        {
                            Pt = pt,
                            N = pt.Y,
                            E = pt.X,
                            Type = HybridCommands.GuessType(pt, tr),
                            Desc = ""
                        });
                }
                tr.Commit();
            }

            _verts = newList;
            RefreshGrid();
        }

        private void InsertUpdate_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < _verts.Count; i++)
            {
                var row = _grid.Rows[i];
                _verts[i] = new VertexInfo
                {
                    Pt = _verts[i].Pt,
                    N = double.Parse(row.Cells["Northing"].Value.ToString()),
                    E = double.Parse(row.Cells["Easting"].Value.ToString()),
                    Type = row.Cells["Type"].Value?.ToString() ?? "",
                    Desc = row.Cells["Desc"].Value?.ToString() ?? ""
                };
            }

            HybridCommands.InsertOrUpdate(_verts, _chkHybrid.Checked);
            HybridCommands.WriteVertexData(
                _currentPlId,
                AcadApp.DocumentManager.MdiActiveDocument.Database,
                _verts
            );
            RefreshGrid();
        }

        private void RefreshGrid()
        {
            _grid.Rows.Clear();
            var col = (DataGridViewComboBoxColumn)_grid.Columns["Type"];
            foreach (var v in _verts)
            {
                if (!string.IsNullOrEmpty(v.Type) && !col.Items.Contains(v.Type))
                    col.Items.Add(v.Type);
                _grid.Rows.Add(
                    _grid.Rows.Count + 1,
                    v.N.ToString("F2"),
                    v.E.ToString("F2"),
                    v.Type,
                    v.Desc
                );
            }
        }

        private void Grid_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (_grid.Columns[e.ColumnIndex].Name == "Type")
            {
                var newVal = e.FormattedValue?.ToString() ?? "";
                var col = (DataGridViewComboBoxColumn)_grid.Columns["Type"];
                if (!col.Items.Contains(newVal))
                    col.Items.Add(newVal);
                _grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = newVal;
            }
        }

        private void Combo_Validating(object sender, CancelEventArgs e)
        {
            if (_grid.CurrentCell.OwningColumn.Name == "Type"
                && _grid.EditingControl is ComboBox cb)
            {
                _grid.CurrentCell.Value = cb.Text;
            }
        }
    }

    internal class VertexInfo
    {
        public Point3d Pt;
        public double N, E;
        public string Type, Desc;
        public int ID;
    }


    internal class Point3dEquality : IEqualityComparer<Point3d>
    {
        private static readonly Tolerance _tol = new Tolerance(1e-4, 1e-4);
        public bool Equals(Point3d a, Point3d b) => a.IsEqualTo(b, _tol);
        public int GetHashCode(Point3d p) => (p.X.GetHashCode() * 397) ^ p.Y.GetHashCode();
    }
}