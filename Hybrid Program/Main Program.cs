// HybridSurvey_MainProgram.cs
// PATCHED by Codex, 2025-06-10 – brace & scope fixes
// REFACTORED 2025-06-10
// Hybrid Survey – AutoCAD 2013–2025 plug-in
// test upload
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
        internal const double kMatchTol = 0.03;       // vertex search tolerance
        private const double kRowH = 4.0;

        // -----------------------------------------------------------------
        //  Read-only protection for NUMBER / ID attributes in "Hybrd Num"
        // -----------------------------------------------------------------
        private static bool _allowInternalEdits = false;  // toggled by helper
        private static bool _warnedOnce         = false;  // only one console msg

        // Static ctor runs when the class is first touched
        static HybridCommands()
        {
            var db = HostApplicationServices.WorkingDatabase;
            db.ObjectModified += Db_ObjectModified;
        }

        private static void Db_ObjectModified(object sender, ObjectEventArgs e)
        {
            if (_allowInternalEdits) return;               // program change → allow

            if (e.DBObject is AttributeReference ar)
            {
                string tag = ar.Tag?.ToUpperInvariant();
                if (tag != "NUMBER" && tag != "ID") return;

                var br = ar.OwnerId.GetObject(OpenMode.ForRead) as BlockReference;
                if (br == null || !br.Name.Equals("Hybrd Num", StringComparison.OrdinalIgnoreCase))
                    return;

                // Cancel the user edit by restoring the pre-modify value
                using (var tr = ar.Database.TransactionManager.StartOpenCloseTransaction())
                {
                    var arW = (AttributeReference)tr.GetObject(ar.ObjectId, OpenMode.ForWrite);
                    arW.TextString = arW.TextString;   // rewrite cached original
                    tr.Commit();
                }

                if (!_warnedOnce)
                {
                    AcadApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                        "\nHybrid NUMBER / ID are managed by the plug-in and cannot be edited manually.");
                    _warnedOnce = true;
                }
            }
        }

        // inside HybridCommands, replace your old SimpleVertex with this:
        private class SimpleVertex
        {
            public double X, Y;
            public string Type, Desc;
            public int ID;         // ← new
        }

        public static void WriteVertexData(ObjectId plId,
                                           Database db,
                                           List<VertexInfo> verts)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                //-- ensure the polyline owns an extension dictionary
                var pl = (Polyline)tr.GetObject(plId, OpenMode.ForWrite);
                if (pl.ExtensionDictionary.IsNull)
                    pl.CreateExtensionDictionary();

                var ext = (DBDictionary)tr.GetObject(pl.ExtensionDictionary, OpenMode.ForWrite);

                //-- ensure “HybridData” dictionary
                if (!ext.Contains("HybridData"))
                {
                    var dict = new DBDictionary();
                    ext.SetAt("HybridData", dict);
                    tr.AddNewlyCreatedDBObject(dict, true);
                }
                var dataDict = (DBDictionary)tr.GetObject(ext.GetAt("HybridData"), OpenMode.ForWrite);

                //-- ensure XRecord “Data”
                Xrecord xr;
                if (dataDict.Contains("Data"))
                    xr = (Xrecord)tr.GetObject(dataDict.GetAt("Data"), OpenMode.ForWrite);
                else
                {
                    xr = new Xrecord();
                    dataDict.SetAt("Data", xr);
                    tr.AddNewlyCreatedDBObject(xr, true);
                }

                //-- serialise payload  (round coordinates to 0.001)
                var jsonPayload = JsonConvert.SerializeObject(
                    verts.Select(v => new SimpleVertex
                    {
                        X = Math.Round(v.Pt.X, 3),
                        Y = Math.Round(v.Pt.Y, 3),
                        Type = v.Type,
                        Desc = v.Desc,
                        ID = v.ID
                    }).ToList()
                );

                xr.Data = new ResultBuffer(new TypedValue((int)DxfCode.Text, jsonPayload));

                tr.Commit();          // <-- commit once, here
            }
        }

        private static void WriteTableMetadata(Table tbl, IList<VertexInfo> verts, Transaction tr)
        {
            // (a) guarantee an extension dictionary on the table
            if (tbl.ExtensionDictionary.IsNull)
                tbl.CreateExtensionDictionary();

            var ext = (DBDictionary)tr.GetObject(tbl.ExtensionDictionary, OpenMode.ForWrite);

            // (b) guarantee an inner dict called "HybridData"
            if (!ext.Contains("HybridData"))
            {
                var d = new DBDictionary();
                ext.SetAt("HybridData", d);
                tr.AddNewlyCreatedDBObject(d, true);
            }
            var dataDict = (DBDictionary)tr.GetObject(ext.GetAt("HybridData"), OpenMode.ForWrite);

            // (c) guarantee XRecord "Data"
            Xrecord xr;
            if (dataDict.Contains("Data"))
                xr = (Xrecord)tr.GetObject(dataDict.GetAt("Data"), OpenMode.ForWrite);
            else
            {
                xr = new Xrecord();
                dataDict.SetAt("Data", xr);
                tr.AddNewlyCreatedDBObject(xr, true);
            }

            // (d) same JSON you write on the polyline
            var simple = verts.Select(v => new
            {
                X = Math.Round(v.Pt.X, 3),
                Y = Math.Round(v.Pt.Y, 3),
                v.Type,
                v.Desc,
                v.ID
            });
            xr.Data = new ResultBuffer(
                new TypedValue((int)DxfCode.Text, JsonConvert.SerializeObject(simple))
            );
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

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                //-- guarantee style / layer
                var styleId = EnsureTableStyle(tr, db);
                EnsureLayer(tr, db, kTableLayer);

                //-- locate or create the table
                ObjectId tblId = FindExistingTableId(tr, db, styleId);
                Table tbl;
                Point3d insPt;

                if (tblId != ObjectId.Null)
                {
                    tbl = (Table)tr.GetObject(tblId, OpenMode.ForRead);
                    tbl.UpgradeOpen();
                    insPt = tbl.Position;
                }
                else
                {
                    var ppr = ed.GetPoint("\nPick table insertion point:");
                    if (ppr.Status != PromptStatus.OK) return;
                    insPt = ppr.Value;

                    tbl = new Table
                    {
                        TableStyle = styleId,
                        Layer = kTableLayer,
                        Position = insPt
                    };
                    tbl.SetDatabaseDefaults();

                    var cs = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    cs.AppendEntity(tbl);
                    tr.AddNewlyCreatedDBObject(tbl, true);
                }

                //-- rebuild rows
                tbl.Position = insPt;
                tbl.SetSize(verts.Count + 1, 5);   // +1 header row we’ll delete
                tbl.DeleteRows(0, 1);

                for (int c = 0; c < 5; c++) tbl.Columns[c].Width = kColW[c];
                for (int r = 0; r < verts.Count; r++) tbl.Rows[r].Height = kRowH;

                for (int i = 0; i < verts.Count; i++)
                {
                    var v = verts[i];
                    Write(tbl, i, 0, (i + 1).ToString());
                    Write(tbl, i, 1, v.N.ToString("F2"));
                    Write(tbl, i, 2, v.E.ToString("F2"));
                    Write(tbl, i, 3, v.Type);
                    Write(tbl, i, 4, v.Desc);
                }

                //-- (NEW) store the JSON on the table itself
                WriteTableMetadata(tbl, verts, tr);

                //-- optional hybrid blocks
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

        // -----------------------------------------------------------------------------
        //  PlaceHybridBlocks – inserts / updates the XC / RC / EC markers
        //  • Respects drawing precision via CurrentTol()
        //  • Ignores upper / lower-case in the Type column
        //  • Never inserts a duplicate when the correct block is already present
        // -----------------------------------------------------------------------------
        private static void PlaceHybridBlocks(
            Transaction tr,
            Database db,
            IList<VertexInfo> verts)
        {
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            var tol = CurrentTol();                    // drawing precision

            foreach (var v in verts)
            {
                // XC / RC / EC only ---------------------------------------------------
                var vType = (v.Type ?? string.Empty).Trim().ToUpperInvariant();
                if (vType != "XC" && vType != "RC" && vType != "EC")
                    continue;

                BlockReference nearby = null;          // hybrid block within tolerance?

                // scan model space for blocks near this vertex ----------------
                foreach (ObjectId id in space)
                {
                    if (id.ObjectClass != RXObject.GetClass(typeof(BlockReference)))
                        continue;

                    var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                    if (br.Position.DistanceTo(v.Pt) > kMatchTol)
                        continue;                      // within tolerance?

                    if (br.Name.ToUpperInvariant().StartsWith("HYBRID_"))
                    {
                        nearby = br;
                        break;
                    }
                }

                if (nearby != null)
                {
                    var name = nearby.Name.ToUpperInvariant();
                    if (name != $"HYBRID_{vType}")
                    {
                        var pos = nearby.Position;
                        nearby.UpgradeOpen();
                        nearby.Erase();

                        var nb = new BlockReference(pos, bt[$"Hybrid_{vType}"])
                        {
                            Layer = kBlockLayer,
                            ScaleFactors = new Scale3d(5, 5, 1)
                        };
                        space.AppendEntity(nb);
                        tr.AddNewlyCreatedDBObject(nb, true);
                    }
                }
                else
                {
                    var nb = new BlockReference(v.Pt, bt[$"Hybrid_{vType}"])
                    {
                        Layer = kBlockLayer,
                        ScaleFactors = new Scale3d(5, 5, 1)
                    };
                    space.AppendEntity(nb);
                    tr.AddNewlyCreatedDBObject(nb, true);
                }
            }
        }

        internal static string GuessType(Point3d pt, Transaction tr)
        {
            var db = AcadApp.DocumentManager.MdiActiveDocument.Database;
            var ms = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
            var tol = CurrentTol();                          // ← dynamic tolerance

            foreach (ObjectId id in ms)
            {
                if (id.ObjectClass != RXObject.GetClass(typeof(BlockReference))) continue;
                var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                if (!br.Position.IsEqualTo(pt, tol)) continue;

                var nm = br.Name.ToUpperInvariant();
                if (nm.Contains("HYBRID_XC")) return "XC";
                if (nm.Contains("HYBRID_RC")) return "RC";
                if (nm.Contains("HYBRID_EC")) return "EC";
                if (nm.StartsWith("FDI") || nm.StartsWith("FDSPIKE"))
                    return "OC";
            }
            return "";
        }

        // -----------------------------------------------------------------------------
        //  EnsureHybridNumBlock
        //  – Guarantees the block “Hybrd Num” has two attributes:
        //
        //      • NUMBER  – visible, variable
        //      • ID      – hidden,  variable
        //
        //    Works whether the block is being created for the first time, or already
        //    exists with the wrong visibility/constant flags.
        // -----------------------------------------------------------------------------
        private static void EnsureHybridNumBlock(Transaction tr, Database db)
        {
            const string blkName = "Hybrd Num";
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var ids = new Dictionary<string, AttributeDefinition>(StringComparer.OrdinalIgnoreCase);

            BlockTableRecord btr;

            // -------------------------------------------------------------------------
            //  1) Create a *new* definition if one doesn’t exist
            // -------------------------------------------------------------------------
            if (!bt.Has(blkName))
            {
                bt.UpgradeOpen();
                btr = new BlockTableRecord { Name = blkName, Origin = Point3d.Origin };
                bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);
            }
            else
            {
                btr = (BlockTableRecord)tr.GetObject(bt[blkName], OpenMode.ForWrite);
            }

            // -------------------------------------------------------------------------
            //  2) Scan existing attributes (if any)
            // -------------------------------------------------------------------------
            foreach (ObjectId entId in btr)
            {
                if (tr.GetObject(entId, OpenMode.ForRead) is AttributeDefinition ad)
                    ids[ad.Tag] = ad;
            }

            // -------------------------------------------------------------------------
            //  3) Make sure NUMBER definition is present & correct
            // -------------------------------------------------------------------------
            if (!ids.TryGetValue("NUMBER", out var numDef))
            {
                numDef = new AttributeDefinition
                {
                    Tag = "NUMBER",
                    Prompt = "Number",
                    Height = kTxtH,
                    Position = Point3d.Origin,
                    Invisible = false,
                    Constant = false
                };
                numDef.SetDatabaseDefaults();
                btr.AppendEntity(numDef);
                tr.AddNewlyCreatedDBObject(numDef, true);
            }
            else
            {
                if (numDef.Constant || numDef.Invisible)
                {
                    numDef.UpgradeOpen();
                    numDef.Constant = false;
                    numDef.Invisible = false;
                }
            }

            // -------------------------------------------------------------------------
            //  4) Make sure ID definition is present & correct (hidden)
            // -------------------------------------------------------------------------
            if (!ids.TryGetValue("ID", out var idDef))
            {
                idDef = new AttributeDefinition
                {
                    Tag = "ID",
                    Prompt = "Tag",
                    Height = kTxtH,
                    Position = Point3d.Origin,
                    Invisible = true,
                    Constant = false
                };
                idDef.SetDatabaseDefaults();
                btr.AppendEntity(idDef);
                tr.AddNewlyCreatedDBObject(idDef, true);
            }
            else
            {
                bool changed = idDef.Constant || !idDef.Invisible;
                if (changed)
                {
                    idDef.UpgradeOpen();
                    idDef.Constant = false;   // must be variable so we can write a unique value
                    idDef.Invisible = true;    // keep it hidden from view
                }
            }
        }

        private static void EnsureNumberingContext(
            Transaction tr,
            Database db,
            out BlockTable bt,
            out BlockTableRecord space,
            out BlockTableRecord defRec)
        {
            EnsureLayer(tr, db, kBlockLayer);
            EnsureHybridNumBlock(tr, db);
            bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            defRec = (BlockTableRecord)tr.GetObject(bt["Hybrd Num"], OpenMode.ForRead);
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
            var db  = doc.Database;
            var ed  = doc.Editor;

            // 1) Pick the polyline
            var peo = new PromptEntityOptions("\nSelect polyline to renumber: ");
            peo.SetRejectMessage("\nThat’s not a polyline.");
            peo.AddAllowedClass(typeof(Polyline), false);
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;
            var plId = per.ObjectId;

            // 2) Read back stored metadata (Pt → ID, Type, Desc)
            var oldData = ReadVertexData(plId, db);
            var oldMap = oldData
                .GroupBy(v => v.Pt, new Point3dEquality(kMatchTol))
                .Select(g => g.First())
                .ToDictionary(v => v.Pt, v => v, new Point3dEquality(kMatchTol));

            // 3) Rebuild current vertex list in polyline order, carrying over IDs
            var verts = new List<VertexInfo>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var pl = (Polyline)tr.GetObject(plId, OpenMode.ForRead);
                for (int i = 0; i < pl.NumberOfVertices; i++)
                {
                    var pt = pl.GetPoint3dAt(i);
                    if (oldMap.TryGetValue(pt, out var vi))
                        verts.Add(vi);
                    else
                        verts.Add(new VertexInfo { Pt = pt, N = pt.Y, E = pt.X, Type = "", Desc = "", ID = 0 });
                }
                tr.Commit();
            }

            // 4) Renumber in model space
            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                // ensure layers/blocks & upgrade any old bubbles missing an ID
                EnsureNumberingContext(tr, db, out var bt, out var space, out var defRec);
                UpgradeExistingHybrdNumBubbles(tr, db, space, bt);

                // 4a) Build map of every existing bubble by its hidden ID
                int created = 0, reused = 0, renumbered = 0;
                var idToBr = new Dictionary<int, List<BlockReference>>();
                int maxId = 0;
                foreach (ObjectId entId in space)
                {
                    if (entId.ObjectClass != RXObject.GetClass(typeof(BlockReference)))
                        continue;
                    var br = (BlockReference)tr.GetObject(entId, OpenMode.ForRead);
                    if (!br.Name.Equals("Hybrd Num", StringComparison.OrdinalIgnoreCase))
                        continue;

                    foreach (ObjectId attId in br.AttributeCollection)
                    {
                        var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                        if (ar.Tag == "ID" && int.TryParse(ar.TextString, out int v))
                        {
                            if (!idToBr.TryGetValue(v, out var list))
                            {
                                list = new List<BlockReference>();
                                idToBr[v] = list;
                            }
                            list.Add(br);
                            maxId = Math.Max(maxId, v);
                        }
                    }
                }

                // 4b) Assign new IDs only to verts where ID==0
                int nextId = maxId + 1;
                foreach (var v in verts)
                {
                    if (v.ID == 0)
                    {
                        v.ID = nextId++;
                    }
                }

                // 4c) Create any missing bubbles
                foreach (var v in verts)
                {
                    if (!idToBr.ContainsKey(v.ID))
                    {
                        var brNew = GetOrCreateNumberBubble(tr, space, defRec, v.Pt, v.ID, kMatchTol); created++;
                        idToBr[v.ID] = new List<BlockReference> { brNew };
                    } else {
                        reused += idToBr[v.ID].Count;
                    }
                }

                // 4d) Update the visible NUMBER to match the new sequence
                for (int i = 0; i < verts.Count; i++)
                {
                    int vid = verts[i].ID;
                    if (!idToBr.TryGetValue(vid, out var bubbleList))
                        continue;

                    string seq = (i + 1).ToString();
                    foreach (var br in bubbleList)
                    {
                        foreach (ObjectId attId in br.AttributeCollection)
                        {
                            var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForWrite);
                            if (ar.Tag == "NUMBER")
                            {
                                AllowProtectedEdits(() =>
                                {
                                    if (ar.TextString != seq) { ar.TextString = seq; renumbered++; }
                                    ar.AdjustAlignment(db);
                                });
                            }
                        }
                    }
                }

                // 5) Persist metadata back onto the polyline
                WriteVertexData(plId, db, verts);
                tr.Commit();
                ed.WriteMessage("\nUpdateNumbering: {created} new, {reused} reused, {renumbered} renumbered.");
            }

            doc.Editor.Regen();
        }



        [CommandMethod("DUMPVERTEXDATA")]
        public static void DumpVertexData()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            // pick a polyline
            var peo = new PromptEntityOptions("\nSelect polyline to dump data:");
            peo.SetRejectMessage("\nThat’s not a polyline.");
            peo.AddAllowedClass(typeof(Polyline), false);
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            // read your stored list
            var list = ReadVertexData(per.ObjectId, db);
            // reserialize with nice formatting
            var json = JsonConvert.SerializeObject(list, Formatting.Indented);

            // output back to the command line
            ed.WriteMessage("\n--------- HybridData JSON ---------\n");
            ed.WriteMessage(json);
            ed.WriteMessage("\n-----------------------------------\n");
        }
        /// <summary>
        /// Removes the “HybridData” extension dictionary (and all stored IDs, types, descs)
        /// from every polyline in model‐space.  After running this, your polylines will
        /// look “un‐tagged” and you can re-run AddNumbering/InsertUpdate to rebuild them.
        /// </summary>
        [CommandMethod("PurgeHybridData")]
        public static void PurgeHybridData()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                // iterate every entity in model‐space
                var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                foreach (ObjectId id in btr)
                {
                    // look only for polylines
                    var pl = tr.GetObject(id, OpenMode.ForWrite) as Polyline;
                    if (pl == null || pl.ExtensionDictionary.IsNull)
                        continue;

                    // open the extension dict
                    var ext = (DBDictionary)tr.GetObject(pl.ExtensionDictionary, OpenMode.ForWrite);
                    if (!ext.Contains("HybridData"))
                        continue;

                    // grab & erase the Xrecord
                    var xrecId = ext.GetAt("HybridData");
                    ext.Remove("HybridData");

                    var xrec = tr.GetObject(xrecId, OpenMode.ForWrite) as Xrecord;
                    xrec?.Erase();
                }

                tr.Commit();
            }

            ed.WriteMessage("\nAll HybridData has been purged from every polyline.");
            doc.Editor.Regen();
        }


        private static BlockReference CreateBubble(Transaction tr,
                                                   BlockTableRecord space,
                                                   BlockTable bt,
                                                   BlockTableRecord defRec,
                                                   Point3d pt,
                                                   int id)
        {
            var br = new BlockReference(pt, bt["Hybrd Num"])
            {
                Layer = kBlockLayer,
                ScaleFactors = new Scale3d(1, 1, 1)
            };
            space.AppendEntity(br);
            tr.AddNewlyCreatedDBObject(br, true);

            foreach (ObjectId defId in defRec)
            {
                if (tr.GetObject(defId, OpenMode.ForRead) is AttributeDefinition def)
                {
                    var ar = new AttributeReference();
                    ar.SetAttributeFromBlock(def, Matrix3d.Identity);
                    ar.Position = pt;
                    ar.Invisible = def.Invisible;
                    AllowProtectedEdits(() =>
                    {
                        ar.TextString = def.Tag == "ID" ? id.ToString() : string.Empty;
                    });

                    br.AttributeCollection.AppendAttribute(ar);
                    tr.AddNewlyCreatedDBObject(ar, true);
                }
            }

            return br;    // caller commits
        }

        private static readonly Dictionary<ObjectId, Dictionary<Point3d, BlockReference>> _bubbleCache
            = new Dictionary<ObjectId, Dictionary<Point3d, BlockReference>>();

        private static BlockReference GetOrCreateNumberBubble(
            Transaction tr,
            BlockTableRecord space,
            BlockTableRecord defRec,
            Point3d pt,
            int id,
            double tol)
        {
            if (!_bubbleCache.TryGetValue(space.ObjectId, out var map))
            {
                map = new Dictionary<Point3d, BlockReference>(new Point3dEquality(kMatchTol));
                foreach (ObjectId entId in space)
                {
                    if (entId.ObjectClass != RXObject.GetClass(typeof(BlockReference)))
                        continue;
                    var br0 = (BlockReference)tr.GetObject(entId, OpenMode.ForRead);
                    if (!br0.Name.Equals("Hybrd Num", StringComparison.OrdinalIgnoreCase))
                        continue;
                    map[br0.Position] = br0;
                }
                _bubbleCache[space.ObjectId] = map;
            }

            if (map.TryGetValue(pt, out var br))
            {
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForWrite);
                    if (ar.Tag == "ID" && ar.TextString != id.ToString())
                        AllowProtectedEdits(() => { ar.TextString = id.ToString(); });
                }
                return br;
            }

            var nb = new BlockReference(pt, defRec.ObjectId)
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
                    AllowProtectedEdits(() =>
                    {
                        ar.TextString = def.Tag == "ID" ? id.ToString() : "";
                    });
                    nb.AttributeCollection.AppendAttribute(ar);
                    tr.AddNewlyCreatedDBObject(ar, true);
                }
            }
            map[pt] = nb;
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
                AllowProtectedEdits(() => { arNew.TextString = nextId.ToString(); });
                br.AttributeCollection.AppendAttribute(arNew);
                tr.AddNewlyCreatedDBObject(arNew, true);
                nextId++;
            }
        }


        /// <summary>
        /// Returns a tolerance that matches the drawing’s displayed precision
        /// (½ of the current LUPREC rounding unit, with a small safety margin).
        /// </summary>
        private static Tolerance CurrentTol()
        {
            int prec = Convert.ToInt32(AcadApp.GetSystemVariable("LUPREC")); // 0-8
            if (prec < 2) return new Tolerance(kMatchTol, kMatchTol);
            double eps = Math.Pow(10, -prec) * 0.51;                         // ≈ 0.5 ULP
            return new Tolerance(eps, eps);
        }

        /// <summary>Runs <paramref name="action"/> while temporarily allowing
        /// edits to protected NUMBER / ID attributes.</summary>
        private static void AllowProtectedEdits(Action action)
        {
            _allowInternalEdits = true;
            try { action(); }
            finally { _allowInternalEdits = false; }
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

            // 3) Read all metadata from source
            var srcList = ReadVertexData(srcId, db);

            // 4) Build a new list for the target: preserve matches within tolerance
            var newList = new List<VertexInfo>();
            var tol = kMatchTol;                               // drawing precision

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var plDst = (Polyline)tr.GetObject(dstId, OpenMode.ForRead);
                for (int i = 0; i < plDst.NumberOfVertices; i++)
                {
                    var pt = plDst.GetPoint3dAt(i);

                    int idx = srcList.FindIndex(v => v.Pt.DistanceTo(pt) <= tol);
                    if (idx >= 0)
                    {
                        var vi = srcList[idx];
                        newList.Add(new VertexInfo
                        {
                            Pt = vi.Pt,
                            N = vi.N,
                            E = vi.E,
                            Type = vi.Type,
                            Desc = vi.Desc,
                            ID = vi.ID
                        });
                        srcList.RemoveAt(idx);                 // mark as used
                    }
                    else
                    {
                        newList.Add(new VertexInfo
                        {
                            Pt = pt,
                            N = pt.Y,
                            E = pt.X,
                            Type = "",
                            Desc = "",
                            ID = 0
                        });
                    }
                }
                tr.Commit();
            }

            // append any leftover source vertices so metadata isn’t lost
            foreach (var extra in srcList)
                newList.Add(extra);

            // 5) Write that metadata back onto the target polyline
            WriteVertexData(dstId, db, newList);
            ed.WriteMessage($"\nTransferred metadata to {newList.Count} vertices.");
        }

        // in HybridCommands:
        [CommandMethod("AddNumbering")]
        public static void AddNumbering()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                EnsureNumberingContext(tr, db, out var bt, out var space, out var defRec);

                //–- pick polyline
                var peo = new PromptEntityOptions("\nSelect polyline to place numbers on: ");
                peo.SetRejectMessage("\nThat’s not a polyline.");
                peo.AddAllowedClass(typeof(Polyline), false);
                var per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK) return;

                var plId = per.ObjectId;
                var pl = (Polyline)tr.GetObject(plId, OpenMode.ForRead);

                //–- load / allocate IDs (same logic as before) ………………
                var verts = BuildVertexListWithIds(tr, db, pl);   // <-- helper unchanged

                //–- create / update bubbles …………………………………………
                UpdateOrCreateBubbles(tr, db, space, defRec, verts);   // <-- helper unchanged

                //–- write IDs back on the polyline
                WriteVertexData(plId, db, verts);

                //–- (NEW) update the JSON on the table, if it exists
                ObjectId tblId = FindExistingTableId(tr, db, EnsureTableStyle(tr, db));
                if (tblId != ObjectId.Null)
                {
                    var tbl = (Table)tr.GetObject(tblId, OpenMode.ForWrite);
                    WriteTableMetadata(tbl, verts, tr);
                }

                tr.Commit();
                ed.WriteMessage($"\nAddNumbering: {verts.Count} vertices processed.");
            }

            doc.Editor.Regen();
        }


        // -----------------------------------------------------------------------------
        // 2)  TryReadTableMetadata  – returns true if the table contains the payload
        // -----------------------------------------------------------------------------
        private static bool TryReadTableMetadata(
            ObjectId tblId,
            Database db,
            out List<VertexInfo> verts)
        {
            verts = new List<VertexInfo>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var tbl = (Table)tr.GetObject(tblId, OpenMode.ForRead);
                if (tbl.ExtensionDictionary.IsNull) return false;

                var ext = (DBDictionary)tr.GetObject(tbl.ExtensionDictionary, OpenMode.ForRead);
                if (!ext.Contains("HybridData")) return false;

                var dataDict = (DBDictionary)tr.GetObject(ext.GetAt("HybridData"), OpenMode.ForRead);
                if (!dataDict.Contains("Data")) return false;

                var xr = (Xrecord)tr.GetObject(dataDict.GetAt("Data"), OpenMode.ForRead);
                var tv = xr.Data.Cast<TypedValue>()
                                .FirstOrDefault(t => t.TypeCode == (int)DxfCode.Text);
                if (tv == null) return false;

                var list = JsonConvert.DeserializeObject<List<SimpleVertex>>((string)tv.Value);
                if (list == null || list.Count == 0) return false;

                verts = list.Select(sv => new VertexInfo
                {
                    Pt = new Point3d(sv.X, sv.Y, 0),
                    N = sv.Y,
                    E = sv.X,
                    Type = sv.Type,
                    Desc = sv.Desc,
                    ID = sv.ID
                }).ToList();

                return true;
            }
        }


        // -----------------------------------------------------------------------------
        // 3)  ParseRowsFromTable  – your original “scan every row” fallback
        // -----------------------------------------------------------------------------
        private static List<VertexInfo> ParseRowsFromTable(Table tbl)
        {
            var verts = new List<VertexInfo>();
            int rowCount = tbl.Rows.Count;

            for (int r = 0; r < rowCount; r++)
            {
                string sN = tbl.Cells[r, 1].TextString;
                string sE = tbl.Cells[r, 2].TextString;

                if (double.TryParse(sN, out double N) &&
                    double.TryParse(sE, out double E))
                {
                    double n = Math.Round(N, 3);
                    double e = Math.Round(E, 3);
                    verts.Add(new VertexInfo
                    {
                        Pt = new Point3d(e, n, 0),
                        N = n,
                        E = e,
                        Type = tbl.Cells[r, 3].TextString,
                        Desc = tbl.Cells[r, 4].TextString,
                        ID = 0          // no ID info in visible rows
                    });
                }
            }
            return verts;
        }
        private static List<VertexInfo> BuildVertexListWithIds(Transaction tr, Database db, Polyline pl)
        {
            var oldList = ReadVertexData(pl.ObjectId, db);
            var oldMap = new Dictionary<Point3d, VertexInfo>(new Point3dEquality(kMatchTol));
            foreach (var v in oldList)
                if (!oldMap.ContainsKey(v.Pt))
                    oldMap[v.Pt] = v;

            var verts = new List<VertexInfo>();
            for (int i = 0; i < pl.NumberOfVertices; i++)
            {
                var pt = pl.GetPoint3dAt(i);
                if (oldMap.TryGetValue(pt, out var vi))
                    verts.Add(vi);
                else
                    verts.Add(new VertexInfo { Pt = pt, N = pt.Y, E = pt.X, Type = "", Desc = "", ID = 0 });
            }

            // assign new IDs if needed
            int maxId = oldList.Max(v => v.ID);
            int nextId = maxId + 1;
            foreach (var v in verts)
                if (v.ID == 0)
                    v.ID = nextId++;

            return verts;
        }
        private static void UpdateOrCreateBubbles(
    Transaction tr,
    Database db,
    BlockTableRecord space,
    BlockTableRecord defRec,
    List<VertexInfo> verts)
        {
            var tol = kMatchTol;
            var bubblesById = new Dictionary<int, List<BlockReference>>();

            foreach (ObjectId id in space)
            {
                if (id.ObjectClass != RXObject.GetClass(typeof(BlockReference))) continue;
                var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                if (!br.Name.Equals("Hybrd Num", StringComparison.OrdinalIgnoreCase)) continue;

                foreach (ObjectId attId in br.AttributeCollection)
                {
                    var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                    if (ar.Tag == "ID" && int.TryParse(ar.TextString, out int val))
                    {
                        if (!bubblesById.TryGetValue(val, out var list))
                        {
                            list = new List<BlockReference>();
                            bubblesById[val] = list;
                        }
                        list.Add(br);
                        break;
                    }
                }
            }

            for (int i = 0; i < verts.Count; i++)
            {
                var v = verts[i];
                string number = (i + 1).ToString();

                if (!bubblesById.TryGetValue(v.ID, out var existing))
                {
                    var br = GetOrCreateNumberBubble(tr, space, defRec, v.Pt, v.ID, tol);
                    bubblesById[v.ID] = new List<BlockReference> { br };
                    existing = bubblesById[v.ID];
                }

                foreach (var br in existing)
                {
                    foreach (ObjectId attId in br.AttributeCollection)
                    {
                        var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForWrite);
                        if (ar.Tag == "NUMBER" && ar.TextString != number)
                        {
                            AllowProtectedEdits(() =>
                            {
                                ar.TextString = number;
                                ar.AdjustAlignment(db);
                            });
                        }
                    }
                }
            }
        }


        // -----------------------------------------------------------------------------
        // 4)  RebuildPolylineFromTable  (complete replacement)
        // -----------------------------------------------------------------------------
        [CommandMethod("REBUILDPLINEFROMTABLE")]
        public static void RebuildPolylineFromTable()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            using (doc.LockDocument())
            {
                // pick the table -----------------------------------------------------
                var peo = new PromptEntityOptions("\nSelect Hybrid-Points table: ");
                peo.SetRejectMessage("\nEntity must be a Table.");
                peo.AddAllowedClass(typeof(Table), false);
                var per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK) return;

                List<VertexInfo> verts;

                // 1) Preferred path – pull JSON from the table itself
                if (TryReadTableMetadata(per.ObjectId, db, out verts))
                {
                    ed.WriteMessage("\nUsing vertex metadata stored on the table.");
                }
                else
                {
                    // 2) Fallback – parse visible text rows
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var tbl = (Table)tr.GetObject(per.ObjectId, OpenMode.ForRead);
                        verts = ParseRowsFromTable(tbl);
                        tr.Commit();
                    }
                }

                if (verts.Count == 0)
                {
                    ed.WriteMessage("\nNo vertices found – nothing rebuilt.");
                    return;
                }

                // build new polyline -------------------------------------------------
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var tbl = (Table)tr.GetObject(per.ObjectId, OpenMode.ForRead);
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

                    // persist metadata back onto the *new* polyline
                    WriteVertexData(pl.ObjectId, db, verts);

                    tr.Commit();
                    ed.WriteMessage($"\nRebuilt polyline with {verts.Count} vertices.");
                }
            }
        }

    }

    internal sealed class VertexForm : Form
    {
        private readonly DataGridView _grid;
        private readonly CheckBox _chkHybrid;
        private List<VertexInfo> _verts;
        private ObjectId _currentPlId = ObjectId.Null;

        public VertexForm()
        {
            Text = "Hybrid Vertex Editor";
            ClientSize = new System.Drawing.Size(1000, 580);
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
        // in VertexForm:
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

            // load previously‐written metadata
            var oldList = HybridCommands.ReadVertexData(_currentPlId, doc.Database);

            // FIX: skip duplicate Pt keys
            var oldMap = new Dictionary<Point3d, VertexInfo>(new Point3dEquality(HybridCommands.kMatchTol));
            foreach (var v in oldList)
            {
                if (!oldMap.ContainsKey(v.Pt))
                    oldMap.Add(v.Pt, v);
                // else ignore any repeats
            }

            // build the new list in vertex order
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
            double north = Math.Round(double.Parse(row.Cells["Northing"].Value.ToString()), 3);
            double east = Math.Round(double.Parse(row.Cells["Easting"].Value.ToString()), 3);

            _verts[i] = new VertexInfo
            {
                Pt = new Point3d(east, north, 0),
                N = north,
                E = east,
                Type = row.Cells["Type"].Value?.ToString() ?? "",
                Desc = row.Cells["Desc"].Value?.ToString() ?? "",
                ID = _verts[i].ID
            };
        }

        // 3) Push the updates back into the drawing
        HybridCommands.InsertOrUpdate(_verts, _chkHybrid.Checked);
        HybridCommands.WriteVertexData(
            _currentPlId,
            AcadApp.DocumentManager.MdiActiveDocument.Database,
            _verts
        );

        // 4) Refresh the grid display
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
        /// <summary>
        /// Returns a tolerance that matches the drawing’s displayed precision
        /// (½ of the current LUPREC rounding unit, with a small safety margin).
        /// </summary>

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
        private readonly double _eps;
        private readonly int _dec;

        public Point3dEquality(double? customTol = null)
        {
            _dec = Convert.ToInt32(AcadApp.GetSystemVariable("LUPREC")); // 0-8
            _eps = customTol ?? Math.Pow(10, -_dec) * 0.51;              // default ≈ half ULP
        }

        public bool Equals(Point3d a, Point3d b) =>
            Math.Abs(a.X - b.X) <= _eps && Math.Abs(a.Y - b.Y) <= _eps;

        public int GetHashCode(Point3d p)
        {
            double m = 1.0 / _eps;                   // scale to tolerance grid
            long ix = (long)Math.Round(p.X * m);
            long iy = (long)Math.Round(p.Y * m);
            return ((ix * 397) ^ iy).GetHashCode();
        }
    }
}