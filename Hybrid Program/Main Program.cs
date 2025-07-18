// HybridSurvey_MainProgram.cs
// PATCHED by Codex, 2025-06-10 – brace & scope fixes
// REFACTORED 2025-06-10
// Hybrid Survey – AutoCAD 2013–2025 plug-in
// test upload
#region usings
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;       // ← ADDED: gives you Size, Point, Color...
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Security.Cryptography;
using System.Text;          // StringBuilder in copy/paste helpers
using System.Windows.Forms;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using AcColor = Autodesk.AutoCAD.Colors.Color;
using System.IO;
#endregion

[assembly: CommandClass(typeof(HybridSurvey.HybridCommands))]

namespace HybridSurvey
{
    // ----------------------------------------------------------------------
    //  Basic data types used across the plug-in
    // ----------------------------------------------------------------------
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
            _eps = customTol ?? Math.Pow(10, -_dec) * 0.51; // ≈ half ULP
        }

        public bool Equals(Point3d a, Point3d b) =>
            Math.Abs(a.X - b.X) <= _eps && Math.Abs(a.Y - b.Y) <= _eps;

        public int GetHashCode(Point3d p)
        {
            double m = 1.0 / _eps; // scale to tolerance grid
            long ix = (long)Math.Round(p.X * m);
            long iy = (long)Math.Round(p.Y * m);
            return ((ix * 397) ^ iy).GetHashCode();
        }
    }

    // -----------------------------------------------------------------------------
    //  HybridGuard  – prevents manual edits to “Hybrd Num” bubbles
    // -----------------------------------------------------------------------------
    public sealed class HybridGuard : IExtensionApplication
    {
        /* ===============================  state  =============================== */
        private static readonly HashSet<IntPtr> _wired = new HashSet<IntPtr>();
        private static readonly Dictionary<ObjectId, string> _orig = new Dictionary<ObjectId, string>();
        private static DateTime _lastPopup = DateTime.MinValue;
        private const int POPUP_COOLDOWN_SEC = 60;
        // ─────────────────────────────────────────────────────────────────────────
        //  NEW: commands that should *pause* the guard while they run
        // ─────────────────────────────────────────────────────────────────────────
        private static readonly HashSet<string> _passThroughCmds =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        // attribute-safe edits
        "ATTSYNC",
        "MOVE","COPY","ROTATE","SCALE","MIRROR","STRETCH",
        "SAVE","QSAVE","SAVEAS","QUIT","EXIT","CLOSE","CLOSEALL","CLOSEALLOTHER",

        // grip / interactive
        "GRIP_MOVE","GRIP_STRETCH","GRIP_SCALE","GRIP_ROTATE","GRIP_MIRROR",
        "DRAGMOVE","DRAG","NUDGE",          // ← NEW: mouse drag & nudge

        // clipboard / insert
        "PASTECLIP","PASTEORIG",
        "INSERT","-INSERT","PINSERT",
        "COPYCLIP","COPYBASE",

        // undo / redo family
        "UNDO","U","REDO","OOPS"
    };
        /* ---  internal “suspend” flag ----------------------------------------- */
        private static int _suspendDepth;
        internal static IDisposable Suspend()
        {
            _suspendDepth++;
            return new SuspendCookie();
        }
        private sealed class SuspendCookie : IDisposable
        {
            public void Dispose() => _suspendDepth--;
        }
        private static bool Suspended => _suspendDepth > 0;
        // ──────────────────────────────────────────────────────────────────────────
        //  Helper – safely test if a Database is still being loaded (property exists
        //  only in AutoCAD 2025+). Returns false on older releases.
        // ──────────────────────────────────────────────────────────────────────────
        private static bool IsDbLoading(Database db)
        {
            var pi = typeof(Database).GetProperty("IsBeingLoaded");
            return pi != null && pi.PropertyType == typeof(bool) && (bool)pi.GetValue(db);
        }

        /* =======================  IExtensionApplication ======================= */
        public void Initialize()
        {
            // ── 1) crash black-box ─────────────────────────────────────────────
            AppDomain.CurrentDomain.UnhandledException += (s, a) =>
            {
                try
                {
                    var ex = (System.Exception)a.ExceptionObject;  // ← fully-qualified
                    string dump =
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}]\n" +
                        $"Unhandled: {ex}\n" +
                        $"Inner:     {ex.InnerException}\n\n";

                    System.IO.File.AppendAllText(
                        System.IO.Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                            "HybridGuard_log.txt"),
                        dump);
                }
                catch { /* absolutely nothing */ }
            };

            // ── 2) original initialise body ────────────────────────────────────
            foreach (Document doc in AcadApp.DocumentManager)
            {
                hook(doc.Database);
                WireCommandHooks(doc);
            }

            AcadApp.DocumentManager.DocumentCreated += OnDocCreated;
            AcadApp.DocumentManager.DocumentActivated += OnDocActivated;
        }

        public void Terminate()
        {
            foreach (Document doc in AcadApp.DocumentManager)
            {
                unhook(doc.Database);
                UnwireCommandHooks(doc);
            }

            AcadApp.DocumentManager.DocumentCreated -= OnDocCreated;
            AcadApp.DocumentManager.DocumentActivated -= OnDocActivated;

            _wired.Clear();
            _orig.Clear();
        }

        /* ============================  wiring  ================================= */
        private static void hook(Database db)
        {
            if (db == null || db.IsDisposed || _wired.Contains(db.UnmanagedObject))
                return;

            db.ObjectOpenedForModify += onOpened;
            db.ObjectModified += onModified;
            _wired.Add(db.UnmanagedObject);
        }
        private static void unhook(Database db)
        {
            if (db == null || !_wired.Contains(db.UnmanagedObject)) return;
            db.ObjectOpenedForModify -= onOpened;
            db.ObjectModified -= onModified;
            _wired.Remove(db.UnmanagedObject);
        }

        /* ---------- document‑level helpers ------------------------------------ */
        private static void WireCommandHooks(Document doc)
        {
            doc.CommandWillStart += CmdWillStart;
            doc.CommandEnded += CmdEnded;
            doc.CommandCancelled += CmdEnded;
            doc.CommandFailed += CmdEnded;
        }
        private static void UnwireCommandHooks(Document doc)
        {
            doc.CommandWillStart -= CmdWillStart;
            doc.CommandEnded -= CmdEnded;
            doc.CommandCancelled -= CmdEnded;
            doc.CommandFailed -= CmdEnded;
        }
        private static void OnDocCreated(object sender, DocumentCollectionEventArgs e)
        {
            hook(e.Document.Database);
            WireCommandHooks(e.Document);
        }
        private static void OnDocActivated(object sender, DocumentCollectionEventArgs e)
        {
            hook(e.Document.Database);   // idempotent
        }


        /* ============================  events  ================================= */
        // Caches original NUMBER text when the attribute is first opened for modify
        private static void onOpened(object sender, ObjectEventArgs e)
        {
            var db = e.DBObject?.Database;
            if (db == null || IsDbLoading(db)) return;   // bail while drawing is still loading
            if (Suspended) return;

            if (e.DBObject is AttributeReference ar &&
                ar.Tag.Equals("NUMBER", StringComparison.OrdinalIgnoreCase))
            {
                _orig[ar.ObjectId] = ar.TextString;
            }
        }

        // Reverts unauthorised edits to NUMBER attributes & rate-limits the warning
        private static void onModified(object sender, ObjectEventArgs e)
        {
            var db = e.DBObject?.Database;
            if (db == null || IsDbLoading(db)) return;   // drawing still streaming in
            if (Suspended) return;                       // guard temporarily paused

            var ar = e.DBObject as AttributeReference;
            if (ar == null || !ar.Tag.Equals("NUMBER", StringComparison.OrdinalIgnoreCase))
                return;

            /* 1) make sure the attribute is inside a Hybrd Num block */
            var tr = ar.Database.TransactionManager;
            var br = ar.OwnerId.IsValid
                   ? tr.GetObject(ar.OwnerId, OpenMode.ForRead) as BlockReference
                   : null;
            if (br == null) return;

            string blkName = br.IsDynamicBlock
                ? ((BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead)).Name
                : br.Name;

            if (!blkName.Equals("Hybrd Num", StringComparison.OrdinalIgnoreCase))
                return;

            /* 2) roll back manual edits */
            if (_orig.TryGetValue(ar.ObjectId, out string oldVal) && oldVal != ar.TextString)
            {
                if (!ar.IsWriteEnabled) ar.UpgradeOpen();
                ar.TextString = oldVal;
                ar.AdjustAlignment(ar.Database);
            }
            _orig.Remove(ar.ObjectId);

            /* 3) gentle reminder (rate-limited & only when UI is interactive) */
            if ((DateTime.Now - _lastPopup).TotalSeconds >= POPUP_COOLDOWN_SEC &&
                AcadApp.IsQuiescent)
            {
                _lastPopup = DateTime.Now;
                AcadApp.ShowAlertDialog(
                    "PLEASE DON’T EDIT THIS BLOCK MANUALLY!\n\n" +
                    "The “NUMBER” attribute is maintained automatically and any manual " +
                    "change you make is immediately reverted.");
            }
        }
        private static string Canon(string cmd)
        {
            int i = 0;
            while (i < cmd.Length && (cmd[i] == '.' || cmd[i] == '_' ||
                                      cmd[i] == '*' || cmd[i] == '\''))
                i++;
            return cmd.Substring(i);
        }

        /* ---------- command hooks: suspend guard during ATTSYNC --------------- */
        private static void CmdWillStart(object s, CommandEventArgs e)
        {
            if (_passThroughCmds.Contains(Canon(e.GlobalCommandName)))
                _suspendDepth++;              // pause the guard
        }

        private static void CmdEnded(object s, CommandEventArgs e)
        {
            if (_passThroughCmds.Contains(Canon(e.GlobalCommandName)) && _suspendDepth > 0)
                _suspendDepth--;              // resume
        }

    }

    internal static class HybridCommands
    {
        private const string kTableStyle = "Induction Bend";
        private const string kTableLayer = "Hybrid_Points_TBL";
        private const string kBlockLayer = "L-MON";
        private const double kTxtH = 2.5;
        private static readonly double[] kColW = { 40, 60, 60, 40, 120 };
        internal const double kMatchTol = 0.03;       // vertex search tolerance
        private const double kRowH = 4.0;

        // inside HybridCommands, replace your old SimpleVertex with this:
        private class SimpleVertex
        {
            public double X, Y;
            public string Type, Desc;
            public int ID;         // ← new
        }

        // ──────────────────────────────────────────────────────────────────────────────
        // 1)  WriteVertexData  – now safely handles a polyline that was ERASED / UNDOed
        // ──────────────────────────────────────────────────────────────────────────────
        public static void WriteVertexData(ObjectId plId,
                                           Database db,
                                           List<VertexInfo> verts)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                // ── SAFE OPEN ──
                Polyline pl;
                try
                {
                    // ‘false’ → don’t allow access to an object that is already erased
                    pl = (Polyline)tr.GetObject(plId, OpenMode.ForWrite, false);
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex) when (IsWasErased(ex))
                {
                    // the polyline was deleted/undone – silently bail
                    return;
                }

                if (pl == null) return;          // sanity (shouldn’t happen)

                //-- ensure the polyline owns an extension dictionary
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

                tr.Commit();     // commit once, here
            }
        }
        // put this helper somewhere in HybridCommands, HybridGuard, or a utilities file
        private static bool IsWasErased(Autodesk.AutoCAD.Runtime.Exception ex)
        {
            const int kWasErased = 52;   // ErrorStatus.eWasErased in all full runtimes
            return (int)ex.ErrorStatus == kWasErased;
        }

        internal static void WriteTableMetadata(Table tbl, IList<VertexInfo> verts, Transaction tr)
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
        /// <summary>
        /// Ensures the hidden ID attribute equals <paramref name="id"/>.
        /// </summary>
        private static void UpdateHiddenIdIfNeeded(
            Transaction tr,
            BlockReference br,
            int id)
        {
            foreach (ObjectId attId in br.AttributeCollection)
            {
                var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForWrite);
                if (ar.Tag == "ID" && ar.TextString != id.ToString())
                {
                    ar.TextString = id.ToString();
                    return;
                }
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
            // If an editor form already exists, just show/activate it
            if (VertexForm.ActiveInstance != null && !VertexForm.ActiveInstance.IsDisposed)
            {
                var f = VertexForm.ActiveInstance;
                if (f.WindowState == FormWindowState.Minimized) f.WindowState = FormWindowState.Normal;
                f.BringToFront();
                f.Focus();
                return;
            }

            // …otherwise create a brand-new one
            var form = new VertexForm();
            form.Show();    // ActiveInstance is set inside the constructor
        }

        public static void InsertOrUpdate(IList<VertexInfo> verts, bool insertHybrid)
        {
            // ‑‑‑ do NOT fire the guard while we touch attributes
            using (HybridGuard.Suspend())
            {
                if (verts == null || verts.Count == 0) return;

                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                var db = doc.Database;
                var ed = doc.Editor;

                using (doc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    /* ---- existing body is IDENTICAL from here down ---- */
                    var styleId = EnsureTableStyle(tr, db);
                    EnsureLayer(tr, db, kTableLayer);

                    ObjectId tblId = FindExistingTableId(tr, db, styleId, kTableLayer);
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

                    tbl.Position = insPt;
                    tbl.SetSize(verts.Count + 1, 5);
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

                    WriteTableMetadata(tbl, verts, tr);

                    if (insertHybrid)
                    {
                        EnsureAllHybridBlocks(tr, db);
                        PlaceHybridBlocks(tr, db, verts);
                    }

                    tr.Commit();
                }
            } // ← guard resumes
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

        internal static ObjectId EnsureTableStyle(Transaction tr, Database db)
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

                // ⬇️ Only touch these flags if it is *not* the current layer
                if (db.Clayer != ltr.ObjectId)
                {
                    ltr.IsFrozen = false;
                    ltr.IsLocked = false;
                }
            }
        }

        /// <summary>
        /// Returns the ObjectId of the first table that
        ///   • has the requested TableStyle, and
        ///   • sits on the given layer.
        /// If none found, returns ObjectId.Null.
        /// </summary>
        internal static ObjectId FindExistingTableId(
            Transaction tr,
            Database db,
            ObjectId styleId,
            string layerName)
        {
            var ms = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

            foreach (ObjectId entId in ms)
            {
                var tbl = tr.GetObject(entId, OpenMode.ForRead) as Table;
                if (tbl == null) continue;

                bool styleMatch = tbl.TableStyle == styleId;
                bool layerMatch = tbl.Layer.Equals(layerName, StringComparison.OrdinalIgnoreCase);

                if (styleMatch && layerMatch)
                    return entId;               // ← found the one we should update
            }

            return ObjectId.Null;               // ← none on that layer ⇒ create new
        }

        private static void EnsureAllHybridBlocks(Transaction tr, Database db)
        {
            EnsureHybridBlock(tr, db, "Hybrid_XC", AcColor.FromColorIndex(ColorMethod.ByAci, 1));
            EnsureHybridBlock(tr, db, "Hybrid_RC", AcColor.FromColorIndex(ColorMethod.ByAci, 2));
            EnsureHybridBlock(tr, db, "Hybrid_EC", AcColor.FromColorIndex(ColorMethod.ByAci, 3));
        }

        private static void EnsureHybridBlock(Transaction tr, Database db,
                                              string name, AcColor col)   // <─ use the alias
        {
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            if (bt.Has(name)) return;
            bt.UpgradeOpen();

            var btr = new BlockTableRecord { Name = name, Origin = Point3d.Origin };
            bt.Add(btr);
            tr.AddNewlyCreatedDBObject(btr, true);

            // 2) nothing else changes – the variable ‘col’ is already an AcColor
            double half = 2.5;
            var l1 = new Line(new Point3d(-half, 0, 0), new Point3d(half, 0, 0)) { Color = col };
            var l2 = new Line(new Point3d(0, -half, 0), new Point3d(0, half, 0)) { Color = col };
            btr.AppendEntity(l1); tr.AddNewlyCreatedDBObject(l1, true);
            btr.AppendEntity(l2); tr.AddNewlyCreatedDBObject(l2, true);
        }

        // -----------------------------------------------------------------------------
        // PlaceHybridBlocks – keeps Hybrid_XC/RC/EC markers 100 % in-sync
        // -----------------------------------------------------------------------------
        // HybridCommands.cs  – replace the whole method
        // ***  C# 7.3-compatible  ***
        // HybridCommands.cs  – FULL replacement
        // Keeps XC / RC / EC markers in-sync but leaves every other Hybrid_* block alone
        private static void PlaceHybridBlocks(
            Transaction tr,
            Database db,
            IList<VertexInfo> verts)
        {
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

            /* ---------- 1) What we *want* to see on the polyline ------------------ */
            var want = new Dictionary<Point3d, string>(new Point3dEquality(kMatchTol));

            foreach (var v in verts)
            {
                string t = (v.Type ?? string.Empty).Trim().ToUpperInvariant();
                if (t == "XC" || t == "RC" || t == "EC")
                    want[v.Pt] = t;                   // last win inside tolerance bucket
            }

            /* ---------- 2)  Scan existing Hybrid_* blocks ------------------------- */
            var toErase = new List<BlockReference>();

            foreach (ObjectId id in space)
            {
                if (id.ObjectClass != RXObject.GetClass(typeof(BlockReference))) continue;
                var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);

                if (!br.Name.StartsWith("Hybrid_", StringComparison.OrdinalIgnoreCase)) continue;

                /* 2a) Ignore **detail-area** symbols that are NOT located on a vertex */
                if (!want.TryGetValue(br.Position, out string rightType))
                    continue;

                /* 2b) Vertex match – keep only if the subtype matches ----------------*/
                string existingType = br.Name.Substring(7).ToUpperInvariant(); // XC/RC/EC/OC…
                if (rightType != existingType)
                    toErase.Add(br);               // wrong subtype → replace later
            }

            foreach (var br in toErase)
            {
                br.UpgradeOpen();
                br.Erase();
            }

            /* ---------- 3)  Insert bubbles that are still missing ----------------- */
            foreach (var v in verts)
            {
                string t = (v.Type ?? string.Empty).Trim().ToUpperInvariant();
                if (!(t == "XC" || t == "RC" || t == "EC")) continue;

                bool present = false;
                foreach (ObjectId id in space)
                {
                    if (id.ObjectClass != RXObject.GetClass(typeof(BlockReference))) continue;
                    var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);

                    if (br.Position.DistanceTo(v.Pt) <= kMatchTol &&
                        br.Name.Equals("Hybrid_" + t, StringComparison.OrdinalIgnoreCase))
                    {
                        present = true;
                        break;
                    }
                }
                if (present) continue;

                var nb = new BlockReference(v.Pt, bt["Hybrid_" + t])
                {
                    Layer = kBlockLayer,
                    ScaleFactors = new Scale3d(5, 5, 1)
                };
                space.AppendEntity(nb);
                tr.AddNewlyCreatedDBObject(nb, true);

                // hidden ID stays in sync so bubbles survive moves
                UpdateHiddenIdIfNeeded(tr, nb, v.ID);
            }
        }
        // -----------------------------------------------------------------------------
        //  DeleteStrayNumberBubbles
        //  – Erases every Hybrd Num block whose hidden ID is *not* in keepIds.
        // -----------------------------------------------------------------------------
        internal static void DeleteStrayNumberBubbles(
            Transaction tr,
            BlockTableRecord space,
            HashSet<int> keepIds)
        {
            foreach (ObjectId id in space)
            {
                if (id.ObjectClass != RXObject.GetClass(typeof(BlockReference))) continue;
                var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                if (!br.Name.Equals("Hybrd Num", StringComparison.OrdinalIgnoreCase)) continue;

                /* fetch hidden ID (-1 if missing / unreadable) */
                int idVal = -1;
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                    if (ar.Tag == "ID" && int.TryParse(ar.TextString, out int v))
                    {
                        idVal = v;
                        break;
                    }
                }

                if (!keepIds.Contains(idVal))        // stray → delete
                {
                    br.UpgradeOpen();
                    br.Erase();
                }
            }
        }

        internal static string GuessType(Point3d pt, Transaction tr)
        {
            var db = AcadApp.DocumentManager.MdiActiveDocument.Database;
            var ms = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

            foreach (ObjectId id in ms)
            {
                if (id.ObjectClass != RXObject.GetClass(typeof(BlockReference))) continue;

                var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                if (br.Position.DistanceTo(pt) > kMatchTol) continue;   // ⬅︎ 0.03 m snap

                string nm = br.Name.ToUpperInvariant();
                if (nm.Contains("HYBRID_XC")) return "XC";
                if (nm.Contains("HYBRID_RC")) return "RC";
                if (nm.Contains("HYBRID_EC")) return "EC";
                if (nm.StartsWith("FDI") || nm.StartsWith("FDSPIKE")) return "OC";
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
        /// <summary>
        /// Overwrites the vertex coordinates of <paramref name="plId"/> with the
        /// values supplied in <paramref name="verts"/>.  
        /// • Expands or trims the vertex list as required.  
        /// • Unlocks / relocks the polyline’s layer if it happens to be locked.  
        /// • All changes are committed in a single transaction.
        /// </summary>
        public static void UpdatePolylineGeometry(ObjectId plId, IList<VertexInfo> verts)
        {
            if (plId == ObjectId.Null || !plId.IsValid || verts == null) return;

            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                // ---- open the polyline for write (fail fast if it was erased) -------
                if (!(tr.GetObject(plId, OpenMode.ForWrite, false) is Polyline pl))
                    return;

                // ---- make sure we are allowed to edit the layer ----------------------
                var lyr = (LayerTableRecord)tr.GetObject(pl.LayerId, OpenMode.ForWrite);
                bool relock = false;
                if (lyr.IsLocked)
                {
                    lyr.IsLocked = false;
                    relock = true;
                }

                // ---- update, add or remove vertices ---------------------------------
                int existing = pl.NumberOfVertices;
                int target = verts.Count;
                int common = Math.Min(existing, target);

                // edit the ones that already exist
                for (int i = 0; i < common; i++)
                    pl.SetPointAt(i, new Point2d(verts[i].E, verts[i].N));

                // add any new ones
                for (int i = existing; i < target; i++)
                    pl.AddVertexAt(i, new Point2d(verts[i].E, verts[i].N), 0, 0, 0);

                // remove any surplus ones (work backwards!)
                for (int i = existing - 1; i >= target; i--)
                    pl.RemoveVertexAt(i);

                // put the layer back the way we found it
                if (relock) lyr.IsLocked = true;

                tr.Commit();
            }
        }

        /// <summary>
        /// Makes sure we can safely create / edit “Hybrd Num” bubbles:
        ///   • guarantees L-MON exists and is temporarily unlocked  
        ///   • hops away from a *locked* current layer (L-MON or 0)  
        ///   • falls back to a private parking layer “_HS_TMP” if even layer 0 is locked  
        ///   • returns the key objects needed for bubble work
        /// </summary>
        internal static void EnsureNumberingContext(
            Transaction tr, Database db,
            out BlockTable bt, out BlockTableRecord space, out BlockTableRecord defRec)
        {
            /* ---------- 0) make sure L-MON itself exists ---------- */
            EnsureLayer(tr, db, kBlockLayer);

            /* ---------- 1) prepare layer handles we’ll need ---------- */
            var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

            ObjectId idLayer0 = lt["0"];
            var layer0Rec = (LayerTableRecord)tr.GetObject(idLayer0, OpenMode.ForWrite);

            /* create / fetch a hidden, always-unlocked parking layer */
            ObjectId idTmp;
            if (!lt.Has("_HS_TMP"))
            {
                lt.UpgradeOpen();
                var tmp = new LayerTableRecord
                {
                    Name = "_HS_TMP",
                    IsOff = false,
                    IsLocked = false,
                    Color = AcColor.FromColorIndex(ColorMethod.ByAci, 7)   // light-gray
                };
                idTmp = lt.Add(tmp);
                tr.AddNewlyCreatedDBObject(tmp, true);
            }
            else
            {
                idTmp = lt["_HS_TMP"];
            }

            /* ---------- 2) unlock L-MON (kBlockLayer) if required ---------- */
            ObjectId oldClayer = db.Clayer;
            var ltrMon = (LayerTableRecord)tr.GetObject(lt[kBlockLayer], OpenMode.ForWrite);
            bool relock = false;

            try
            {
                // If the *current* layer is the one we need to unlock, hop away first
                if (db.Clayer == ltrMon.ObjectId && ltrMon.IsLocked)
                {
                    db.Clayer = layer0Rec.IsLocked ? idTmp : idLayer0;
                }

                if (ltrMon.IsLocked)
                {
                    ltrMon.IsLocked = false;   // unlock just for the duration
                    relock = true;
                }

                /* ---------- 3) original body ---------- */
                EnsureHybridNumBlock(tr, db);

                bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                defRec = (BlockTableRecord)tr.GetObject(bt["Hybrd Num"], OpenMode.ForRead);
            }
            finally
            {
                // …restore user’s layer
                if (db.Clayer != oldClayer) db.Clayer = oldClayer;

                // …and relock L-MON if we unlocked it
                if (relock && !ltrMon.IsLocked)
                    ltrMon.IsLocked = true;
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
        // -----------------------------------------------------------------------------
        // ADD NUMBERING
        // -----------------------------------------------------------------------------
        [CommandMethod("AddNumbering")]
        public static void AddNumbering()
        {
            // prevent HybridGuard pop-ups while we manipulate attributes
            using (HybridGuard.Suspend())
            {
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                var db = doc.Database;
                var ed = doc.Editor;

                using (doc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    // make sure L-MON exists, is (temporarily) unlocked, etc.
                    EnsureNumberingContext(tr, db,
                        out var bt, out var space, out var defRec);

                    _bubbleCache.Remove(space.ObjectId);               // clear per-drawing cache
                    UpgradeExistingHybrdNumBubbles(tr, db, space, bt); // ← NEW: add IDs to legacy bubbles

                    /* ----- let the user pick the target polyline ----- */
                    var peo = new PromptEntityOptions(
                        "\nSelect polyline to place numbers on: ");
                    peo.SetRejectMessage("\nThat’s not a polyline.");
                    peo.AddAllowedClass(typeof(Polyline), false);

                    var per = ed.GetEntity(peo);
                    if (per.Status != PromptStatus.OK) return;

                    var plId = per.ObjectId;
                    var pl = (Polyline)tr.GetObject(plId, OpenMode.ForRead);

                    /* ----- build vertex list (re-using any existing IDs) ----- */
                    var verts = BuildVertexListWithIds(tr, db, pl);

                    /* ----- (re)create bubbles and NUMBER attributes ----- */
                    UpdateOrCreateBubbles(tr, db, space, defRec, verts);

                    /* ----- persist metadata on the polyline ----- */
                    WriteVertexData(plId, db, verts);

                    /* ----- update (or create) the table if it exists ----- */
                    ObjectId tblId = FindExistingTableId(
                        tr, db, EnsureTableStyle(tr, db), kTableLayer);

                    if (tblId != ObjectId.Null)
                    {
                        var tbl = (Table)tr.GetObject(tblId, OpenMode.ForWrite);
                        WriteTableMetadata(tbl, verts, tr);
                    }

                    tr.Commit();
                    ed.WriteMessage(
                        $"\nAddNumbering: {verts.Count} vertices processed.");
                }

                doc.Editor.Regen();
            }
        }

        // -----------------------------------------------------------------------------
        // UPDATE NUMBERING
        // -----------------------------------------------------------------------------
        [CommandMethod("UpdateNumbering")]
        public static void UpdateNumbering()
        {
            using (HybridGuard.Suspend())   // prevent guard pop‑ups
            {
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                var db = doc.Database;
                var ed = doc.Editor;

                var peo = new PromptEntityOptions("\nSelect polyline to renumber: ");
                peo.SetRejectMessage("\nThat’s not a polyline.");
                peo.AddAllowedClass(typeof(Polyline), false);
                var per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK) return;
                var plId = per.ObjectId;

                /* -------- build current vertex list (unchanged code) -------- */
                var oldData = ReadVertexData(plId, db);
                var oldMap = oldData
                    .GroupBy(v => v.Pt, new Point3dEquality(kMatchTol))
                    .Select(g => g.First())
                    .ToDictionary(v => v.Pt, v => v, new Point3dEquality(kMatchTol));

                var verts = new List<VertexInfo>();
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var pl = (Polyline)tr.GetObject(plId, OpenMode.ForRead);
                    for (int i = 0; i < pl.NumberOfVertices; i++)
                    {
                        var pt = pl.GetPoint3dAt(i);
                        verts.Add(oldMap.TryGetValue(pt, out var vi)
                                  ? vi
                                  : new VertexInfo { Pt = pt, N = pt.Y, E = pt.X, Type = "", Desc = "", ID = 0 });
                    }
                    tr.Commit();
                }

                using (doc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    EnsureNumberingContext(tr, db, out var bt, out var space, out var defRec);

                    _bubbleCache.Remove(space.ObjectId);      // ← clear per‑drawing cache

                    UpgradeExistingHybrdNumBubbles(tr, db, space, bt);

                    /* -------- remainder of original logic unchanged --------- */

                    var idToBr = new Dictionary<int, List<BlockReference>>();
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
                                if (!idToBr.TryGetValue(v, out var list)) idToBr[v] = list = new List<BlockReference>();
                                list.Add(br);
                                maxId = Math.Max(maxId, v);
                            }
                        }
                    }

                    int nextId = maxId + 1;
                    foreach (var v in verts) if (v.ID == 0) v.ID = nextId++;

                    int created = 0, reused = 0, renumbered = 0;

                    foreach (var v in verts)
                    {
                        if (!idToBr.ContainsKey(v.ID))
                        {
                            var brNew = GetOrCreateNumberBubble(tr, space, defRec, v.Pt, v.ID, kMatchTol);
                            idToBr[v.ID] = new List<BlockReference> { brNew };
                            created++;
                        }
                        else
                        {
                            reused += idToBr[v.ID].Count;
                        }
                    }

                    for (int i = 0; i < verts.Count; i++)
                    {
                        string seq = (i + 1).ToString();
                        foreach (var br in idToBr[verts[i].ID])
                        {
                            foreach (ObjectId attId in br.AttributeCollection)
                            {
                                var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForWrite);
                                if (ar.Tag == "NUMBER" && ar.TextString != seq)
                                {
                                    ar.TextString = seq;
                                    ar.AdjustAlignment(db);
                                    renumbered++;
                                }
                            }
                        }
                    }

                    WriteVertexData(plId, db, verts);
                    tr.Commit();
                    ed.WriteMessage($"\nUpdateNumbering: {created} new, {reused} reused, {renumbered} renumbered.");
                }

                doc.Editor.Regen();
            }
        }
        // -----------------------------------------------------------------------------
        //  DeleteBubblesAboveNumber
        //  – Erases every Hybrd Num block whose visible NUMBER attribute exceeds
        //    <maxNumber>.  ID is ignored; position is ignored.
        // -----------------------------------------------------------------------------
        internal static void DeleteBubblesAboveNumber(
            Transaction tr,
            BlockTableRecord space,
            int maxNumber,
            Database db)
        {
            foreach (ObjectId id in space)
            {
                if (id.ObjectClass != RXObject.GetClass(typeof(BlockReference))) continue;
                var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                if (!br.Name.Equals("Hybrd Num", StringComparison.OrdinalIgnoreCase)) continue;

                foreach (ObjectId attId in br.AttributeCollection)
                {
                    var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                    if (ar.Tag.Equals("NUMBER", StringComparison.OrdinalIgnoreCase) &&
                        int.TryParse(ar.TextString, out int seq) &&
                        seq > maxNumber)
                    {
                        br.UpgradeOpen();
                        br.Erase();
                        break;                      // one NUMBER per block – next block
                    }
                }
            }
        }



        // -----------------------------------------------------------------------------
        //  DUMPVERTEXDATA  – now writes a *.txt file as well
        // -----------------------------------------------------------------------------
        [CommandMethod("DUMPVERTEXDATA")]
        public static void DumpVertexData()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            /* 1) let the user pick the source polyline */
            var peo = new PromptEntityOptions("\nSelect polyline to dump data:");
            peo.SetRejectMessage("\nThat’s not a polyline.");
            peo.AddAllowedClass(typeof(Polyline), false);

            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            /* 2) read the stored HybridData list */
            var list = ReadVertexData(per.ObjectId, db);
            if (list.Count == 0)
            {
                ed.WriteMessage("\nNo HybridData found on that polyline.");
                return;
            }

            /* 3) pretty-print JSON */
            string json = JsonConvert.SerializeObject(list, Formatting.Indented);

            /* 4) decide where to save it – same folder as DWG if possible */
            string dwgPath = db.Filename;
            string folder = string.IsNullOrEmpty(dwgPath)
                             ? Path.GetTempPath()
                             : Path.GetDirectoryName(dwgPath);

            string baseName = string.IsNullOrEmpty(dwgPath)
                              ? "UnsavedDrawing"
                              : Path.GetFileNameWithoutExtension(dwgPath);

            string fileName = $"{baseName}_HybridData_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
            string fullPath = Path.Combine(folder, fileName);

            /* 5) write the file, catching any IO errors */
            try
            {
                File.WriteAllText(fullPath, json, Encoding.UTF8);
                ed.WriteMessage($"\nHybridData written to: {fullPath}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n⚠  Could not write file – {ex.Message}");
            }

            /* 6) still show a short preview in the command window */
            ed.WriteMessage("\n--------- HybridData JSON (first 20 lines) ---------\n");
            using (var sr = new StringReader(json))
            {
                for (int i = 0; i < 20; i++)
                {
                    string line = sr.ReadLine();
                    if (line == null) break;
                    ed.WriteMessage(line + "\n");
                }
                if (sr.Peek() != -1)
                    ed.WriteMessage("... (truncated – see file for full content)\n");
            }
            ed.WriteMessage("----------------------------------------------------\n");
        }
        /// <summary>
        /// Removes the “HybridData” extension dictionary (and all stored IDs, types, descs)
        /// from every polyline in model‐space.  After running this, your polylines will
        /// look “un‐tagged” and you can re-run AddNumbering/InsertUpdate to rebuild them.
        /// </summary>
        /// <summary>
        /// Removes the “HybridData” Xrecord from every polyline in model space,
        /// unlocking / relocking any layers it needs to touch so that the purge
        /// always succeeds and never crashes with eOnLockedLayer.
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
                var btrMs = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

                int purged = 0;

                foreach (ObjectId entId in btrMs)
                {
                    // --- only consider polylines -------------------------------------------------
                    if (!(tr.GetObject(entId, OpenMode.ForRead) is Polyline pl)) continue;

                    // layer info ­- we might have to unlock it temporarily
                    var lyrRec = (LayerTableRecord)tr.GetObject(pl.LayerId, OpenMode.ForWrite);
                    bool relock = false;

                    if (lyrRec.IsLocked)
                    {
                        lyrRec.IsLocked = false;   // ← safe: the layer may be current
                        relock = true;
                    }

                    // switch the polyline to write-mode *after* we’re sure the layer is unlocked
                    pl.UpgradeOpen();

                    if (!pl.ExtensionDictionary.IsNull)
                    {
                        var ext = (DBDictionary)tr.GetObject(pl.ExtensionDictionary, OpenMode.ForWrite);
                        if (ext.Contains("HybridData"))
                        {
                            // erase the Xrecord + remove the entry from the dict
                            var xrecId = ext.GetAt("HybridData");
                            ext.Remove("HybridData");

                            if (tr.GetObject(xrecId, OpenMode.ForWrite) is Xrecord xr)
                                xr.Erase();

                            purged++;
                        }
                    }

                    // put the layer back the way we found it
                    if (relock) lyrRec.IsLocked = true;
                }

                tr.Commit();
                ed.WriteMessage($"\nPurged HybridData from {purged} polyline{(purged == 1 ? "" : "s")}.");
            }

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
                    ar.TextString = def.Tag == "ID" ? id.ToString() : string.Empty;

                    br.AttributeCollection.AppendAttribute(ar);
                    tr.AddNewlyCreatedDBObject(ar, true);
                }
            }

            return br;    // caller commits
        }

        private static readonly Dictionary<ObjectId, Dictionary<Point3d, BlockReference>> _bubbleCache
            = new Dictionary<ObjectId, Dictionary<Point3d, BlockReference>>();

        /// <summary>
        /// Scans <paramref name="space" /> once and returns a map
        ///   key   = bubble insertion point (tolerance aware)
        ///   value = BlockReference for the bubble
        /// The scan is very fast (&lt;1 ms for several thousand entities).
        /// </summary>
        private static Dictionary<Point3d, BlockReference> BuildBubbleMap(
            Transaction tr,
            BlockTableRecord space)
        {
            var map = new Dictionary<Point3d, BlockReference>(
                new Point3dEquality(kMatchTol));

            foreach (ObjectId id in space)
            {
                if (id.ObjectClass != RXObject.GetClass(typeof(BlockReference)))
                    continue;
                var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                if (!br.Name.Equals("Hybrd Num", StringComparison.OrdinalIgnoreCase))
                    continue;

                map[br.Position] = br; // duplicates collapse via comparer
            }
            return map;
        }

        // -----------------------------------------------------------------------------
        // GetOrCreateNumberBubble  – uses cache; validates type to avoid stale reuse
        // -----------------------------------------------------------------------------
        private static BlockReference GetOrCreateNumberBubble(
            Transaction tr,
            BlockTableRecord space,
            BlockTableRecord defRec,
            Point3d pt,
            int id,
            double tol)
        {
            /* ---------- 1) get (or build) cache for this model space ---------- */
            if (!_bubbleCache.TryGetValue(space.ObjectId, out var map))
            {
                map = BuildBubbleMap(tr, space);          // one-time fast scan
                _bubbleCache[space.ObjectId] = map;
            }

            /* ---------- 2) existing live bubble at this point? ---------------- */
            if (map.TryGetValue(pt, out var cached))
            {
                bool alive = !cached.IsErased &&
                              cached.ObjectId.IsValid &&
                              cached.Name.Equals("Hybrd Num", StringComparison.OrdinalIgnoreCase);  // ← NEW check

                if (alive)
                {
                    UpdateHiddenIdIfNeeded(tr, cached, id);   // keep ID in sync
                    return cached;                            // ← reuse it
                }
                map.Remove(pt);                               // stale entry – purge from cache
            }

            /* ---------- 3) nothing usable – make a brand-new bubble ----------- */
            var nb = CreateBubble(
                tr,
                space,
                (BlockTable)tr.GetObject(space.Database.BlockTableId, OpenMode.ForRead),
                defRec,
                pt,
                id);

            map[pt] = nb;                                    // add to cache
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
            int maxId = oldList.Any() ? oldList.Max(v => v.ID) : 0;
            int nextId = maxId + 1;
            foreach (var v in verts)
                if (v.ID == 0)
                    v.ID = nextId++;

            return verts;
        }
        // HybridCommands.cs  – FULL replacement
        // Re-numbers existing bubbles, creates missing ones **and** purges any
        // bubble that is not sitting on the correct vertex (even if it has a valid ID).
        // HybridCommands.cs  – FULL replacement
        // Keeps bubbles and vertices in perfect 1-to-1 sync **by ID only**.
        // Position is ignored; the bubble can live anywhere in the drawing.
        // ──────────────────────────────────────────────────────────────────────────────
        //  UpdateOrCreateBubbles  – FINAL 2025-06-XX (ID-only logic, robust parsing)
        // ──────────────────────────────────────────────────────────────────────────────
        internal static void UpdateOrCreateBubbles(
                    Transaction tr,
            Database db,
            BlockTableRecord space,
            BlockTableRecord defRec,
            List<VertexInfo> verts)
        {
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

            // ---------------------------------------------------------------------[1]
            //  Scan every Hybrd Num block in model-space ONCE
            //      goodIds :  ID → List<BlockReference>
            //      junk    :  blocks that have no readable integer ID  → delete later
            // -------------------------------------------------------------------------
            var goodIds = new Dictionary<int, List<BlockReference>>();
            var junk = new List<BlockReference>();

            foreach (ObjectId id in space)
            {
                if (id.ObjectClass != RXObject.GetClass(typeof(BlockReference))) continue;
                var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                if (!br.Name.Equals("Hybrd Num", StringComparison.OrdinalIgnoreCase)) continue;

                bool ok = false;
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                    if (ar.Tag.Equals("ID", StringComparison.OrdinalIgnoreCase) &&
                        int.TryParse(ar.TextString?.Trim(), out int v))
                    {
                        if (!goodIds.TryGetValue(v, out var list))
                            goodIds[v] = list = new List<BlockReference>();
                        list.Add(br);
                        ok = true;
                        break;
                    }
                }
                if (!ok) junk.Add(br);        // no parseable ID → mark for deletion
            }

            // ---------------------------------------------------------------------[2]
            //  Guarantee *exactly one* bubble per vertex-ID, reuse if it exists
            // -------------------------------------------------------------------------
            foreach (var v in verts)
            {
                if (!goodIds.TryGetValue(v.ID, out var list))            // none at all
                {
                    var brNew = CreateBubble(tr, space, bt, defRec, v.Pt, v.ID);
                    goodIds[v.ID] = list = new List<BlockReference> { brNew };
                }
                else if (list.Count > 1)                                 // duplicates
                {
                    for (int i = 1; i < list.Count; i++)
                    {
                        list[i].UpgradeOpen();
                        list[i].Erase();
                    }
                    list.RemoveRange(1, list.Count - 1);
                }

                /* ----- update visible NUMBER text to current sequence ----- */
                string seq = (verts.IndexOf(v) + 1).ToString();
                var brKeep = list[0];
                foreach (ObjectId attId in brKeep.AttributeCollection)
                {
                    var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForWrite);
                    if (ar.Tag.Equals("NUMBER", StringComparison.OrdinalIgnoreCase) &&
                        ar.TextString != seq)
                    {
                        ar.TextString = seq;
                        ar.AdjustAlignment(db);
                    }
                }
            }

            // ---------------------------------------------------------------------[3]
            //  Purge everything that shouldn’t be here
            //      • bubbles whose ID is NOT in the vertex list
            //      • bubbles that had no readable ID (junk list)
            // -------------------------------------------------------------------------
            var validIds = new HashSet<int>(verts.Select(v => v.ID));

            foreach (var kvp in goodIds)              // kvp.Key == ID
            {
                if (validIds.Contains(kvp.Key)) continue;   // still used → keep
                foreach (var br in kvp.Value)
                {
                    br.UpgradeOpen();
                    br.Erase();
                }
            }

            foreach (var br in junk)                  // orphaned / corrupt ID
            {
                br.UpgradeOpen();
                br.Erase();
            }
        }

        /* helper – safely read the hidden ID attribute, returns false if absent / bad */
        private static bool TryGetId(Transaction tr, BlockReference br, out int idVal)
        {
            foreach (ObjectId attId in br.AttributeCollection)
            {
                var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                if (ar.Tag == "ID" && int.TryParse(ar.TextString?.Trim(), out idVal))
                    return true;
            }
            idVal = -1;
            return false;
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
        // ==========  fields  ==========
        private readonly DataGridView _grid;
        private readonly CheckBox _chkHybrid;
        private List<VertexInfo> _verts;
        private ObjectId _currentPlId = ObjectId.Null;
        private string _homeDwgPath;        // now writable!
        private bool _locked;               // UI disabled when drawing changes
        internal static VertexForm ActiveInstance { get; private set; }
        // ==========  stable-anchor helper  ==========
        private static string AnchorFor(Document doc) =>
            string.IsNullOrEmpty(doc.Database.Filename) ? doc.Name
                                                        : doc.Database.Filename;
        public VertexForm()
        {
            ActiveInstance = this;          
            _homeDwgPath = AnchorFor(AcadApp.DocumentManager.MdiActiveDocument);
            Text = $"Hybrid Vertex Editor  —  {_homeDwgPath}";
            Text += $"  —  {_homeDwgPath}";
            ClientSize = new System.Drawing.Size(1000, 580);
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = true;

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
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
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

            /* ---------- command buttons ---------- */
            var btnGet = new Button { Text = "Get Polyline", Width = 120 };
            var btnUpd = new Button { Text = "Update Polyline", Width = 120 };
            var btnAddNum = new Button { Text = "Add Numbering", Width = 120 };
            var btnUpdNum = new Button { Text = "Update Numbering", Width = 120 };
            var btnRebuild = new Button { Text = "Rebuild Polyline", Width = 140 };
            var btnTransfer = new Button { Text = "Transfer Data", Width = 120 };
            var btnOK = new Button { Text = "Insert/Update", Width = 100 };
            _chkHybrid = new CheckBox { Text = "Hybrid Blocks", Checked = true };
            var btnRemove = new Button { Text = "Remove Point", Width = 120 };

            /*   ➜  add all the “regular” buttons first   */
            pnl.Controls.AddRange(new Control[] {
        btnGet, btnUpd, btnAddNum, btnUpdNum,
        btnRebuild, btnTransfer, _chkHybrid, btnOK
    });

            /*   ➜  finally append  “Remove Point” so it floats to the right   */
            pnl.Controls.Add(btnRemove);

            /* ---------- event wiring ---------- */
            btnRemove.Click += RemoveSelected_Click;
            btnGet.Click += (s, e) => PickAndPopulate();
            btnUpd.Click += (s, e) => PickAndPopulate();
            btnAddNum.Click += (s, e) => HybridCommands.AddNumbering();
            btnUpdNum.Click += (s, e) => HybridCommands.UpdateNumbering();
            btnRebuild.Click += (s, e) => HybridCommands.RebuildPolylineFromTable();
            btnTransfer.Click += (s, e) => HybridCommands.TransferVertexData();
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
            _grid.KeyDown += Grid_KeyDown;   // copy/paste shortcuts

            _verts = new List<VertexInfo>();
            ToggleUiLock(false);             // we're still in the “home” drawing
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            // already there – keeps UI in sync when user switches drawings
            AcadApp.DocumentManager.DocumentActivated += OnDocActivated;

            // NEW – closes the form if its parent drawing is about to vanish
            AcadApp.DocumentManager.DocumentToBeDestroyed += OnDocToBeDestroyed;
        }
        /// <summary>
        /// Auto-closes the editor if its “home” drawing is being closed,
        /// preventing a dangling WinForms handle that can crash AutoCAD.
        /// </summary>
        private void OnDocToBeDestroyed(object sender, DocumentCollectionEventArgs e)
        {
            string doomed = AnchorFor(e.Document);               // helper already exists
            if (string.Equals(doomed, _homeDwgPath, StringComparison.OrdinalIgnoreCase))
            {
                // marshal back to UI thread – the event fires on AutoCAD’s thread
                BeginInvoke((MethodInvoker)(() => Close()));
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            ActiveInstance = null;

            AcadApp.DocumentManager.DocumentActivated -= OnDocActivated;
            AcadApp.DocumentManager.DocumentToBeDestroyed -= OnDocToBeDestroyed;   // ← NEW

            base.OnFormClosed(e);
        }

        private void OnDocActivated(object sender, DocumentCollectionEventArgs e)
        {
            string currentAnchor = AnchorFor(e.Document);

            // If this form was opened for “Drawing1” and that doc was just saved,
            // adopt its real pathname so the UI stays unlocked.
            if (!_locked &&
                string.Equals(_homeDwgPath, e.Document.Name, StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrEmpty(e.Document.Database.Filename))
            {
                _homeDwgPath = currentAnchor;
            }

            bool shouldLock = !string.Equals(currentAnchor, _homeDwgPath,
                                             StringComparison.OrdinalIgnoreCase);

            if (shouldLock != _locked)
            {
                _locked = shouldLock;
                ToggleUiLock(_locked);
            }
        }
        private void ToggleUiLock(bool state)
        {
            string msg = state
                ? "⚠  Form locked — active drawing does not match the one this form was opened for."
                : "Hybrid Vertex Editor";

            // Disable/enable the whole grid and the bottom panel
            _grid.Enabled = !state;
            foreach (Control ctl in Controls)
                if (ctl is FlowLayoutPanel) ctl.Enabled = !state;

            // Update window caption
            Text = msg + $"  —  {_homeDwgPath}";
        }
        // add inside VertexForm

        /// <summary>
        /// Re-sorts all existing "Hybrd Num" blocks along the picked polyline
        /// and reassigns only the NUMBER attribute (ID stays fixed).
        /// </summary>
        // VertexForm.cs  – replace the whole PickAndPopulate method
        private void PickAndPopulate()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            /* ────────────────── let the user pick a polyline ────────────────── */
            var opts = new PromptEntityOptions("\nSelect polyline");
            opts.SetRejectMessage("\nMust be a polyline");
            opts.AddAllowedClass(typeof(Polyline), false);
            var res = ed.GetEntity(opts);
            if (res.Status != PromptStatus.OK) return;

            _currentPlId = res.ObjectId;

            /* ─── [MISSING JSON CHECK] ────────────────────────────────────────── */
            var jsonList = HybridCommands.ReadVertexData(_currentPlId, db);
            bool payloadMissing;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var pl = (Polyline)tr.GetObject(_currentPlId, OpenMode.ForRead);
                payloadMissing = jsonList.Count == 0 && pl.NumberOfVertices > 0;
                tr.Commit();
            }

            if (payloadMissing)
            {
                var answer = MessageBox.Show(
                    "This polyline has vertices but no Hybrid-Data metadata.\n\n" +
                    "Re-build the table and bubbles from scratch?",
                    "Hybrid Survey",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (answer == DialogResult.Yes)
                {
                    // blank list → InsertOrUpdate will create a fresh table + bubbles
                    _verts = new List<VertexInfo>();      // nothing to pre-populate
                    HybridCommands.InsertOrUpdate(_verts, insertHybrid: true);
                    ed.WriteMessage("\nRebuilt table & bubbles with empty Type/Desc.");
                    return;   // nothing more to show in the grid
                }
                else
                {
                    return;   // user declined; abort safely
                }
            }
            /* ─────────────────────────────────────────────────────────────────── */

            // --- existing logic -------------------------------------------------
            // skip duplicate Pt keys
            var oldMap = new Dictionary<Point3d, VertexInfo>(
                new Point3dEquality(HybridCommands.kMatchTol));
            foreach (var v in jsonList)
                if (!oldMap.ContainsKey(v.Pt))
                    oldMap.Add(v.Pt, v);

            // build the new list in vertex order
            var newList = new List<VertexInfo>();
            using (var tr = db.TransactionManager.StartTransaction())
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

        // VertexForm.cs  – FULL replacement of InsertUpdate_Click
        private void InsertUpdate_Click(object sender, EventArgs e)
        {
            /* 1) make sure the cached polyline still exists ------------------------ */
            if (_currentPlId == ObjectId.Null || !_currentPlId.IsValid || _currentPlId.IsErased)
            {
                MessageBox.Show("Pick a polyline with “Get Polyline” first.",
                                "Hybrid Vertex Editor", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            /* 2) keep _verts length in sync with grid ------------------------------ */
            while (_verts.Count < _grid.Rows.Count) _verts.Add(new VertexInfo { ID = 0 });
            while (_verts.Count > _grid.Rows.Count) _verts.RemoveAt(_verts.Count - 1);

            /* 3) copy grid → _verts (store to 3 dp) -------------------------------- */
            for (int i = 0; i < _grid.Rows.Count; i++)
            {
                var r = _grid.Rows[i];
                double north = Math.Round(double.Parse(r.Cells["Northing"].Value?.ToString() ?? "0"), 3);
                double east = Math.Round(double.Parse(r.Cells["Easting"].Value?.ToString() ?? "0"), 3);

                _verts[i] = new VertexInfo
                {
                    Pt = new Point3d(east, north, 0),
                    N = north,
                    E = east,
                    Type = r.Cells["Type"].Value?.ToString() ?? "",
                    Desc = r.Cells["Desc"].Value?.ToString() ?? "",
                    ID = _verts[i].ID          // preserve whatever ID we already had
                };
            }

            /* 3 b) guarantee every vertex has a unique non-zero ID ----------------- */
            int nextId = Math.Max(1, _verts.Max(v => v.ID) + 1);
            foreach (var v in _verts) if (v.ID == 0) v.ID = nextId++;

            /* 4) push geometry + metadata + table ---------------------------------- */
            HybridCommands.UpdatePolylineGeometry(_currentPlId, _verts);
            HybridCommands.InsertOrUpdate(_verts, _chkHybrid.Checked);
            HybridCommands.WriteVertexData(
                _currentPlId,
                AcadApp.DocumentManager.MdiActiveDocument.Database,
                _verts);

            /* 5) **PURGE NUMBER > row-count bubbles** ------------------------------ */
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                HybridCommands.EnsureNumberingContext(
                    tr, db, out _, out var space, out _);
                HybridCommands.DeleteBubblesAboveNumber(tr, space, _verts.Count, db);
                tr.Commit();
            }

            RefreshGrid();
        }
        // ---------------------------------------------------------------------
        //  Excel‑style Ctrl‑C / Ctrl‑V for the DataGridView
        // ---------------------------------------------------------------------

        /// <summary>
        /// Removes every row that is fully-selected in the grid **and**
        /// deletes the matching vertex, Hybrid-Data payload, and Hybrd Num
        /// bubble.  After the purge, the polyline geometry, metadata,
        /// numbering bubbles, and table rows are all brought back into
        /// perfect sync and the visible NUMBER sequence is gap-free.
        /// </summary>
        // VertexForm.cs – FULL replacement
        // VertexForm.cs – FULL replacement of RemoveSelected_Click
        private void RemoveSelected_Click(object sender, EventArgs e)
        {
            if (_grid.SelectedRows.Count == 0) return;
            if (_currentPlId == ObjectId.Null || !_currentPlId.IsValid || _currentPlId.IsErased) return;

            using (HybridGuard.Suspend())           // silence guard
            {
                /* 1) delete rows + remember orphaned IDs --------------------------- */
                var goneIds = new HashSet<int>();
                foreach (DataGridViewRow r in _grid.SelectedRows)
                {
                    int idx = r.Index;
                    if (idx < _verts.Count) goneIds.Add(_verts[idx].ID);
                }
                foreach (DataGridViewRow r in _grid.SelectedRows.Cast<DataGridViewRow>()
                                                                .OrderByDescending(r => r.Index))
                {
                    _verts.RemoveAt(r.Index);
                    _grid.Rows.Remove(r);
                }

                /* 2) update polyline + metadata ----------------------------------- */
                HybridCommands.UpdatePolylineGeometry(_currentPlId, _verts);
                HybridCommands.WriteVertexData(
                    _currentPlId,
                    AcadApp.DocumentManager.MdiActiveDocument.Database,
                    _verts);

                /* 3) bubble maintenance (ID purge + NUMBER>count purge) ------------ */
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                var db = doc.Database;

                using (doc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    HybridCommands.EnsureNumberingContext(
                        tr, db, out _, out var space, out var defRec);

                    // 3 a) delete bubbles whose hidden ID is now gone
                    HybridCommands.DeleteStrayNumberBubbles(tr, space,
                                            new HashSet<int>(_verts.Select(v => v.ID)));

                    // 3 b) delete bubbles whose NUMBER > row-count
                    HybridCommands.DeleteBubblesAboveNumber(tr, space, _verts.Count, db);

                    // 3 c) re-sync NUMBER text / create missing bubbles
                    HybridCommands.UpdateOrCreateBubbles(tr, db, space, defRec, _verts);

                    tr.Commit();
                }

                RefreshGrid();
            }
        }





        private void Grid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                CopySelectionToClipboard();
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.V)
            {
                PasteClipboardToGrid();
                e.Handled = true;
            }
        }

        private void CopySelectionToClipboard()
        {
            if (_grid.GetCellCount(DataGridViewElementStates.Selected) == 0) return;
            var dataObj = _grid.GetClipboardContent();
            if (dataObj != null) Clipboard.SetDataObject(dataObj);
        } 
        // VertexForm.cs  – replace the whole method
        // ***  C# 7.3-compatible  ***
        private void Grid_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            string colName = _grid.Columns[e.ColumnIndex].Name;
            string value = e.FormattedValue == null ? string.Empty : e.FormattedValue.ToString();

            /* ───── Northing / Easting ───── */
            if (colName == "Northing" || colName == "Easting")
            {
                // 1) Block comma decimal separators up-front
                if (value.IndexOf(',') >= 0)
                {
                    e.Cancel = true;
                    MessageBox.Show(
                        "Use a dot (.) for decimals, not a comma.",
                        "Invalid coordinate",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 2) Must parse with invariant culture so “.” is the only decimal mark
                double dummy;
                if (!double.TryParse(
                        value,
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out dummy))
                {
                    e.Cancel = true;
                    MessageBox.Show(
                        "Please enter a valid numeric coordinate.",
                        "Invalid coordinate",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return;                 // nothing else to check for these columns
            }

            /* ───── Type column – keep existing behaviour ───── */
            if (colName == "Type")
            {
                var col = (DataGridViewComboBoxColumn)_grid.Columns["Type"];
                if (!col.Items.Contains(value))
                    col.Items.Add(value);

                _grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = value;
            }
        }

        // VertexForm.cs  – REPLACE the entire PasteClipboardToGrid method
        // (no other changes required)

        private void PasteClipboardToGrid()
        {
            string text = Clipboard.GetText();
            if (string.IsNullOrEmpty(text)) return;

            /* ───── find insertion start (first selected, else [row0,col1]) ───── */
            var startCell = _grid.SelectedCells.Count > 0
                ? _grid.SelectedCells.Cast<DataGridViewCell>()
                       .OrderBy(c => c.RowIndex).ThenBy(c => c.ColumnIndex).First()
                : _grid[1, 0];               // skip read-only “#” column

            int startRow = startCell.RowIndex;
            int startCol = startCell.ColumnIndex;

            var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

            /* helper to grow the backing list + grid rows */
            void EnsureRow(int r)
            {
                while (r >= _grid.Rows.Count)
                {
                    _verts.Add(new VertexInfo
                    {
                        Pt = Point3d.Origin,
                        N = 0,
                        E = 0,
                        Type = "",
                        Desc = "",
                        ID = 0
                    });
                    _grid.Rows.Add(_grid.Rows.Count + 1, "", "", "", "");
                }
            }

            /* ───────── paste loop with comma-decimal guard ───────── */
            for (int i = 0; i < lines.Length; i++)
            {
                if (string.IsNullOrEmpty(lines[i])) continue;
                string[] cells = lines[i].Split('\t');
                EnsureRow(startRow + i);

                for (int j = 0; j < cells.Length; j++)
                {
                    int col = startCol + j;
                    if (col >= _grid.ColumnCount) break;      // ignore overflow
                    if (_grid.Columns[col].ReadOnly) continue; // e.g. “#”

                    string val = cells[j];

                    // Reject comma decimals for numeric columns
                    string colName = _grid.Columns[col].Name;
                    bool numericCol = (colName == "Northing" || colName == "Easting");
                    if (numericCol && val.IndexOf(',') >= 0)
                    {
                        MessageBox.Show(
                            "Use a dot (.) for decimals, not a comma.",
                            "Invalid coordinate",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        continue;   // skip this offending cell
                    }

                    _grid[col, startRow + i].Value = val;
                }
            }
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


        private void Combo_Validating(object sender, CancelEventArgs e)
        {
            if (_grid.CurrentCell.OwningColumn.Name == "Type"
                && _grid.EditingControl is ComboBox cb)
            {
                _grid.CurrentCell.Value = cb.Text;
            }
        }
    }
}

