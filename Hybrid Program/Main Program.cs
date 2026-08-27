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
using System.Globalization;
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
            // Distance-based equality can cross any fixed bucket boundary.  A
            // constant hash is the only fully correct hash for this comparer;
            // Hybrid vertex collections are small, so the linear scan is cheap.
            return 0;
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
            // Subscribe to DocumentToBeDestroyed so we can remove hooks and caches
            AcadApp.DocumentManager.DocumentToBeDestroyed += OnDocToBeDestroyed;
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
            AcadApp.DocumentManager.DocumentToBeDestroyed -= OnDocToBeDestroyed;

            _wired.Clear();
            _orig.Clear();
        }

        /// <summary>
        /// Cleans up resources and caches when a document is about to be destroyed.
        /// This handler unhooks database event handlers and command hooks and clears
        /// any cached data associated with the drawing.  Without this cleanup,
        /// stale ObjectIds can linger in static collections and cause instability.
        /// </summary>
        private static void OnDocToBeDestroyed(object sender, DocumentCollectionEventArgs e)
        {
            var doc = e.Document;
            if (doc == null) return;

            // Unhook database events and command hooks for this document
            unhook(doc.Database);
            UnwireCommandHooks(doc);

            // Clear per‑document caches in HybridCommands
            HybridCommands.ClearCaches();
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
                ar.Tag.Equals("NUMBER", StringComparison.OrdinalIgnoreCase) &&
                IsHybridNumberAttribute(ar))
            {
                _orig[ar.ObjectId] = ar.TextString;
            }
        }

        private static bool IsHybridNumberAttribute(AttributeReference ar)
        {
            if (ar == null || ar.Database == null || ar.OwnerId.IsNull || !ar.OwnerId.IsValid)
                return false;

            try
            {
                // ObjectOpenedForModify is raised inside AutoCAD's owning
                // transaction.  Resolve the owner once here; ObjectModified must
                // not re-enter the transaction manager while a scale/property
                // edit is being applied.
                var tr = ar.Database.TransactionManager.TopTransaction;
                if (tr == null) return false;

                var br = tr.GetObject(ar.OwnerId, OpenMode.ForRead, false) as BlockReference;
                if (br == null) return false;

                string blockName = br.Name;
                if (br.IsDynamicBlock)
                {
                    var definition = tr.GetObject(
                        br.DynamicBlockTableRecord,
                        OpenMode.ForRead,
                        false) as BlockTableRecord;
                    blockName = definition?.Name ?? blockName;
                }

                return string.Equals(
                    (blockName ?? string.Empty).Trim(),
                    "Hybrd Num",
                    StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(
                        (blockName ?? string.Empty).Trim(),
                        "Hybrid Num",
                        StringComparison.OrdinalIgnoreCase);
            }
            catch (System.Exception)
            {
                return false;
            }
        }

        // Reverts unauthorised edits to NUMBER attributes & rate-limits the warning
        private static void onModified(object sender, ObjectEventArgs e)
        {
            var db = e.DBObject?.Database;
            if (db == null || IsDbLoading(db)) return;   // drawing still streaming in
            if (Suspended)
            {
                if (e.DBObject != null) _orig.Remove(e.DBObject.ObjectId);
                return;                                  // guard temporarily paused
            }

            var ar = e.DBObject as AttributeReference;
            if (ar == null || !ar.Tag.Equals("NUMBER", StringComparison.OrdinalIgnoreCase))
                return;

            // The owner was verified in onOpened.  Avoid opening any additional
            // DBObjects from ObjectModified: block X/Y/Z scale edits can modify
            // attribute geometry and this re-entrant open was the crash path.
            if (!_orig.TryGetValue(ar.ObjectId, out string oldVal)) return;
            _orig.Remove(ar.ObjectId);

            bool textChanged = !string.Equals(oldVal, ar.TextString, StringComparison.Ordinal);
            if (textChanged)
            {
                if (!ar.IsWriteEnabled) ar.UpgradeOpen();
                ar.TextString = oldVal;
                ar.AdjustAlignment(ar.Database);
            }

            /* gentle reminder only for a real text edit, never for scale changes */
            if (textChanged &&
                (DateTime.Now - _lastPopup).TotalSeconds >= POPUP_COOLDOWN_SEC &&
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

    internal static partial class HybridCommands
    {
        private const string kTableStyle = "Induction Bend";
        private const string kTableLayer = "Hybrid_Points_TBL";
        private const string kBlockLayer = "L-MON";
        private const double kTxtH = 2.5;
        private const double kTableTxtH = 10.0;
        private static readonly double[] kColW = { 100, 150, 150, 100, 450 };
        internal const double kMatchTol = 0.03;       // vertex search tolerance
        private const double kRowH = 25.0;
        internal const int kStorageDecimals = HybridValuePolicy.StorageDecimals;
        internal const int kDisplayDecimals = HybridValuePolicy.DisplayDecimals;

        internal static double StoreCoordinate(double value)
        {
            return HybridValuePolicy.StoreCoordinate(value);
        }

        internal static Point3d StorePoint(Point3d point)
        {
            return new Point3d(
                StoreCoordinate(point.X),
                StoreCoordinate(point.Y),
                StoreCoordinate(point.Z));
        }

        internal static VertexInfo VertexAt(Point3d point, VertexInfo metadata = null)
        {
            var stored = StorePoint(point);
            return new VertexInfo
            {
                Pt = stored,
                N = stored.Y,
                E = stored.X,
                Type = metadata?.Type ?? string.Empty,
                Desc = metadata?.Desc ?? string.Empty,
                ID = metadata?.ID ?? 0
            };
        }

        internal static string DisplayCoordinate(double value)
        {
            return HybridValuePolicy.DisplayCoordinate(value);
        }

        internal static bool TryParseCoordinate(object value, out double coordinate)
        {
            return HybridValuePolicy.TryParseCoordinate(value, out coordinate);
        }

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
            if (!TryRequireUtmMode("Write Hybrid Vertex Data")) return;

            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null || db == null || doc.Database != db) return;

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                // ── SAFE OPEN ──
                Polyline pl;
                try
                {
                    // Open read-only before temporarily unlocking the layer.
                    // Opening an entity for write while its layer is locked can
                    // fail before the code gets a chance to repair that state.
                    pl = (Polyline)tr.GetObject(plId, OpenMode.ForRead, false);
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex) when (IsWasErased(ex))
                {
                    // the polyline was deleted/undone – silently bail
                    return;
                }

                if (pl == null) return;          // sanity (shouldn’t happen)
                using (UnlockLayer(tr, pl.LayerId))
                {
                    pl.UpgradeOpen();
                    WriteVertexData(tr, pl, verts);
                }

                tr.Commit();     // commit once, here
            }
        }

        private static void WriteVertexData(
            Transaction tr,
            Polyline pl,
            IList<VertexInfo> verts)
        {
            if (tr == null || pl == null) return;

            if (!pl.IsWriteEnabled) pl.UpgradeOpen();
            if (pl.ExtensionDictionary.IsNull) pl.CreateExtensionDictionary();

            var ext = (DBDictionary)tr.GetObject(pl.ExtensionDictionary, OpenMode.ForWrite);
            if (!ext.Contains("HybridData"))
            {
                var dict = new DBDictionary();
                ext.SetAt("HybridData", dict);
                tr.AddNewlyCreatedDBObject(dict, true);
            }

            var dataDict = (DBDictionary)tr.GetObject(ext.GetAt("HybridData"), OpenMode.ForWrite);
            Xrecord xr;
            if (dataDict.Contains("Data"))
                xr = (Xrecord)tr.GetObject(dataDict.GetAt("Data"), OpenMode.ForWrite);
            else
            {
                xr = new Xrecord();
                dataDict.SetAt("Data", xr);
                tr.AddNewlyCreatedDBObject(xr, true);
            }

            var jsonPayload = JsonConvert.SerializeObject(
                (verts ?? new List<VertexInfo>()).Select(v => new SimpleVertex
                {
                    X = StoreCoordinate(v.E),
                    Y = StoreCoordinate(v.N),
                    Type = v.Type ?? string.Empty,
                    Desc = v.Desc ?? string.Empty,
                    ID = v.ID
                }).ToList());

            xr.Data = new ResultBuffer(new TypedValue((int)DxfCode.Text, jsonPayload));
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
                X = StoreCoordinate(v.E),
                Y = StoreCoordinate(v.N),
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
                if (ar.Tag.Equals("ID", StringComparison.OrdinalIgnoreCase) &&
                    ar.TextString != id.ToString())
                {
                    ar.TextString = id.ToString();
                    return;
                }
            }
        }

        public static List<VertexInfo> ReadVertexData(ObjectId plId, Database db)
        {
            var result = new List<VertexInfo>();
            if (db == null || plId.IsNull || !plId.IsValid || plId.IsErased) return result;

            try
            {
                using (var tr = db.TransactionManager.StartOpenCloseTransaction())
                {
                    var pl = tr.GetObject(plId, OpenMode.ForRead, false) as Polyline;
                    if (pl != null) result = ReadVertexData(tr, pl);
                    tr.Commit();
                }
            }
            catch (System.Exception) { }

            return result;
        }

        private static List<VertexInfo> ReadVertexData(Transaction tr, Polyline pl)
        {
            var result = new List<VertexInfo>();
            if (tr == null || pl == null || pl.ExtensionDictionary.IsNull) return result;

            var ext = tr.GetObject(pl.ExtensionDictionary, OpenMode.ForRead, false) as DBDictionary;
            if (ext == null || !ext.Contains("HybridData")) return result;

            var dataDict = tr.GetObject(ext.GetAt("HybridData"), OpenMode.ForRead, false) as DBDictionary;
            if (dataDict == null || !dataDict.Contains("Data")) return result;

            var xr = tr.GetObject(dataDict.GetAt("Data"), OpenMode.ForRead, false) as Xrecord;
            if (xr?.Data == null) return result;

            foreach (TypedValue tv in xr.Data)
            {
                if (tv.TypeCode != (int)DxfCode.Text || !(tv.Value is string json)) continue;

                var list = JsonConvert.DeserializeObject<List<SimpleVertex>>(json);
                if (list == null) continue;

                result.AddRange(list.Select(sv => VertexAt(
                    new Point3d(sv.X, sv.Y, 0),
                    new VertexInfo
                    {
                        Type = sv.Type,
                        Desc = sv.Desc,
                        ID = sv.ID
                    })));
            }

            return NormalizeVertices(result);
        }

        [CommandMethod("HYBRIDPLAN")]
        public static void HybridPlan()
        {
            if (!TryRequireUtmMode("Open Hybrid Vertex Editor")) return;

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
            if (!TryRequireUtmMode("Save Hybrid Table Data")) return;

            // ‑‑‑ do NOT fire the guard while we touch attributes
            using (HybridGuard.Suspend())
            {
                var normalizedVertices = NormalizeVertices(verts);
                if (normalizedVertices.Count == 0) return;

                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                var db = doc.Database;
                var ed = doc.Editor;

                using (doc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var styleId = EnsureTableStyle(tr, db);
                    EnsureLayer(tr, db, kTableLayer);

                    // Always work in Model Space
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var btrModel = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    // Find existing table in MODEL SPACE with the requested style + layer
                    ObjectId tblId = ObjectId.Null;
                    foreach (ObjectId entId in btrModel)
                    {
                        var tblTry = tr.GetObject(entId, OpenMode.ForRead) as Table;
                        if (tblTry == null) continue;
                        bool styleMatch = tblTry.TableStyle == styleId;
                        bool layerMatch = tblTry.Layer.Equals(kTableLayer, StringComparison.OrdinalIgnoreCase);
                        if (styleMatch && layerMatch)
                        {
                            tblId = entId;
                            break;
                        }
                    }

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
                        var ppr = ed.GetPoint("\nPick table insertion point (Model Space):");
                        if (ppr.Status != PromptStatus.OK) return;
                        insPt = ppr.Value;

                        tbl = new Table
                        {
                            TableStyle = styleId,
                            Layer = kTableLayer,
                            Position = insPt
                        };
                        tbl.SetDatabaseDefaults();

                        btrModel.AppendEntity(tbl);
                        tr.AddNewlyCreatedDBObject(tbl, true);
                    }

                    // Fill table
                    tbl.Position = insPt;
                    tbl.SetSize(normalizedVertices.Count + 1, 5);
                    tbl.DeleteRows(0, 1);

                    ApplyHybridTableLayout(tbl);

                    for (int i = 0; i < normalizedVertices.Count; i++)
                    {
                        var v = normalizedVertices[i];
                        Write(tbl, i, 0, (i + 1).ToString());
                        Write(tbl, i, 1, DisplayCoordinate(v.N));
                        Write(tbl, i, 2, DisplayCoordinate(v.E));
                        Write(tbl, i, 3, v.Type);
                        Write(tbl, i, 4, v.Desc);
                    }

                    WriteTableMetadata(tbl, normalizedVertices, tr);
                    tbl.GenerateLayout();

                    if (insertHybrid)
                    {
                        EnsureAllHybridBlocks(tr, db);
                        PlaceHybridBlocks(
                            tr,
                            db,
                            normalizedVertices,
                            "TABLE-" + tbl.Handle);   // legacy table-only API
                    }

                    tr.Commit();
                }
            } // guard resumes
        }
        private static void Write(Table tbl, int row, int col, string txt)
        {
            var cell = tbl.Cells[row, col];
            cell.TextHeight = kTableTxtH;
            cell.TextString = txt ?? "";
        }

        private static void ApplyHybridTableLayout(Table table)
        {
            int columnCount = Math.Min(kColW.Length, table.Columns.Count);
            for (int column = 0; column < columnCount; column++)
                table.Columns[column].Width = kColW[column];

            for (int row = 0; row < table.Rows.Count; row++)
            {
                table.Rows[row].Height = kRowH;
                for (int column = 0; column < columnCount; column++)
                    table.Cells[row, column].TextHeight = kTableTxtH;
            }
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
            var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var ms = (BlockTableRecord)tr.GetObject(
                blockTable[BlockTableRecord.ModelSpace],
                OpenMode.ForRead);

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
            IList<VertexInfo> verts,
            string ownerKey)
        {
            SyncHybridMarkerBlocks(
                tr,
                db,
                ownerKey,
                new List<VertexInfo>(),
                verts,
                insertOrUpdate: true);
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

                // recognize both Hybrd Num and legacy Hybrid Num, incl. dynamic
                string nm;
                if (br.IsDynamicBlock)
                {
                    var dyn = (BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead);
                    nm = dyn?.Name ?? br.Name;
                }
                else nm = br.Name;
                nm = (nm ?? string.Empty).Trim();
                if (!(nm.Equals("Hybrd Num", StringComparison.OrdinalIgnoreCase) ||
                      nm.Equals("Hybrid Num", StringComparison.OrdinalIgnoreCase)))
                    continue;

                int idVal = -1;
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                    if (ar == null) continue;
                    if ((ar.Tag ?? string.Empty).Equals("ID", StringComparison.OrdinalIgnoreCase) &&
                        int.TryParse(ar.TextString?.Trim(), out int v))
                    {
                        idVal = v;
                        break;
                    }
                }

                if (!keepIds.Contains(idVal))
                {
                    // safe erase (unlock layer temporarily if needed)
                    var lyr = (LayerTableRecord)tr.GetObject(br.LayerId, OpenMode.ForWrite);
                    bool relock = false;
                    if (lyr.IsLocked) { lyr.IsLocked = false; relock = true; }
                    br.UpgradeOpen();
                    br.Erase();
                    if (relock) lyr.IsLocked = true;
                }
            }
        }

        internal static string GuessType(Point3d pt, Transaction tr)
        {
            var db = AcadApp.DocumentManager.MdiActiveDocument.Database;
            var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var ms = (BlockTableRecord)tr.GetObject(
                blockTable[BlockTableRecord.ModelSpace],
                OpenMode.ForRead);

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

            // OWNER scopes otherwise-repeating vertex IDs to one polyline.  This
            // prevents numbering or deletion on one Hybrid polyline from touching
            // number bubbles that belong to another.
            if (!ids.TryGetValue("OWNER", out var ownerDef))
            {
                ownerDef = new AttributeDefinition
                {
                    Tag = "OWNER",
                    Prompt = "Owner",
                    Height = kTxtH,
                    Position = Point3d.Origin,
                    Invisible = true,
                    Constant = false
                };
                ownerDef.SetDatabaseDefaults();
                btr.AppendEntity(ownerDef);
                tr.AddNewlyCreatedDBObject(ownerDef, true);
            }
            else if (ownerDef.Constant || !ownerDef.Invisible)
            {
                ownerDef.UpgradeOpen();
                ownerDef.Constant = false;
                ownerDef.Invisible = true;
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
            if (!TryRequireUtmMode("Update Hybrid Polyline Geometry")) return;

            if (plId == ObjectId.Null || !plId.IsValid || verts == null) return;

            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Open read-only first so a locked layer cannot reject the object
                // before we have a chance to unlock it.
                if (!(tr.GetObject(plId, OpenMode.ForRead, false) is Polyline pl))
                    return;

                using (UnlockLayer(tr, pl.LayerId))
                {
                    pl.UpgradeOpen();
                    ApplyPolylineGeometry(pl, NormalizeVertices(verts));
                }

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

                /* ---------- 3) original body, but force Model Space ---------- */
                EnsureHybridNumBlock(tr, db);

                bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var msId = bt[BlockTableRecord.ModelSpace];
                space = (BlockTableRecord)tr.GetObject(msId, OpenMode.ForWrite);   // <- MODEL SPACE ONLY
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
            SyncNumberingCommand();
        }

        // -----------------------------------------------------------------------------
        // UPDATE NUMBERING
        // -----------------------------------------------------------------------------
        [CommandMethod("UpdateNumbering")]
        public static void UpdateNumbering()
        {
            SyncNumberingCommand();
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

                // recognize both Hybrd Num and legacy Hybrid Num, incl. dynamic
                string nm;
                if (br.IsDynamicBlock)
                {
                    var dyn = (BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead);
                    nm = dyn?.Name ?? br.Name;
                }
                else nm = br.Name;
                nm = (nm ?? string.Empty).Trim();
                if (!(nm.Equals("Hybrd Num", StringComparison.OrdinalIgnoreCase) ||
                      nm.Equals("Hybrid Num", StringComparison.OrdinalIgnoreCase)))
                    continue;

                foreach (ObjectId attId in br.AttributeCollection)
                {
                    var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                    if (ar == null) continue;

                    if ((ar.Tag ?? string.Empty).Equals("NUMBER", StringComparison.OrdinalIgnoreCase) &&
                        int.TryParse(ar.TextString?.Trim(), out int seq) &&
                        seq > maxNumber)
                    {
                        var lyr = (LayerTableRecord)tr.GetObject(br.LayerId, OpenMode.ForWrite);
                        bool relock = false;
                        if (lyr.IsLocked) { lyr.IsLocked = false; relock = true; }
                        br.UpgradeOpen();
                        br.Erase();
                        if (relock) lyr.IsLocked = true;
                        break; // one NUMBER per block – next block
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
            if (!TryRequireUtmMode("Dump Hybrid Vertex Data")) return;

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
            if (!TryRequireUtmMode("Purge Hybrid Data")) return;

            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            try
            {
                using (doc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var btrMs = (BlockTableRecord)tr.GetObject(
                        blockTable[BlockTableRecord.ModelSpace],
                        OpenMode.ForRead);

                    int purged = 0;

                    foreach (ObjectId entId in btrMs)
                    {
                        // --- only consider polylines ---------------------------------------------
                        if (!(tr.GetObject(entId, OpenMode.ForRead) is Polyline pl)) continue;

                        // layer info ­- we might have to unlock it temporarily
                        var lyrRec = (LayerTableRecord)tr.GetObject(pl.LayerId, OpenMode.ForWrite);
                        bool relock = false;

                        if (lyrRec.IsLocked)
                        {
                            lyrRec.IsLocked = false;   // safe: the layer may be current
                            relock = true;
                        }

                        // switch the polyline to write-mode after its layer is unlocked
                        pl.UpgradeOpen();

                        if (!pl.ExtensionDictionary.IsNull)
                        {
                            var ext = (DBDictionary)tr.GetObject(pl.ExtensionDictionary, OpenMode.ForWrite);
                            if (ext.Contains("HybridData"))
                            {
                                // Remove and erase the owned HybridData dictionary. It
                                // owns both Data and Owner Xrecords, so erasing only an
                                // assumed Xrecord here would leak/corrupt metadata.
                                var dataId = ext.GetAt("HybridData");
                                var dataDictionary = tr.GetObject(
                                    dataId,
                                    OpenMode.ForWrite,
                                    false) as DBDictionary;
                                ext.Remove("HybridData");
                                dataDictionary?.Erase();

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
            catch (System.Exception ex)
            {
                ed.WriteMessage(
                    "\nHybrid data was not purged. No partial drawing changes were committed. " +
                    ex.Message);
            }
        }


        private static BlockReference CreateBubble(Transaction tr,
                                                   BlockTableRecord space,
                                                   BlockTable bt,
                                                   BlockTableRecord defRec,
                                                   Point3d pt,
                                                   int id,
                                                   string ownerKey = "")
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
                    ar.SetAttributeFromBlock(def, br.BlockTransform);
                    ar.Invisible = def.Invisible;
                    if (def.Tag.Equals("ID", StringComparison.OrdinalIgnoreCase))
                        ar.TextString = id.ToString();
                    else if (def.Tag.Equals("OWNER", StringComparison.OrdinalIgnoreCase))
                        ar.TextString = ownerKey ?? string.Empty;
                    else
                        ar.TextString = string.Empty;

                    br.AttributeCollection.AppendAttribute(ar);
                    tr.AddNewlyCreatedDBObject(ar, true);
                }
            }

            return br;    // caller commits
        }

        private static readonly Dictionary<ObjectId, Dictionary<Point3d, BlockReference>> _bubbleCache
            = new Dictionary<ObjectId, Dictionary<Point3d, BlockReference>>();

        /// <summary>
        /// Clears all cached bubble maps.  Call this method when a drawing
        /// is closing or when the plug‑in is being unloaded to avoid leaving
        /// stale ObjectIds in the static cache.  Failing to clear this cache
        /// can result in references to disposed databases and stability issues.
        /// </summary>
        internal static void ClearCaches()
        {
            _bubbleCache.Clear();
        }

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
            if (!TryRequireUtmMode("Transfer Hybrid Data")) return;

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
            if (srcId == dstId)
            {
                ed.WriteMessage("\nSource and target must be different polylines.");
                return;
            }

            try
            {
                using (doc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var source = tr.GetObject(srcId, OpenMode.ForRead, false) as Polyline;
                    var target = tr.GetObject(dstId, OpenMode.ForRead, false) as Polyline;
                    if (source == null || target == null ||
                        !IsModelSpaceEntity(tr, db, source) ||
                        !IsModelSpaceEntity(tr, db, target))
                    {
                        ed.WriteMessage("\nBoth source and target must be Model Space 2D polylines.");
                        return;
                    }

                    var sourceVertices = ReadVertexData(tr, source);
                    if (sourceVertices.Count == 0)
                    {
                        ed.WriteMessage("\nThe source polyline has no stored Hybrid data.");
                        return;
                    }

                    var previousTargetVertices = ReadVertexData(tr, target);

                    var unusedSource = new List<VertexInfo>(sourceVertices);
                    var transferred = new List<VertexInfo>();
                    for (int i = 0; i < target.NumberOfVertices; i++)
                    {
                        Point3d point = StorePoint(target.GetPoint3dAt(i));
                        int matchIndex = unusedSource.FindIndex(
                            vertex => vertex.Pt.DistanceTo(point) <= kMatchTol);
                        if (matchIndex >= 0)
                        {
                            // Metadata follows the match, but coordinates always
                            // remain the target polyline's actual rounded point.
                            transferred.Add(VertexAt(point, unusedSource[matchIndex]));
                            unusedSource.RemoveAt(matchIndex);
                        }
                        else
                        {
                            transferred.Add(VertexAt(point));
                        }
                    }

                    transferred = NormalizeVertices(transferred);
                    using (UnlockLayer(tr, target.LayerId))
                        WriteVertexData(tr, target, transferred);

                    UpdateOwnedTable(
                        tr,
                        db,
                        ed,
                        transferred,
                        target.Handle.ToString(),
                        createIfMissing: false,
                        legacyMatchVertices: previousTargetVertices);

                    // Do not create marker/number blocks as a side effect of a
                    // transfer. The explicit Save and Numbering actions own those.
                    tr.Commit();
                    ed.WriteMessage(
                        $"\nTransferred Hybrid data to {transferred.Count} target vertices; " +
                        $"{unusedSource.Count} unmatched source rows were not appended.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nHybrid data was not transferred. " + ex.Message);
            }
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
            if (tbl == null || tbl.Columns.Count < 5) return verts;
            int rowCount = tbl.Rows.Count;

            for (int r = 0; r < rowCount; r++)
            {
                string sN = tbl.Cells[r, 1].TextString;
                string sE = tbl.Cells[r, 2].TextString;

                if (TryParseCoordinate(sN, out double N) &&
                    TryParseCoordinate(sE, out double E))
                {
                    verts.Add(new VertexInfo
                    {
                        Pt = new Point3d(E, N, 0),
                        N = N,
                        E = E,
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

            // Expected map for the *current* working polyline
            var expectedSeq = new Dictionary<int, string>();
            var idToPt = new Dictionary<int, Point3d>();
            for (int i = 0; i < verts.Count; i++)
            {
                var v = verts[i];
                if (v == null) continue;
                if (v.ID == 0) continue; // defensive

                if (!expectedSeq.ContainsKey(v.ID))
                {
                    expectedSeq[v.ID] = (i + 1).ToString();
                    idToPt[v.ID] = v.Pt;
                }
            }

            // Helpers (local)
            bool IsHybrdNum(BlockReference br)
            {
                string nm;
                if (br.IsDynamicBlock)
                {
                    var dyn = (BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead);
                    nm = dyn?.Name ?? br.Name;
                }
                else
                {
                    nm = br.Name;
                }
                nm = (nm ?? string.Empty).Trim();
                return nm.Equals("Hybrd Num", StringComparison.OrdinalIgnoreCase)
                    || nm.Equals("Hybrid Num", StringComparison.OrdinalIgnoreCase); // legacy support
            }

            void SafeErase(BlockReference br)
            {
                var lyr = (LayerTableRecord)tr.GetObject(br.LayerId, OpenMode.ForWrite);
                bool relock = false;
                if (lyr.IsLocked) { lyr.IsLocked = false; relock = true; }
                br.UpgradeOpen();
                br.Erase();
                if (relock) lyr.IsLocked = true;
            }

            // Scan Hybrd/Hybrid Num blocks in the given space
            var byId = new Dictionary<int, List<Tuple<BlockReference, string>>>();
            var junk = new List<BlockReference>();

            foreach (ObjectId entId in space)
            {
                if (entId.ObjectClass != RXObject.GetClass(typeof(BlockReference))) continue;
                var br = (BlockReference)tr.GetObject(entId, OpenMode.ForRead);
                if (!IsHybrdNum(br)) continue;

                int idVal = -1; // definite assignment
                bool okId = false;
                string numberText = string.Empty;

                foreach (ObjectId attId in br.AttributeCollection)
                {
                    var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                    if (ar == null) continue;

                    var tag = ar.Tag ?? string.Empty;

                    if (tag.Equals("ID", StringComparison.OrdinalIgnoreCase) &&
                        int.TryParse(ar.TextString?.Trim(), out idVal))
                    {
                        okId = true;
                    }
                    else if (tag.Equals("NUMBER", StringComparison.OrdinalIgnoreCase))
                    {
                        numberText = (ar.TextString ?? string.Empty).Trim();
                    }
                }

                if (!okId)
                {
                    junk.Add(br);
                    continue;
                }

                if (!byId.TryGetValue(idVal, out var list))
                    byId[idVal] = list = new List<Tuple<BlockReference, string>>();

                list.Add(Tuple.Create(br, numberText));
            }

            var validIds = new HashSet<int>(expectedSeq.Keys);

            // 1) Delete everything whose ID is not part of the current vertex set
            foreach (var kvp in byId)
            {
                if (validIds.Contains(kvp.Key)) continue;
                foreach (var pair in kvp.Value)
                    SafeErase(pair.Item1);
            }

            // 2) For each expected ID, keep exactly one bubble; ensure NUMBER matches; delete duplicates
            foreach (int id in validIds)
            {
                string wantSeq = expectedSeq[id];

                byId.TryGetValue(id, out var existingForId);
                BlockReference keeper = null;

                if (existingForId != null && existingForId.Count > 0)
                {
                    // Prefer a bubble that already has the correct NUMBER
                    foreach (var pair in existingForId)
                    {
                        if (string.Equals(pair.Item2, wantSeq, StringComparison.Ordinal))
                        {
                            keeper = pair.Item1;
                            break;
                        }
                    }
                    // Else, take the first
                    if (keeper == null)
                        keeper = existingForId[0].Item1;

                    // Delete the rest (duplicates)
                    foreach (var pair in existingForId)
                    {
                        var br = pair.Item1;
                        if (br.ObjectId == keeper.ObjectId) continue;
                        SafeErase(br);
                    }

                    // Ensure keeper has correct ID + NUMBER
                    foreach (ObjectId attId in keeper.AttributeCollection)
                    {
                        var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForWrite);
                        if (ar == null) continue;

                        var tag = ar.Tag ?? string.Empty;

                        if (tag.Equals("ID", StringComparison.OrdinalIgnoreCase))
                        {
                            string wantIdText = id.ToString();
                            if (!string.Equals(ar.TextString?.Trim(), wantIdText, StringComparison.Ordinal))
                                ar.TextString = wantIdText;
                        }
                        else if (tag.Equals("NUMBER", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!string.Equals(ar.TextString?.Trim(), wantSeq, StringComparison.Ordinal))
                            {
                                ar.TextString = wantSeq;
                                ar.AdjustAlignment(db);
                            }
                        }
                    }
                }
                else
                {
                    // No bubble at all for this ID → create one (numbering operations only)
                    var nb = CreateBubble(
                        tr,
                        space,
                        (BlockTable)tr.GetObject(space.Database.BlockTableId, OpenMode.ForRead),
                        defRec,
                        idToPt[id],
                        id);

                    // Stamp NUMBER on the new bubble
                    foreach (ObjectId attId in nb.AttributeCollection)
                    {
                        var ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForWrite);
                        if (ar == null) continue;
                        if ((ar.Tag ?? string.Empty).Equals("NUMBER", StringComparison.OrdinalIgnoreCase))
                        {
                            ar.TextString = wantSeq;
                            ar.AdjustAlignment(db);
                        }
                    }
                }
            }

            // 3) Purge all bubbles without a readable ID
            foreach (var br in junk)
                SafeErase(br);
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
            if (!TryRequireUtmMode("Rebuild Polyline From Table")) return;

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

                using (var validation = db.TransactionManager.StartOpenCloseTransaction())
                {
                    var selectedTable = validation.GetObject(
                        per.ObjectId,
                        OpenMode.ForRead,
                        false) as Table;
                    if (selectedTable == null ||
                        !IsModelSpaceEntity(validation, db, selectedTable))
                    {
                        ed.WriteMessage("\nSelect a Hybrid table from Model Space.");
                        return;
                    }
                    validation.Commit();
                }

                List<VertexInfo> verts;

                // Preferred path: pull 0.001 metadata from the table itself.
                bool readMetadata = false;
                try
                {
                    readMetadata = TryReadTableMetadata(per.ObjectId, db, out verts);
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage("\nStored table metadata is unreadable; using visible rows. " + ex.Message);
                    verts = new List<VertexInfo>();
                }

                if (readMetadata)
                {
                    ed.WriteMessage("\nUsing vertex metadata stored on the table.");
                }
                else
                {
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

                try
                {
                    verts = NormalizeVertices(verts);
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage("\nThe table contains an invalid coordinate. " + ex.Message);
                    return;
                }

                try
                {
                    // Build and persist the new polyline atomically in Model Space.
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var tbl = (Table)tr.GetObject(per.ObjectId, OpenMode.ForRead, false);
                        var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var model = (BlockTableRecord)tr.GetObject(
                            blockTable[BlockTableRecord.ModelSpace],
                            OpenMode.ForWrite);

                        Polyline pl;
                        using (UnlockLayer(tr, db.Clayer))
                        {
                            pl = new Polyline();
                            pl.SetDatabaseDefaults(db);
                            pl.LayerId = db.Clayer;
                            for (int i = 0; i < verts.Count; i++)
                            {
                                pl.AddVertexAt(
                                    i,
                                    new Point2d(verts[i].E, verts[i].N),
                                    0,
                                    0,
                                    0);
                            }
                            model.AppendEntity(pl);
                            tr.AddNewlyCreatedDBObject(pl, true);
                            WriteVertexData(tr, pl, verts);
                        }

                        using (UnlockLayer(tr, tbl.LayerId))
                        {
                            tbl.UpgradeOpen();
                            ApplyHybridTableLayout(tbl);
                            WriteTableMetadata(tbl, verts, tr);
                            WriteOwnerMetadata(tr, tbl, pl.Handle.ToString());
                            tbl.GenerateLayout();
                        }

                        tr.Commit();
                        ed.WriteMessage($"\nRebuilt Model Space polyline with {verts.Count} vertices.");
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage("\nPolyline was not rebuilt. " + ex.Message);
                }
            }
        }

        [CommandMethod("REBUILDPOLYLINEFROMTABLE")]
        public static void RebuildPolylineFromTableAlias()
        {
            RebuildPolylineFromTable();
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
        private DataGridViewCell _editingCoordinateCell;
        private double _editingCoordinateOriginal;
        private string _editingCoordinateDisplay;
        private bool _hasEditingCoordinateOriginal;
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
            ClientSize = new System.Drawing.Size(1240, 640);
            MinimumSize = new System.Drawing.Size(900, 500);
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = true;

            var pnl = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 88,
                Padding = new Padding(5),
                WrapContents = true,
                AutoScroll = true,
                FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight
            };

            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = true
            };

            _grid.Columns.Add("#", "#");
            _grid.Columns["#"].ReadOnly = true;
            _grid.Columns["#"].FillWeight = 35;

            var coordinateStyle = new DataGridViewCellStyle
            {
                Format = "F2",
                FormatProvider = CultureInfo.InvariantCulture,
                NullValue = string.Empty
            };
            _grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Northing",
                HeaderText = "Northing",
                ValueType = typeof(double),
                DefaultCellStyle = coordinateStyle,
                FillWeight = 85
            });
            _grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Easting",
                HeaderText = "Easting",
                ValueType = typeof(double),
                DefaultCellStyle = coordinateStyle.Clone() as DataGridViewCellStyle,
                FillWeight = 85
            });

            var typeCol = new DataGridViewComboBoxColumn
            {
                Name = "Type",
                HeaderText = "Type",
                FlatStyle = FlatStyle.Flat,
                DropDownWidth = 80,
                FillWeight = 55
            };
            typeCol.Items.AddRange("", "XC", "RC", "EC", "OC", "Other");
            _grid.Columns.Add(typeCol);
            _grid.Columns.Add("Desc", "Description");

            // Row order is the polyline's vertex order. Header sorting would
            // visually reorder values without reordering the backing IDs.
            foreach (DataGridViewColumn column in _grid.Columns)
                column.SortMode = DataGridViewColumnSortMode.NotSortable;

            Controls.Add(_grid);
            Controls.Add(pnl);

            /* ---------- command buttons ---------- */
            Button ActionButton(string text) => new Button
            {
                Text = text,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Height = 30,
                Margin = new Padding(3)
            };

            var btnGet = ActionButton("Select Polyline With Hybrid Data");
            var btnNumbering = ActionButton("Add / Update Numbering");
            var btnRebuild = ActionButton("Rebuild Polyline From Table");
            var btnTransfer = ActionButton("Transfer Hybrid Data From a Polyline to Another");
            var btnOK = ActionButton("Save and Update Data to Table & Polyline");
            _chkHybrid = new CheckBox
            {
                Text = "Update XC / EC Blocks",
                Checked = true,
                AutoSize = true,
                Margin = new Padding(8, 8, 8, 3)
            };
            var btnRemove = ActionButton("Delete Selected Point(s)");

            /*   ➜  add all the “regular” buttons first   */
            pnl.Controls.AddRange(new Control[] {
        btnGet, btnNumbering, btnRebuild, btnTransfer,
        _chkHybrid, btnOK, btnRemove
    });

            /* ---------- event wiring ---------- */
            btnRemove.Click += RemoveSelected_Click;
            btnGet.Click += (s, e) => PickAndPopulate();
            btnNumbering.Click += AddOrUpdateNumbering_Click;
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
            _grid.CellBeginEdit += Grid_CellBeginEdit;
            _grid.CellEndEdit += Grid_CellEndEdit;
            _grid.CellParsing += Grid_CellParsing;
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
                if (!IsDisposed && IsHandleCreated)
                    BeginInvoke((MethodInvoker)(() =>
                    {
                        if (!IsDisposed) Close();
                    }));
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
        // VertexForm.cs – full replacement for PickAndPopulate()
        // VertexForm.cs – revised PickAndPopulate
        private void PickAndPopulate()
        {
            if (!HybridCommands.TryRequireUtmMode("Select Polyline With Hybrid Data")) return;

            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            var opts = new PromptEntityOptions("\nSelect Model Space polyline with Hybrid data: ");
            opts.SetRejectMessage("\nSelect a 2D polyline.");
            opts.AddAllowedClass(typeof(Polyline), false);
            var res = ed.GetEntity(opts);
            if (res.Status != PromptStatus.OK) return;

            if (!HybridCommands.LoadPolylineVertices(
                    res.ObjectId,
                    out var loaded,
                    out var hadStoredData,
                    out var error))
            {
                MessageBox.Show(
                    error,
                    "Hybrid Vertex Editor",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            // Selection/reload is deliberately read-only.  It reconciles current
            // polyline vertices with stored metadata in memory; the explicit Save
            // action is the only UI path that writes geometry/table/block data.
            _currentPlId = res.ObjectId;
            _verts = loaded;
            RefreshGrid();

            if (!hadStoredData && loaded.Count > 0)
                ed.WriteMessage(
                    "\nNo stored Hybrid metadata was found. Current vertices were loaded; use Save when the rows are ready.");
        }

        private void AddOrUpdateNumbering_Click(object sender, EventArgs e)
        {
            if (_currentPlId.IsNull || !_currentPlId.IsValid || _currentPlId.IsErased)
            {
                MessageBox.Show(
                    "Select a polyline with Hybrid data first.",
                    "Hybrid Vertex Editor",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }
            if (!HybridCommands.TryRequireUtmMode("Add / Update Numbering")) return;

            if (HybridCommands.AddOrUpdateNumbering(
                    _currentPlId,
                    out var updatedVertices,
                    out var error))
            {
                _verts = updatedVertices;
                RefreshGrid();
            }
            else
            {
                MessageBox.Show(
                    error,
                    "Hybrid Vertex Editor",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private bool TryReadGridVertices(out List<VertexInfo> vertices, out string error)
        {
            vertices = new List<VertexInfo>();
            error = string.Empty;
            _grid.EndEdit();

            for (int i = 0; i < _grid.Rows.Count; i++)
            {
                var row = _grid.Rows[i];
                if (!HybridCommands.TryParseCoordinate(row.Cells["Northing"].Value, out var north) ||
                    !HybridCommands.TryParseCoordinate(row.Cells["Easting"].Value, out var east))
                {
                    error = "Row " + (i + 1) + " needs valid Northing and Easting values using a dot (.) decimal separator.";
                    return false;
                }

                var existing = i < _verts.Count ? _verts[i] : null;
                vertices.Add(new VertexInfo
                {
                    Pt = new Point3d(east, north, existing?.Pt.Z ?? 0.0),
                    N = north,
                    E = east,
                    Type = Convert.ToString(row.Cells["Type"].Value) ?? string.Empty,
                    Desc = Convert.ToString(row.Cells["Desc"].Value) ?? string.Empty,
                    ID = existing?.ID ?? 0
                });
            }

            return true;
        }

        // VertexForm.cs  – FULL replacement of InsertUpdate_Click
        private void InsertUpdate_Click(object sender, EventArgs e)
        {
            if (_currentPlId == ObjectId.Null || !_currentPlId.IsValid || _currentPlId.IsErased)
            {
                MessageBox.Show("Select a polyline with Hybrid data first.",
                                "Hybrid Vertex Editor", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!HybridCommands.TryRequireUtmMode("Save and Update Hybrid Data")) return;

            if (!TryReadGridVertices(out var pendingVertices, out var validationError))
            {
                MessageBox.Show(
                    validationError,
                    "Invalid Hybrid data",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (HybridCommands.SaveAndUpdate(
                    _currentPlId,
                    pendingVertices,
                    _chkHybrid.Checked,
                    out var savedVertices,
                    out var saveError))
            {
                _verts = savedVertices;
                RefreshGrid();
                return;
            }

            if (!string.IsNullOrWhiteSpace(saveError))
                MessageBox.Show(
                    saveError,
                    "Hybrid data was not saved",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
        }

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
            if (_currentPlId == ObjectId.Null || !_currentPlId.IsValid || _currentPlId.IsErased)
            {
                MessageBox.Show(
                    "Select a polyline with Hybrid data first.",
                    "Hybrid Vertex Editor",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }
            if (!HybridCommands.TryRequireUtmMode("Delete Hybrid Point")) return;

            var selectedIndexes = _grid.SelectedCells
                .Cast<DataGridViewCell>()
                .Select(cell => cell.RowIndex)
                .Concat(_grid.SelectedRows.Cast<DataGridViewRow>().Select(row => row.Index))
                .Where(index => index >= 0 && index < _grid.Rows.Count)
                .Distinct()
                .OrderBy(index => index)
                .ToList();
            if (selectedIndexes.Count == 0 && _grid.CurrentCell != null)
                selectedIndexes.Add(_grid.CurrentCell.RowIndex);
            if (selectedIndexes.Count == 0) return;

            string noun = selectedIndexes.Count == 1 ? "point" : "points";
            if (MessageBox.Show(
                    "Delete the selected " + noun + " from the table, polyline, and any managed blocks?",
                    "Delete Hybrid " + noun,
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning) != DialogResult.Yes)
                return;

            if (!TryReadGridVertices(out var pendingVertices, out var validationError))
            {
                MessageBox.Show(
                    validationError,
                    "Invalid Hybrid data",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            var removeSet = new HashSet<int>(selectedIndexes);
            var remaining = pendingVertices
                .Where((vertex, index) => !removeSet.Contains(index))
                .ToList();

            if (HybridCommands.DeletePointsAndUpdate(
                    _currentPlId,
                    remaining,
                    _chkHybrid.Checked,
                    out var savedVertices,
                    out var error))
            {
                _verts = savedVertices;
                if (_verts.Count == 0) _currentPlId = ObjectId.Null;
                RefreshGrid();
                return;
            }

            MessageBox.Show(
                error,
                "Hybrid point was not deleted",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }





        private void Grid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                CopySelectionToClipboard();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.Control && e.KeyCode == Keys.V)
            {
                PasteClipboardToGrid();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void CopySelectionToClipboard()
        {
            var selected = _grid.SelectedCells
                .Cast<DataGridViewCell>()
                .Where(cell => cell.RowIndex >= 0 && cell.ColumnIndex >= 0)
                .ToList();
            if (selected.Count == 0) return;

            int minRow = selected.Min(cell => cell.RowIndex);
            int maxRow = selected.Max(cell => cell.RowIndex);
            int minColumn = selected.Min(cell => cell.ColumnIndex);
            int maxColumn = selected.Max(cell => cell.ColumnIndex);
            var selectedKeys = new HashSet<int>(
                selected.Select(cell =>
                    (cell.RowIndex * _grid.ColumnCount) + cell.ColumnIndex));
            var clipboard = new StringBuilder();

            for (int row = minRow; row <= maxRow; row++)
            {
                if (row > minRow) clipboard.AppendLine();
                for (int column = minColumn; column <= maxColumn; column++)
                {
                    if (column > minColumn) clipboard.Append('\t');
                    int key = (row * _grid.ColumnCount) + column;
                    if (!selectedKeys.Contains(key)) continue;

                    var cell = _grid[column, row];
                    string value;
                    if (IsCoordinateColumn(cell.OwningColumn.Name) &&
                        HybridCommands.TryParseCoordinate(cell.Value, out var coordinate))
                    {
                        // Clipboard retains storage precision even though the
                        // grid deliberately renders these values with F2.
                        value = HybridValuePolicy.StorageCoordinateText(coordinate);
                    }
                    else
                    {
                        value = Convert.ToString(
                            cell.FormattedValue,
                            CultureInfo.InvariantCulture) ?? string.Empty;
                        value = value.Replace('\t', ' ')
                                     .Replace('\r', ' ')
                                     .Replace('\n', ' ');
                    }

                    clipboard.Append(value);
                }
            }

            Clipboard.SetText(clipboard.ToString());
        }

        private static bool IsCoordinateColumn(string columnName)
        {
            return columnName == "Northing" || columnName == "Easting";
        }

        private void Grid_CellBeginEdit(
            object sender,
            DataGridViewCellCancelEventArgs e)
        {
            _editingCoordinateCell = null;
            _editingCoordinateDisplay = string.Empty;
            _hasEditingCoordinateOriginal = false;

            if (e.RowIndex < 0 || e.ColumnIndex < 0 ||
                !IsCoordinateColumn(_grid.Columns[e.ColumnIndex].Name))
                return;

            var cell = _grid[e.ColumnIndex, e.RowIndex];
            if (!HybridCommands.TryParseCoordinate(cell.Value, out var original))
                return;

            _editingCoordinateCell = cell;
            _editingCoordinateOriginal = original;
            _editingCoordinateDisplay = Convert.ToString(
                cell.FormattedValue,
                CultureInfo.InvariantCulture) ??
                HybridCommands.DisplayCoordinate(original);
            _hasEditingCoordinateOriginal = true;
        }

        private void Grid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            _editingCoordinateCell = null;
            _editingCoordinateDisplay = string.Empty;
            _hasEditingCoordinateOriginal = false;
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
                if (!HybridCommands.TryParseCoordinate(value, out _))
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

        private void Grid_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
        {
            string columnName = _grid.Columns[e.ColumnIndex].Name;
            if (!IsCoordinateColumn(columnName)) return;

            if (_hasEditingCoordinateOriginal &&
                ReferenceEquals(
                    _editingCoordinateCell,
                    _grid[e.ColumnIndex, e.RowIndex]) &&
                HybridValuePolicy.TryResolveEditedCoordinate(
                    e.Value,
                    _editingCoordinateOriginal,
                    _editingCoordinateDisplay,
                    out var resolvedCoordinate))
            {
                // Entering and leaving a coordinate cell exposes only its F2
                // text. Preserve the original hidden third decimal unless the
                // user actually changes that text; changed text is normalized
                // through the same 0.001 storage policy.
                e.Value = resolvedCoordinate;
                e.ParsingApplied = true;
                return;
            }

            if (HybridCommands.TryParseCoordinate(e.Value, out var coordinate))
            {
                // Keep the underlying 0.001 value while the column's F2 format
                // displays 0.00.  A routine save therefore cannot discard the
                // hidden third decimal.
                e.Value = coordinate;
                e.ParsingApplied = true;
            }
        }

        // VertexForm.cs  – REPLACE the entire PasteClipboardToGrid method
        // (no other changes required)

        private void PasteClipboardToGrid()
        {
            string text = Clipboard.GetText();
            if (string.IsNullOrEmpty(text)) return;

            text = text.TrimEnd('\r', '\n');
            if (_grid.Rows.Count == 0)
            {
                _verts.Add(new VertexInfo());
                _grid.Rows.Add(1, 0.0, 0.0, string.Empty, string.Empty);
            }

            string[] sourceLines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            string[][] sourceCells = sourceLines.Select(line => line.Split('\t')).ToArray();

            // Excel-style fill: a copied single cell is broadcast to every
            // selected editable cell, without copying or replacing whole rows.
            if (sourceCells.Length == 1 && sourceCells[0].Length == 1 &&
                _grid.SelectedCells.Count > 1)
            {
                string fillValue = sourceCells[0][0];
                foreach (DataGridViewCell target in _grid.SelectedCells)
                    TrySetPastedCell(target, fillValue, showError: false);
                return;
            }

            /* ───── find insertion start (first selected, else [row0,col1]) ───── */
            var startCell = _grid.SelectedCells.Count > 0
                ? _grid.SelectedCells.Cast<DataGridViewCell>()
                       .OrderBy(c => c.RowIndex).ThenBy(c => c.ColumnIndex).First()
                : _grid[1, 0];               // skip read-only “#” column

            int startRow = startCell.RowIndex;
            int startCol = startCell.ColumnIndex;

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
                    _grid.Rows.Add(_grid.Rows.Count + 1, 0.0, 0.0, "", "");
                }
            }

            for (int i = 0; i < sourceCells.Length; i++)
            {
                EnsureRow(startRow + i);

                for (int j = 0; j < sourceCells[i].Length; j++)
                {
                    int col = startCol + j;
                    if (col >= _grid.ColumnCount) break;      // ignore overflow
                    if (_grid.Columns[col].ReadOnly) continue; // e.g. “#”
                    TrySetPastedCell(
                        _grid[col, startRow + i],
                        sourceCells[i][j],
                        showError: true);
                }
            }
        }

        private bool TrySetPastedCell(
            DataGridViewCell cell,
            string value,
            bool showError)
        {
            if (cell == null || cell.ReadOnly) return false;
            string columnName = cell.OwningColumn.Name;

            if (columnName == "Northing" || columnName == "Easting")
            {
                if (value.IndexOf(',') >= 0 ||
                    !HybridCommands.TryParseCoordinate(value, out var coordinate))
                {
                    if (showError)
                        MessageBox.Show(
                            "Coordinate values must be finite numbers using a dot (.) decimal separator.",
                            "Invalid coordinate",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                    return false;
                }

                cell.Value = coordinate;
                return true;
            }

            if (columnName == "Type")
            {
                var typeColumn = (DataGridViewComboBoxColumn)_grid.Columns["Type"];
                if (!typeColumn.Items.Contains(value)) typeColumn.Items.Add(value);
            }

            cell.Value = value;
            return true;
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
                    HybridCommands.StoreCoordinate(v.N),
                    HybridCommands.StoreCoordinate(v.E),
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

