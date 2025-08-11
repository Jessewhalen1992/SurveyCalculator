// SurveyCalculator - ALS workflow commands for AutoCAD Map 3D (.NET 4.8)
// Commands: PLANLINK, PLANEXTRACT, EVILINK, ALSADJ
// Build x64; refs: acdbmgd.dll, acmgd.dll, ManagedMapApi.dll (Copy Local = false)
// Framework refs used: System.Windows.Forms, System.Web.Extensions (JSON)

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Windows.Forms;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(SurveyCalculator.Commands))]

namespace SurveyCalculator
{
    // ---------------------------- Config & Helpers ----------------------------
    internal static class Config
    {
        public const string LayerPdf = "X_PDF_PLAN";
        public const string LayerAdjusted = "BOUNDARY_ADJ";
        public const string LayerResiduals = "QA_RESIDUALS";

        // Tight tolerances (m)
        public const double ClosureWarningMeters = 0.03;
        public const double ResidualGreen = 0.015;
        public const double ResidualYellow = 0.02;
        public const double ResidualArrowMinLen = 0.01;
        public const double ResidualArrowHead = 0.25;

        // EVILINK auto PlanID snap
        public const double NearestVertexMaxMeters = 5.0;

        public static string CurrentDwgFolder()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            string n = string.IsNullOrEmpty(doc?.Name) ? "" : doc.Name;
            return string.IsNullOrEmpty(n) ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop) : Path.GetDirectoryName(n);
        }
        public static string Stem()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            string n = string.IsNullOrEmpty(doc?.Name) ? "Drawing" : Path.GetFileNameWithoutExtension(doc.Name);
            return string.IsNullOrEmpty(n) ? "Drawing" : n;
        }
        public static string PlanJsonPath() => Path.Combine(CurrentDwgFolder(), $"{Stem()}_PlanGraph.json");
        public static string EvidenceJsonPath() => Path.Combine(CurrentDwgFolder(), $"{Stem()}_EvidenceLinks.json");
        public static string ReportCsvPath() => Path.Combine(CurrentDwgFolder(), $"{Stem()}_AdjustmentReport.csv");
    }

    internal static class Json
    {
        static readonly JavaScriptSerializer Ser = new JavaScriptSerializer() { MaxJsonLength = int.MaxValue };
        public static void Save<T>(string path, T obj) => File.WriteAllText(path, Ser.Serialize(obj), Encoding.UTF8);
        public static T Load<T>(string path) => new JavaScriptSerializer().Deserialize<T>(File.ReadAllText(path, Encoding.UTF8));
    }

    internal static class Cad
    {
        public static Editor Ed => Application.DocumentManager.MdiActiveDocument.Editor;
        public static Database Db => HostApplicationServices.WorkingDatabase;

        public static ObjectId EnsureLayer(string name, short aci = 7, bool lockAfter = false)
        {
            ObjectId id;
            using (var tr = Db.TransactionManager.StartTransaction())
            {
                var lt = (LayerTable)tr.GetObject(Db.LayerTableId, OpenMode.ForRead);
                if (!lt.Has(name))
                {
                    lt.UpgradeOpen();
                    var rec = new LayerTableRecord { Name = name, Color = Color.FromColorIndex(ColorMethod.ByAci, aci) };
                    id = lt.Add(rec); tr.AddNewlyCreatedDBObject(rec, true);
                }
                else id = lt[name];
                tr.Commit();
            }
            if (lockAfter) LockLayer(name, true);
            return id;
        }
        public static void LockLayer(string name, bool locked)
        {
            using var tr = Db.TransactionManager.StartTransaction();
            var lt = (LayerTable)tr.GetObject(Db.LayerTableId, OpenMode.ForRead);
            if (!lt.Has(name)) return;
            var rec = (LayerTableRecord)tr.GetObject(lt[name], OpenMode.ForRead);
            rec.UpgradeOpen(); rec.IsLocked = locked; tr.Commit();
        }
        public static ObjectId AddToModelSpace(Entity e)
        {
            using var tr = Db.TransactionManager.StartTransaction();
            var bt = (BlockTable)tr.GetObject(Db.BlockTableId, OpenMode.ForRead);
            var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
            var id = ms.AppendEntity(e); tr.AddNewlyCreatedDBObject(e, true); tr.Commit(); return id;
        }
        public static void TransformEntities(IEnumerable<ObjectId> ids, Matrix3d m)
        {
            using var tr = Db.TransactionManager.StartTransaction();
            foreach (var id in ids) ((Entity)tr.GetObject(id, OpenMode.ForWrite)).TransformBy(m);
            tr.Commit();
        }
        public static void SetLayer(IEnumerable<ObjectId> ids, string layer)
        {
            EnsureLayer(layer);
            using var tr = Db.TransactionManager.StartTransaction();
            foreach (var id in ids) ((Entity)tr.GetObject(id, OpenMode.ForWrite)).Layer = layer;
            tr.Commit();
        }
        public static PromptSelectionResult PromptSelect(string msg, params TypedValue[] filter)
        {
            var opts = new PromptSelectionOptions { MessageForAdding = "\n" + msg + ": " };
            return (filter != null && filter.Length > 0)
                ? Ed.GetSelection(opts, new SelectionFilter(filter))
                : Ed.GetSelection(opts);
        }
        public static PromptPointResult GetPoint(string msg) => Ed.GetPoint("\n" + msg + ": ");
        public static double Distance(Point2d a, Point2d b) { var dx = a.X - b.X; var dy = a.Y - b.Y; return Math.Sqrt(dx * dx + dy * dy); }
        public static double ToAzimuthRad(Point2d a, Point2d b)
        {
            double th = Math.Atan2(b.Y - a.Y, b.X - a.X); if (th < 0) th += 2 * Math.PI; return th;
        }
        public static string FormatBearing(double az)
        {
            az %= 2 * Math.PI; if (az < 0) az += 2 * Math.PI;
            bool north = az <= Math.PI / 2 || az >= 3 * Math.PI / 2; bool east = az < Math.PI;
            double theta = north ? (east ? az : Math.PI - az) : (east ? 2 * Math.PI - az : az - Math.PI);
            double deg = theta * 180.0 / Math.PI; int d = (int)Math.Floor(deg + 1e-9);
            double remM = (deg - d) * 60.0; int m = (int)Math.Floor(remM + 1e-9); double s = (remM - m) * 60.0;
            return $"{(north ? "N" : "S")} {d:00}°{m:00}'{s:00.##}" {(east ? "E" : "W")}";
        }
    }

    // ---------------------------- Data Models ----------------------------
    [Serializable] public class XY { public double X; public double Y; public XY() { } public XY(double x, double y){X=x;Y=y;} }

    [Serializable]
    public class PlanEdge { public string FromId; public string ToId; public double Distance; public double BearingRad; }

    [Serializable]
    public class PlanData
    {
        public List<string> VertexIds = new List<string>();  // P1..Pn
        public List<XY> VertexXY = new List<XY>();           // actual polyline XY for nearest-vertex mapping
        public List<PlanEdge> Edges = new List<PlanEdge>();  // distances + bearings (geometry or plan numbers)
        public bool Closed = true;
        public int Count => VertexIds.Count;
        public int IndexOf(string id) => VertexIds.FindIndex(v => string.Equals(v, id, StringComparison.OrdinalIgnoreCase));
        public PlanEdge EdgeAt(int i) => Edges[i];
    }

    [Serializable]
    public class EvidencePoint
    {
        public string PlanId = "";      // e.g., "P7"
        public string Handle = "";      // drawing handle
        public double X; public double Y;
        public string EvidenceType = ""; // block name
        public bool Held = false;
        public int Priority = 0;        // reserved
        public Point2d XY() => new Point2d(X, Y);
    }

    [Serializable] public class EvidenceLinks { public List<EvidencePoint> Points = new List<EvidencePoint>(); }

    // ---------------------------- Parsing ----------------------------
    internal static class BearingParser
    {
        static readonly Regex Dms = new Regex(@"([NS])\s*(\d{1,3})°\s*(\d{1,2})?['’]?\s*(\d{1,2}(\.\d+)?)?[""”]?\s*([EW])",
                                              RegexOptions.Compiled | RegexOptions.IgnoreCase);
        public static bool TryParse(string text, out double az)
        {
            az = 0; if (string.IsNullOrWhiteSpace(text)) return false;
            text = text.Replace('’', '\'').Replace('″', '"');
            var m = Dms.Match(text); if (!m.Success) return false;
            int d = int.Parse(m.Groups[2].Value);
            int min = string.IsNullOrEmpty(m.Groups[3].Value) ? 0 : int.Parse(m.Groups[3].Value);
            double sec = 0; if (!string.IsNullOrEmpty(m.Groups[4].Value)) double.TryParse(m.Groups[4].Value, out sec);
            double ang = (d + min / 60.0 + sec / 3600.0) * Math.PI / 180.0;
            bool north = char.ToUpper(m.Groups[1].Value[0]) == 'N'; bool east = char.ToUpper(m.Groups[6].Value[0]) == 'E';
            az = north ? (east ? ang : Math.PI - ang) : (east ? 2 * Math.PI - ang : ang - Math.PI);
            return true;
        }
    }

    // ---------------------------- Math (Similarity + Traverse) ----------------------------
    internal static class SimilarityFit
    {
        public static bool Solve(IReadOnlyList<(Point2d from, Point2d to, double w)> pairs,
                                 bool lockScale, out double k, out double c, out double s, out Vector2d t)
        {
            k = 1; c = 1; s = 0; t = new Vector2d(0, 0);
            if (pairs == null || pairs.Count < 2) return false;
            double W = 0, axs = 0, ays = 0, bxs = 0, bys = 0;
            foreach (var (a, b, w) in pairs){ W += w; axs += w * a.X; ays += w * a.Y; bxs += w * b.X; bys += w * b.Y; }
            var ca = new Point2d(axs / W, ays / W); var cb = new Point2d(bxs / W, bys / W);

            double Sxx=0,Sxy=0,Syx=0,Syy=0,S1=0;
            foreach (var (a, b, w) in pairs)
            {
                var A = a - ca; var B = b - cb;
                Sxx += w * A.X * B.X; Sxy += w * A.X * B.Y;
                Syx += w * A.Y * B.X; Syy += w * A.Y * B.Y;
                S1  += w * (A.X * A.X + A.Y * A.Y);
            }
            double A1 = Sxx + Syy; double B1 = Sxy - Syx; double norm = Math.Sqrt(A1 * A1 + B1 * B1);
            if (norm < 1e-12) return false;
            c = A1 / norm; s = B1 / norm;

            double kNum = (Sxx * c + Sxy * s) + (Syy * c - Syx * s);
            k = lockScale ? 1.0 : Math.Max(1e-12, kNum / S1);

            t = new Vector2d(
                cb.X - (k * (c * ca.X - s * ca.Y)),
                cb.Y - (k * (s * ca.X + c * ca.Y))
            );
            return true;
        }
        public static Matrix3d ToMatrix(double k, double c, double s, Vector2d t)
        {
            var m = Matrix3d.Scaling(k, Point3d.Origin);
            m = Matrix3d.Rotation(Math.Atan2(s, c), Vector3d.ZAxis, Point3d.Origin) * m;
            m = Matrix3d.Displacement(new Vector3d(t.X, t.Y, 0)) * m;
            return m;
        }
    }

    internal static class TraverseSolver
    {
        // Compass‑rule closure along a segment chain
        public static List<Point2d> DriveBetween(Point2d startHeld, Point2d endHeld,
            IReadOnlyList<PlanEdge> chain, double closureTol, out double closure)
        {
            var raw = new List<Point2d> { startHeld };
            var deltas = new List<Vector2d>(); double sumL = 0;
            foreach (var e in chain)
            {
                double dx = e.Distance * Math.Cos(e.BearingRad);
                double dy = e.Distance * Math.Sin(e.BearingRad);
                deltas.Add(new Vector2d(dx, dy)); sumL += e.Distance;
                raw.Add(new Point2d(raw[^1].X + dx, raw[^1].Y + dy));
            }
            double Cx = endHeld.X - raw[^1].X; double Cy = endHeld.Y - raw[^1].Y;
            closure = Math.Sqrt(Cx * Cx + Cy * Cy);
            var closed = new List<Point2d> { startHeld };
            for (int i = 0; i < deltas.Count; i++)
            {
                double f = chain[i].Distance / sumL;
                var adj = new Vector2d(deltas[i].X + f * Cx, deltas[i].Y + f * Cy);
                closed.Add(new Point2d(closed[^1].X + adj.X, closed[^1].Y + adj.Y));
            }
            return closed;
        }
    }

    // ---------------------------- EVILINK Form ----------------------------
    public class EvilinkForm : Form
    {
        public EvidenceLinks Links { get; private set; }
        DataGridView grid = new DataGridView();
        Button btnSave = new Button();
        Button btnCancel = new Button();
        Label lbl = new Label();

        public EvilinkForm(EvidenceLinks links)
        {
            Text = "Evidence ↔ Plan Linker (EVILINK)";
            Width = 820; Height = 520; StartPosition = FormStartPosition.CenterScreen;
            Links = links;

            grid.Dock = DockStyle.Top; grid.Height = 400; grid.AutoGenerateColumns = false;
            grid.AllowUserToAddRows = false; grid.AllowUserToDeleteRows = false; grid.DataSource = Links.Points;

            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Handle", DataPropertyName = "Handle", ReadOnly = true, Width = 120 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "X", DataPropertyName = "X", ReadOnly = true, Width = 120 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Y", DataPropertyName = "Y", ReadOnly = true, Width = 120 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "PlanID (P#)", DataPropertyName = "PlanId", Width = 100 });
            grid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "Held", DataPropertyName = "Held", Width = 60 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "EvidenceType (block)", DataPropertyName = "EvidenceType", ReadOnly = true, Width = 180 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Priority", DataPropertyName = "Priority", Width = 60 });

            btnSave.Text = "Save"; btnSave.Width = 100; btnSave.Top = 430; btnSave.Left = 600;
            btnSave.Click += (s, e) => { DialogResult = DialogResult.OK; Close(); };
            btnCancel.Text = "Cancel"; btnCancel.Width = 100; btnCancel.Top = 430; btnCancel.Left = 710;
            btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

            lbl.Text = "Tip: Auto PlanID from nearest plan vertex. FDI auto‑Held, fdspike not held by default.";
            lbl.Top = 405; lbl.Left = 10; lbl.Width = 560;

            Controls.Add(grid); Controls.Add(btnSave); Controls.Add(btnCancel); Controls.Add(lbl);
        }
    }

    // ---------------------------- Commands ----------------------------
    public class Commands
    {
        // -------- PLANLINK --------
        [CommandMethod("PLANLINK")]
        public void PlanLink()
        {
            var ed = Cad.Ed;
            ed.WriteMessage("\nPLANLINK — align imported plan entities to field control (similarity fit).");

            var psr = Cad.PromptSelect("Select imported plan entities (from PDFIMPORT)");
            if (psr.Status != PromptStatus.OK) return;
            var ids = psr.Value.GetObjectIds();
            if (ids == null || ids.Length == 0) { ed.WriteMessage("\nNothing selected."); return; }

            Cad.EnsureLayer(Config.LayerPdf, aci: 4);
            Cad.SetLayer(ids, Config.LayerPdf);

            var pairs = new List<(Point2d from, Point2d to, double w)>();
            for (int i = 1; i <= 5; i++)
            {
                var p1 = Cad.GetPoint($"Pick SOURCE point ON PLAN (pair {i}, Enter to finish)");
                if (p1.Status != PromptStatus.OK)
                { if (pairs.Count >= 2) break; ed.WriteMessage("\nNeed at least 2 pairs."); return; }
                var p2 = Cad.GetPoint($"Pick MATCHING FIELD point (target for pair {i})");
                if (p2.Status != PromptStatus.OK) { ed.WriteMessage("\nCancelled."); return; }
                pairs.Add((new Point2d(p1.Value.X, p1.Value.Y), new Point2d(p2.Value.X, p2.Value.Y), 1.0));
            }

            // Free-scale to detect unit mismatch and potential improvement
            if (!SimilarityFit.Solve(pairs, false, out double kf, out double cf, out double sf, out Vector2d tf))
            { ed.WriteMessage("\nCould not solve similarity transform (check control pairs)."); return; }

            // Decide with ALSGuard
            var (k, c, s, t, usedScale) = DecideScalingALS(pairs, kf, cf, sf, tf, ed);
            var m = SimilarityFit.ToMatrix(k, c, s, t);
            Cad.TransformEntities(ids, m);

            Cad.LockLayer(Config.LayerPdf, true);
            ed.WriteMessage(usedScale
                ? $"\nPLANLINK done (scale applied). k={k:0.000000}, rot={Math.Atan2(s,c)*180/Math.PI:0.000}°, t=({t.X:0.###},{t.Y:0.###}). Layer {Config.LayerPdf} locked."
                : $"\nPLANLINK done (scale locked to 1.0). rot={Math.Atan2(s,c)*180/Math.PI:0.000}°, t=({t.X:0.###},{t.Y:0.###}). Layer {Config.LayerPdf} locked.");
        }

        private static double ComputeRms(IReadOnlyList<(Point2d from, Point2d to, double w)> pairs,
                                         double k, double c, double s, Vector2d t)
        {
            double ss = 0, W = 0;
            foreach (var (a, b, w) in pairs)
            {
                double x = k * (c * a.X - s * a.Y) + t.X;
                double y = k * (s * a.X + c * a.Y) + t.Y;
                double dx = x - b.X, dy = y - b.Y;
                ss += w * (dx * dx + dy * dy); W += w;
            }
            return Math.Sqrt(ss / Math.Max(1e-12, W));
        }

        private (double k, double c, double s, Vector2d t, bool scaled) DecideScalingALS(
            List<(Point2d from, Point2d to, double w)> pairs, double kFree, double cFree, double sFree, Vector2d tFree, Editor ed)
        {
            // Locked-scale fit for comparison
            SimilarityFit.Solve(pairs, true, out double kl, out double cl, out double sl, out Vector2d tl);

            double rmsFree = ComputeRms(pairs, kFree, cFree, sFree, tFree);
            double rmsLock = ComputeRms(pairs, 1.0, cl, sl, tl);

            bool unitsMismatch = Math.Abs(kFree - 0.3048) < 0.02 || Math.Abs(kFree - 3.28084) < 0.05;

            if (unitsMismatch)
            {
                var opt = new PromptKeywordOptions($"\nALSGuard: scale {kFree:0.######} looks like FEET↔METRES. Apply scaling? [Yes/No] <Yes>: ");
                opt.Keywords.Add("Yes"); opt.Keywords.Add("No"); opt.AllowNone = true;
                var ans = ed.GetKeywords(opt); bool yes = !(ans.Status == PromptStatus.OK && ans.StringResult == "No");
                return yes ? (kFree, cFree, sFree, tFree, true) : (1, cl, sl, tl, false);
            }

            bool smallScale = Math.Abs(kFree - 1.0) <= 0.001;
            bool enoughHeld = pairs.Count >= 3;
            bool bigImprovement = rmsFree < 0.5 * rmsLock && rmsLock > Config.ResidualYellow;

            if (enoughHeld && smallScale && bigImprovement)
            {
                var opt = new PromptKeywordOptions(
                    $"\nALSGuard: free-scale RMS {rmsFree:0.###} m vs locked {rmsLock:0.###} m with tiny scale Δ={Math.Abs(kFree - 1):0.#####}. Apply scaling? [Yes/No] <No>: ");
                opt.Keywords.Add("Yes"); opt.Keywords.Add("No"); opt.AllowNone = true;
                var ans = ed.GetKeywords(opt); bool yes = (ans.Status == PromptStatus.OK && ans.StringResult == "Yes");
                return yes ? (kFree, cFree, sFree, tFree, true) : (1, cl, sl, tl, false);
            }
            return (1, cl, sl, tl, false);
        }

        // -------- PLANEXTRACT --------
        [CommandMethod("PLANEXTRACT")]
        public void PlanExtract()
        {
            var ed = Cad.Ed;
            ed.WriteMessage("\nPLANEXTRACT — select the plan boundary polyline on layer X_PDF_PLAN.");

            var peo = new PromptEntityOptions("\nSelect plan boundary polyline: ");
            peo.SetRejectMessage("\nOnly Polyline entities are supported.");
            peo.AddAllowedClass(typeof(Polyline), true);
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            var plan = new PlanData();
            using (var tr = Cad.Db.TransactionManager.StartTransaction())
            {
                var pl = (Polyline)tr.GetObject(per.ObjectId, OpenMode.ForRead);
                if (!pl.Closed) ed.WriteMessage("\nWarning: selected polyline is not closed; proceeding.");
                int n = pl.NumberOfVertices;
                if (n < 3) { ed.WriteMessage("\nPolyline has <3 vertices; aborting."); return; }

                var verts = new List<Point2d>(n);
                for (int i = 0; i < n; i++) verts.Add(pl.GetPoint2dAt(i));

                for (int i = 0; i < n; i++)
                {
                    string vid = $"P{i + 1}";
                    plan.VertexIds.Add(vid);
                    plan.VertexXY.Add(new XY(verts[i].X, verts[i].Y));
                }
                for (int i = 0; i < n; i++)
                {
                    int j = (i + 1) % n;
                    double dist = Cad.Distance(verts[i], verts[j]);
                    double az = Cad.ToAzimuthRad(verts[i], verts[j]);
                    plan.Edges.Add(new PlanEdge { FromId = plan.VertexIds[i], ToId = plan.VertexIds[j], Distance = dist, BearingRad = az });
                }
                tr.Commit();
            }

            var path = Config.PlanJsonPath();
            Json.Save(path, plan);
            ed.WriteMessage($"\nWrote plan to: {path}");
            DrawPlanSummaryMText(plan);
        }

        private static void DrawPlanSummaryMText(PlanData plan)
        {
            var sb = new StringBuilder();
            sb.AppendLine("\\LPlan Extract Summary\\l");
            sb.AppendLine($"Vertices: {plan.Count}");
            sb.AppendLine($"Perimeter (geom): {plan.Edges.Sum(e => e.Distance):0.###} m");
            Cad.EnsureLayer(Config.LayerPdf, aci: 4);
            var mt = new MText { Contents = sb.ToString(), Location = new Point3d(0, 0, 0), TextHeight = 2.5, Layer = Config.LayerPdf };
            Cad.AddToModelSpace(mt);
        }

        // -------- EVILINK --------
        [CommandMethod("EVILINK")]
        public void Evilink()
        {
            var ed = Cad.Ed;
            ed.WriteMessage("\nEVILINK — select evidence blocks (no attributes needed).");

            string planPath = Config.PlanJsonPath();
            if (!File.Exists(planPath)) { ed.WriteMessage($"\nPlan JSON not found: {planPath}\nRun PLANEXTRACT first."); return; }
            var plan = Json.Load<PlanData>(planPath);
            if (plan.VertexXY == null || plan.VertexXY.Count != plan.Count) { ed.WriteMessage("\nPlan JSON missing VertexXY; re-run PLANEXTRACT."); return; }

            var tv = new TypedValue[] { new TypedValue((int)DxfCode.Start, "INSERT") };
            var psr = Cad.PromptSelect("Select evidence blocks", tv);
            if (psr.Status != PromptStatus.OK) return;

            var links = new EvidenceLinks();
            using (var tr = Cad.Db.TransactionManager.StartTransaction())
            {
                foreach (var id in psr.Value.GetObjectIds())
                {
                    var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference; if (br == null) continue;
                    string blockName = GetBlockName(br, tr);
                    var evXY = new Point2d(br.Position.X, br.Position.Y);
                    string planId = NearestPlanId(evXY, plan, Config.NearestVertexMaxMeters);

                    bool strong = string.Equals(blockName, "FDI", StringComparison.OrdinalIgnoreCase);
                    bool weak   = string.Equals(blockName, "fdspike", StringComparison.OrdinalIgnoreCase);
                    int pri = strong ? 2 : (weak ? 0 : 1);
                    bool heldDefault = strong;

                    links.Points.Add(new EvidencePoint
                    {
                        PlanId = planId,
                        Handle = br.Handle.ToString(),
                        X = evXY.X, Y = evXY.Y,
                        EvidenceType = blockName,
                        Held = heldDefault,
                        Priority = pri
                    });
                }
                tr.Commit();
            }

            var frm = new EvilinkForm(links);
            var res = Application.ShowModalDialog(frm);
            if (res == DialogResult.OK)
            {
                var path = Config.EvidenceJsonPath();
                Json.Save(path, links);
                ed.WriteMessage($"\nSaved links: {path}");
            }
            else ed.WriteMessage("\nCancelled (no changes saved).");
        }

        private static string GetBlockName(BlockReference br, Transaction tr)
        {
            ObjectId btrId = br.IsDynamicBlock ? br.DynamicBlockTableRecord : br.BlockTableRecord;
            var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
            return btr.Name;
        }
        private static string NearestPlanId(Point2d p, PlanData plan, double maxDist)
        {
            int bestIdx = -1; double best = double.MaxValue;
            for (int i = 0; i < plan.Count; i++)
            {
                var v = plan.VertexXY[i]; var q = new Point2d(v.X, v.Y);
                double d = Cad.Distance(p, q); if (d < best) { best = d; bestIdx = i; }
            }
            return (bestIdx >= 0 && best <= maxDist) ? plan.VertexIds[bestIdx] : "";
        }

        // -------- ALSADJ --------
        [CommandMethod("ALSADJ")]
        public void AlsAdj()
        {
            var ed = Cad.Ed;
            ed.WriteMessage("\nALSADJ — hold anchors, re-drive between anchors by plan numbers with compass closure.");

            string planPath = Config.PlanJsonPath();
            if (!File.Exists(planPath)) { ed.WriteMessage($"\nMissing: {planPath}. Run PLANEXTRACT."); return; }
            var plan = Json.Load<PlanData>(planPath);
            if (plan.Count < 3 || plan.Edges.Count < 3) { ed.WriteMessage("\nPlan data invalid."); return; }

            string evPath = Config.EvidenceJsonPath();
            if (!File.Exists(evPath)) { ed.WriteMessage($"\nMissing: {evPath}. Run EVILINK."); return; }
            var links = Json.Load<EvidenceLinks>(evPath);
            var held = links.Points.Where(p => p.Held && !string.IsNullOrWhiteSpace(p.PlanId)).ToList();
            if (held.Count < 2) { ed.WriteMessage("\nNeed at least two Held evidence points."); return; }

            // Build synthetic legal loop by traversing edges from origin
            var synth = TraverseWhole(plan);

            // Fit from HELD pairs only (ALSGuard included here too)
            var pairs = new List<(Point2d from, Point2d to, double w)>();
            var idToIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < plan.Count; i++) idToIndex[plan.VertexIds[i]] = i;

            foreach (var hp in held)
            {
                if (!idToIndex.TryGetValue(hp.PlanId, out int idx)) { ed.WriteMessage($"\nHeld PlanID {hp.PlanId} not in plan; skipping."); continue; }
                pairs.Add((synth[idx], hp.XY(), 1.0));
            }
            if (pairs.Count < 2) { ed.WriteMessage("\nInsufficient valid Held pairs."); return; }

            // Free-scale; then ALSGuard decision
            if (!SimilarityFit.Solve(pairs, false, out double kf, out double cf, out double sf, out Vector2d tf))
            { ed.WriteMessage("\nCould not solve similarity transform from Held pairs."); return; }
            var (k, c, s, t, usedScale) = DecideScalingALS(pairs, kf, cf, sf, tf, ed);

            // Transform entire synthetic loop to drawing frame
            var planXY = new List<Point2d>(synth.Count);
            for (int i = 0; i < synth.Count; i++)
            {
                var p = synth[i];
                double x = k * (c * p.X - s * p.Y) + t.X;
                double y = k * (s * p.X + c * p.Y) + t.Y;
                planXY.Add(new Point2d(x, y));
            }

            // Prepare adjusted coords; seed held directly from evidence
            var adj = new Point2d[plan.Count];
            foreach (var h in held) if (idToIndex.TryGetValue(h.PlanId, out int idx)) adj[idx] = h.XY();

            // Walk arcs between consecutive held vertices (wrap)
            var heldIdx = held.Where(h => idToIndex.ContainsKey(h.PlanId))
                              .Select(h => idToIndex[h.PlanId])
                              .Distinct().OrderBy(i => i).ToList();
            if (!plan.Closed) { ed.WriteMessage("\nPlan not closed; expected closed loop."); return; }

            for (int hIdx = 0; hIdx < heldIdx.Count; hIdx++)
            {
                int start = heldIdx[hIdx];
                int end = heldIdx[(hIdx + 1) % heldIdx.Count];

                var chain = new List<PlanEdge>(); int i = start;
                while (i != end) { chain.Add(plan.EdgeAt(i)); i = (i + 1) % plan.Count; }

                var startXY = adj[start]; var endXY = adj[end];
                double closure;
                var drove = TraverseSolver.DriveBetween(startXY, endXY, chain, Config.ClosureWarningMeters, out closure);
                if (closure > Config.ClosureWarningMeters)
                    Cad.Ed.WriteMessage($"\nWarning: closure {closure:0.###} m between {plan.VertexIds[start]} and {plan.VertexIds[end]} exceeds {Config.ClosureWarningMeters:0.###} m.");

                int kidx = 0; int j = start;
                while (j != end) { adj[j] = drove[kidx]; kidx++; j = (j + 1) % plan.Count; }
                adj[end] = drove[kidx];
            }

            // Build adjusted polyline
            Cad.EnsureLayer(Config.LayerAdjusted, aci: 2);
            var plAdj = new Polyline { Layer = Config.LayerAdjusted, Closed = true };
            for (int i = 0; i < plan.Count; i++) plAdj.AddVertexAt(i, adj[i], 0, 0, 0);
            Cad.AddToModelSpace(plAdj);

            // Residual arrows
            Cad.EnsureLayer(Config.LayerResiduals, aci: 1);
            int residualCount = 0;
            var dictEv = links.Points.Where(p => !string.IsNullOrWhiteSpace(p.PlanId))
                                     .GroupBy(p => p.PlanId, StringComparer.OrdinalIgnoreCase)
                                     .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < plan.Count; i++)
            {
                string id = plan.VertexIds[i];
                if (!dictEv.TryGetValue(id, out var ev)) continue;
                var evPt = ev.XY(); var adPt = adj[i]; double d = Cad.Distance(evPt, adPt);
                if (d < Config.ResidualArrowMinLen) continue; residualCount++; DrawResidualArrow(evPt, adPt, d);
            }

            // CSV
            WriteCsvReport(plan, links, adj, k, Math.Atan2(s, c));

            ed.WriteMessage($"\nALSADJ done. {(usedScale ? "Scale applied." : "Scale locked.")} Adjusted boundary on {Config.LayerAdjusted}. Residuals on {Config.LayerResiduals} ({residualCount}).\nCSV: {Config.ReportCsvPath()}");
        }

        private static List<Point2d> TraverseWhole(PlanData plan)
        {
            var pts = new List<Point2d>(plan.Count);
            var p = new Point2d(0, 0); pts.Add(p);
            for (int i = 0; i < plan.Edges.Count; i++)
            {
                var e = plan.Edges[i];
                double dx = e.Distance * Math.Cos(e.BearingRad);
                double dy = e.Distance * Math.Sin(e.BearingRad);
                p = new Point2d(p.X + dx, p.Y + dy);
                if (i < plan.Count - 1) pts.Add(p);
            }
            if (pts.Count != plan.Count) { pts = new List<Point2d>(plan.Count); for (int i = 0; i < plan.Count; i++) pts.Add(new Point2d(i, 0)); }
            return pts;
        }

        private static void DrawResidualArrow(Point2d fromEv, Point2d toAdj, double len)
        {
            short aci = (len < Config.ResidualGreen) ? (short)3 : (len < Config.ResidualYellow) ? (short)2 : (short)1;
            var color = Color.FromColorIndex(ColorMethod.ByAci, aci);

            var ln = new Line(new Point3d(fromEv.X, fromEv.Y, 0), new Point3d(toAdj.X, toAdj.Y, 0)) { Layer = Config.LayerResiduals, Color = color };
            Cad.AddToModelSpace(ln);

            var v = new Vector2d(fromEv.X - toAdj.X, fromEv.Y - toAdj.Y);
            double vlen = Math.Sqrt(v.X * v.X + v.Y * v.Y); if (vlen < 1e-6) return;
            var dir = new Vector2d(v.X / vlen, v.Y / vlen);
            double ah = Math.Min(Config.ResidualArrowHead, len * 0.2);
            double ang = 20.0 * Math.PI / 180.0;
            var r1 = new Vector2d(Math.Cos(ang) * dir.X - Math.Sin(ang) * dir.Y, Math.Sin(ang) * dir.X + Math.Cos(ang) * dir.Y);
            var r2 = new Vector2d(Math.Cos(-ang) * dir.X - Math.Sin(-ang) * dir.Y, Math.Sin(-ang) * dir.X + Math.Cos(-ang) * dir.Y);
            var a1 = new Point2d(toAdj.X + ah * r1.X, toAdj.Y + ah * r1.Y);
            var a2 = new Point2d(toAdj.X + ah * r2.X, toAdj.Y + ah * r2.Y);
            Cad.AddToModelSpace(new Line(new Point3d(toAdj.X, toAdj.Y, 0), new Point3d(a1.X, a1.Y, 0)) { Layer = Config.LayerResiduals, Color = color });
            Cad.AddToModelSpace(new Line(new Point3d(toAdj.X, toAdj.Y, 0), new Point3d(a2.X, a2.Y, 0)) { Layer = Config.LayerResiduals, Color = color });
        }

        private static void WriteCsvReport(PlanData plan, EvidenceLinks links, Point2d[] adj, double scale, double rotRad)
        {
            var dictEv = links.Points.Where(p => !string.IsNullOrWhiteSpace(p.PlanId))
                                     .GroupBy(p => p.PlanId, StringComparer.OrdinalIgnoreCase)
                                     .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
            var sb = new StringBuilder();
            sb.AppendLine("PlanID,Held,Field_X,Field_Y,Adj_X,Adj_Y,Residual_m,EvidenceType,Seg_Distance_m,Seg_Bearing,Scale,Rotation_deg");
            for (int i = 0; i < plan.Count; i++)
            {
                string id = plan.VertexIds[i]; var a = adj[i];
                bool held = false; double fx = double.NaN, fy = double.NaN, resid = double.NaN; string evType = "";
                if (dictEv.TryGetValue(id, out var ev)) { held = ev.Held; fx = ev.X; fy = ev.Y; resid = Math.Sqrt((fx - a.X) * (fx - a.X) + (fy - a.Y) * (fy - a.Y)); evType = ev.EvidenceType ?? ""; }
                var e = plan.EdgeAt(i); string bearing = Cad.FormatBearing(e.BearingRad);
                sb.AppendLine($"{id},{(held ? "Y" : "N")},{(double.IsNaN(fx) ? "" : fx.ToString("0.###"))},{(double.IsNaN(fy) ? "" : fy.ToString("0.###"))},{a.X:0.###},{a.Y:0.###},{(double.IsNaN(resid) ? "" : resid.ToString("0.###"))},{evType},{e.Distance:0.###},{bearing},{scale:0.000000},{(rotRad*180.0/Math.PI):0.000}");
            }
            File.WriteAllText(Config.ReportCsvPath(), sb.ToString(), Encoding.UTF8);
        }
    }
}
