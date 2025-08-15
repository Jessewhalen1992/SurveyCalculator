// SurveyCalculator - ALS workflow commands for AutoCAD Map 3D (.NET Framework 4.8)
// Commands: PLANLINK, PLANEXTRACT, EVILINK, ALSADJ, PLANCOGO, PLANCOGOUI, ALSWIZARD, PLANFROMTEXT, PLANLOCK, EVIHARVEST
// Build x64; refs: acdbmgd.dll, acmgd.dll, ManagedMapApi.dll (Copy Local = false)
// Framework refs used: System.Windows.Forms

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using SurveyCalculator;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

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
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            string n = string.IsNullOrEmpty(doc?.Name) ? "" : doc.Name;
            return string.IsNullOrEmpty(n) ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop) : Path.GetDirectoryName(n);
        }
        public static string Stem()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            string n = string.IsNullOrEmpty(doc?.Name) ? "Drawing" : Path.GetFileNameWithoutExtension(doc.Name);
            return string.IsNullOrEmpty(n) ? "Drawing" : n;
        }
        public static string PlanJsonPath()
        {
            string folder = CurrentDwgFolder();
            string primary = Path.Combine(folder, $"{Stem()}_PlanGraph.json");
            string[] fallbacks =
            {
                Path.Combine(folder, "COGO_PlanGraph.json"),
                Path.Combine(folder, "COGO.json"),
                Path.Combine(folder, "PlanGraph.json")
            };
            foreach (var fb in fallbacks)
                if (File.Exists(fb))
                    return fb;
            return primary;
        }

        public static string EvidenceJsonPath()
        {
            string folder = CurrentDwgFolder();
            string primary = Path.Combine(folder, $"{Stem()}_EvidenceLinks.json");
            string[] fallbacks =
            {
                Path.Combine(folder, "COGO_EvidenceLinks.json"),
                Path.Combine(folder, "EvidenceLinks.json")
            };
            foreach (var fb in fallbacks)
                if (File.Exists(fb))
                    return fb;
            return primary;
        }
        public static string ReportCsvPath() => Path.Combine(CurrentDwgFolder(), $"{Stem()}_AdjustmentReport.csv");
    }
    internal class UserCogoInput
    {
        public int Line;
        public string Point1 = "";
        public string Point2 = "";
        public string Bearing = "";
        public double Distance;
        public string Azimuth = "";
        public string Deflection = "";
    }

    internal static class CogoFile
    {
        public static void SaveUserCogoInput(string filePath, List<UserCogoInput> entries)
        {
            using (var sw = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                sw.WriteLine("Line,Point1,Point2,Bearing,Distance,Azimuth,Deflection");
                foreach (var e in entries)
                {
                    sw.WriteLine(string.Format(CultureInfo.InvariantCulture,
                        "{0},{1},{2},{3},{4},{5},{6}",
                        e.Line,
                        e.Point1,
                        e.Point2,
                        e.Bearing,
                        e.Distance,
                        e.Azimuth ?? string.Empty,
                        e.Deflection ?? string.Empty));
                }
            }
        }

        public static List<UserCogoInput> LoadUserCogoInput(string filePath)
        {
            var list = new List<UserCogoInput>();
            if (!File.Exists(filePath)) return list;
            var lines = File.ReadAllLines(filePath);
            for (int i = 1; i < lines.Length; i++)
            {
                var parts = lines[i].Split(',');
                if (parts.Length < 7) continue;
                double dist;
                double.TryParse(parts[4], NumberStyles.Float, CultureInfo.InvariantCulture, out dist);
                int lineNo;
                int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out lineNo);
                list.Add(new UserCogoInput
                {
                    Line = lineNo,
                    Point1 = parts[1],
                    Point2 = parts[2],
                    Bearing = parts[3],
                    Distance = dist,
                    Azimuth = parts[5],
                    Deflection = parts[6]
                });
            }
            return list;
        }
    }

    // ---------------------------- PLAN COGO UI (Form) ----------------------------
    public class PlanCogoForm : Form
    {
        private class CogoLegRow
        {
            public string Bearing { get; set; } = "";
            public double Distance { get; set; } = 0.0;
            public bool Locked { get; set; } = false;   // keep this leg's distance fixed during adjustment
        }
        private class VertexMapRow
        {
            public string VertexId { get; set; } = "";  // P1, P2, ...
            public int EvidenceNo { get; set; } = 0;    // 0 = (none), otherwise 1..N
            public bool Held { get; set; } = false;
        }
        // Plan name dropdown (allows typing a new name)
        private readonly ComboBox cboPlan = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDown, // user can type a new name
            Width = 320
        };
        private readonly Button btnNewPlan = new Button { Text = "New...", Width = 70 };

        private class ComboItem { public int Index { get; set; } public string Label { get; set; } = ""; }

        // ---------- Document ownership & safety ----------
        private readonly Document ownerDoc;
        private readonly DocumentCollection docCol;

        private Editor OwnerEd => ownerDoc.Editor;

        // ---------- Grids & UI ----------
        private readonly DataGridView gridLegs = new DataGridView();
        private readonly DataGridView gridVerts = new DataGridView();

        private readonly BindingList<CogoLegRow> legs = new BindingList<CogoLegRow>();
        private readonly BindingList<VertexMapRow> verts = new BindingList<VertexMapRow>();

        private readonly TextBox txtCsf = new TextBox { Width = 80, Text = "1.0" };
        private readonly CheckBox chkApplyCsf = new CheckBox { Text = "Apply CSF to distances on Save" };
        private readonly Button btnAddLeg = new Button { Text = "Add Leg" };
        private readonly Button btnDelLeg = new Button { Text = "Delete Selected" };
        private readonly Button btnCloseLeg = new Button { Text = "Compute Closing Leg" };
        private readonly Button btnRefreshVerts = new Button { Text = "Refresh Vertices" };

        private readonly Button btnSave = new Button { Text = "Save" };
        private readonly Button btnSaveAdj = new Button { Text = "Save & Adjust" };
        private readonly Button btnCancel = new Button { Text = "Cancel" };
        private readonly Button btnDraw = new Button { Text = "Draw" }; // draw preview at pick

        private readonly Label lblSummary = new Label { AutoSize = true };
        private readonly Label lblDocGuard = new Label { AutoSize = true, ForeColor = System.Drawing.Color.Firebrick };

        private EvidenceLinks links = new EvidenceLinks();
        private List<ComboItem> evidenceOptions = new List<ComboItem>();

        public bool RunAdjustAfterSave { get; private set; } = false;

        // --- Units + distance column (class-level so header can change) ---
        private readonly DataGridViewTextBoxColumn colDist =
            new DataGridViewTextBoxColumn
            {
                HeaderText = "Distance (m)",
                DataPropertyName = "Distance",
                Width = 180,
                DefaultCellStyle = { Format = "0.###" }
            };

        private enum InputUnits { Meters = 0, Feet = 1, Chains = 2 }
        private InputUnits currentUnits = InputUnits.Meters;

        private readonly ComboBox cboUnits = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 110 };
        private readonly Label lblUnits = new Label { AutoSize = true, Text = "Input Units:" };

        public PlanCogoForm(Document owner)
        {
            ownerDoc = owner ?? AcadApp.DocumentManager.MdiActiveDocument;
            docCol = AcadApp.DocumentManager;

            this.Text = "Plan COGO Editor";
            this.Width = 1000;
            this.Height = 680;
            this.StartPosition = FormStartPosition.CenterScreen;

            // Bind lifecycle to OWNER document
            HookDocumentLifetime();

            // Left: COGO legs
            var left = new GroupBox { Text = "COGO Legs", Left = 10, Top = 40, Width = 560, Height = 520 };
            gridLegs.Parent = left; gridLegs.Left = 10; gridLegs.Top = 50; gridLegs.Width = 540; gridLegs.Height = 420;
            gridLegs.AutoGenerateColumns = false; gridLegs.AllowUserToAddRows = false; gridLegs.RowHeadersVisible = false;
            gridLegs.DataSource = legs;

            var colLegNo = new DataGridViewTextBoxColumn { HeaderText = "#", ReadOnly = true, Width = 40 };
            var colFrom = new DataGridViewTextBoxColumn { HeaderText = "From", ReadOnly = true, Width = 60 };
            var colTo = new DataGridViewTextBoxColumn { HeaderText = "To", ReadOnly = true, Width = 60 };
            var colBearing = new DataGridViewTextBoxColumn { HeaderText = "Bearing", DataPropertyName = "Bearing", Width = 170 };

            var colLock = new DataGridViewCheckBoxColumn
            {
                HeaderText = "Lock",
                DataPropertyName = "Locked",
                Width = 52,
                ThreeState = false
            };

            gridLegs.Columns.AddRange(colLegNo, colFrom, colTo, colBearing, colLock, colDist);
            ConfigureLegsGrid(); // selection mode, no sorting, Delete key handler
            gridLegs.DataBindingComplete += (s, e) => RefreshLegNumbersAndEnds();

            // Commit checkbox edits immediately
            gridLegs.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (gridLegs.IsCurrentCellDirty)
                    gridLegs.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };

            // Leg tool row
            var pnlLegTools = new Panel { Parent = left, Left = 10, Top = 20, Width = 540, Height = 28 };
            btnAddLeg.Parent = pnlLegTools; btnAddLeg.Left = 0; btnAddLeg.Width = 90;
            btnDelLeg.Parent = pnlLegTools; btnDelLeg.Left = 100; btnDelLeg.Width = 120;
            btnCloseLeg.Parent = pnlLegTools; btnCloseLeg.Left = 230; btnCloseLeg.Width = 150;
            btnRefreshVerts.Parent = pnlLegTools; btnRefreshVerts.Left = 390; btnRefreshVerts.Width = 140;

            btnAddLeg.Click += (s, e) => { legs.Add(new CogoLegRow()); RefreshLegNumbersAndEnds(); };
            btnDelLeg.Click += DeleteSelectedLegs;
            btnCloseLeg.Click += (s, e) => ComputeClosingLeg();
            btnRefreshVerts.Click += (s, e) => RebuildVertexList(preserve: true);

            // Right: Vertex ↔ Evidence#
            var right = new GroupBox { Text = "Vertex ↔ Evidence # (and Held)", Left = 580, Top = 40, Width = 400, Height = 520 };
            gridVerts.Parent = right; gridVerts.Left = 10; gridVerts.Top = 20; gridVerts.Width = 380; gridVerts.Height = 450;
            gridVerts.AutoGenerateColumns = false; gridVerts.AllowUserToAddRows = false; gridVerts.RowHeadersVisible = false;
            gridVerts.DataSource = verts;

            var colVertex = new DataGridViewTextBoxColumn { HeaderText = "Vertex", DataPropertyName = "VertexId", ReadOnly = true, Width = 70 };
            var colEv = new DataGridViewComboBoxColumn
            {
                HeaderText = "Evidence #",
                DataPropertyName = "EvidenceNo",
                Width = 180,
                DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            };
            var colHeld = new DataGridViewCheckBoxColumn { HeaderText = "Held", DataPropertyName = "Held", Width = 70, ThreeState = false };
            gridVerts.Columns.AddRange(colVertex, colEv, colHeld);

            gridVerts.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (gridVerts.IsCurrentCellDirty) gridVerts.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };

            gridVerts.CellValueChanged += (s, e) =>
            {
                if (e.RowIndex >= 0 && e.ColumnIndex == 1 && e.RowIndex < verts.Count)
                {
                    var vm = verts[e.RowIndex];
                    var ep = GetEvidenceByNo(vm.EvidenceNo);
                    if (ep != null)
                    {
                        vm.Held = ep.Held; // copy default; user can still change
                        gridVerts.InvalidateRow(e.RowIndex);
                    }
                }
            };

            // Top: doc guard banner
            lblDocGuard.Parent = this;
            lblDocGuard.Left = 12;
            lblDocGuard.Top = 10;
            lblDocGuard.Text = $"Bound to: {SafeDocName(ownerDoc)}";
            lblDocGuard.AutoEllipsis = true;
            lblDocGuard.Width = 950;

            // Bottom panel (CSF + summary + buttons)
            var bottom = new Panel { Left = 10, Top = 570, Width = 970, Height = 80, Parent = this };
            var lblCsf = new Label { Text = "CSF:", Left = 10, Top = 8, AutoSize = true, Parent = bottom };
            txtCsf.Parent = bottom; txtCsf.Left = 50; txtCsf.Top = 5;
            chkApplyCsf.Parent = bottom; chkApplyCsf.Left = 140; chkApplyCsf.Top = 6; chkApplyCsf.Width = 260;
            lblSummary.Parent = bottom; lblSummary.Left = 420; lblSummary.Top = 8;

            btnDraw.Parent = bottom; btnDraw.Left = 570; btnDraw.Top = 35; btnDraw.Width = 90;
            btnDraw.Click += (s, e) => DrawAtPick();

            btnSave.Parent = bottom; btnSave.Left = 680; btnSave.Top = 35; btnSave.Width = 90;
            btnSaveAdj.Parent = bottom; btnSaveAdj.Left = 780; btnSaveAdj.Top = 35; btnSaveAdj.Width = 110;
            btnCancel.Parent = bottom; btnCancel.Left = 900; btnCancel.Top = 35; btnCancel.Width = 80;

            btnSave.Click += (s, e) =>
            {
                if (!EnsureOwnerIsCurrent()) return;
                if (SaveAll())
                {
                    RunAdjustAfterSave = false;
                    if (this.Modal) this.DialogResult = DialogResult.OK;
                    this.Close();
                }
            };

            btnSaveAdj.Click += (s, e) =>
            {
                if (!EnsureOwnerIsCurrent()) return;
                if (SaveAll())
                {
                    RunAdjustAfterSave = true;
                    if (!this.Modal)
                        Commands.RunAlsAdjFromUI();
                    if (this.Modal) this.DialogResult = DialogResult.OK;
                    this.Close();
                }
            };

            btnCancel.Click += (s, e) =>
            {
                if (this.Modal) this.DialogResult = DialogResult.Cancel;
                this.Close();
            };

            // Units selector + modeless handlers
            SetupUnitsUI(bottom);
            WireModelessMode();
            InitPlanDropdownUI();
            // add groups to form
            this.Controls.Add(left);
            this.Controls.Add(right);

            // Load existing data if available
            LoadEvidence();
            LoadPlanIfAny();
            if (legs.Count == 0)
                LoadCogoCsvIfExists();
            WireEvidenceComboColumn();
            RebuildVertexList(preserve: false);
            UpdateSummary();
        }

        private string SafeDocName(Document d)
            => string.IsNullOrEmpty(d?.Name) ? "(unnamed)" : Path.GetFileName(d.Name);

        private void HookDocumentLifetime()
        {
            // Close the form if the owner drawing is closing/destroyed OR if user switches drawings.
            docCol.DocumentToBeDestroyed += DocCol_DocumentToBeDestroyed;
            docCol.DocumentDestroyed += DocCol_DocumentDestroyed;
            docCol.DocumentBecameCurrent += DocCol_DocumentBecameCurrent;

            this.FormClosed += (s, e) =>
            {
                // Unhook
                docCol.DocumentToBeDestroyed -= DocCol_DocumentToBeDestroyed;
                docCol.DocumentDestroyed -= DocCol_DocumentDestroyed;
                docCol.DocumentBecameCurrent -= DocCol_DocumentBecameCurrent;
            };
        }

        private void DocCol_DocumentToBeDestroyed(object sender, DocumentCollectionEventArgs e)
        {
            if (e.Document == ownerDoc)
            {
                try { this.BeginInvoke(new Action(() => this.Close())); }
                catch { /* ignore */ }
            }
        }
        private void DocCol_DocumentDestroyed(object sender, DocumentDestroyedEventArgs e)
        {
            if (string.Equals(e.FileName, ownerDoc?.Name, StringComparison.OrdinalIgnoreCase))
            {
                try { this.BeginInvoke(new Action(() => this.Close())); }
                catch { /* ignore */ }
            }
        }
        private void DocCol_DocumentBecameCurrent(object sender, DocumentCollectionEventArgs e)
        {
            // If the current drawing is NOT the owner, close this form to eliminate any cross‑doc UI risk.
            if (e.Document != ownerDoc)
            {
                try { this.BeginInvoke(new Action(() => this.Close())); } catch { /* ignore */ }
            }
        }

        private bool EnsureOwnerIsCurrent()
        {
            var cur = AcadApp.DocumentManager.MdiActiveDocument;
            if (cur != ownerDoc)
            {
                MessageBox.Show($"This editor belongs to “{SafeDocName(ownerDoc)}”.\nSwitch back to that drawing to use it.", "Wrong Drawing Active",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            return true;
        }

        private void RefreshLegNumbersAndEnds()
        {
            for (int i = 0; i < gridLegs.Rows.Count; i++)
            {
                var r = gridLegs.Rows[i];
                r.Cells[0].Value = (i + 1).ToString();
                r.Cells[1].Value = $"P{i + 1}";
                r.Cells[2].Value = $"P{i + 2}";
            }
            UpdateSummary();
        }

        private void UpdateSummary()
        {
            double sum = 0, sx = 0, sy = 0;
            for (int i = 0; i < legs.Count; i++)
            {
                var row = legs[i];
                sum += Math.Abs(row.Distance);

                if (BearingParserEx.TryParseToAzimuthRad(row.Bearing, out double az))
                {
                    double d = row.Distance;
                    if (d < 0) { d = -d; az = Commands.Normalize(az + Math.PI); }
                    sx += d * Math.Cos(az);
                    sy += d * Math.Sin(az);
                }
            }
            double closure = Math.Sqrt(sx * sx + sy * sy);
            int lockedCount = legs.Count(l => l.Locked);
            lblSummary.Text = $"Legs: {legs.Count} (locked: {lockedCount})   Perimeter (raw): {sum:0.###} {UnitAbbrev(currentUnits)}   Open-closure: {closure:0.###} {UnitAbbrev(currentUnits)}";
        }

        private void ComputeClosingLeg()
        {
            if (!EnsureOwnerIsCurrent()) return;

            double sx = 0, sy = 0;
            for (int i = 0; i < legs.Count; i++)
            {
                var row = legs[i];
                if (!BearingParserEx.TryParseToAzimuthRad(row.Bearing, out double az)) { MessageBox.Show($"Invalid bearing on leg {i + 1}."); return; }
                double d = row.Distance;
                if (!(d != 0)) { MessageBox.Show($"Invalid distance on leg {i + 1}."); return; }
                if (d < 0) { d = -d; az = Commands.Normalize(az + Math.PI); }
                sx += d * Math.Cos(az);
                sy += d * Math.Sin(az);
            }
            double dx = -sx, dy = -sy;
            double dist = Math.Sqrt(dx * dx + dy * dy);
            double azClose = Math.Atan2(dy, dx);
            if (azClose < 0) azClose += 2 * Math.PI;

            // Pick style that matches the majority of existing bearings
            bool useCompact =
                legs.Count > 0 &&
                legs.Count(l => BearingParserEx.LooksQuadCompact(l.Bearing)) >
                (legs.Count / 2);

            string btxt = useCompact ? Cad.FormatBearingCompact(azClose)
                                     : Cad.FormatBearing(azClose);

            // distance is already in the current input units; round to UI precision
            legs.Add(new CogoLegRow { Bearing = btxt, Distance = Math.Round(dist, 3), Locked = false });
            RefreshLegNumbersAndEnds();
        }

        private void LoadEvidence()
        {
            string evPath = Config.EvidenceJsonPath();
            if (File.Exists(evPath))
            {
                links = Json.Load<EvidenceLinks>(evPath) ?? new EvidenceLinks();
                Cad.RefreshEvidencePositionsFromDwg(links);
                Json.Save(evPath, links);
            }

            evidenceOptions = new List<ComboItem> { new ComboItem { Index = 0, Label = "(none)" } };
            for (int i = 0; i < links.Points.Count; i++)
            {
                var p = links.Points[i];
                string tag = $"#{i + 1}";
                if (!string.IsNullOrWhiteSpace(p.EvidenceType)) tag += $" {p.EvidenceType}";
                evidenceOptions.Add(new ComboItem { Index = i + 1, Label = tag });
            }
        }

        private EvidencePoint GetEvidenceByNo(int no)
            => (no >= 1 && no <= links.Points.Count) ? links.Points[no - 1] : null;

        private void WireEvidenceComboColumn()
        {
            if (gridVerts.Columns.Count < 2) return;
            if (gridVerts.Columns[1] is DataGridViewComboBoxColumn c)
            {
                c.DataSource = evidenceOptions;
                c.ValueMember = "Index";
                c.DisplayMember = "Label";
            }
        }

        private void LoadPlanIfAny()
        {
            string planPath = Config.PlanJsonPath();
            if (!File.Exists(planPath)) return;
            ApplyPlanToUI(PlanIO.LoadActive(planPath));
        }

        private void LoadCogoCsvIfExists()
        {
            try
            {
                string folderPath = Path.GetDirectoryName(ownerDoc?.Name);
                if (string.IsNullOrEmpty(folderPath)) return;

                // Try COGO (no extension) first, then legacy COGO.csv
                string[] candidates = {
                    Path.Combine(folderPath, "COGO"),
                    Path.Combine(folderPath, "COGO.csv")
                };
                string file = candidates.FirstOrDefault(File.Exists);
                if (string.IsNullOrEmpty(file)) return;

                var entries = CogoFile.LoadUserCogoInput(file);
                if (entries == null || entries.Count == 0) return;

                legs.Clear();
                foreach (var e in entries)
                    legs.Add(new CogoLegRow { Bearing = e.Bearing, Distance = e.Distance, Locked = false });
            }
            catch
            {
                // ignore
            }
        }

        private void RebuildVertexList(bool preserve)
        {
            var old = verts.ToDictionary(v => v.VertexId, v => v);
            verts.Clear();
            int vcount = Math.Max(1, legs.Count + 1);
            for (int i = 0; i < vcount; i++)
            {
                string vid = $"P{i + 1}";
                if (preserve && old.TryGetValue(vid, out var o))
                    verts.Add(new VertexMapRow { VertexId = vid, EvidenceNo = o.EvidenceNo, Held = o.Held });
                else
                    verts.Add(new VertexMapRow { VertexId = vid, EvidenceNo = 0, Held = false });
            }
        }

        private void ConfigureLegsGrid()
        {
            gridLegs.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            gridLegs.MultiSelect = true;

            foreach (DataGridViewColumn col in gridLegs.Columns)
                col.SortMode = DataGridViewColumnSortMode.NotSortable;

            gridLegs.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Delete)
                {
                    DeleteSelectedLegs(s, EventArgs.Empty);
                    e.Handled = true;
                }
            };
        }

        private void DeleteSelectedLegs(object? sender, EventArgs e)
        {
            gridLegs.EndEdit();

            var toRemove = new HashSet<CogoLegRow>();

            foreach (DataGridViewRow r in gridLegs.SelectedRows)
                if (!r.IsNewRow && r.DataBoundItem is CogoLegRow rowObj)
                    toRemove.Add(rowObj);

            if (toRemove.Count == 0)
            {
                foreach (DataGridViewCell c in gridLegs.SelectedCells)
                    if (!c.OwningRow.IsNewRow && c.OwningRow.DataBoundItem is CogoLegRow rowObj)
                        toRemove.Add(rowObj);
            }

            if (toRemove.Count == 0 &&
                gridLegs.CurrentRow != null &&
                !gridLegs.CurrentRow.IsNewRow &&
                gridLegs.CurrentRow.DataBoundItem is CogoLegRow curObj)
            {
                toRemove.Add(curObj);
            }

            if (toRemove.Count == 0) return;

            foreach (var r in toRemove)
                legs.Remove(r);

            RefreshLegNumbersAndEnds();
            RebuildVertexList(preserve: true);
        }

        private bool TryBuildPlan(out PlanData plan, out double closureLen)
        {
            plan = null; closureLen = 0;
            if (!EnsureOwnerIsCurrent()) return false;

            double csf = 1.0;
            if (!double.TryParse(txtCsf.Text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out csf) || csf <= 0.5 || csf >= 1.5)
            {
                MessageBox.Show("Invalid CSF. Enter a value between 0.5 and 1.5.");
                return false;
            }

            double u2m = UnitToMeters(currentUnits); // convert from selected input units to metres

            plan = new PlanData { CombinedScaleFactor = csf };
            var cur = new Point2d(0, 0);
            plan.VertexIds.Add("P1");
            plan.VertexXY.Add(new XY(cur.X, cur.Y));

            for (int i = 0; i < legs.Count; i++)
            {
                var r = legs[i];

                if (!BearingParserEx.TryParseToAzimuthRadWithFlags(r.Bearing, out double az, out bool hadNegBearing))
                {
                    MessageBox.Show($"Invalid bearing on leg {i + 1}: \"{r.Bearing}\"");
                    plan = null; return false;
                }

                if (!(r.Distance != 0))
                {
                    MessageBox.Show($"Invalid distance on leg {i + 1}.");
                    plan = null; return false;
                }
                bool distNegative = r.Distance < 0;

                if (hadNegBearing && distNegative)
                {
                    MessageBox.Show($"Leg {i + 1}: cannot use '-' on BOTH bearing and distance.");
                    plan = null; return false;
                }

                double localAz = az;
                double dAbs = Math.Abs(r.Distance);
                if (distNegative) localAz = Commands.Normalize(localAz + Math.PI);

                double dMetres = dAbs * u2m;
                double dUse = chkApplyCsf.Checked ? dMetres * csf : dMetres;

                string fromId = $"P{i + 1}";
                string toId = $"P{i + 2}";

                cur = new Point2d(cur.X + dUse * Math.Cos(localAz), cur.Y + dUse * Math.Sin(localAz));
                plan.Edges.Add(new PlanEdge { FromId = fromId, ToId = toId, Distance = dUse, BearingRad = localAz, Locked = r.Locked, BearingText = r.Bearing });
                plan.VertexIds.Add(toId);
                plan.VertexXY.Add(new XY(cur.X, cur.Y));
            }

            var first = new Point2d(plan.VertexXY[0].X, plan.VertexXY[0].Y);
            var lastXY = plan.VertexXY[plan.VertexXY.Count - 1];
            var last = new Point2d(lastXY.X, lastXY.Y);
            closureLen = Cad.Distance(first, last);
            plan.Closed = closureLen < 1e-6;
            return true;
        }

        private bool SaveAll()
        {
            if (!TryBuildPlan(out var plan, out double closure))
                return false;

            string folderPath = Path.GetDirectoryName(ownerDoc?.Name);
            if (!string.IsNullOrEmpty(folderPath))
            {
                var entries = new List<UserCogoInput>();
                for (int i = 0; i < legs.Count; i++)
                {
                    entries.Add(new UserCogoInput
                    {
                        Line = i + 1,
                        Point1 = "P" + (i + 1),
                        Point2 = "P" + (i + 2),
                        Bearing = legs[i].Bearing,
                        Distance = legs[i].Distance,
                        Azimuth = string.Empty,
                        Deflection = string.Empty
                    });
                }
                CogoFile.SaveUserCogoInput(Path.Combine(folderPath, "COGO"), entries);
            }

            // Derive a name from the dropdown (user may type a new one)
            var planName = string.IsNullOrWhiteSpace(cboPlan.Text)
                ? $"Session {DateTime.Now:yyyyMMdd_HHmm}"
                : cboPlan.Text.Trim();

            PlanBuild.SaveAndDraw(plan, closure, tag: "COGO-Editor", planName: planName);

            // Persist evidence links (unchanged from your version)
            foreach (var p in links.Points) p.PlanId = "";
            var byNo = new Dictionary<int, EvidencePoint>();
            for (int i = 0; i < links.Points.Count; i++) byNo[i + 1] = links.Points[i];

            foreach (var v in verts)
            {
                if (v.EvidenceNo <= 0) continue;
                if (!byNo.TryGetValue(v.EvidenceNo, out var ep)) continue;
                ep.PlanId = v.VertexId;
                ep.Held = v.Held;
            }

            Json.Save(Config.EvidenceJsonPath(), links);
            Cad.WriteMessageSafe($"\nSaved evidence links: {Config.EvidenceJsonPath()}");

            // Refresh the pulldown to show the newly saved/updated name
            ReloadPlanDropdown();
            LoadSelectedPlan();

            return true;
        }

        private void DrawAtPick()
        {
            if (!EnsureOwnerIsCurrent()) return;
            if (!TryBuildPlan(out var plan, out double _)) return;

            var pr = OwnerEd.GetPoint(new PromptPointOptions("\nPick insertion point (P1) for COGO: "));
            if (pr.Status != PromptStatus.OK) return;

            var basePt = pr.Value;

            Cad.EnsureLayer(Config.LayerPdf, aci: 4);
            var pl = new Polyline { Layer = Config.LayerPdf, Closed = plan.Closed };
            for (int i = 0; i < plan.VertexXY.Count; i++)
            {
                var v = plan.VertexXY[i];
                pl.AddVertexAt(i, new Point2d(basePt.X + v.X, basePt.Y + v.Y), 0, 0, 0);
            }
            Cad.AddToModelSpace(pl);

            OwnerEd.WriteMessage($"\nCOGO drawn at ({basePt.X:0.###}, {basePt.Y:0.###}). " +
                                (chkApplyCsf.Checked ? "CSF applied." : "CSF not applied."));
        }

        private void InitPlanDropdownUI()
        {
            // Small bar under the doc-guard label
            var planBar = new Panel { Left = 10, Top = 28, Width = 560, Height = 26, Parent = this };
            var lblPlan = new Label { Parent = planBar, Left = 0, Top = 6, AutoSize = true, Text = "Named COGO:" };
            cboPlan.Parent = planBar; cboPlan.Left = 100; cboPlan.Top = 2; cboPlan.DropDownStyle = ComboBoxStyle.DropDown;

            // make the dropdown friendlier
            cboPlan.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cboPlan.AutoCompleteSource = AutoCompleteSource.ListItems;

            // "New..." button to create a new named COGO
            btnNewPlan.Parent = planBar;
            btnNewPlan.Left = cboPlan.Right + 8;
            btnNewPlan.Top = 0;
            btnNewPlan.Click += (s, e) =>
            {
                cboPlan.SelectedIndex = -1;
                cboPlan.Text = string.Empty;
                cboPlan.Focus();
                MessageBox.Show("Type a new name in the box, then click Save or Save & Adjust.",
                    "New COGO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };
            planBar.Resize += (s, e) => btnNewPlan.Left = cboPlan.Right + 8;

            ReloadPlanDropdown();
            LoadSelectedPlan();
            cboPlan.SelectedIndexChanged += (s, e) => LoadSelectedPlan();
        }

        private void ReloadPlanDropdown()
        {
            string path = Config.PlanJsonPath();
            var names = PlanIO.Names(path).ToList();
            cboPlan.Items.Clear();
            foreach (var n in names) cboPlan.Items.Add(n);

            var store = PlanIO.LoadStore(path);
            string active = store.ActiveName;

            if (!string.IsNullOrWhiteSpace(active))
            {
                int idx = names.FindIndex(n => string.Equals(n, active, StringComparison.OrdinalIgnoreCase));
                if (idx >= 0) cboPlan.SelectedIndex = idx;
            }
            else if (names.Count > 0)
            {
                cboPlan.SelectedIndex = 0;
            }
        }

        private void LoadSelectedPlan()
        {
            string name = cboPlan.Text?.Trim();
            var store = PlanIO.LoadStore(Config.PlanJsonPath());
            var np = (!string.IsNullOrWhiteSpace(name) ? store.Find(name) : null) ?? store.Find(store.ActiveName);
            if (np != null) ApplyPlanToUI(np.Data);
        }

        private void ApplyPlanToUI(PlanData plan)
        {
            legs.Clear();
            double m2u = 1.0 / UnitToMeters(currentUnits);

            for (int i = 0; i < plan.Edges.Count; i++)
            {
                var e = plan.EdgeAt(i);
                string btxt = string.IsNullOrWhiteSpace(e.BearingText) ? Cad.FormatBearing(e.BearingRad) : e.BearingText;

                legs.Add(new CogoLegRow
                {
                    Bearing = btxt,
                    Distance = e.Distance * m2u,
                    Locked = e.Locked
                });
            }
            txtCsf.Text = plan.CombinedScaleFactor.ToString("0.######", CultureInfo.InvariantCulture);
            RefreshLegNumbersAndEnds();

            // Rebuild vertex grid to match leg count; keep any evidence/Held if possible
            var idToEvidenceNo = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < links.Points.Count; i++)
            {
                var p = links.Points[i];
                if (!string.IsNullOrWhiteSpace(p.PlanId)) idToEvidenceNo[p.PlanId] = i + 1;
            }

            verts.Clear();
            int vcount = Math.Max(1, legs.Count + 1);
            for (int i = 0; i < vcount; i++)
            {
                string vid = $"P{i + 1}";
                int eno = idToEvidenceNo.TryGetValue(vid, out var k) ? k : 0;
                bool held = GetEvidenceByNo(eno)?.Held ?? false;
                verts.Add(new VertexMapRow { VertexId = vid, EvidenceNo = eno, Held = held });
            }
        }


        // ---------- Units helpers (PlanCogoForm only) ----------
        private static string UnitAbbrev(InputUnits u)
            => u == InputUnits.Meters ? "m" : (u == InputUnits.Feet ? "ft" : "ch");

        private static double UnitToMeters(InputUnits u)
            => u == InputUnits.Meters ? 1.0
               : (u == InputUnits.Feet ? 0.3048
               : 20.1168); // 1 chain = 66 ft

        private void UpdateDistanceColumnHeader()
        {
            colDist.HeaderText = $"Distance ({UnitAbbrev(currentUnits)})";
        }

        private void ConvertExistingDistances(InputUnits fromU, InputUnits toU)
        {
            if (fromU == toU) return;
            double f = UnitToMeters(fromU) / UnitToMeters(toU);
            for (int i = 0; i < legs.Count; i++)
                legs[i].Distance = legs[i].Distance * f;

            gridLegs.Refresh();
            UpdateSummary();
        }

        private void SetupUnitsUI(Control bottom)
        {
            lblUnits.Parent = bottom; lblUnits.Left = 410; lblUnits.Top = 8;
            cboUnits.Parent = bottom; cboUnits.Left = lblUnits.Left + 80; cboUnits.Top = 4;

            if (lblSummary.Left < (cboUnits.Left + cboUnits.Width + 20))
                lblSummary.Left = cboUnits.Left + cboUnits.Width + 20;

            cboUnits.Items.AddRange(new object[] { "m (meters)", "ft (feet)", "ch (chains)" });
            cboUnits.SelectedIndex = (int)InputUnits.Meters;
            UpdateDistanceColumnHeader();

            cboUnits.SelectedIndexChanged += (s, e) =>
            {
                var newU = (InputUnits)cboUnits.SelectedIndex;
                var oldU = currentUnits;
                if (newU == oldU) return;

                if (legs.Count > 0)
                {
                    var dr = MessageBox.Show(
                        "Convert existing distances to the new unit?\n\nYes: preserve lengths (values will be rescaled)\nNo: keep numbers (only the unit label changes)",
                        "Change Input Units",
                        MessageBoxButtons.YesNoCancel,
                        MessageBoxIcon.Question);

                    if (dr == DialogResult.Cancel)
                    {
                        cboUnits.SelectedIndex = (int)oldU;
                        return;
                    }
                    if (dr == DialogResult.Yes)
                        ConvertExistingDistances(oldU, newU);
                }

                currentUnits = newU;
                UpdateDistanceColumnHeader();
                UpdateSummary();
            };
        }

        private void WireModelessMode()
        {
            this.KeyPreview = true;
            this.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape)
                {
                    btnCancel.PerformClick();
                    e.Handled = true;
                }
                else if (e.KeyCode == Keys.F5)
                {
                    RebuildVertexList(preserve: true);
                    e.Handled = true;
                }
            };

            this.TopMost = false;
            this.ShowInTaskbar = false;
        }
    }

    // ---------------------------- Commands: PLANCOGOUI + updated ALSWIZARD ----------------------------
    public partial class Commands
    {
        [CommandMethod("PLANCOGOUI", CommandFlags.Session)]
        public void PlanCogoUI()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var frm = new PlanCogoForm(doc);
            AcadApp.ShowModelessDialog(frm);  // .NET 4.8 signature
        }

        // Wizard: Harvest → COGO UI → ALSADJ (on Save & Adjust)
        [CommandMethod("ALSWIZARD", CommandFlags.Session)]
        public void AlsWizard()
        {
            Cad.WriteMessageSafe("\nALSWIZARD — Harvest → PLANCOGOUI → ALSADJ (on Save & Adjust).");

            EvidenceHarvest(); // uses your existing implementation

            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var frm = new PlanCogoForm(doc);
            AcadApp.ShowModelessDialog(frm);  // .NET 4.8 signature
            Cad.WriteMessageSafe("\nCOGO editor opened modeless. Use Save to write JSON; Save & Adjust to run ALSADJ.");
        }
    }

    internal static class Json
    {
        static readonly JavaScriptSerializer Jss = new JavaScriptSerializer();

        public static void Save<T>(string path, T obj)
            => File.WriteAllText(path, Jss.Serialize(obj), Encoding.UTF8);

        public static T Load<T>(string path)
            => Jss.Deserialize<T>(File.ReadAllText(path, Encoding.UTF8));
    }

    internal static class Cad
    {
        public static Editor Ed => AcadApp.DocumentManager.MdiActiveDocument.Editor;
        public static Database Db => HostApplicationServices.WorkingDatabase;

        public static void WriteMessageSafe(string msg)
        {
            try { Ed.WriteMessage(msg); }
            catch { /* eNotApplicable if wrong doc state — ignore */ }
        }

        public static void RefreshEvidencePositionsFromDwg(EvidenceLinks links)
        {
            if (links?.Points == null || links.Points.Count == 0) return;

            using (var tr = Db.TransactionManager.StartTransaction())
            {
                foreach (var p in links.Points)
                {
                    if (string.IsNullOrWhiteSpace(p.Handle)) continue;

                    try
                    {
                        long hval = Convert.ToInt64(p.Handle, 16);
                        var h = new Handle(hval);
                        ObjectId id = Db.GetObjectId(false, h, 0);
                        if (!id.IsNull && id.IsValid)
                        {
                            if (tr.GetObject(id, OpenMode.ForRead, false) is BlockReference br)
                            {
                                p.X = br.Position.X;
                                p.Y = br.Position.Y;
                            }
                        }
                    }
                    catch
                    {
                        // Handle not found / entity erased / xref, etc.
                    }
                }
                tr.Commit();
            }
        }

        public static bool UnlockIfLocked(string layerName)
        {
            bool wasLocked = false;
            WithLockedDoc(() =>
            {
                using (var tr = Db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(Db.LayerTableId, OpenMode.ForRead);
                    if (lt.Has(layerName))
                    {
                        var rec = (LayerTableRecord)tr.GetObject(lt[layerName], OpenMode.ForRead);
                        if (rec.IsLocked)
                        {
                            rec.UpgradeOpen();
                            rec.IsLocked = false;
                            wasLocked = true;
                        }
                    }
                    tr.Commit();
                }
            });
            return wasLocked;
        }

        public static void WithLockedDoc(Action action)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) { action(); return; }

            var oldDb = HostApplicationServices.WorkingDatabase;
            try
            {
                HostApplicationServices.WorkingDatabase = doc.Database;

                if (doc.LockMode == DocumentLockMode.NotLocked)
                {
                    using (doc.LockDocument())
                        action();
                }
                else
                {
                    action();
                }
            }
            finally
            {
                HostApplicationServices.WorkingDatabase = oldDb;
            }
        }

        public static void WithLockedDoc(Document doc, Action action)
        {
            if (doc == null) { action(); return; }

            var oldDb = HostApplicationServices.WorkingDatabase;
            try
            {
                HostApplicationServices.WorkingDatabase = doc.Database;

                if (doc.LockMode == DocumentLockMode.NotLocked)
                {
                    using (doc.LockDocument())
                        action();
                }
                else
                {
                    action();
                }
            }
            finally
            {
                HostApplicationServices.WorkingDatabase = oldDb;
            }
        }

        public static ObjectId EnsureLayer(string name, short aci = 7, bool lockAfter = false)
        {
            ObjectId id = ObjectId.Null;
            WithLockedDoc(() =>
            {
                using (var tr = Db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(Db.LayerTableId, OpenMode.ForRead);
                    if (!lt.Has(name))
                    {
                        lt.UpgradeOpen();
                        var rec = new LayerTableRecord
                        {
                            Name = name,
                            Color = Color.FromColorIndex(ColorMethod.ByAci, aci)
                        };
                        id = lt.Add(rec);
                        tr.AddNewlyCreatedDBObject(rec, true);
                    }
                    else
                    {
                        id = lt[name];
                    }
                    tr.Commit();

                    if (lockAfter) LockLayer(name, true);
                }
            });
            return id;
        }

        public static void LockLayer(string name, bool locked)
        {
            WithLockedDoc(() =>
            {
                using (var tr = Db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(Db.LayerTableId, OpenMode.ForRead);
                    if (!lt.Has(name)) return;
                    var rec = (LayerTableRecord)tr.GetObject(lt[name], OpenMode.ForRead);
                    rec.UpgradeOpen();
                    rec.IsLocked = locked;
                    tr.Commit();
                }
            });
        }

        public static ObjectId AddToModelSpace(Entity e)
        {
            ObjectId id = ObjectId.Null;
            WithLockedDoc(() =>
            {
                using (var tr = Db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(Db.BlockTableId, OpenMode.ForRead);
                    var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    id = ms.AppendEntity(e);
                    tr.AddNewlyCreatedDBObject(e, true);
                    tr.Commit();
                }
            });
            return id;
        }

        public static void TransformEntities(IEnumerable<ObjectId> ids, Matrix3d m)
        {
            WithLockedDoc(() =>
            {
                using (var tr = Db.TransactionManager.StartTransaction())
                {
                    foreach (var id in ids)
                        ((Entity)tr.GetObject(id, OpenMode.ForWrite)).TransformBy(m);
                    tr.Commit();
                }
            });
        }

        public static void SetLayer(IEnumerable<ObjectId> ids, string layer)
        {
            EnsureLayer(layer);
            WithLockedDoc(() =>
            {
                using (var tr = Db.TransactionManager.StartTransaction())
                {
                    foreach (var id in ids)
                        ((Entity)tr.GetObject(id, OpenMode.ForWrite)).Layer = layer;
                    tr.Commit();
                }
            });
        }
        public static PromptSelectionResult PromptSelect(string msg, params TypedValue[] filter)
        {
            var opts = new PromptSelectionOptions { MessageForAdding = "\n" + msg + ": " };
            return (filter != null && filter.Length > 0)
                ? Ed.GetSelection(opts, new SelectionFilter(filter))
                : Ed.GetSelection(opts);
        }
        public static PromptPointResult GetPoint(string msg, bool allowNone = false)
        {
            var opts = new PromptPointOptions("\n" + msg + ": ")
            {
                AllowNone = allowNone
            };
            return Ed.GetPoint(opts);
        }
        public static double Distance(Point2d a, Point2d b) { var dx = a.X - b.X; var dy = a.Y - b.Y; return Math.Sqrt(dx * dx + dy * dy); }
        public static double ToAzimuthRad(Point2d a, Point2d b)
        {
            double th = Math.Atan2(b.Y - a.Y, b.X - a.X); if (th < 0) th += 2 * Math.PI; return th;
        }
        public static string FormatBearingCompact(double az)
        {
            az %= 2 * Math.PI; if (az < 0) az += 2 * Math.PI;
            bool north = az <= Math.PI / 2 || az >= 3 * Math.PI / 2;
            bool east  = az < Math.PI;

            double theta = north ? (east ? az : Math.PI - az)
                                 : (east ? 2 * Math.PI - az : az - Math.PI);

            double deg = theta * 180.0 / Math.PI;
            int d = (int)Math.Floor(deg + 1e-9);
            double remM = (deg - d) * 60.0;
            int m = (int)Math.Floor(remM + 1e-9);
            double sExact = (remM - m) * 60.0;
            int s = (int)Math.Round(sExact);

            if (s == 60) { s = 0; m += 1; }
            if (m == 60) { m = 0; d += 1; }

            return $"{(north ? "N" : "S")}{d:00}.{m:00}{s:00}{(east ? "E" : "W")}";
        }
        public static string FormatBearing(double az)
        {
            az %= 2 * Math.PI; if (az < 0) az += 2 * Math.PI;
            bool north = az <= Math.PI / 2 || az >= 3 * Math.PI / 2; bool east = az < Math.PI;
            double theta = north ? (east ? az : Math.PI - az) : (east ? 2 * Math.PI - az : az - Math.PI);
            double deg = theta * 180.0 / Math.PI; int d = (int)Math.Floor(deg + 1e-9);
            double remM = (deg - d) * 60.0; int m = (int)Math.Floor(remM + 1e-9); double s = (remM - m) * 60.0;
            return $"{(north ? "N" : "S")} {d:00}°{m:00}'{s:00.##}\" {(east ? "E" : "W")}";
        }
    }

    // ---------------------------- Data Models ----------------------------
    [Serializable] public class XY { public double X; public double Y; public XY() { } public XY(double x, double y) { X = x; Y = y; } }

    [Serializable]
    public class PlanEdge
    {
        public string FromId;
        public string ToId;
        public double Distance;
        public double BearingRad;

        public bool Locked;

        // preserve user-entered bearing text exactly
        public string BearingText = "";
    }

    [Serializable]
    public class PlanData
    {
        public List<string> VertexIds = new List<string>();  // P1..Pn
        public List<XY> VertexXY = new List<XY>();           // actual polyline XY for nearest-vertex mapping
        public List<PlanEdge> Edges = new List<PlanEdge>();  // distances + bearings (geometry or plan numbers)
        public bool Closed = true;
        public double CombinedScaleFactor = 1.0;
        public int Count => VertexIds.Count;
        public int IndexOf(string id) => VertexIds.FindIndex(v => string.Equals(v, id, StringComparison.OrdinalIgnoreCase));
        public PlanEdge EdgeAt(int i) => Edges[i];
    }

    [Serializable]
    public class NamedPlan
    {
        public string Name = "Default";
        public DateTime ModifiedUtc = DateTime.UtcNow;
        public PlanData Data = new PlanData();
    }

    [Serializable]
    public class PlanStore
    {
        public string ActiveName = "Default";
        public List<NamedPlan> Plans = new List<NamedPlan>();

        public NamedPlan Find(string name) =>
            Plans.FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
    }

    internal static class PlanIO
    {
        public static PlanStore LoadStore(string path)
        {
            if (!File.Exists(path)) return new PlanStore();

            // Try new multi-plan schema
            try
            {
                var s = Json.Load<PlanStore>(path);
                if (s != null && s.Plans != null && s.Plans.Count > 0)
                    return s;
            }
            catch { /* fall through */ }

            // Fallback: legacy single PlanData file -> migrate
            try
            {
                var old = Json.Load<PlanData>(path);
                if (old != null)
                {
                    return new PlanStore
                    {
                        ActiveName = "Default",
                        Plans = new List<NamedPlan>
                    {
                        new NamedPlan { Name = "Default", ModifiedUtc = DateTime.UtcNow, Data = old }
                    }
                    };
                }
            }
            catch { /* return empty */ }

            return new PlanStore();
        }

        public static void SaveStore(string path, PlanStore store) => Json.Save(path, store);

        public static PlanData LoadActive(string path)
        {
            var store = LoadStore(path);
            var active = store.Find(store.ActiveName) ?? store.Plans.FirstOrDefault();
            return active?.Data ?? new PlanData();
        }

        public static void Upsert(string path, string name, PlanData plan)
        {
            if (string.IsNullOrWhiteSpace(name)) name = "Active";
            var store = LoadStore(path);
            var np = store.Find(name);
            if (np == null)
                store.Plans.Add(new NamedPlan { Name = name, ModifiedUtc = DateTime.UtcNow, Data = plan });
            else
            {
                np.Data = plan;
                np.ModifiedUtc = DateTime.UtcNow;
            }
            store.ActiveName = name;
            SaveStore(path, store);
        }

        public static IEnumerable<string> Names(string path) =>
            LoadStore(path).Plans
                           .OrderByDescending(p => p.ModifiedUtc)
                           .Select(p => p.Name);
    }


    [Serializable]
    public class EvidencePoint
    {
        public string PlanId { get; set; } = "";     // e.g., "P7"
        public string Handle { get; set; } = "";     // drawing handle
        public double X { get; set; }
        public double Y { get; set; }
        public string EvidenceType { get; set; } = "";
        public bool Held { get; set; } = false;
        public int Priority { get; set; } = 0;

        public Point2d XY() => new Point2d(X, Y);
    }

    [Serializable] public class EvidenceLinks { public List<EvidencePoint> Points = new List<EvidencePoint>(); }

    // ---------------------------- Parsing ----------------------------
    internal static class BearingParserEx
    {
        // Quadrant bearing: N dd°mm'ss.s" E
        static readonly Regex Quad = new Regex(
            @"^\s*([NS])\s*([0-9]{1,3})\s*(?:[D°])\s*([0-9]{1,2})?\s*['’]?\s*([0-9]{1,2}(?:\.\d+)?)?\s*[""”]?\s*([EW])\s*$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // Quadrant compact: N dd.mm[ss] E  (N45.3045E or S12.30W)
        static readonly Regex QuadCompact = new Regex(
            @"^\s*([NS])\s*([0-9]{1,3})(?:\.([0-9]{2}|[0-9]{4}))?\s*([EW])\s*$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // Azimuth-from-north DMS
        static readonly Regex AzDms = new Regex(
            @"^\s*(?:AZ\s*)?([0-9]{1,3})\s*(?:[D°])\s*([0-9]{1,2})?\s*['’]?\s*([0-9]{1,2}(?:\.\d+)?)?\s*[""”]?\s*$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // General numeric: D.MM[SS] (2 or 4 digits) or decimal degrees
        static readonly Regex AzNumeric = new Regex(
            @"^\s*(?:AZ\s*)?([0-9]+)(?:\.([0-9]+))?\s*$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static bool LooksQuadCompact(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            return QuadCompact.IsMatch(text.Trim());
        }

        private static bool TryParseDegreesMmSs(string degStr, string fracStr, out double degOut)
        {
            degOut = 0;
            if (!int.TryParse(degStr, NumberStyles.Integer, CultureInfo.InvariantCulture, out int d)) return false;

            if (string.IsNullOrEmpty(fracStr))
            {
                degOut = d;
                return true;
            }

            if (fracStr.Length == 2)
            {
                if (!int.TryParse(fracStr, NumberStyles.Integer, CultureInfo.InvariantCulture, out int mm)) return false;
                if (mm < 0 || mm >= 60) return false;
                degOut = d + (mm / 60.0);
                return true;
            }
            if (fracStr.Length == 4)
            {
                if (!int.TryParse(fracStr.Substring(0, 2), NumberStyles.Integer, CultureInfo.InvariantCulture, out int mm)) return false;
                if (!int.TryParse(fracStr.Substring(2, 2), NumberStyles.Integer, CultureInfo.InvariantCulture, out int ss)) return false;
                if (mm < 0 || mm >= 60 || ss < 0 || ss >= 60) return false;
                degOut = d + (mm / 60.0) + (ss / 3600.0);
                return true;
            }
            return false;
        }

        /// <summary>
        /// Parse bearing text to azimuth (east-of-x axis, radians). Supports:
        /// - Quadrant forms (tokenized or compact D.MM[SS])
        /// - Azimuth tokenized DMS
        /// - Pure numeric: D.MM[SS] or decimal degrees
        /// - Optional leading '-' meaning "add 180°"
        /// </summary>
        public static bool TryParseToAzimuthRadWithFlags(string text, out double azEastRad, out bool hadLeadingMinus)
        {
            azEastRad = 0;
            hadLeadingMinus = false;
            if (string.IsNullOrWhiteSpace(text)) return false;
            text = text.Trim().Replace('’', '\'').Replace('″', '"');

            if (text.StartsWith("-", StringComparison.Ordinal))
            {
                hadLeadingMinus = true;
                text = text.Substring(1).TrimStart();
            }
            else if (text.StartsWith("+", StringComparison.Ordinal))
            {
                text = text.Substring(1).TrimStart();
            }

            var mq = Quad.Match(text);
            if (mq.Success)
            {
                bool north = char.ToUpperInvariant(mq.Groups[1].Value[0]) == 'N';
                bool east = char.ToUpperInvariant(mq.Groups[5].Value[0]) == 'E';
                int d = int.Parse(mq.Groups[2].Value, CultureInfo.InvariantCulture);
                int m = string.IsNullOrEmpty(mq.Groups[3].Value) ? 0 : int.Parse(mq.Groups[3].Value, CultureInfo.InvariantCulture);
                double s = string.IsNullOrEmpty(mq.Groups[4].Value) ? 0 : double.Parse(mq.Groups[4].Value, CultureInfo.InvariantCulture);

                double thetaDeg = d + m / 60.0 + s / 3600.0;
                double azN = 0;
                if (north && east) azN = thetaDeg;
                if (north && !east) azN = 360.0 - thetaDeg;
                if (!north && east) azN = 180.0 - thetaDeg;
                if (!north && !east) azN = 180.0 + thetaDeg;

                azEastRad = (Math.PI / 2) - (azN * Math.PI / 180.0);
                while (azEastRad < 0) azEastRad += 2 * Math.PI;
                while (azEastRad >= 2 * Math.PI) azEastRad -= 2 * Math.PI;

                if (hadLeadingMinus) azEastRad = Commands.Normalize(azEastRad + Math.PI);
                return true;
            }

            var mqc = QuadCompact.Match(text);
            if (mqc.Success)
            {
                bool north = char.ToUpperInvariant(mqc.Groups[1].Value[0]) == 'N';
                bool east = char.ToUpperInvariant(mqc.Groups[4].Value[0]) == 'E';
                string dstr = mqc.Groups[2].Value;
                string fstr = mqc.Groups[3].Success ? mqc.Groups[3].Value : "";
                if (!TryParseDegreesMmSs(dstr, fstr, out double thetaDeg)) return false;

                double azN = 0;
                if (north && east) azN = thetaDeg;
                if (north && !east) azN = 360.0 - thetaDeg;
                if (!north && east) azN = 180.0 - thetaDeg;
                if (!north && !east) azN = 180.0 + thetaDeg;

                azEastRad = (Math.PI / 2) - (azN * Math.PI / 180.0);
                while (azEastRad < 0) azEastRad += 2 * Math.PI;
                while (azEastRad >= 2 * Math.PI) azEastRad -= 2 * Math.PI;

                if (hadLeadingMinus) azEastRad = Commands.Normalize(azEastRad + Math.PI);
                return true;
            }

            var md = AzDms.Match(text);
            if (md.Success)
            {
                int d = int.Parse(md.Groups[1].Value, CultureInfo.InvariantCulture);
                int m = string.IsNullOrEmpty(md.Groups[2].Value) ? 0 : int.Parse(md.Groups[2].Value, CultureInfo.InvariantCulture);
                double s = string.IsNullOrEmpty(md.Groups[3].Value) ? 0 : double.Parse(md.Groups[3].Value, CultureInfo.InvariantCulture);
                double azNdeg = d + m / 60.0 + s / 3600.0;

                azEastRad = (Math.PI / 2) - (azNdeg * Math.PI / 180.0);
                while (azEastRad < 0) azEastRad += 2 * Math.PI;
                while (azEastRad >= 2 * Math.PI) azEastRad -= 2 * Math.PI;

                if (hadLeadingMinus) azEastRad = Commands.Normalize(azEastRad + Math.PI);
                return true;
            }

            var mn = AzNumeric.Match(text);
            if (mn.Success)
            {
                string degStr = mn.Groups[1].Value;
                string fracStr = mn.Groups[2].Success ? mn.Groups[2].Value : "";

                double degVal;
                if (!string.IsNullOrEmpty(fracStr) && (fracStr.Length == 2 || fracStr.Length == 4))
                {
                    if (!TryParseDegreesMmSs(degStr, fracStr, out degVal))
                        return false;
                }
                else
                {
                    if (!double.TryParse(degStr + (string.IsNullOrEmpty(fracStr) ? "" : "." + fracStr),
                                         NumberStyles.Float, CultureInfo.InvariantCulture, out degVal))
                        return false;
                }

                double azNdeg = degVal % 360.0;
                if (azNdeg < 0) azNdeg += 360.0;

                azEastRad = (Math.PI / 2) - (azNdeg * Math.PI / 180.0);
                while (azEastRad < 0) azEastRad += 2 * Math.PI;
                while (azEastRad >= 2 * Math.PI) azEastRad -= 2 * Math.PI;

                if (hadLeadingMinus) azEastRad = Commands.Normalize(azEastRad + Math.PI);
                return true;
            }

            return false;
        }

        public static bool TryParseToAzimuthRad(string text, out double azEastRad)
            => TryParseToAzimuthRadWithFlags(text, out azEastRad, out _);

        static readonly Regex BearingDistLine = new Regex(@"^\s*(?<b>.+?)\s*(?:,|\s)\s*(?<d>[+-]?\d+(?:\.\d+)?)\s*$", RegexOptions.Compiled);
        public static bool TryParseBearingAndDistance(string line, out double azEast, out double dist)
        {
            azEast = 0; dist = 0;
            if (string.IsNullOrWhiteSpace(line)) return false;
            var m = BearingDistLine.Match(line);
            if (!m.Success) return false;

            var btxt = m.Groups["b"].Value;
            var dtxt = m.Groups["d"].Value;

            if (!TryParseToAzimuthRadWithFlags(btxt, out double az, out bool hadNegBearing)) return false;
            if (!double.TryParse(dtxt, NumberStyles.Float, CultureInfo.InvariantCulture, out double d)) return false;

            bool distNegative = d < 0;
            if (hadNegBearing && distNegative) return false;

            if (distNegative) { d = -d; az = Commands.Normalize(az + Math.PI); }

            azEast = az; dist = d;
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
            foreach (var (a, b, w) in pairs) { W += w; axs += w * a.X; ays += w * a.Y; bxs += w * b.X; bys += w * b.Y; }
            var ca = new Point2d(axs / W, ays / W); var cb = new Point2d(bxs / W, bys / W);

            double Sxx = 0, Sxy = 0, Syx = 0, Syy = 0, S1 = 0;
            foreach (var (a, b, w) in pairs)
            {
                var A = a - ca; var B = b - cb;
                Sxx += w * A.X * B.X; Sxy += w * A.X * B.Y;
                Syx += w * A.Y * B.X; Syy += w * A.Y * B.Y;
                S1 += w * (A.X * A.X + A.Y * A.Y);
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
        public static List<Point2d> DriveOpenFrom(Point2d startHeld, IReadOnlyList<PlanEdge> chain)
        {
            var pts = new List<Point2d> { startHeld };
            var p = startHeld;
            foreach (var e in chain)
            {
                p = new Point2d(
                    p.X + e.Distance * Math.Cos(e.BearingRad),
                    p.Y + e.Distance * Math.Sin(e.BearingRad));
                pts.Add(p);
            }
            return pts;
        }
    }

    internal enum SolveTag { Held, BearingDistance, BearingIntersection, Proportion, CompassTraverse, SimilarityBetween, SimilarityBetweenLocked }

    internal static class PlanBuild
    {
        // Convenience overload: compute closure internally
        public static void SaveAndDraw(PlanData plan, string tag, string planName = null)
        {
            double closure = 0.0;
            if (plan.VertexXY.Count > 0)
            {
                var a = new Point2d(plan.VertexXY[0].X, plan.VertexXY[0].Y);
                var last = plan.VertexXY[plan.VertexXY.Count - 1];
                var b = new Point2d(last.X, last.Y);
                closure = Cad.Distance(a, b);
            }
            SaveAndDraw(plan, closure, tag, planName);
        }
        // Build a PlanData from a sequence of edges (length in metres, azimuth in radians, east-of-X axis).
        // Returns the open-closure length between the last and first vertex through the out parameter.
        public static PlanData BuildPlanFromEdges(
            IEnumerable<(double L, double az, string btxt)> edges,
            out double closureLen)
        {
            var plan = new PlanData
            {
                CombinedScaleFactor = 1.0,
                Closed = false
            };

            // Start at origin with P1
            var curX = 0.0;
            var curY = 0.0;
            plan.VertexIds.Add("P1");
            plan.VertexXY.Add(new XY(curX, curY));

            int i = 0;
            foreach (var (L, az, btxt) in edges)
            {
                string fromId = $"P{i + 1}";
                string toId = $"P{i + 2}";

                // step forward
                curX += L * Math.Cos(az);
                curY += L * Math.Sin(az);

                plan.Edges.Add(new PlanEdge
                {
                    FromId = fromId,
                    ToId = toId,
                    Distance = L,
                    BearingRad = az,
                    Locked = false,
                    BearingText = btxt ?? string.Empty
                });

                plan.VertexIds.Add(toId);
                plan.VertexXY.Add(new XY(curX, curY));
                i++;
            }

            // open-closure (distance last→first)
            var lastV = plan.VertexXY[plan.VertexXY.Count - 1];
            double dx = lastV.X - plan.VertexXY[0].X;
            double dy = lastV.Y - plan.VertexXY[0].Y;
            closureLen = Math.Sqrt(dx * dx + dy * dy);

            // consider it "closed" if it actually returns to start
            plan.Closed = closureLen < 1e-6;
            return plan;
        }

        // Keep old signature for existing call sites
        public static void SaveAndDraw(PlanData plan, double closureLen, string tag)
            => SaveAndDraw(plan, closureLen, tag, null);

        // New: name-aware save
        public static void SaveAndDraw(PlanData plan, double closureLen, string tag, string planName)
        {
            string path = Config.PlanJsonPath();

            // Write/merge into the named-plan store (auto-migrates legacy file on first save)
            PlanIO.Upsert(path, planName ?? "Active", plan);
            Cad.WriteMessageSafe($"\nWrote plan to: {path}");

            // Draw boundary + summary (unchanged)
            Cad.EnsureLayer(Config.LayerPdf, aci: 4);

            var pl = new Polyline { Layer = Config.LayerPdf, Closed = plan.Closed };
            for (int i = 0; i < plan.VertexXY.Count; i++)
            {
                var v = plan.VertexXY[i];
                pl.AddVertexAt(i, new Point2d(v.X, v.Y), 0, 0, 0);
            }
            Cad.AddToModelSpace(pl);

            var sb = new StringBuilder();
            sb.AppendLine($"\\LPlan ({tag}) Summary\\l");
            sb.AppendLine($"Vertices: {plan.Count}");
            sb.AppendLine($"Perimeter: {plan.Edges.Sum(e => e.Distance):0.###} m");
            sb.AppendLine($"Closure: {closureLen:0.###} m");
            var mt = new MText { Contents = sb.ToString(), Location = new Point3d(0, 0, 0), TextHeight = 2.5, Layer = Config.LayerPdf };
            Cad.AddToModelSpace(mt);
        }
    }

    // ---------------------------- EVILINK Form ----------------------------
    public class EvilinkForm : Form
    {
        public EvidenceLinks Links { get; private set; }

        private readonly DataGridView grid = new DataGridView();
        private readonly Button btnSave = new Button();
        private readonly Button btnCancel = new Button();
        private readonly Label lbl = new Label();

        public EvilinkForm(EvidenceLinks links)
        {
            this.Text = "Evidence ↔ Plan Linker (EVILINK)";
            this.Width = 860;
            this.Height = 540;
            this.StartPosition = FormStartPosition.CenterScreen;

            Links = links ?? new EvidenceLinks();

            grid.Dock = DockStyle.Top;
            grid.Height = 420;
            grid.AutoGenerateColumns = false;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.AllowUserToOrderColumns = false;
            grid.RowHeadersVisible = false;
            grid.DataSource = Links.Points;

            var colNo = new DataGridViewTextBoxColumn { HeaderText = "#", ReadOnly = true, Width = 48 };
            grid.Columns.Add(colNo);

            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Handle", DataPropertyName = "Handle", ReadOnly = true, Width = 120 });

            var colX = new DataGridViewTextBoxColumn { HeaderText = "X", DataPropertyName = "X", ReadOnly = true, Width = 120 };
            colX.DefaultCellStyle.Format = "0.###";
            colX.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            grid.Columns.Add(colX);

            var colY = new DataGridViewTextBoxColumn { HeaderText = "Y", DataPropertyName = "Y", ReadOnly = true, Width = 120 };
            colY.DefaultCellStyle.Format = "0.###";
            colY.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            grid.Columns.Add(colY);

            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "PlanID (P#)", DataPropertyName = "PlanId", Width = 110 });
            grid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "Held", DataPropertyName = "Held", Width = 60 });

            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "EvidenceType (block)", DataPropertyName = "EvidenceType", ReadOnly = true, Width = 180 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Priority", DataPropertyName = "Priority", Width = 70 });

            grid.DataBindingComplete += (s, e) =>
            {
                for (int i = 0; i < Links.Points.Count && i < grid.Rows.Count; i++)
                    grid.Rows[i].Cells[0].Value = (i + 1).ToString();
            };

            btnSave.Text = "Save";
            btnSave.Width = 110;
            btnSave.Top = 450;
            btnSave.Left = this.Width - 270;
            btnSave.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnSave.Click += (s, e) => { this.DialogResult = DialogResult.OK; this.Close(); };

            btnCancel.Text = "Cancel";
            btnCancel.Width = 110;
            btnCancel.Top = 450;
            btnCancel.Left = this.Width - 150;
            btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            lbl.Text = "Tip: Use the #s shown on-screen and in the drawing labels to assign PlanIDs (P1, P2, ...).";
            lbl.Top = 425;
            lbl.Left = 10;
            lbl.Width = this.Width - 320;
            lbl.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            this.Controls.Add(grid);
            this.Controls.Add(btnSave);
            this.Controls.Add(btnCancel);
            this.Controls.Add(lbl);
        }
    }

    // ---------------------------- Commands ----------------------------
    public partial class Commands
    {
        internal static double Normalize(double a)
        {
            while (a < 0) a += 2 * Math.PI;
            while (a >= 2 * Math.PI) a -= 2 * Math.PI;
            return a;
        }
        static double AngleBetween(double a1, double a2)
        {
            double d = Math.Abs(Normalize(a1 - a2));
            if (d > Math.PI) d = 2 * Math.PI - d;
            return d;
        }
        static bool TryLineIntersection(Point2d A, Vector2d dirA, Point2d B, Vector2d dirB, out Point2d P)
        {
            double cross = dirA.X * dirB.Y - dirA.Y * dirB.X;
            if (Math.Abs(cross) < 1e-12) { P = default; return false; }
            var c = new Vector2d(B.X - A.X, B.Y - A.Y);
            double t = (c.X * dirB.Y - c.Y * dirB.X) / cross;
            P = new Point2d(A.X + t * dirA.X, A.Y + t * dirA.Y);
            return true;
        }
        private static bool TryParseDistance(string raw, out double val)
        {
            val = 0;
            if (string.IsNullOrWhiteSpace(raw)) return false;
            var cleaned = Regex.Replace(raw, @"[^\d\.\-+]", "");
            return double.TryParse(cleaned, NumberStyles.Float, CultureInfo.InvariantCulture, out val);
        }
        private static string WriteTextWithFallback(string primaryPath, string contents)
        {
            try
            {
                File.WriteAllText(primaryPath, contents, Encoding.UTF8);
                return primaryPath;
            }
            catch (IOException ex)
            {
                Cad.WriteMessageSafe($"\nReport file locked: {Path.GetFileName(primaryPath)} — {ex.Message}");

                string dir = Path.GetDirectoryName(primaryPath) ?? Config.CurrentDwgFolder();
                string name = Path.GetFileNameWithoutExtension(primaryPath);
                string ext = Path.GetExtension(primaryPath);
                string ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string alt = Path.Combine(dir, $"{name}_{ts}{ext}");

                try { File.WriteAllText(alt, contents, Encoding.UTF8); return alt; }
                catch (IOException)
                {
                    using (var sfd = new SaveFileDialog
                    {
                        Title = "Save Adjustment Report CSV",
                        Filter = "CSV files (*.csv)|*.csv|All files|*.*",
                        FileName = Path.GetFileName(primaryPath),
                        InitialDirectory = dir
                    })
                    {
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            File.WriteAllText(sfd.FileName, contents, Encoding.UTF8);
                            return sfd.FileName;
                        }

                        string tmp = Path.Combine(Path.GetTempPath(), $"{name}_{ts}{ext}");
                        File.WriteAllText(tmp, contents, Encoding.UTF8);
                        Cad.WriteMessageSafe($"\nWrote report to temp: {tmp}");
                        return tmp;
                    }
                }
            }
        }
        [CommandMethod("EVIHARVEST")]
        public void EvidenceHarvest()
        {
            var ed = Cad.Ed;
            Cad.WriteMessageSafe("\nEVIHARVEST — select evidence blocks, whitelist filter, number, label, and save.");

            var defaultNames = new[]
            {
                "PLI","plhub","plspike","PLIBAR","FDI","TEMP","fdhub","FDIBAR","fdspike","WITF"
            };
            var allowed = new HashSet<string>(defaultNames, StringComparer.OrdinalIgnoreCase);

            string wlPath = Path.Combine(Config.CurrentDwgFolder(), $"{Config.Stem()}_BlockWhitelist.json");
            try
            {
                if (File.Exists(wlPath))
                {
                    var persisted = Json.Load<List<string>>(wlPath) ?? new List<string>();
                    foreach (var n in persisted) if (!string.IsNullOrWhiteSpace(n)) allowed.Add(n.Trim());
                }
            }
            catch { /* ignore */ }

            var addOpt = new PromptStringOptions("\nAdditional block names to recognize (comma-separated, Enter for none): ")
            { AllowSpaces = true };
            var addRes = ed.GetString(addOpt);
            var addedNames = new List<string>();
            if (addRes.Status == PromptStatus.OK && !string.IsNullOrWhiteSpace(addRes.StringResult))
            {
                foreach (var raw in addRes.StringResult.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var n = raw.Trim();
                    if (n.Length == 0) continue;
                    if (allowed.Add(n)) addedNames.Add(n);
                }
                if (addedNames.Count > 0) Json.Save(wlPath, allowed.ToList());
            }

            var psr = Cad.PromptSelect("Select evidence blocks (window/crossing recommended)",
                new TypedValue((int)DxfCode.Start, "INSERT"));
            if (psr.Status != PromptStatus.OK) return;

            string evPath = Config.EvidenceJsonPath();
            EvidenceLinks links = File.Exists(evPath) ? (Json.Load<EvidenceLinks>(evPath) ?? new EvidenceLinks())
                                                      : new EvidenceLinks();

            var byHandle = new Dictionary<string, EvidencePoint>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in links.Points) if (!string.IsNullOrEmpty(p.Handle)) byHandle[p.Handle] = p;

            const string LabelLayer = "EVIDENCE_TAGS";
            const double LabelHeight = 5.0;
            Cad.EnsureLayer(LabelLayer, aci: 2);

            int added = 0, labeled = 0;
            Cad.WithLockedDoc(() =>
            {
                using (var tr = Cad.Db.TransactionManager.StartTransaction())
                {
                    var ms = (BlockTableRecord)tr.GetObject(
                        ((BlockTable)tr.GetObject(Cad.Db.BlockTableId, OpenMode.ForRead))[BlockTableRecord.ModelSpace],
                        OpenMode.ForWrite);

                    bool HasLabelNear(string text, Point2d p, double tol)
                    {
                        double tol2 = tol * tol;
                        foreach (ObjectId eid in ms)
                        {
                            if (tr.GetObject(eid, OpenMode.ForRead) is not DBText dt) continue;
                            if (!string.Equals(dt.Layer, LabelLayer, StringComparison.OrdinalIgnoreCase)) continue;
                            if (!string.Equals((dt.TextString ?? "").Trim(), text, StringComparison.Ordinal)) continue;
                            var pos3 = dt.AlignmentPoint.IsEqualTo(Point3d.Origin) ? dt.Position : dt.AlignmentPoint;
                            double dx = pos3.X - p.X, dy = pos3.Y - p.Y;
                            if (dx * dx + dy * dy <= tol2) return true;
                        }
                        return false;
                    }

                    foreach (var id in psr.Value.GetObjectIds())
                    {
                        var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                        if (br == null) continue;

                        string bname = GetBlockName(br, tr);
                        if (!allowed.Contains(bname)) continue;

                        string h = br.Handle.ToString();
                        if (!byHandle.TryGetValue(h, out var ep))
                        {
                            bool strong = string.Equals(bname, "FDI", StringComparison.OrdinalIgnoreCase);
                            bool weak = string.Equals(bname, "fdspike", StringComparison.OrdinalIgnoreCase);

                            ep = new EvidencePoint
                            {
                                PlanId = "",
                                Handle = h,
                                X = br.Position.X,
                                Y = br.Position.Y,
                                EvidenceType = bname,
                                Held = strong,
                                Priority = strong ? 2 : (weak ? 0 : 1)
                            };
                            links.Points.Add(ep);
                            byHandle[h] = ep;
                            added++;
                        }

                        int num = links.Points.FindIndex(p => string.Equals(p.Handle, h, StringComparison.OrdinalIgnoreCase)) + 1;
                        if (num <= 0) num = 1;

                        var p2 = new Point2d(ep.X, ep.Y);
                        string labelText = $"#{num}";
                        if (!HasLabelNear(labelText, p2, tol: LabelHeight * 0.35))
                        {
                            var txt = new DBText
                            {
                                Layer = LabelLayer,
                                TextString = labelText,
                                Height = LabelHeight,
                                Justify = AttachmentPoint.MiddleCenter,
                                AlignmentPoint = new Point3d(p2.X, p2.Y, 0),
                                Position = new Point3d(p2.X, p2.Y, 0)
                            };
                            ms.AppendEntity(txt);
                            tr.AddNewlyCreatedDBObject(txt, true);
                            txt.AdjustAlignment(Cad.Db);
                            labeled++;
                        }
                    }

                    tr.Commit();
                }
            });

            Json.Save(evPath, links);
            Cad.WriteMessageSafe($"\nHarvested {added} new evidence; labeled {labeled}. Saved: {evPath}\n" +
                            "Run EVILINK (table) to assign PlanIDs and Held.");
        }

        public static void RunAlsAdjFromUI()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc != null)
                doc.SendStringToExecute("_.ALSADJ ", true, false, false);
        }

        [CommandMethod("PLANLINK")]
        public void PlanLink()
        {
            var ed = Cad.Ed;
            Cad.WriteMessageSafe("\nPLANLINK — align imported plan entities to field control (similarity fit).");

            var psr = Cad.PromptSelect("Select imported plan entities (from PDFIMPORT)");
            if (psr.Status != PromptStatus.OK) return;
            var ids = psr.Value.GetObjectIds();
            if (ids == null || ids.Length == 0) { Cad.WriteMessageSafe("\nNothing selected."); return; }

            Cad.EnsureLayer(Config.LayerPdf, aci: 4);
            bool wasLocked = Cad.UnlockIfLocked(Config.LayerPdf);

            try
            {
                Cad.SetLayer(ids, Config.LayerPdf);

                var pairs = new List<(Point2d from, Point2d to, double w)>();
                for (int i = 1; i <= 5; i++)
                {
                    var p1 = Cad.GetPoint($"Pick SOURCE point ON PLAN (pair {i}, Enter to finish)", allowNone: true);
                    if (p1.Status != PromptStatus.OK)
                    {
                        if (pairs.Count >= 2) break;
                        Cad.WriteMessageSafe("\nNeed at least 2 pairs.");
                        return;
                    }
                    var p2 = Cad.GetPoint($"Pick MATCHING FIELD point (target for pair {i})");
                    if (p2.Status != PromptStatus.OK) { Cad.WriteMessageSafe("\nCancelled."); return; }

                    pairs.Add((new Point2d(p1.Value.X, p1.Value.Y),
                               new Point2d(p2.Value.X, p2.Value.Y),
                               1.0));
                }

                if (!SimilarityFit.Solve(pairs, false, out double kf, out double cf, out double sf, out Vector2d tf))
                {
                    Cad.WriteMessageSafe("\nCould not solve similarity transform (check control pairs).");
                    return;
                }

                var (k, c, s, t, usedScale) = DecideScalingALS(pairs, kf, cf, sf, tf, ed);

                var m = SimilarityFit.ToMatrix(k, c, s, t);
                Cad.TransformEntities(ids, m);

                Cad.WriteMessageSafe(usedScale
                    ? $"\nPLANLINK done (scale applied). k={k:0.000000}, rot={Math.Atan2(s, c) * 180 / Math.PI:0.000}°, t=({t.X:0.###},{t.Y:0.###})."
                    : $"\nPLANLINK done (scale locked to 1.0). rot={Math.Atan2(s, c) * 180 / Math.PI:0.000}°, t=({t.X:0.###},{t.Y:0.###}).");
            }
            finally
            {
                Cad.LockLayer(Config.LayerPdf, true);
                Cad.WriteMessageSafe($" Layer {Config.LayerPdf} locked.");
            }
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

        [CommandMethod("PLANCOGO")]
        public void PlanCogo()
        {
            var ed = Cad.Ed;
            Cad.WriteMessageSafe("\nPLANCOGO — build plan graph from numeric bearings/distances (grid azimuths).");

            var mode = new PromptKeywordOptions("\nMode [Manual/CSV] <Manual>: ");
            mode.Keywords.Add("Manual"); mode.Keywords.Add("CSV"); mode.AllowNone = true;
            var km = ed.GetKeywords(mode);
            bool useCsv = (km.Status == PromptStatus.OK && km.StringResult == "CSV");

            var plan = new PlanData { Closed = true, CombinedScaleFactor = 1.0 };
            double csf = 1.0;
            var rawEntries = new List<UserCogoInput>();

            if (useCsv)
            {
                using (var dlg = new OpenFileDialog { Title = "Select COGO CSV", Filter = "CSV files (*.csv)|*.csv|All files|*.*", Multiselect = false })
                {
                    if (dlg.ShowDialog() != DialogResult.OK) { Cad.WriteMessageSafe("\nCancelled."); return; }

                    string[] lines = File.ReadAllLines(dlg.FileName);
                    if (lines.Length == 0) { Cad.WriteMessageSafe("\nCSV is empty."); return; }

                    foreach (var raw in lines.Take(10))
                    {
                        var t = raw.Trim();
                        if (!t.StartsWith("#", StringComparison.Ordinal)) continue;
                        var kv = t.TrimStart('#').Trim();
                        var parts = kv.Split(new[] { '=' }, 2);
                        for (int i = 0; i < parts.Length; i++) parts[i] = parts[i].Trim();
                        if (parts.Length == 2 && parts[0].Equals("CSF", StringComparison.OrdinalIgnoreCase))
                        {
                            if (double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var v) && v > 0.5 && v < 1.5) csf = v;
                        }
                    }
                    plan.CombinedScaleFactor = csf;

                    bool applyCsf = true;
                    if (Math.Abs(csf - 1.0) > 1e-6)
                    {
                        var k = new PromptKeywordOptions($"\nCSV header CSF={csf:0.######}. Apply to distances? [Yes/No] <Yes>: ");
                        k.Keywords.Add("Yes"); k.Keywords.Add("No"); k.AllowNone = true;
                        var kr = ed.GetKeywords(k);
                        applyCsf = !(kr.Status == PromptStatus.OK && kr.StringResult == "No");
                    }

                    var cur = new Point2d(0, 0);
                    int autoIdx = 1;

                    foreach (var raw in lines)
                    {
                        var line = raw.Trim();
                        if (line.Length == 0 || line.StartsWith("#")) continue;

                        var parts = line.Split(',').Select(s => s.Trim()).ToArray();

                        string fromId = "", toId = "", btxt = "", dtxt = "";

                        if (parts.Length == 4)
                        {
                            fromId = parts[0]; toId = parts[1]; btxt = parts[2]; dtxt = parts[3];
                        }
                        else if (parts.Length == 2)
                        {
                            fromId = $"P{autoIdx}"; toId = $"P{autoIdx + 1}"; btxt = parts[0]; dtxt = parts[1];
                        }
                        else
                        {
                            Cad.WriteMessageSafe($"\nSkip row (need 2 or 4 columns): {line}");
                            continue;
                        }

                        if (!BearingParserEx.TryParseToAzimuthRadWithFlags(btxt, out double az, out bool hadNegBearing))
                        { Cad.WriteMessageSafe($"\nInvalid bearing: {btxt}"); return; }

                        if (!TryParseDistance(dtxt, out double dist) || Math.Abs(dist) <= 0)
                        { Cad.WriteMessageSafe($"\nInvalid distance: {dtxt}"); return; }

                        bool distNegative = dist < 0;
                        if (hadNegBearing && distNegative)
                        { Cad.WriteMessageSafe($"\nCannot use '-' on BOTH bearing and distance in: {line}"); return; }

                        double localAz = az;
                        double dAbs = Math.Abs(dist);
                        if (distNegative) localAz = Normalize(localAz + Math.PI);

                        double dUse = applyCsf ? dAbs * csf : dAbs;

                        if (plan.VertexIds.Count == 0)
                        {
                            plan.VertexIds.Add(fromId);
                            plan.VertexXY.Add(new XY(cur.X, cur.Y));
                        }

                        cur = new Point2d(cur.X + dUse * Math.Cos(localAz), cur.Y + dUse * Math.Sin(localAz));
                        plan.Edges.Add(new PlanEdge { FromId = fromId, ToId = toId, Distance = dUse, BearingRad = localAz, Locked = false, BearingText = btxt });
                        plan.VertexIds.Add(toId);
                        plan.VertexXY.Add(new XY(cur.X, cur.Y));

                        autoIdx++;
                    }
                }
            }
            else
            {
                var s = ed.GetString(new PromptStringOptions("\nCombined Scale Factor to apply to distances (Enter for 1.0): ") { AllowSpaces = false });
                if (s.Status == PromptStatus.OK && double.TryParse(s.StringResult, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) && v > 0.5 && v < 1.5)
                    csf = v;
                plan.CombinedScaleFactor = csf;

                var cur = new Point2d(0, 0);
                int leg = 1;

                plan.VertexIds.Add("P1");
                plan.VertexXY.Add(new XY(cur.X, cur.Y));

                while (true)
                {
                    var ps = new PromptStringOptions($"\nLeg {leg} — enter \"bearing distance\" (e.g., 256.0440 550.50  or  N45.3045E 550.50). Enter to finish.")
                    { AllowSpaces = true };
                    var pr = ed.GetString(ps);

                    if (pr.Status != PromptStatus.OK || string.IsNullOrWhiteSpace(pr.StringResult))
                    {
                        if (plan.Edges.Count >= 2) break;
                        Cad.WriteMessageSafe("\nNeed at least 2 legs."); return;
                    }

                    double az, dist;
                    string btxtForSave;

                    var m = new Regex(@"^\s*(?<b>.+?)\s*(?:,|\s)\s*(?<d>[+-]?\d+(?:\.\d+)?)\s*$").Match(pr.StringResult);
                    if (m.Success)
                    {
                        btxtForSave = m.Groups["b"].Value.Trim();
                        var dtxt = m.Groups["d"].Value.Trim();

                        if (!BearingParserEx.TryParseToAzimuthRadWithFlags(btxtForSave, out az, out bool hadNegBearing))
                        { Cad.WriteMessageSafe("\nCould not parse bearing."); return; }
                        if (!double.TryParse(dtxt, NumberStyles.Float, CultureInfo.InvariantCulture, out dist))
                        { Cad.WriteMessageSafe("\nCould not parse distance."); return; }

                        bool distNeg = dist < 0;
                        if (hadNegBearing && distNeg) { Cad.WriteMessageSafe("\nCannot use '-' on BOTH bearing and distance."); return; }
                        if (distNeg) { dist = -dist; az = Normalize(az + Math.PI); }
                    }
                    else
                    {
                        btxtForSave = pr.StringResult.Trim();
                        if (!BearingParserEx.TryParseToAzimuthRadWithFlags(btxtForSave, out az, out bool hadNegBearing))
                        { Cad.WriteMessageSafe("\nCould not parse bearing."); return; }
                        var dd = ed.GetDouble(new PromptDoubleOptions($"Distance for leg {leg} (m): ") { AllowNegative = true, AllowZero = false });
                        if (dd.Status != PromptStatus.OK) return;
                        dist = dd.Value;
                        bool distNeg = dist < 0;
                        if (hadNegBearing && distNeg) { Cad.WriteMessageSafe("\nCannot use '-' on BOTH bearing and distance."); return; }
                        if (distNeg) { dist = -dist; az = Normalize(az + Math.PI); }
                    }

                    rawEntries.Add(new UserCogoInput
                    {
                        Line = leg,
                        Point1 = $"P{leg}",
                        Point2 = $"P{leg + 1}",
                        Bearing = btxtForSave,
                        Distance = dist,
                        Azimuth = string.Empty,
                        Deflection = string.Empty
                    });

                    double dUse = dist * csf;

                    string fromId = $"P{leg}";
                    string toId = $"P{leg + 1}";

                    cur = new Point2d(cur.X + dUse * Math.Cos(az), cur.Y + dUse * Math.Sin(az));
                    plan.Edges.Add(new PlanEdge { FromId = fromId, ToId = toId, Distance = dUse, BearingRad = az, Locked = false, BearingText = btxtForSave });
                    plan.VertexIds.Add(toId);
                    plan.VertexXY.Add(new XY(cur.X, cur.Y));

                    leg++;
                }

                var kopt = new PromptKeywordOptions("\nClose the figure? [Yes/No] <Yes>: ");
                kopt.Keywords.Add("Yes"); kopt.Keywords.Add("No"); kopt.AllowNone = true;
                var kr = ed.GetKeywords(kopt);
                plan.Closed = !(kr.Status == PromptStatus.OK && kr.StringResult == "No");
            }
            string folderPath2 = Path.GetDirectoryName(AcadApp.DocumentManager.MdiActiveDocument?.Name);
            if (!string.IsNullOrEmpty(folderPath2))
            {
                if (rawEntries.Count > 0)
                {
                    // Manual mode: save the RAW user entries
                    CogoFile.SaveUserCogoInput(Path.Combine(folderPath2, "COGO"), rawEntries);
                }
                else
                {
                    // CSV mode: if no raw entries collected, save from edges (as-is)
                    var entries2 = new List<UserCogoInput>();
                    for (int i = 0; i < plan.Edges.Count; i++)
                    {
                        var e2 = plan.EdgeAt(i);
                        entries2.Add(new UserCogoInput
                        {
                            Line = i + 1,
                            Point1 = e2.FromId,
                            Point2 = e2.ToId,
                            Bearing = e2.BearingText,
                            Distance = e2.Distance,
                            Azimuth = string.Empty,
                            Deflection = string.Empty
                        });
                    }
                    CogoFile.SaveUserCogoInput(Path.Combine(folderPath2, "COGO"), entries2);
                }
            }

            PlanBuild.SaveAndDraw(plan, tag: "COGO");
            Cad.WriteMessageSafe("\nNext step: EVILINK → assign Held evidence (P#) → ALSADJ.");
        }

        private string GetTextContent(ObjectId id)
        {
            using (var tr = Cad.Db.TransactionManager.StartTransaction())
            {
                var obj = tr.GetObject(id, OpenMode.ForRead, false);
                if (obj is DBText dbt) return dbt.TextString.Trim();
                if (obj is MText mt) return mt.Contents.Trim();
                return "";
            }
        }

        [CommandMethod("PLANFROMTEXT")]
        public void PlanFromText()
        {
            var ed = Cad.Ed;
            Cad.WriteMessageSafe("\nPLANFROMTEXT — extract bearing/distances by selecting text objects.");

            var calls = new List<(double azRad, double dist, string btxt, string dtxt)>();
            int leg = 1;

            while (true)
            {
                var peo = new PromptEntityOptions($"\nLeg {leg}: Select BEARING text (Enter to finish)");
                peo.SetRejectMessage("\n  Only text (DBText/MText) allowed.");
                peo.AddAllowedClass(typeof(DBText), true);
                peo.AddAllowedClass(typeof(MText), true);
                var perB = ed.GetEntity(peo);

                if (perB.Status != PromptStatus.OK)
                {
                    if (calls.Count >= 2) break;
                    Cad.WriteMessageSafe("\nNeed at least 2 calls before finishing.");
                    return;
                }

                string bearingRaw = GetTextContent(perB.ObjectId);
                if (!BearingParserEx.TryParseToAzimuthRadWithFlags(bearingRaw, out double az, out bool hadNegBearing))
                {
                    Cad.WriteMessageSafe($"\nCannot parse bearing from \"{bearingRaw}\".");
                    return;
                }

                var ped = new PromptEntityOptions($"Leg {leg}: Select DISTANCE text");
                ped.SetRejectMessage("\n  Only text (DBText/MText) allowed.");
                ped.AddAllowedClass(typeof(DBText), true);
                ped.AddAllowedClass(typeof(MText), true);
                var perD = ed.GetEntity(ped);
                if (perD.Status != PromptStatus.OK) { Cad.WriteMessageSafe("\nCancelled."); return; }

                string distRaw = GetTextContent(perD.ObjectId);
                if (!TryParseDistance(distRaw, out double distVal) || Math.Abs(distVal) <= 0)
                {
                    Cad.WriteMessageSafe($"\nCannot parse distance from \"{distRaw}\".");
                    return;
                }
                bool distNeg = distVal < 0;
                if (hadNegBearing && distNeg) { Cad.WriteMessageSafe("\nCannot use '-' on BOTH bearing and distance."); return; }
                if (distNeg) { distVal = -distVal; az = Normalize(az + Math.PI); }

                calls.Add((az, distVal, bearingRaw, distRaw));
                leg++;
            }

            var closeOpt = new PromptKeywordOptions("\nClose the figure? [Yes/No] <Yes>: ");
            closeOpt.Keywords.Add("Yes");
            closeOpt.Keywords.Add("No");
            closeOpt.AllowNone = true;
            var cr = ed.GetKeywords(closeOpt);
            bool closeIt = !(cr.Status == PromptStatus.OK && cr.StringResult == "No");

            var csvRows = new List<string>();
            var edges = new List<(double L, double az, string btxt)>();
            for (int i = 0; i < calls.Count; i++)
            {
                var c = calls[i];
                string fromId = $"P{i + 1}";
                string toId = $"P{i + 2}";
                edges.Add((c.dist, c.azRad, c.btxt));
                csvRows.Add($"{fromId},{toId},{c.btxt},{c.dtxt}");
            }

            if (closeIt)
            {
                double sumX = 0, sumY = 0;
                for (int i = 0; i < calls.Count; i++)
                {
                    sumX += calls[i].dist * Math.Cos(calls[i].azRad);
                    sumY += calls[i].dist * Math.Sin(calls[i].azRad);
                }
                double dx = -sumX;
                double dy = -sumY;
                double closureDist = Math.Sqrt(dx * dx + dy * dy);
                double closureAz = Math.Atan2(dy, dx);
                if (closureAz < 0) closureAz += 2 * Math.PI;

                bool useCompact = calls.Count(c => BearingParserEx.LooksQuadCompact(c.btxt)) > calls.Count / 2;
                string closTxt = useCompact ? Cad.FormatBearingCompact(closureAz) : Cad.FormatBearing(closureAz);
                edges.Add((closureDist, closureAz, closTxt));
                string fromLast = $"P{calls.Count + 1}";
                csvRows.Add($"{fromLast},P1,{closTxt},{closureDist.ToString("0.###", CultureInfo.InvariantCulture)}");
            }

            string csvPath = Path.Combine(Config.CurrentDwgFolder(), $"{Config.Stem()}_SelectedCalls.csv");
            File.WriteAllText(csvPath, "FromId,ToId,Bearing,Distance\n" + string.Join("\n", csvRows));
            Cad.WriteMessageSafe($"\nWrote calls CSV: {csvPath}");

            var pdata = PlanBuild.BuildPlanFromEdges(edges, out double closureLen);

            string folderPath3 = Path.GetDirectoryName(AcadApp.DocumentManager.MdiActiveDocument?.Name);
            if (!string.IsNullOrEmpty(folderPath3))
            {
                var entries3 = new List<UserCogoInput>();
                for (int i = 0; i < edges.Count; i++)
                {
                    var e3 = edges[i];
                    entries3.Add(new UserCogoInput
                    {
                        Line = i + 1,
                        Point1 = "P" + (i + 1),
                        Point2 = "P" + (i + 2),
                        Bearing = e3.btxt,
                        Distance = e3.L,
                        Azimuth = string.Empty,
                        Deflection = string.Empty
                    });
                }
                CogoFile.SaveUserCogoInput(Path.Combine(folderPath3, "COGO"), entries3);
            }

            PlanBuild.SaveAndDraw(pdata, closureLen, "TextSelection");

            Cad.WriteMessageSafe("\nNext: run EVILINK, assign Held evidence to P# IDs, then ALSADJ.");
        }

        [CommandMethod("PLANEXTRACT")]
        public void PlanExtract()
        {
            var ed = Cad.Ed;
            Cad.WriteMessageSafe("\nPLANEXTRACT (legacy) — PDF linework extraction. Prefer PLANCOGO or PLANFROMTEXT.");

            var peo = new PromptEntityOptions("\nSelect plan boundary polyline: ");
            peo.SetRejectMessage("\nOnly Polyline entities are supported.");
            peo.AddAllowedClass(typeof(Polyline), true);
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            var plan = new PlanData();
            using (var tr = Cad.Db.TransactionManager.StartTransaction())
            {
                var pl = (Polyline)tr.GetObject(per.ObjectId, OpenMode.ForRead);
                if (!pl.Closed) Cad.WriteMessageSafe("\nWarning: selected polyline is not closed; proceeding.");
                int n = pl.NumberOfVertices;
                if (n < 3) { Cad.WriteMessageSafe("\nPolyline has <3 vertices; aborting."); return; }

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
                    plan.Edges.Add(new PlanEdge { FromId = plan.VertexIds[i], ToId = plan.VertexIds[j], Distance = dist, BearingRad = az, Locked = false, BearingText = Cad.FormatBearing(az) });
                }
                tr.Commit();
            }

            PlanBuild.SaveAndDraw(plan, tag: "Extract");
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

        [CommandMethod("EVILINK")]
        public void Evilink()
        {
            Cad.WriteMessageSafe("\nEVILINK — edit evidence table (assign PlanIDs / Held).");

            string evPath = Config.EvidenceJsonPath();

            if (File.Exists(evPath))
            {
                var links = Json.Load<EvidenceLinks>(evPath) ?? new EvidenceLinks();
                if (links.Points.Count > 0)
                {
                    using (var frm = new EvilinkForm(links))
                    {
                        var res = AcadApp.ShowModalDialog(frm);  // .NET 4.8 signature
                        if (res == DialogResult.OK)
                        {
                            Json.Save(evPath, frm.Links);
                            Cad.WriteMessageSafe($"\nSaved links: {evPath}");
                        }
                        else
                        {
                            Cad.WriteMessageSafe("\nCancelled (no changes saved).");
                        }
                    }
                    return;
                }
            }

            Cad.WriteMessageSafe("\nNo evidence list found. Run EVIHARVEST first, then EVILINK.");
        }

        private bool ShowEvilinkTableAndSave(EvidenceLinks links)
        {
            using (var frm = new EvilinkForm(links))
            {
                var res = AcadApp.ShowModalDialog(frm);
                if (res == DialogResult.OK)
                {
                    Json.Save(Config.EvidenceJsonPath(), frm.Links);
                    return true;
                }
                return false;
            }
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

        [CommandMethod("PLANLOCK")]
        public void PlanLockToggle()
        {
            var ed = Cad.Ed;
            string planPath = Config.PlanJsonPath();
            if (!File.Exists(planPath))
            {
                Cad.WriteMessageSafe($"\nMissing: {planPath}. Run PLANEXTRACT/PLANCOGO/PLANFROMTEXT.");
                return;
            }

            var plan = PlanIO.LoadActive(planPath);
            if (plan.Edges.Count == 0)
            {
                Cad.WriteMessageSafe("\nNo legs.");
                return;
            }

            var opt = new PromptIntegerOptions($"\nLeg # to toggle lock (1..{plan.Edges.Count}): ")
            {
                AllowNone = false,
                LowerLimit = 1,
                UpperLimit = plan.Edges.Count
            };
            var res = ed.GetInteger(opt);
            if (res.Status != PromptStatus.OK) return;

            int idx = res.Value - 1;
            plan.Edges[idx].Locked = !plan.Edges[idx].Locked;

            var store = PlanIO.LoadStore(planPath);
            PlanIO.Upsert(planPath, store.ActiveName, plan);

            var e = plan.Edges[idx];
            Cad.WriteMessageSafe($"\nLeg {res.Value} {e.FromId}->{e.ToId} is now {(e.Locked ? "LOCKED (fixed length)" : "unlocked")}.");
        }

        [CommandMethod("ALSADJ")]
        public void AlsAdj()
        {
            var ed = Cad.Ed;
            Cad.WriteMessageSafe("\nALSADJ — hold anchors, re-drive between anchors; preserves Locked leg lengths.");

            string planPath = Config.PlanJsonPath();
            if (!File.Exists(planPath)) { Cad.WriteMessageSafe($"\nMissing: {planPath}. Run PLANEXTRACT/PLANCOGO/PLANFROMTEXT."); return; }
            var plan = PlanIO.LoadActive(planPath);
            if (plan.Count < 2 || plan.Edges.Count < 1) { Cad.WriteMessageSafe("\nPlan data invalid — need at least one leg and two vertices."); return; }

            string evPath = Config.EvidenceJsonPath();
            if (!File.Exists(evPath)) { Cad.WriteMessageSafe($"\nMissing: {evPath}. Run EVILINK."); return; }
            var links = Json.Load<EvidenceLinks>(evPath) ?? new EvidenceLinks();

            Cad.RefreshEvidencePositionsFromDwg(links);
            try { Json.Save(evPath, links); } catch { /* non-fatal */ }

            var held = links.Points.Where(p => p.Held && !string.IsNullOrWhiteSpace(p.PlanId)).ToList();
            if (held.Count < 2) { Cad.WriteMessageSafe("\nNeed at least two Held evidence points."); return; }

            var synth = TraverseWhole(plan);

            var pairs = new List<(Point2d from, Point2d to, double w)>();
            var idToIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < plan.Count; i++) idToIndex[plan.VertexIds[i]] = i;

            foreach (var hp in held)
            {
                if (idToIndex.TryGetValue(hp.PlanId, out int idx))
                    pairs.Add((synth[idx], hp.XY(), 1.0));
            }
            if (pairs.Count < 2) { Cad.WriteMessageSafe("\nInsufficient valid Held pairs."); return; }

            if (!SimilarityFit.Solve(pairs, false, out double kf, out double cf, out double sf, out Vector2d tf))
            { Cad.WriteMessageSafe("\nCould not solve similarity transform from Held pairs."); return; }
            var (k, c, s, t, usedScale) = DecideScalingALS(pairs, kf, cf, sf, tf, ed);

            var adj = new Point2d[plan.Count];
            var tag = new Dictionary<int, SolveTag>();
            var isHeld = new bool[plan.Count];
            foreach (var h in held)
            {
                if (idToIndex.TryGetValue(h.PlanId, out int idx))
                {
                    adj[idx] = h.XY();
                    isHeld[idx] = true;
                    tag[idx] = SolveTag.Held;
                }
            }

            var heldIdx = held.Where(h => idToIndex.ContainsKey(h.PlanId))
                              .Select(h => idToIndex[h.PlanId])
                              .Distinct()
                              .OrderBy(i => i)
                              .ToList();

            int n = plan.Count;

            for (int hIdx = 0; hIdx < heldIdx.Count - 1; hIdx++)
            {
                int start = heldIdx[hIdx];
                int end = heldIdx[hIdx + 1];

                var chain = new List<PlanEdge>();
                for (int ii = start; ii < end; ii++) chain.Add(plan.EdgeAt(ii));

                var startXY = adj[start];
                var endXY = adj[end];

                List<Point2d> drove;
                try
                {
                    drove = TwoPointBetweenRespectLocks(startXY, endXY, chain);
                }
                catch (System.Exception ex)
                {
                    Cad.WriteMessageSafe($"\nSegment {plan.VertexIds[start]}→{plan.VertexIds[end]} cannot satisfy Locked constraint: {ex.Message}");
                    Cad.WriteMessageSafe("\nALSADJ aborted. (Release a lock or add another anchor across the locked span.)");
                    return;
                }

                int kidx = 0;
                var segTag = chain.Any(e => e.Locked) ? SolveTag.SimilarityBetweenLocked : SolveTag.SimilarityBetween;
                for (int j = start; j < end; j++) { adj[j] = drove[kidx++]; tag[j] = segTag; }
                adj[end] = drove[kidx]; if (!tag.ContainsKey(end)) tag[end] = segTag;
            }

            if (plan.Closed && heldIdx.Count >= 2)
            {
                int start = heldIdx[heldIdx.Count - 1];
                int end = heldIdx[0];

                var chain = new List<PlanEdge>();
                int i = start;
                while (i != end)
                {
                    chain.Add(plan.EdgeAt(i));
                    i = (i + 1) % n;
                }

                var startXY = adj[start];
                var endXY = adj[end];

                List<Point2d> drove;
                try
                {
                    drove = TwoPointBetweenRespectLocks(startXY, endXY, chain);
                }
                catch (System.Exception ex)
                {
                    Cad.WriteMessageSafe($"\nWrap segment {plan.VertexIds[start]}→{plan.VertexIds[end]} cannot satisfy Locked constraint: {ex.Message}");
                    Cad.WriteMessageSafe("\nALSADJ aborted. (Release a lock or add another anchor across the locked span.)");
                    return;
                }

                int kidx = 0; i = start;
                var segTag = chain.Any(e => e.Locked) ? SolveTag.SimilarityBetweenLocked : SolveTag.SimilarityBetween;
                while (i != end) { adj[i] = drove[kidx++]; tag[i] = segTag; i = (i + 1) % n; }
                adj[end] = drove[kidx]; if (!tag.ContainsKey(end)) tag[end] = segTag;
            }

            if (!plan.Closed && heldIdx.Count >= 1)
            {
                int firstHeld = heldIdx[0];
                int lastHeld = heldIdx[heldIdx.Count - 1];

                if (firstHeld > 0)
                {
                    var revChain = new List<PlanEdge>();
                    for (int i2 = firstHeld - 1; i2 >= 0; i2--)
                    {
                        var e = plan.EdgeAt(i2);
                        revChain.Add(new PlanEdge { Distance = e.Distance, BearingRad = Normalize(e.BearingRad + Math.PI), Locked = e.Locked, BearingText = e.BearingText });
                    }
                    var droveLeft = TraverseSolver.DriveOpenFrom(adj[firstHeld], revChain);
                    for (int kidx = 1; kidx < droveLeft.Count; kidx++)
                    {
                        int vi = firstHeld - kidx;
                        adj[vi] = droveLeft[kidx];
                        if (!tag.ContainsKey(vi)) tag[vi] = SolveTag.BearingDistance;
                    }
                }

                if (lastHeld < n - 1)
                {
                    var fwdChain = new List<PlanEdge>();
                    for (int i2 = lastHeld; i2 < n - 1; i2++) fwdChain.Add(plan.EdgeAt(i2));
                    var droveRight = TraverseSolver.DriveOpenFrom(adj[lastHeld], fwdChain);
                    for (int kidx = 1; kidx < droveRight.Count; kidx++)
                    {
                        int vi = lastHeld + kidx;
                        adj[vi] = droveRight[kidx];
                        if (!tag.ContainsKey(vi)) tag[vi] = SolveTag.BearingDistance;
                    }
                }
            }

            Cad.EnsureLayer(Config.LayerAdjusted, aci: 2);
            var plAdj = new Polyline { Layer = Config.LayerAdjusted, Closed = plan.Closed };
            for (int i = 0; i < plan.Count; i++) plAdj.AddVertexAt(i, adj[i], 0, 0, 0);
            Cad.AddToModelSpace(plAdj);

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

            var methodById = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < plan.Count; i++)
            {
                if (!tag.TryGetValue(i, out var tg)) tg = SolveTag.SimilarityBetween;
                methodById[plan.VertexIds[i]] = tg.ToString();
            }

            var outPath = WriteCsvReport(plan, links, adj, k, Math.Atan2(s, c), methodById);

            Cad.WriteMessageSafe($"\nALSADJ done. {(usedScale ? "Scale applied." : "Scale locked.")} " +
                            $"Adjusted boundary on {Config.LayerAdjusted}. Residuals on {Config.LayerResiduals} ({residualCount})." +
                            $"\nCSV: {outPath}");
        }

        private static List<Point2d> TwoPointSimilarityBetween(Point2d startHeld, Point2d endHeld,
            IReadOnlyList<PlanEdge> chain)
        {
            var raw = TraverseSolver.DriveOpenFrom(startHeld, chain);
            var pEndRaw = raw[raw.Count - 1];

            var vRaw = new Vector2d(pEndRaw.X - startHeld.X, pEndRaw.Y - startHeld.Y);
            var v = new Vector2d(endHeld.X - startHeld.X, endHeld.Y - startHeld.Y);

            double k = v.Length / Math.Max(1e-12, vRaw.Length);
            double rot = Math.Atan2(v.Y, v.X) - Math.Atan2(vRaw.Y, vRaw.X);
            double c = Math.Cos(rot), s = Math.Sin(rot);

            var outPts = new List<Point2d>(raw.Count);
            foreach (var p in raw)
            {
                double dx = p.X - startHeld.X, dy = p.Y - startHeld.Y;
                double x = k * (c * dx - s * dy) + startHeld.X;
                double y = k * (s * dx + c * dy) + startHeld.Y;
                outPts.Add(new Point2d(x, y));
            }
            return outPts;
        }

        private static List<Point2d> TwoPointBetweenRespectLocks(Point2d startHeld, Point2d endHeld, IReadOnlyList<PlanEdge> chain)
        {
            Vector2d Vlocked = new Vector2d(0, 0);
            Vector2d Vvar = new Vector2d(0, 0);

            for (int i = 0; i < chain.Count; i++)
            {
                var e = chain[i];
                double cx = Math.Cos(e.BearingRad), sy = Math.Sin(e.BearingRad);
                var u = new Vector2d(cx, sy);

                if (e.Locked) Vlocked = new Vector2d(Vlocked.X + e.Distance * u.X, Vlocked.Y + e.Distance * u.Y);
                else Vvar = new Vector2d(Vvar.X + e.Distance * u.X, Vvar.Y + e.Distance * u.Y);
            }

            var v = new Vector2d(endHeld.X - startHeld.X, endHeld.Y - startHeld.Y);
            double vLen = Math.Sqrt(v.X * v.X + v.Y * v.Y);

            double a = (Vvar.X * Vvar.X + Vvar.Y * Vvar.Y);
            double s = 1.0;
            double phi;

            if (a < 1e-12)
            {
                double Llocked = Math.Sqrt(Vlocked.X * Vlocked.X + Vlocked.Y * Vlocked.Y);
                if (Math.Abs(Llocked - vLen) > 1e-6)
                    throw new InvalidOperationException("no free (unlocked) legs to absorb span difference");
                double angA = Math.Atan2(Vlocked.Y, Vlocked.X);
                double angV = Math.Atan2(v.Y, v.X);
                phi = angV - angA;
            }
            else
            {
                double b = 2.0 * (Vlocked.X * Vvar.X + Vlocked.Y * Vvar.Y);
                double cc = (Vlocked.X * Vlocked.X + Vlocked.Y * Vlocked.Y) - vLen * vLen;

                double disc = b * b - 4.0 * a * cc;
                if (disc < 0) disc = 0;

                double sqrtD = Math.Sqrt(disc);
                double s1 = (-b + sqrtD) / (2.0 * a);
                double s2 = (-b - sqrtD) / (2.0 * a);

                double best = double.PositiveInfinity;
                void consider(double cand)
                {
                    if (cand > 0)
                    {
                        double score = Math.Abs(cand - 1.0);
                        if (score < best) { best = score; s = cand; }
                    }
                }
                consider(s1); consider(s2);
                if (double.IsNaN(s) || double.IsInfinity(s)) s = 1.0;

                double Ax = Vlocked.X + s * Vvar.X;
                double Ay = Vlocked.Y + s * Vvar.Y;
                double angA = Math.Atan2(Ay, Ax);
                double angV = Math.Atan2(v.Y, v.X);
                phi = angV - angA;
            }

            var outPts = new List<Point2d>(chain.Count + 1) { startHeld };
            var p = startHeld;
            for (int i = 0; i < chain.Count; i++)
            {
                var e = chain[i];
                double len = e.Locked ? e.Distance : (e.Distance * s);
                double th = Normalize(e.BearingRad + phi);
                p = new Point2d(p.X + len * Math.Cos(th), p.Y + len * Math.Sin(th));
                outPts.Add(p);
            }
            return outPts;
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

        private static string WriteCsvReport(
            PlanData plan,
            EvidenceLinks links,
            Point2d[] adj,
            double scale,
            double rotRad,
            Dictionary<string, string> methodById = null)
        {
            var dictEv = links.Points
                .Where(p => !string.IsNullOrWhiteSpace(p.PlanId))
                .GroupBy(p => p.PlanId, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            var sb = new StringBuilder();
            bool hasMethod = methodById != null && methodById.Count > 0;
            sb.AppendLine("PlanID,Held,Field_X,Field_Y,Adj_X,Adj_Y,Residual_m,EvidenceType,Seg_Distance_m,Seg_Bearing,Scale,Rotation_deg" + (hasMethod ? ",SolveMethod" : ""));

            static string Csv(string s)
                => string.IsNullOrEmpty(s) ? "" :
                   (s.IndexOfAny(new[] { ',', '"', '\n', '\r' }) >= 0 ? "\"" + s.Replace("\"", "\"\"") + "\"" : s);

            for (int i = 0; i < plan.Count; i++)
            {
                string id = plan.VertexIds[i];
                var a = adj[i];

                bool held = false;
                double fx = double.NaN, fy = double.NaN, resid = double.NaN;
                string evType = "";

                if (dictEv.TryGetValue(id, out var ev))
                {
                    held = ev.Held;
                    fx = ev.X;
                    fy = ev.Y;
                    resid = Math.Sqrt((fx - a.X) * (fx - a.X) + (fy - a.Y) * (fy - a.Y));
                    evType = ev.EvidenceType ?? "";
                }

                string segDist = "";
                string segBearing = "";
                if (i < plan.Edges.Count)
                {
                    var e = plan.EdgeAt(i);
                    segDist = e.Distance.ToString("0.###", CultureInfo.InvariantCulture);
                    segBearing = !string.IsNullOrWhiteSpace(e.BearingText) ? e.BearingText : Cad.FormatBearing(e.BearingRad);
                }

                var row = string.Join(",",
                    id,
                    held ? "Y" : "N",
                    double.IsNaN(fx) ? "" : fx.ToString("0.###", CultureInfo.InvariantCulture),
                    double.IsNaN(fy) ? "" : fy.ToString("0.###", CultureInfo.InvariantCulture),
                    a.X.ToString("0.###", CultureInfo.InvariantCulture),
                    a.Y.ToString("0.###", CultureInfo.InvariantCulture),
                    double.IsNaN(resid) ? "" : resid.ToString("0.###", CultureInfo.InvariantCulture),
                    Csv(evType),
                    segDist,
                    Csv(segBearing),
                    scale.ToString("0.000000", CultureInfo.InvariantCulture),
                    (rotRad * 180.0 / Math.PI).ToString("0.000", CultureInfo.InvariantCulture)
                );

                if (hasMethod && methodById.TryGetValue(id, out var mth))
                    row += "," + Csv(mth);

                sb.AppendLine(row);
            }

            string usedPath = WriteTextWithFallback(Config.ReportCsvPath(), sb.ToString());
            return usedPath;
        }
    }
}
