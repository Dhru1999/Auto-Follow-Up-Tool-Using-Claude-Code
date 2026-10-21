using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OXS = DocumentFormat.OpenXml.Spreadsheet;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Compression;
using System.Xml.Linq;
using FollowUpTool.Core;

namespace FollowUpTool
{
    // ─────────────────────────── Data model ───────────────────────────

    internal class JobApplication
    {
        public string Company { get; set; } = "";
        public string Role { get; set; } = "";
        public DateTime? AppliedDate { get; set; }
        public DateTime? FollowUpDate { get; set; }
        public string CompanyEmail { get; set; } = "";
        public string RecruiterEmail { get; set; } = "";
        public string HrEmail { get; set; } = "";
        public string RecruiterName { get; set; } = "";
        public string Status { get; set; } = "";
        public string Notes { get; set; } = "";
        public int FollowUpCount { get; set; }
        public string SheetName { get; set; } = "";
        public int ExcelRowNumber { get; set; }
    }

    internal class AppSettings
    {
        public string ExcelPath { get; set; } = "";
        public string SenderEmail { get; set; } = "";
        public string SenderName { get; set; } = "";
        [JsonIgnore]
        public string AppPassword { get; set; } = "";          // plain text, memory only
        public string EncryptedAppPassword { get; set; } = ""; // DPAPI Base64, written to disk
        [JsonIgnore]
        public string ApiKey { get; set; } = "";               // plain text, memory only
        public string EncryptedAnthropicApiKey { get; set; } = ""; // DPAPI Base64, written to disk

        // Earlier versions wrote the API key to settings.json in plain text under
        // this name. It is read once for migration and never written again.
        [JsonPropertyName("AnthropicApiKey")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? LegacyPlainTextApiKey { get; set; }
    }

    /// <summary>Windows DPAPI, scoped to the current Windows user.</summary>
    internal sealed class DpapiProtector : ISecretProtector
    {
        public byte[] Protect(byte[] plain) => ProtectedData.Protect(plain, null, DataProtectionScope.CurrentUser);
        public byte[] Unprotect(byte[] cipher) => ProtectedData.Unprotect(cipher, null, DataProtectionScope.CurrentUser);
    }

    // ─────────────────────────── Main Form ───────────────────────────

    public class MainForm : Form
    {
        // ── Controls ──────────────────────────────────────────────────
        private TextBox txtFilePath = null!;
        private Button btnBrowse = null!;
        private Button btnLoad = null!;
        private Label lblFoundCount = null!;

        private DataGridView dgv = null!;

        private TextBox txtTo = null!;
        private TextBox txtSubject = null!;
        private RichTextBox rtbBody = null!;
        private Button btnGenerate = null!;
        private Button btnSend = null!;

        private TextBox txtSenderEmail = null!;
        private TextBox txtSenderName = null!;
        private TextBox txtAppPassword = null!;
        private TextBox txtApiKey = null!;
        private Button btnSaveSettings = null!;

        private StatusStrip statusStrip = null!;
        private ToolStripStatusLabel statusLabel = null!;

        // ── State ──────────────────────────────────────────────────────
        private List<JobApplication> _apps = new();
        private JobApplication? _selected;
        private AppSettings _settings = new();
        private static readonly ISecretProtector Protector = new DpapiProtector();
        private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(60) };
        private readonly string _settingsPath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                         "FollowUpTool", "settings.json");

        // ── Colors ─────────────────────────────────────────────────────
        private static readonly Color C_DARK   = Color.FromArgb(31, 78, 121);
        private static readonly Color C_BLUE   = Color.FromArgb(41, 128, 185);
        private static readonly Color C_GREEN  = Color.FromArgb(39, 174, 96);
        private static readonly Color C_ORANGE = Color.FromArgb(230, 126, 34);
        private static readonly Color C_BG     = Color.FromArgb(245, 248, 250);
        private static readonly Color C_WHITE  = Color.White;

        // ═══════════════════════════════════════════════════════════════
        public MainForm()
        {
            LoadSettings();
            BuildUI();
        }

        // ═══════════════════════════════════════════════════════════════
        //  UI Construction
        // ═══════════════════════════════════════════════════════════════

        private void BuildUI()
        {
            Text = "Job Application Follow-Up Tool";
            Size = new Size(1500, 950);
            MinimumSize = new Size(1200, 800);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = C_BG;
            Font = new Font("Segoe UI", 9.5f);

            // ── Title bar ──────────────────────────────────────────────
            var titlePanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 52,
                BackColor = C_DARK,
                Padding = new Padding(16, 0, 0, 0)
            };
            var titleLabel = new Label
            {
                Text = "  Job Application Follow-Up Tool",
                ForeColor = C_WHITE,
                Font = new Font("Segoe UI", 14f, FontStyle.Bold),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            titlePanel.Controls.Add(titleLabel);

            // ── File picker row ────────────────────────────────────────
            var filePanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 66,
                BackColor = C_WHITE,
                Padding = new Padding(12, 10, 12, 10)
            };
            filePanel.Paint += (s, e) => e.Graphics.DrawLine(
                new Pen(Color.FromArgb(220, 220, 220)), 0, filePanel.Height - 1, filePanel.Width, filePanel.Height - 1);

            var lblFile = new Label { Text = "Excel File:", AutoSize = true, Top = 16, Left = 8, ForeColor = Color.FromArgb(80, 80, 80) };
            txtFilePath = new TextBox { Left = 80, Top = 16, Width = 680, Height = 34, Text = _settings.ExcelPath };
            btnBrowse = MakeButton("Browse…", C_BLUE, 770, 14, 80);
            btnLoad = MakeButton("Load Due Follow-ups  ↺", C_GREEN, 860, 14, 200);
            lblFoundCount = new Label { Left = 1070, Top = 20, AutoSize = true, ForeColor = C_DARK, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };

            btnBrowse.Click += BtnBrowse_Click;
            btnLoad.Click += BtnLoad_Click;

            foreach (Control c in new Control[] { lblFile, txtFilePath, btnBrowse, btnLoad, lblFoundCount })
                filePanel.Controls.Add(c);

            // ── Main split (left = grid, right = composer) ─────────────
            var splitMain = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                BackColor = C_BG
            };
            // SplitterDistance and min sizes must be set after the form has a real width
            Load += (s, e) =>
            {
                splitMain.Panel1MinSize = 420;
                splitMain.Panel2MinSize = 400;
                splitMain.SplitterDistance = Math.Max(420, (splitMain.Width - splitMain.SplitterWidth) * 50 / 100);
                if (!string.IsNullOrEmpty(_settings.ExcelPath) && File.Exists(_settings.ExcelPath))
                    BtnLoad_Click(null, EventArgs.Empty);
            };

            // Left panel — applications grid
            BuildGridPanel(splitMain.Panel1);

            // Right panel — email composer + settings
            BuildComposerPanel(splitMain.Panel2);

            // ── Status bar ─────────────────────────────────────────────
            statusStrip = new StatusStrip { BackColor = C_DARK };
            statusLabel = new ToolStripStatusLabel("Ready — load an Excel file to begin")
            {
                ForeColor = C_WHITE,
                Font = new Font("Segoe UI", 9f)
            };
            statusStrip.Items.Add(statusLabel);

            // ── Assemble ───────────────────────────────────────────────
            Controls.Add(splitMain);
            Controls.Add(statusStrip);
            Controls.Add(filePanel);
            Controls.Add(titlePanel);

            // Pre-fill file path from settings
            if (!string.IsNullOrEmpty(_settings.ExcelPath) && File.Exists(_settings.ExcelPath))
                txtFilePath.Text = _settings.ExcelPath;
        }

        // ── Grid panel ─────────────────────────────────────────────────

        private void BuildGridPanel(SplitterPanel panel)
        {
            panel.BackColor = C_BG;
            panel.Padding = new Padding(10, 10, 5, 10);

            var header = new Label
            {
                Text = "Follow-Ups Due (today or overdue)",
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                ForeColor = C_DARK,
                Dock = DockStyle.Top,
                Height = 28
            };

            dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                BackgroundColor = C_WHITE,
                BorderStyle = BorderStyle.None,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 32,
                RowTemplate = { Height = 30 },
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                GridColor = Color.FromArgb(230, 230, 230),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = C_DARK;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = C_WHITE;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(41, 128, 185);
            dgv.DefaultCellStyle.SelectionForeColor = C_WHITE;
            dgv.EnableHeadersVisualStyles = false;
            dgv.SelectionChanged += Dgv_SelectionChanged;
            dgv.DataBindingComplete += (s, e) => StyleGridRows();

            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Company", HeaderText = "Company", FillWeight = 25, MinimumWidth = 80 });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Role", HeaderText = "Role", FillWeight = 30, MinimumWidth = 80 });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "AppliedDate", HeaderText = "Applied", FillWeight = 14, MinimumWidth = 80 });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "DaysAgo", HeaderText = "Days Ago", FillWeight = 8, MinimumWidth = 72 });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "CompanyEmail", HeaderText = "Company Email", FillWeight = 23, MinimumWidth = 120 });

            panel.Controls.Add(dgv);
            panel.Controls.Add(header);
        }

        // ── Composer panel ─────────────────────────────────────────────

        private void BuildComposerPanel(SplitterPanel panel)
        {
            panel.BackColor = C_BG;
            panel.Padding = new Padding(5, 10, 10, 10);

            var composerBox = new GroupBox
            {
                Text = "Email Composer",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = C_DARK,
                BackColor = C_WHITE,
                Padding = new Padding(10)
            };

            var tbl = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 7,
                Padding = new Padding(8)
            };
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 72));
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));   // To
            tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));   // Subject
            tbl.RowStyles.Add(new RowStyle(SizeType.Percent, 100));   // Body
            tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));   // Buttons row
            tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 12));   // Spacer
            tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 110));  // Settings
            tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));   // Send row

            // Row 0 — To
            tbl.Controls.Add(FieldLabel("To:"), 0, 0);
            txtTo = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 4, 0, 4), PlaceholderText = "recruiter@company.com", Enabled = false };
            tbl.Controls.Add(txtTo, 1, 0);

            // Row 1 — Subject
            tbl.Controls.Add(FieldLabel("Sub:"), 0, 1);
            txtSubject = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 4, 0, 4) };
            tbl.Controls.Add(txtSubject, 1, 1);

            // Row 2 — Body
            tbl.Controls.Add(FieldLabel("Body:"), 0, 2);
            rtbBody = new RichTextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9.5f),
                BorderStyle = BorderStyle.FixedSingle,
                ScrollBars = RichTextBoxScrollBars.Vertical
            };
            tbl.Controls.Add(rtbBody, 1, 2);

            // Row 3 — Generate button
            var btnRow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, Padding = new Padding(0, 4, 0, 0) };
            btnGenerate = MakeButton("Generate AI Draft", C_ORANGE, 0, 0, 160);
            btnGenerate.Click += BtnGenerate_Click;
            btnRow.Controls.Add(btnGenerate);
            tbl.Controls.Add(btnRow, 1, 3);

            // Row 4 — Spacer / separator
            var sep = new Label { Dock = DockStyle.Fill, BorderStyle = BorderStyle.Fixed3D, Height = 1 };
            tbl.SetColumnSpan(sep, 2);
            tbl.Controls.Add(sep, 0, 4);

            // Row 5 — Settings panel (2-col span)
            var settingsPanel = BuildSettingsPanel();
            tbl.SetColumnSpan(settingsPanel, 2);
            tbl.Controls.Add(settingsPanel, 0, 5);

            // Row 7 — Send button (right-aligned)
            var sendRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(0, 4, 0, 0)
            };
            btnSend = MakeButton("Send Email  ✉", C_GREEN, 0, 0, 150);
            btnSend.Font = new Font("Segoe UI", 10f, FontStyle.Bold);
            btnSend.Height = 44;
            btnSend.Click += BtnSend_Click;
            sendRow.Controls.Add(btnSend);
            tbl.Controls.Add(sendRow, 0, 6);
            tbl.SetColumnSpan(sendRow, 2);

            composerBox.Controls.Add(tbl);
            panel.Controls.Add(composerBox);
        }

        private Panel BuildSettingsPanel()
        {
            var pnl = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(248, 250, 252), Padding = new Padding(4) };

            var lbl = new Label
            {
                Text = "Gmail & API Settings",
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                ForeColor = C_DARK,
                AutoSize = true,
                Location = new Point(4, 4)
            };

            var tbl = new TableLayoutPanel
            {
                Left = 0, Top = 22, Width = 580, Height = 68,
                ColumnCount = 4, RowCount = 2,
                AutoSize = true
            };
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96));
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70));
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60));
            tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));

            // Row 0: Sender Email | [txtSenderEmail]  App Password | [txtAppPassword]
            tbl.Controls.Add(FieldLabel("Sender Email:"), 0, 0);
            txtSenderEmail = new TextBox { Dock = DockStyle.Fill, Text = _settings.SenderEmail, Margin = new Padding(0, 2, 4, 2), PlaceholderText = "you@gmail.com" };
            tbl.Controls.Add(txtSenderEmail, 1, 0);
            var lblAppPassword = FieldLabel("App Password:");
            tbl.Controls.Add(lblAppPassword, 2, 0);
            txtAppPassword = new TextBox { Dock = DockStyle.Fill, Text = _settings.AppPassword, Margin = new Padding(0, 2, 0, 2), PlaceholderText = "Gmail App Password", UseSystemPasswordChar = true };
            tbl.Controls.Add(txtAppPassword, 3, 0);

            // Row 1: Sender Name | [txtSenderName]  API Key | [txtApiKey]
            tbl.Controls.Add(FieldLabel("Sender Name:"), 0, 1);
            txtSenderName = new TextBox { Dock = DockStyle.Fill, Text = _settings.SenderName, Margin = new Padding(0, 2, 4, 2), PlaceholderText = "Your name" };
            tbl.Controls.Add(txtSenderName, 1, 1);
            var lblApiKey = FieldLabel("Anthropic Key:");
            tbl.Controls.Add(lblApiKey, 2, 1);
            txtApiKey = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = true, Text = _settings.ApiKey, Margin = new Padding(0, 2, 0, 2), PlaceholderText = "sk-ant-... (or set ANTHROPIC_API_KEY)" };
            tbl.Controls.Add(txtApiKey, 3, 1);

            btnSaveSettings = MakeButton("Save", C_DARK, 590, 34, 70);
            btnSaveSettings.Height = 60;
            btnSaveSettings.Click += BtnSaveSettings_Click;

            tbl.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            pnl.Controls.Add(lbl);
            pnl.Controls.Add(tbl);
            pnl.Controls.Add(btnSaveSettings);
            return pnl;
        }

        // ═══════════════════════════════════════════════════════════════
        //  Excel Reading
        // ═══════════════════════════════════════════════════════════════

        private List<JobApplication> ReadExcel(string path)
        {
            var results = new List<JobApplication>();
            var today = DateTime.Today;
            XNamespace ssNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace rNs  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace pkNs = "http://schemas.openxmlformats.org/package/2006/relationships";

            using var zip = ZipFile.OpenRead(path);

            // Shared string table
            var sst = new List<string>();
            var sstEntry = zip.GetEntry("xl/sharedStrings.xml");
            if (sstEntry != null)
            {
                using var s = sstEntry.Open();
                foreach (var si in XDocument.Load(s).Descendants(ssNs + "si"))
                    sst.Add(string.Concat(si.Descendants(ssNs + "t").Select(t => t.Value)));
            }

            // workbook sheet name → worksheet ZIP path
            var sheetPaths = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string? firstSheetName = null;
            var wbEntry   = zip.GetEntry("xl/workbook.xml");
            var wbRels    = zip.GetEntry("xl/_rels/workbook.xml.rels");
            if (wbEntry != null && wbRels != null)
            {
                using var ws = wbEntry.Open();
                using var rs = wbRels.Open();
                var wbDoc   = XDocument.Load(ws);
                var relsDoc = XDocument.Load(rs);
                var ridMap  = relsDoc.Descendants(pkNs + "Relationship")
                    .ToDictionary(r => r.Attribute("Id")?.Value ?? "", r => r.Attribute("Target")?.Value ?? "");
                foreach (var sh in wbDoc.Descendants(ssNs + "sheet"))
                {
                    var name = sh.Attribute("name")?.Value ?? "";
                    var rid  = sh.Attribute(rNs + "id")?.Value ?? "";
                    if (ridMap.TryGetValue(rid, out var target))
                    {
                        // Targets are relative to xl/ but may also be written as absolute ("/xl/...").
                        target = target.TrimStart('/');
                        sheetPaths[name] = target.StartsWith("xl/") ? target : "xl/" + target;
                        firstSheetName ??= name;
                    }
                }
            }

            // Sheets are read in this order; if none of them exist the first
            // sheet of the workbook is used, so any tracker layout works.
            var preferred = new[] { "Internship- Data Engineer", "Status", "Job Applications" };
            var sheetNames = preferred.Where(sheetPaths.ContainsKey).ToArray();
            if (sheetNames.Length == 0 && firstSheetName != null)
                sheetNames = new[] { firstSheetName };

            foreach (var sheetName in sheetNames)
            {
                if (!sheetPaths.TryGetValue(sheetName, out var wsPath)) continue;
                var wsEntry = zip.GetEntry(wsPath);
                if (wsEntry == null) continue;

                using var wsStream = wsEntry.Open();
                var wsDoc = XDocument.Load(wsStream);
                var rows  = wsDoc.Descendants(ssNs + "row")
                    .OrderBy(r => int.TryParse(r.Attribute("r")?.Value, out int n) ? n : 0)
                    .ToList();
                if (rows.Count == 0) continue;

                // column letter → mapped field name
                var colMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var cell in rows[0].Elements(ssNs + "c"))
                {
                    var col    = ColumnRef.LettersOf(cell.Attribute("r")?.Value);
                    var header = HeaderMapper.Map(ZipCellText(cell, sst, ssNs));
                    if (!string.IsNullOrEmpty(header) && !colMap.ContainsKey(header))
                        colMap[header] = col;
                }
                // invert: field name → column letter
                var fieldCol = colMap.ToDictionary(kv => kv.Value, kv => kv.Key, StringComparer.OrdinalIgnoreCase);
                var colField = colMap; // field→col

                string Field(XElement row, string name)
                {
                    if (!colField.TryGetValue(name, out var col)) return "";
                    var c = row.Elements(ssNs + "c")
                        .FirstOrDefault(x => ColumnRef.LettersOf(x.Attribute("r")?.Value)
                            .Equals(col, StringComparison.OrdinalIgnoreCase));
                    return c != null ? ZipCellText(c, sst, ssNs) : "";
                }

                DateTime? FieldDate(XElement row, string name)
                {
                    if (!colField.TryGetValue(name, out var col)) return null;
                    var c = row.Elements(ssNs + "c")
                        .FirstOrDefault(x => ColumnRef.LettersOf(x.Attribute("r")?.Value)
                            .Equals(col, StringComparison.OrdinalIgnoreCase));
                    return c != null ? CellDateParser.Parse(ZipCellText(c, sst, ssNs)) : null;
                }

                foreach (var row in rows.Skip(1))
                {
                    var company = Field(row, "Company");
                    var role    = Field(row, "Role");
                    if (string.IsNullOrEmpty(company) && string.IsNullOrEmpty(role)) continue;

                    var status       = Field(row, "Status");
                    var appliedDate  = FieldDate(row, "AppliedDate");
                    var followUpDate = FollowUpRules.DueDate(appliedDate, FieldDate(row, "FollowUpDate"));

                    if (!FollowUpRules.NeedsFollowUp(status, Field(row, "FollowupDone"), appliedDate, followUpDate, today))
                        continue;

                    int.TryParse(Field(row, "FollowUpCount"), out int fuCount);
                    int.TryParse(row.Attribute("r")?.Value, out int rowNum);

                    results.Add(new JobApplication
                    {
                        Company       = company,
                        Role          = role,
                        AppliedDate   = appliedDate,
                        FollowUpDate  = followUpDate,
                        CompanyEmail  = Field(row, "CompanyEmail"),
                        RecruiterEmail= Field(row, "RecruiterEmail"),
                        HrEmail       = Field(row, "HrEmail"),
                        RecruiterName = Field(row, "RecruiterName"),
                        Status        = status,
                        Notes         = Field(row, "Notes"),
                        FollowUpCount = fuCount,
                        SheetName     = sheetName,
                        ExcelRowNumber= rowNum
                    });
                }
            }

            return results
                .GroupBy(a => $"{a.Company}|{a.Role}")
                .Select(g => g.First())
                .OrderBy(a => a.AppliedDate)
                .ToList();
        }

        private static string ZipCellText(XElement cell, List<string> sst, XNamespace ns)
        {
            var t = cell.Attribute("t")?.Value;
            var v = cell.Element(ns + "v")?.Value ?? "";
            if (t == "s")
            {
                if (int.TryParse(v, out int i) && i < sst.Count) return sst[i];
                return "";
            }
            if (t == "inlineStr") return cell.Element(ns + "is")?.Element(ns + "t")?.Value ?? "";
            if (t == "b") return v == "1" ? "TRUE" : "FALSE";
            return v;
        }

        // ═══════════════════════════════════════════════════════════════
        //  Event Handlers
        // ═══════════════════════════════════════════════════════════════

        private void BtnBrowse_Click(object? sender, EventArgs e)
        {
            using var dlg = new OpenFileDialog
            {
                Title = "Select Job Applications Excel File",
                Filter = "Excel Workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
                InitialDirectory = Directory.Exists(Path.GetDirectoryName(txtFilePath.Text))
                    ? Path.GetDirectoryName(txtFilePath.Text)
                    : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };
            if (dlg.ShowDialog() == DialogResult.OK)
                txtFilePath.Text = dlg.FileName;
        }

        private void BtnLoad_Click(object? sender, EventArgs e)
        {
            var path = txtFilePath.Text.Trim();
            if (!File.Exists(path))
            {
                MessageBox.Show("File not found:\n" + path, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SetStatus("Loading Excel file…");
            btnLoad.Enabled = false;

            Task.Run(() =>
            {
                try
                {
                    var apps = ReadExcel(path);
                    Invoke(() =>
                    {
                        _apps = apps;
                        PopulateGrid();
                        _settings.ExcelPath = path;
                        SaveSettings();
                        var count = apps.Count;
                        lblFoundCount.Text = count == 0 ? "No follow-ups due" : $"{count} follow-up{(count == 1 ? "" : "s")} due";
                        lblFoundCount.ForeColor = count > 0 ? C_ORANGE : Color.Gray;
                        SetStatus(count == 0
                            ? "No applications are due for follow-up."
                            : $"Found {count} application{(count == 1 ? "" : "s")} due or overdue for follow-up with no response yet.");
                    });
                }
                catch (Exception ex)
                {
                    Invoke(() =>
                    {
                        MessageBox.Show($"Error reading Excel file:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        SetStatus("Error loading file.");
                    });
                }
                finally
                {
                    Invoke(() => btnLoad.Enabled = true);
                }
            });
        }

        private void PopulateGrid()
        {
            dgv.Rows.Clear();
            foreach (var app in _apps)
            {
                var daysAgo = app.AppliedDate.HasValue
                    ? (DateTime.Today - app.AppliedDate.Value).Days.ToString()
                    : "-";
                dgv.Rows.Add(
                    app.Company,
                    app.Role,
                    app.AppliedDate?.ToString("dd MMM yyyy") ?? "-",
                    daysAgo + "d",
                    app.CompanyEmail
                );
            }
            StyleGridRows();
            if (dgv.Rows.Count > 0)
                dgv.Rows[0].Selected = true;
        }

        private void StyleGridRows()
        {
            for (int i = 0; i < dgv.Rows.Count; i++)
                dgv.Rows[i].DefaultCellStyle.BackColor = i % 2 == 0
                    ? C_WHITE
                    : Color.FromArgb(235, 245, 255);
        }

        private void Dgv_SelectionChanged(object? sender, EventArgs e)
        {
            if (dgv.SelectedRows.Count == 0) return;
            var idx = dgv.SelectedRows[0].Index;
            if (idx < 0 || idx >= _apps.Count) return;

            _selected = _apps[idx];
            txtTo.Text = !string.IsNullOrEmpty(_selected.CompanyEmail)
                ? _selected.CompanyEmail
                : _selected.RecruiterEmail;
            txtSubject.Text = $"Follow-up: {_selected.Role} Application";
            rtbBody.Text = BuildDefaultDraft(_selected);
        }

        private string BuildDefaultDraft(JobApplication app)
        {
            var greeting = string.IsNullOrEmpty(app.RecruiterName)
                ? "Hi,"
                : $"Hi {app.RecruiterName.Split(' ')[0]},";
            var applied = app.AppliedDate?.ToString("d MMMM yyyy") ?? "recently";

            return $"{greeting}\n\nI wanted to follow up on my application for the {app.Role} position at {app.Company}, submitted on {applied}. I remain very interested in the role and would love to hear if there are any updates on the process.\n\nPlease let me know if you need any additional information from my side.\n\nRegards,\n{txtSenderName.Text.Trim()}".TrimEnd();
        }

        // ── Generate AI Draft ──────────────────────────────────────────

        private async void BtnGenerate_Click(object? sender, EventArgs e)
        {
            if (_selected == null)
            {
                MessageBox.Show("Please select an application from the list first.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var apiKey = txtApiKey.Text.Trim();
            if (string.IsNullOrEmpty(apiKey))
                apiKey = Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY") ?? "";

            if (string.IsNullOrEmpty(apiKey))
            {
                MessageBox.Show("Enter your Anthropic API key in the settings panel below.", "API Key Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtApiKey.Focus();
                return;
            }

            btnGenerate.Enabled = false;
            SetStatus($"Generating AI draft for {_selected.Company}…");

            try
            {
                var draft = await GenerateAiDraftAsync(_selected, apiKey);
                rtbBody.Text = draft;
                SetStatus("AI draft generated — review and edit before sending.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"AI generation failed:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                SetStatus("AI draft generation failed.");
            }
            finally
            {
                btnGenerate.Enabled = true;
            }
        }

        private async Task<string> GenerateAiDraftAsync(JobApplication app, string apiKey)
        {
            var request = new DraftRequest(app.Company, app.Role, app.RecruiterName, app.AppliedDate,
                                           app.FollowUpCount, app.Notes, txtSenderName.Text.Trim());
            var model = Environment.GetEnvironmentVariable("CLAUDE_MODEL") ?? DraftPrompt.DefaultModel;

            using var message = new HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/messages")
            {
                Content = new StringContent(DraftPrompt.BuildRequestJson(request, model), Encoding.UTF8, "application/json")
            };
            message.Headers.Add("x-api-key", apiKey);
            message.Headers.Add("anthropic-version", DraftPrompt.ApiVersion);

            HttpResponseMessage resp;
            try { resp = await Http.SendAsync(message); }
            catch (TaskCanceledException) { throw new Exception("The Claude API did not respond within 60 seconds. Try again."); }
            catch (HttpRequestException ex) { throw new Exception($"Could not reach the Claude API: {ex.Message}"); }

            using (resp)
            {
                var body = await resp.Content.ReadAsStringAsync();
                if (!resp.IsSuccessStatusCode)
                    throw new Exception(DraftPrompt.DescribeError((int)resp.StatusCode, body));
                return DraftPrompt.ParseResponse(body);
            }
        }

        // ── Send Email ─────────────────────────────────────────────────

        private async void BtnSend_Click(object? sender, EventArgs e)
        {
            var to = txtTo.Text.Trim();
            var subject = txtSubject.Text.Trim();
            var bodyText = rtbBody.Text.Trim();

            if (!MailAddress.TryCreate(to, out _))
            {
                MessageBox.Show("Enter a valid recipient email address in the To field.", "Missing Recipient", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTo.Focus();
                return;
            }
            if (string.IsNullOrEmpty(subject))
            {
                MessageBox.Show("Enter a subject line.", "Missing Subject", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtSubject.Focus();
                return;
            }
            if (string.IsNullOrEmpty(bodyText))
            {
                MessageBox.Show("Email body is empty. Generate or type a draft first.", "Empty Body", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                rtbBody.Focus();
                return;
            }

            var senderEmail = txtSenderEmail.Text.Trim();
            var appPassword = txtAppPassword.Text.Trim();
            var senderName = txtSenderName.Text.Trim();

            if (string.IsNullOrEmpty(senderEmail) || string.IsNullOrEmpty(appPassword))
            {
                MessageBox.Show("Enter your Gmail address and App Password in the settings panel.", "Settings Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var confirm = MessageBox.Show(
                $"Send email to:\n{to}\n\nSubject: {subject}\n\nProceed?",
                "Confirm Send", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;

            btnSend.Enabled = false;
            SetStatus($"Sending email to {to}…");

            var excelPath = txtFilePath.Text.Trim();
            var sentApp = _selected!;

            try
            {
                string? excelWriteError = null;
                await Task.Run(() =>
                {
                    SendSmtp(to, sentApp.HrEmail, subject, bodyText, senderEmail, senderName, appPassword);
                    try { MarkFollowupDone(excelPath, sentApp.SheetName, sentApp.ExcelRowNumber); }
                    catch (Exception ex) { excelWriteError = ex.Message; }
                });

                SetStatus($"Email sent to {to} at {DateTime.Now:HH:mm}.");
                if (excelWriteError != null)
                    MessageBox.Show($"Email sent successfully to {to}.\n\nNote: Could not update Excel file:\n{excelWriteError}\n\nPlease mark the row manually.", "Sent (Excel update failed)", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                else
                    MessageBox.Show($"Email sent successfully to {to}.", "Sent", MessageBoxButtons.OK, MessageBoxIcon.Information);

                _selected = null;
                BtnLoad_Click(null, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to send email:\n{ex.Message}\n\nTip: Use a Gmail App Password (not your regular password).", "Send Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                SetStatus("Email send failed — check settings.");
            }
            finally
            {
                btnSend.Enabled = true;
            }
        }

        private static void SendSmtp(string to, string cc, string subject, string body,
                                     string senderEmail, string senderName, string password)
        {
            var msg = new MailMessage
            {
                From = new MailAddress(senderEmail, senderName),
                Subject = subject,
                Body = body,
                IsBodyHtml = false
            };
            msg.To.Add(new MailAddress(to));
            if (MailAddress.TryCreate(cc, out var ccAddress))
                msg.CC.Add(ccAddress);

            using var smtp = new SmtpClient("smtp.gmail.com", 587)
            {
                EnableSsl = true,
                Credentials = new NetworkCredential(senderEmail, password),
                DeliveryMethod = SmtpDeliveryMethod.Network
            };
            smtp.Send(msg);
        }

        // Uses OpenXML SDK directly so ClosedXML never writes (and never corrupts) the file.
        private static void MarkFollowupDone(string path, string sheetName, int excelRowNumber)
        {
            if (excelRowNumber <= 0 || !File.Exists(path)) return;

            using var doc = SpreadsheetDocument.Open(path, isEditable: true);
            var wbPart = doc.WorkbookPart ?? throw new InvalidOperationException("No workbook part");

            var sheetEntry = wbPart.Workbook.Descendants<OXS.Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
            if (sheetEntry?.Id?.Value == null) return;

            var wsPart = (WorksheetPart)wbPart.GetPartById(sheetEntry.Id.Value);
            var sheetData = wsPart.Worksheet.GetFirstChild<OXS.SheetData>();
            if (sheetData == null) return;

            // Header is the first row in the sheet
            var allRows = sheetData.Elements<OXS.Row>().OrderBy(r => r.RowIndex?.Value ?? 0).ToList();
            if (allRows.Count == 0) return;
            var headerRow = allRows[0];
            uint headerRowIdx = headerRow.RowIndex?.Value ?? 1;

            // Find "Followup done" column or determine where to add it
            int followupCol = -1;
            int maxCol = 0;
            foreach (var c in headerRow.Elements<OXS.Cell>())
            {
                string colLetter = ColumnRef.LettersOf(c.CellReference?.Value);
                int colNum = ColumnRef.ToNumber(colLetter);
                if (colNum > maxCol) maxCol = colNum;
                string val = OxCellText(wbPart, c);
                if (val.Equals("Followup done", StringComparison.OrdinalIgnoreCase) ||
                    val.Equals("followup_done", StringComparison.OrdinalIgnoreCase))
                { followupCol = colNum; break; }
            }

            if (followupCol < 0)
            {
                followupCol = maxCol + 1;
                OxWriteCell(sheetData, headerRow, headerRowIdx, followupCol, "Followup done");
            }

            // Write "Yes" to the data row
            var dataRow = sheetData.Elements<OXS.Row>()
                .FirstOrDefault(r => r.RowIndex?.Value == (uint)excelRowNumber);
            if (dataRow == null)
            {
                dataRow = new OXS.Row { RowIndex = (uint)excelRowNumber };
                var after = sheetData.Elements<OXS.Row>()
                    .FirstOrDefault(r => r.RowIndex?.Value > (uint)excelRowNumber);
                if (after != null) sheetData.InsertBefore(dataRow, after);
                else sheetData.Append(dataRow);
            }
            OxWriteCell(sheetData, dataRow, (uint)excelRowNumber, followupCol, "Yes");

            wsPart.Worksheet.Save();
        }

        private static string OxCellText(WorkbookPart wbPart, OXS.Cell cell)
        {
            if (cell.DataType?.Value == OXS.CellValues.SharedString)
            {
                var sst = wbPart.SharedStringTablePart?.SharedStringTable;
                if (sst != null && int.TryParse(cell.CellValue?.Text, out int i))
                    return sst.Elements<OXS.SharedStringItem>().ElementAtOrDefault(i)?.InnerText ?? "";
            }
            if (cell.DataType?.Value == OXS.CellValues.InlineString)
                return cell.InlineString?.Text?.Text ?? "";
            return cell.CellValue?.Text ?? "";
        }

        private static void OxWriteCell(OXS.SheetData sd, OXS.Row row, uint rowIdx, int colIdx, string value)
        {
            string cellRef = ColumnRef.ToLetters(colIdx) + rowIdx;
            var existing = row.Elements<OXS.Cell>()
                .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
            existing?.Remove();

            var newCell = new OXS.Cell
            {
                CellReference = cellRef,
                DataType = new EnumValue<OXS.CellValues>(OXS.CellValues.InlineString),
                InlineString = new OXS.InlineString { Text = new OXS.Text(value) }
            };

            var refCell = row.Elements<OXS.Cell>().FirstOrDefault(c =>
                ColumnRef.ToNumber(ColumnRef.LettersOf(c.CellReference?.Value)) > colIdx);
            if (refCell != null) row.InsertBefore(newCell, refCell);
            else row.Append(newCell);
        }

        // ── Settings persistence ───────────────────────────────────────

        private void BtnSaveSettings_Click(object? sender, EventArgs e)
        {
            _settings.SenderEmail = txtSenderEmail.Text.Trim();
            _settings.SenderName = txtSenderName.Text.Trim();
            _settings.AppPassword = txtAppPassword.Text.Trim();
            _settings.ApiKey = txtApiKey.Text.Trim();
            _settings.ExcelPath = txtFilePath.Text.Trim();
            SaveSettings();
            SetStatus("Settings saved.");
        }

        private void LoadSettings()
        {
            try
            {
                if (!File.Exists(_settingsPath)) return;
                var json = File.ReadAllText(_settingsPath);
                _settings = JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
                _settings.AppPassword = SettingsSecrets.Decrypt(Protector, _settings.EncryptedAppPassword);
                _settings.ApiKey = SettingsSecrets.Decrypt(Protector, _settings.EncryptedAnthropicApiKey);

                // One-time migration: re-save a key that older versions stored in plain text.
                if (!string.IsNullOrEmpty(_settings.LegacyPlainTextApiKey))
                {
                    if (string.IsNullOrEmpty(_settings.ApiKey))
                        _settings.ApiKey = _settings.LegacyPlainTextApiKey;
                    SaveSettings();
                }
            }
            catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
            {
                _settings = new AppSettings();
            }
        }

        private void SaveSettings()
        {
            try
            {
                _settings.EncryptedAppPassword = SettingsSecrets.Encrypt(Protector, _settings.AppPassword);
                _settings.EncryptedAnthropicApiKey = SettingsSecrets.Encrypt(Protector, _settings.ApiKey);
                _settings.LegacyPlainTextApiKey = null;
                Directory.CreateDirectory(Path.GetDirectoryName(_settingsPath)!);
                File.WriteAllText(_settingsPath, JsonSerializer.Serialize(_settings, new JsonSerializerOptions { WriteIndented = true }));
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or CryptographicException)
            {
                SetStatusSafe($"Could not save settings: {ex.Message}");
            }
        }

        // SaveSettings can run before the status bar exists (during LoadSettings).
        private void SetStatusSafe(string msg)
        {
            if (statusLabel != null) SetStatus(msg);
        }

        // ═══════════════════════════════════════════════════════════════
        //  Helpers
        // ═══════════════════════════════════════════════════════════════

        private static Button MakeButton(string text, Color back, int left, int top, int width)
        {
            return new Button
            {
                Text = text,
                BackColor = back,
                ForeColor = C_WHITE,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 0 },
                Font = new Font("Segoe UI", 9f),
                Location = new Point(left, top),
                Width = width,
                Height = 38,
                Cursor = Cursors.Hand
            };
        }

        private static Label FieldLabel(string text) => new Label
        {
            Text = text,
            TextAlign = ContentAlignment.MiddleRight,
            Dock = DockStyle.Fill,
            ForeColor = Color.FromArgb(80, 80, 80)
        };

        private void SetStatus(string msg)
        {
            if (InvokeRequired) Invoke(() => statusLabel.Text = msg);
            else statusLabel.Text = msg;
        }
    }
}
