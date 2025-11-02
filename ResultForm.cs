using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.IO;

using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using OxyPlot.WindowsForms;

using ClosedXML.Excel;

namespace StyleWatcherWin
{
    // 主题与色板
    static class UIStyle
    {
        public static readonly Font H1   = new Font("Microsoft YaHei UI", 13, FontStyle.Bold);
        public static readonly Font Body = new Font("Microsoft YaHei UI", 10, FontStyle.Regular);

        public static readonly Color HeaderBack = Color.FromArgb(245, 247, 250);
        public static readonly Color TextDark   = Color.FromArgb(47, 47, 47);
        public static readonly Color AccentBlue = Color.FromArgb(25, 145, 235);
        public static readonly Color AccentGreen= Color.FromArgb(18, 183, 106);
        public static readonly Color GridLine   = Color.FromArgb(235, 238, 245);

        public static OxyColor Ox(byte r, byte g, byte b) => OxyColor.FromRgb(r, g, b);
        public static OxyColor Ox(Color c) => OxyColor.FromRgb(c.R, c.G, c.B);
    }

    public class ResultForm : Form
    {
        // 顶部：标题、摘要、操作
        readonly Label _lblTitle = new Label();
        readonly Label _lblSummary = new Label();
        readonly Button _btnQuery = new Button();
        readonly Button _btnExport = new Button();
        readonly TextBox _boxInput = new TextBox();

        // Tab
        readonly TabControl _tabs = new TabControl();

        // 图表
        readonly PlotView _pvTrend = new PlotView();
        readonly PlotView _pvSizeTop = new PlotView();
        readonly PlotView _pvColorTop = new PlotView();

        // 明细
        readonly DataGridView _grid = new DataGridView();
        readonly TextBox _boxSearch = new TextBox();
        BindingSource _binding = new BindingSource();
        List<dynamic> _detailRows = new List<dynamic>(); // 当前表数据源（加工后）

        readonly AppConfig _cfg;
        ParsedPayload _parsed = new ParsedPayload();

        public ResultForm(AppConfig cfg)
        {
            _cfg = cfg;

            Text = "StyleWatcher";
            KeyPreview = true;
            StartPosition = FormStartPosition.Manual;
            DoubleBuffered = true;

            Width = Math.Max(cfg.window.width, 1100);
            Height = Math.Max(cfg.window.height, 720);
            BackColor = Color.White;

            BuildHeader();
            BuildTabs();

            // 快捷键
            KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape) this.Hide();
                else if (e.KeyCode == Keys.Enter) _btnQuery.PerformClick();
            };

            // 缩放后重渲染
            ResizeEnd += (s, e) =>
            {
                if (_parsed != null && _parsed.Records != null)
                {
                    RenderTrend();
                    RenderTopCharts();
                }
            };
        }

        void BuildHeader()
        {
            // 顶部信息区（两行）
            var header = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 96,
                Padding = new Padding(12, 10, 12, 10),
                BackColor = UIStyle.HeaderBack,
                ColumnCount = 2,
                RowCount = 2
            };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            header.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            header.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));

            // 标题（左上）
            _lblTitle.Dock = DockStyle.Fill;
            _lblTitle.AutoEllipsis = true;
            _lblTitle.Text = "—";
            _lblTitle.Font = UIStyle.H1;
            _lblTitle.ForeColor = UIStyle.TextDark;

            // 摘要（右上）
            _lblSummary.Dock = DockStyle.Fill;
            _lblSummary.TextAlign = ContentAlignment.MiddleRight;
            _lblSummary.Font = new Font(UIStyle.Body, FontStyle.Regular);
            _lblSummary.ForeColor = UIStyle.TextDark;

            // 输入 + 重新查询（左下）
            var pnlLeft = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, Padding = new Padding(0), Margin = new Padding(0) };
            var lblInput = new Label { Text = "选中文本：", AutoSize = true, Margin = new Padding(0, 8, 8, 0), Font = UIStyle.Body };
            _boxInput.Width = 520; _boxInput.Font = UIStyle.Body;
            StyleBtn(_btnQuery, "重新查询");
            _btnQuery.Click += async (s, e) =>
            {
                SetLoading("查询中...");
                var textNow = _boxInput.Text.Trim();
                var raw = await ApiHelper.QueryAsync(_cfg, textNow);
                ApplyRawText(textNow, raw);
            };
            pnlLeft.Controls.AddRange(new Control[] { lblInput, _boxInput, _btnQuery });

            // 导出（右下）
            var pnlRight = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(0), Margin = new Padding(0) };
            StyleBtn(_btnExport, "导出Excel");
            _btnExport.Click += (s, e) => ExportExcel();
            pnlRight.Controls.Add(_btnExport);

            header.Controls.Add(_lblTitle, 0, 0);
            header.Controls.Add(_lblSummary, 1, 0);
            header.Controls.Add(pnlLeft, 0, 1);
            header.Controls.Add(pnlRight, 1, 1);
            Controls.Add(header);
        }

        void StyleBtn(Button b, string text)
        {
            b.Text = text;
            b.FlatStyle = FlatStyle.Flat;
            b.FlatAppearance.BorderSize = 0;
            b.Height = 28;
            b.Width = 96;
            b.Margin = new Padding(8, 4, 0, 0);
            b.BackColor = Color.White;
            b.Font = UIStyle.Body;
        }

        void BuildTabs()
        {
            _tabs.Dock = DockStyle.Fill;
            Controls.Add(_tabs);

            // 概览页
            var pageOverview = new TabPage("概览") { BackColor = Color.White };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2, Padding = new Padding(12) };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 45)); // 趋势图 45%
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 55)); // Top10 55%

            _pvTrend.Dock = DockStyle.Fill;
            _pvSizeTop.Dock = DockStyle.Fill;
            _pvColorTop.Dock = DockStyle.Fill;

            layout.Controls.Add(_pvTrend, 0, 0);
            layout.SetColumnSpan(_pvTrend, 2);
            layout.Controls.Add(_pvSizeTop, 0, 1);
            layout.Controls.Add(_pvColorTop, 1, 1);

            pageOverview.Controls.Add(layout);
            _tabs.TabPages.Add(pageOverview);

            // 明细页（加工数据）
            var pageDetail = new TabPage("明细") { BackColor = Color.White };
            var panelDetail = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(12) };
            panelDetail.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            panelDetail.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var bar = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(0) };
            _boxSearch.Width = 220; _boxSearch.Font = UIStyle.Body; _boxSearch.PlaceholderText = "搜索（日期/款式/尺码/颜色）";
            _boxSearch.TextChanged += (s, e) => ApplyFilter(_boxSearch.Text);
            bar.Controls.Add(_boxSearch);

            _grid.Dock = DockStyle.Fill;
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.AllowUserToDeleteRows = false;
            _grid.RowHeadersVisible = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _grid.BackgroundColor = Color.White;
            _grid.BorderStyle = BorderStyle.None;
            _grid.EnableHeadersVisualStyles = false;
            _grid.ColumnHeadersDefaultCellStyle.BackColor = Color.White;
            _grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(50, 50, 50);
            _grid.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei UI", 10, FontStyle.Bold);
            _grid.DefaultCellStyle.Font = UIStyle.Body;
            _grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(249, 251, 253);
            _grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(230, 243, 255);
            _grid.DefaultCellStyle.SelectionForeColor = Color.Black;
            _grid.DataSource = _binding;

            panelDetail.Controls.Add(bar, 0, 0);
            panelDetail.Controls.Add(_grid, 0, 1);
            pageDetail.Controls.Add(panelDetail);
            _tabs.TabPages.Add(pageDetail);
        }

        // 外部调用
        public void ShowAndFocusNearCursor(bool topMost)
        {
            var p = Cursor.Position;
            var targetX = Math.Max(0, Math.Min(Screen.PrimaryScreen.WorkingArea.Width - Width, p.X + 12));
            var targetY = Math.Max(0, Math.Min(Screen.PrimaryScreen.WorkingArea.Height - Height, p.Y + 12));
            Location = new Point(targetX, targetY);

            TopMost = topMost;
            if (!Visible) Show();
            WindowState = FormWindowState.Normal;
            Activate();
            BringToFront();
            Focus();
        }
        public void ShowAndFocusCentered(bool topMost)
        {
            // 取光标所在屏的工作区（排除任务栏）
            var wa = Screen.FromPoint(Cursor.Position).WorkingArea;

            int x = wa.Left + (wa.Width  - Width)  / 2;
            int y = wa.Top  + (wa.Height - Height) / 2;

            Location = new Point(Math.Max(wa.Left, x), Math.Max(wa.Top, y));
            TopMost = topMost;

            if (!Visible) Show();
            WindowState = FormWindowState.Normal;
            Activate();
            BringToFront();
            Focus();
        }

        public void ShowNoActivateAtCursor()
        {
            var p = Cursor.Position;
            var targetX = Math.Max(0, Math.Min(Screen.PrimaryScreen.WorkingArea.Width - Width, p.X + 12));
            var targetY = Math.Max(0, Math.Min(Screen.PrimaryScreen.WorkingArea.Height - Height, p.Y + 12));
            Location = new Point(targetX, targetY);
            if (!Visible) Show();
        }
        public void FocusInput() => _boxInput.Focus();

        public void SetLoading(string message)
        {
            _lblTitle.Text = "—";
            _lblSummary.Text = message ?? "";
            _pvTrend.Model    = new PlotModel { Title = "最近日销量（加载中）" };
            _pvSizeTop.Model  = new PlotModel { Title = "尺码 Top10（加载中）" };
            _pvColorTop.Model = new PlotModel { Title = "颜色 Top10（加载中）" };
            _binding.DataSource = null;
            _detailRows.Clear();
        }

        public void ApplyRawText(string input, string rawMsg)
        {
            _boxInput.Text = input ?? "";
            var pretty = Formatter.Prettify(rawMsg ?? "");

            _parsed = PayloadParser.Parse(pretty ?? "");

            // 标题与摘要
            _lblTitle.Text = string.IsNullOrEmpty(_parsed.Title) ? "—" : _parsed.Title;
            var yesterday = string.IsNullOrEmpty(_parsed.Yesterday) ? "昨日：—" : _parsed.Yesterday;
            var sum7d = _parsed.Sum7d.HasValue ? $"近7天：{_parsed.Sum7d.Value:N0}" : "近7天：—";
            _lblSummary.Text = $"{yesterday}    {sum7d}";

            // 图表
            RenderTrend();
            RenderTopCharts();

            // 明细（只用加工后数据）
            _detailRows = _parsed.Records
                .OrderByDescending(r => r.Date).ThenByDescending(r => r.Qty)
                .Select(r => new
                {
                    日期 = r.Date.ToString("yyyy-MM-dd"),
                    款式 = r.Name,
                    尺码 = string.IsNullOrWhiteSpace(r.Size) ? "(未知)" : r.Size,
                    颜色 = string.IsNullOrWhiteSpace(r.Color) ? "(未知)" : r.Color,
                    数量 = r.Qty
                }).Cast<dynamic>().ToList();

            _binding.DataSource = _detailRows;
        }

        // ============== 图表渲染 ==============
        void RenderTrend()
        {
            if (_parsed.Records == null || _parsed.Records.Count == 0)
            {
                _pvTrend.Model = new PlotModel { Title = "最近日销量（无数据）" };
                return;
            }

            var maxDay = _parsed.Records.Max(x => x.Date).Date;
            var start = maxDay.AddDays(-6);
            var dict = _parsed.Records
                .GroupBy(r => r.Date.Date)
                .ToDictionary(g => g.Key, g => g.Sum(x => x.Qty));
            var days = Enumerable.Range(0, 7).Select(i => start.AddDays(i)).ToList();

            var model = new PlotModel { Title = "最近 7 天销量趋势" };
            model.TextColor = UIStyle.Ox(UIStyle.TextDark);
            model.PlotAreaBorderColor = UIStyle.Ox(220, 224, 230);

            var cat = new CategoryAxis
            {
                Position = AxisPosition.Bottom,
                IsPanEnabled = false,
                IsZoomEnabled = false,
                Angle = -45,
                MajorGridlineStyle = LineStyle.Solid,
                MajorGridlineColor = UIStyle.Ox(UIStyle.GridLine)
            };
            foreach (var d in days) cat.Labels.Add(d.ToString("MM-dd"));
            model.Axes.Add(cat);

            var yAxis = new LinearAxis
            {
                Position = AxisPosition.Left,
                MinorGridlineStyle = LineStyle.Dot,
                MajorGridlineStyle = LineStyle.Solid,
                MajorGridlineColor = UIStyle.Ox(UIStyle.GridLine)
            };
            model.Axes.Add(yAxis);

            var series = new LineSeries
            {
                Color = UIStyle.Ox(UIStyle.AccentBlue),
                MarkerFill = UIStyle.Ox(UIStyle.AccentBlue),
                MarkerType = MarkerType.Circle,
                MarkerSize = 3.5,
                StrokeThickness = 2.2,
                TrackerFormatString = "日期：{2}\n销量：{4}"
            };
            for (int i = 0; i < days.Count; i++)
            {
                var qty = dict.TryGetValue(days[i], out var v) ? v : 0;
                series.Points.Add(new DataPoint(i, qty));
            }
            model.Series.Add(series);

            ApplyResponsiveStyles(model, xTitle: "日期", yTitle: "销量", axisAngle: -45);
            _pvTrend.Model = model;
        }

        void RenderTopCharts()
        {
            // 尺码 Top10
            var bySize = _parsed.Records
                .GroupBy(r => string.IsNullOrWhiteSpace(r.Size) ? "(未知)" : r.Size)
                .Select(g => new { Key = g.Key, Qty = g.Sum(x => x.Qty) })
                .OrderByDescending(x => x.Qty).Take(10).ToList();

            var modelSize = new PlotModel { Title = "尺码 Top10（7天）" };
            modelSize.TextColor = UIStyle.Ox(UIStyle.TextDark);
            modelSize.PlotAreaBorderColor = UIStyle.Ox(220, 224, 230);

            var catSize = new CategoryAxis
            {
                Position = AxisPosition.Left,
                IsPanEnabled = false,
                IsZoomEnabled = false,
                StartPosition = 1, EndPosition = 0, // 最大在上
                MajorGridlineStyle = LineStyle.Solid,
                MajorGridlineColor = UIStyle.Ox(UIStyle.GridLine)
            };
            var serSize = new BarSeries
            {
                LabelPlacement = LabelPlacement.Outside,
                LabelFormatString = "{0}",
                TrackerFormatString = "尺码：{Category}\n销量：{Value}",
                FillColor = UIStyle.Ox(30, 180, 120)
            };
            foreach (var it in bySize)
            {
                catSize.Labels.Add(it.Key);
                serSize.Items.Add(new BarItem(it.Qty));
            }
            modelSize.Axes.Add(catSize);
            modelSize.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, MinorGridlineStyle = LineStyle.Dot, MajorGridlineStyle = LineStyle.Solid, MajorGridlineColor = UIStyle.Ox(UIStyle.GridLine) });
            modelSize.Series.Add(serSize);
            ApplyResponsiveStyles(modelSize, xTitle: "销量", yTitle: "尺码");
            _pvSizeTop.Model = modelSize;

            // 颜色 Top10
            var byColor = _parsed.Records
                .GroupBy(r => string.IsNullOrWhiteSpace(r.Color) ? "(未知)" : r.Color)
                .Select(g => new { Key = g.Key, Qty = g.Sum(x => x.Qty) })
                .OrderByDescending(x => x.Qty).Take(10).ToList();

            var modelColor = new PlotModel { Title = "颜色 Top10（7天）" };
            modelColor.TextColor = UIStyle.Ox(UIStyle.TextDark);
            modelColor.PlotAreaBorderColor = UIStyle.Ox(220, 224, 230);

            var catColor = new CategoryAxis
            {
                Position = AxisPosition.Left,
                IsPanEnabled = false,
                IsZoomEnabled = false,
                StartPosition = 1, EndPosition = 0,
                MajorGridlineStyle = LineStyle.Solid,
                MajorGridlineColor = UIStyle.Ox(UIStyle.GridLine)
            };
            var serColor = new BarSeries
            {
                LabelPlacement = LabelPlacement.Outside,
                LabelFormatString = "{0}",
                TrackerFormatString = "颜色：{Category}\n销量：{Value}",
                FillColor = UIStyle.Ox(UIStyle.AccentBlue)
            };
            foreach (var it in byColor)
            {
                catColor.Labels.Add(it.Key);
                serColor.Items.Add(new BarItem(it.Qty));
            }
            modelColor.Axes.Add(catColor);
            modelColor.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, MinorGridlineStyle = LineStyle.Dot, MajorGridlineStyle = LineStyle.Solid, MajorGridlineColor = UIStyle.Ox(UIStyle.GridLine) });
            modelColor.Series.Add(serColor);
            ApplyResponsiveStyles(modelColor, xTitle: "销量", yTitle: "颜色");
            _pvColorTop.Model = modelColor;
        }

        // 自适配
        void ApplyResponsiveStyles(PlotModel model, string xTitle = null, string yTitle = null, double? fontScale = null, double? axisAngle = null)
        {
            if (model == null) return;

            var scale = fontScale ?? (Width < 800 ? 0.75 : (Width < 1000 ? 0.85 : 1.0));
            if (model.DefaultFontSize == 0) model.DefaultFontSize = 12;
            model.DefaultFontSize *= scale;

            foreach (var ax in model.Axes)
            {
                if (xTitle != null && (ax.Position == AxisPosition.Bottom || ax.Position == AxisPosition.Top))
                    ax.Title = xTitle;
                if (yTitle != null && (ax.Position == AxisPosition.Left || ax.Position == AxisPosition.Right))
                    ax.Title = yTitle;

                if (ax.FontSize == 0) ax.FontSize = model.DefaultFontSize;
                ax.FontSize *= scale;

                if (axisAngle.HasValue && ax is CategoryAxis cat && ax.Position == AxisPosition.Bottom)
                    cat.Angle = axisAngle.Value;
            }
        }

        // 过滤
        void ApplyFilter(string keyword)
        {
            if (_detailRows == null || _detailRows.Count == 0)
            {
                _binding.DataSource = null;
                return;
            }
            keyword = (keyword ?? "").Trim();
            if (keyword.Length == 0)
            {
                _binding.DataSource = _detailRows;
                return;
            }
            var lower = keyword.ToLowerInvariant();
            var filtered = _detailRows.Where(r =>
            {
                string d = r.日期?.ToString()?.ToLowerInvariant() ?? "";
                string n = r.款式?.ToString()?.ToLowerInvariant() ?? "";
                string s = r.尺码?.ToString()?.ToLowerInvariant() ?? "";
                string c = r.颜色?.ToString()?.ToLowerInvariant() ?? "";
                string q = r.数量?.ToString()?.ToLowerInvariant() ?? "";
                return d.Contains(lower) || n.Contains(lower) || s.Contains(lower) || c.Contains(lower) || q.Contains(lower);
            }).ToList();
            _binding.DataSource = filtered;
        }

        // 导出 Excel（ClosedXML）
// 导出 Excel（ClosedXML）
        void ExportExcel()
        {
            try
            {
                var exportsDir = Path.Combine(AppContext.BaseDirectory, "exports");
                Directory.CreateDirectory(exportsDir);
                var file = Path.Combine(exportsDir, $"StyleWatcher_{DateTime.Now:yyyyMMdd_HHmm}.xlsx");

                using var wb = new XLWorkbook();

                // === Sheet 1: 明细 ===
                var ws2 = wb.Worksheets.Add("明细");
                ws2.Cell(1, 1).Value = "日期";
                ws2.Cell(1, 2).Value = "款式";
                ws2.Cell(1, 3).Value = "尺码";
                ws2.Cell(1, 4).Value = "颜色";
                ws2.Cell(1, 5).Value = "数量";
                int r = 2;
                foreach (dynamic row in _detailRows)
                {
                    ws2.Cell(r, 1).Value = row.日期;
                    ws2.Cell(r, 2).Value = row.款式;
                    ws2.Cell(r, 3).Value = row.尺码;
                    ws2.Cell(r, 4).Value = row.颜色;
                    ws2.Cell(r, 5).Value = row.数量;
                    r++;
                }
                ws2.RangeUsed().SetAutoFilter();
                ws2.Columns().AdjustToContents();

                // === Sheet 2: 趋势7天 ===
                var ws1 = wb.Worksheets.Add("趋势7天");
                // 从 _parsed 聚合近 7 天
                var maxDay = _parsed.Records.Any() ? _parsed.Records.Max(x => x.Date).Date : DateTime.Today;
                var start = maxDay.AddDays(-6);
                var dict = _parsed.Records
                    .GroupBy(rr => rr.Date.Date)
                    .ToDictionary(g => g.Key, g => g.Sum(x => x.Qty));

                ws1.Cell(1, 1).Value = "日期";
                ws1.Cell(1, 2).Value = "销量";
                int rr = 2;
                for (int i = 0; i < 7; i++)
                {
                    var d = start.AddDays(i);
                    var qty = dict.TryGetValue(d, out var v) ? v : 0;
                    ws1.Cell(rr, 1).Value = d.ToString("yyyy-MM-dd");
                    ws1.Cell(rr, 2).Value = qty;
                    rr++;
                }
                ws1.Columns().AdjustToContents();

                // 不再导出“摘要”sheet
                wb.SaveAs(file);

                var res = MessageBox.Show($"已导出：\n{file}\n\n是否打开所在文件夹？", "导出成功",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (res == DialogResult.Yes)
                {
                    try { System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{file}\""); } catch { }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}
