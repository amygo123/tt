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
    static class UIStyle
    {
        public static readonly Font H1   = new Font("Microsoft YaHei UI", 13, FontStyle.Bold);
        public static readonly Font Body = new Font("Microsoft YaHei UI", 10, FontStyle.Regular);
        public static readonly Color HeaderBack = Color.FromArgb(245, 247, 250);
        public static readonly Color TextDark   = Color.FromArgb(47, 47, 47);
        public static readonly Color Ok   = Color.FromArgb(32, 178, 170);
        public static readonly Color Warn = Color.FromArgb(255, 165, 0);
        public static readonly Color Danger = Color.FromArgb(220, 20, 60);
        public static readonly Color CardBg = Color.FromArgb(250, 250, 250);
    }

    public class ResultForm : Form
    {
        private readonly AppConfig _cfg;
        private readonly TextBox _input = new TextBox();
        private readonly Button _btnQuery = new Button();
        private readonly Button _btnExport = new Button();
        private readonly TabControl _tabs = new TabControl();

        // KPI
        private FlowLayoutPanel _kpiBar = new FlowLayoutPanel();
        private Panel _kpiSales7 = new Panel();
        private Panel _kpiDoc = new Panel();
        private Panel _kpiInv = new Panel();
        private Panel _kpiMissing = new Panel();

        // 概览图
        private PlotView _plot7d = new PlotView();
        private PlotView _plotSize = new PlotView();
        private PlotView _plotColor = new PlotView();
        private PlotView _plotWarehouse = new PlotView();

        // 趋势窗口切换
        private FlowLayoutPanel _trendSwitch = new FlowLayoutPanel();
        private int _trendWindow = 7;

        // 销售明细
        private DataGridView _grid = new DataGridView();
        private BindingSource _binding = new BindingSource();
        private TextBox _boxSearch = new TextBox();

        // 库存页
        private InventoryTabPage _invPage;

        public ResultForm(AppConfig cfg)
        {
            _cfg = cfg;
            _trendWindow = (_cfg.ui?.trendWindows?.FirstOrDefault() ?? 7);
            InitUI();
        }

        private void InitUI()
        {
            Text = "Style Watcher";
            StartPosition = FormStartPosition.CenterScreen;
            TopMost = _cfg.window.alwaysOnTop;
            Font = new Font("Microsoft YaHei UI", _cfg.window.fontSize, FontStyle.Regular);
            Width = _cfg.window.width;
            Height = _cfg.window.height;

            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, BackColor = Color.White };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 68)); // 顶栏稍增高，避免重叠
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // 内容
            Controls.Add(root);

            // 顶栏
            var top = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, Padding = new Padding(12,10,12,6), BackColor = UIStyle.HeaderBack };
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _input.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            _input.MinimumSize = new Size(360, 30);
            _btnQuery.Text = "重新查询";
            _btnQuery.AutoSize = true; _btnQuery.AutoSizeMode = AutoSizeMode.GrowAndShrink; _btnQuery.Padding = new Padding(10,6,10,6);
            _btnQuery.Click += async (s, e) => { _btnQuery.Enabled = false; try { await ReloadAsync(); } finally { _btnQuery.Enabled = true; } };
            _btnExport.Text = "导出 Excel";
            _btnExport.AutoSize = true; _btnExport.AutoSizeMode = AutoSizeMode.GrowAndShrink; _btnExport.Padding = new Padding(10,6,10,6);
            _btnExport.Click += (s, e) => ExportExcel();
            top.Controls.Add(_input, 0, 0);
            top.Controls.Add(_btnQuery, 1, 0);
            top.Controls.Add(_btnExport, 2, 0);
            root.Controls.Add(top, 0, 0);

            // 内容容器：KPI + Tabs
            var content = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
            content.RowStyles.Add(new RowStyle(SizeType.Absolute, 118)); // 再抬高，彻底避免文字裁剪
            content.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.Controls.Add(content, 0, 1);

            // KPI 条
            _kpiBar.Dock = DockStyle.Fill;
            _kpiBar.Height = 118;
            _kpiBar.FlowDirection = FlowDirection.LeftToRight;
            _kpiBar.WrapContents = true;
            _kpiBar.Padding = new Padding(12, 8, 12, 8);
            _kpiBar.AutoScroll = false;
            _kpiBar.Controls.Add(MakeKpi(_kpiSales7, "近7日销量", "—", Aggregations.AlertLevel.Unknown));
            _kpiBar.Controls.Add(MakeKpi(_kpiInv, "可用库存(总)", "—", Aggregations.AlertLevel.Unknown));
            _kpiBar.Controls.Add(MakeKpi(_kpiDoc, "库存天数(DoC)", "—", Aggregations.AlertLevel.Unknown));
            _kpiBar.Controls.Add(MakeKpi(_kpiMissing, "缺尺码数", "—", Aggregations.AlertLevel.Unknown));
            content.Controls.Add(_kpiBar, 0, 0);

            // Tabs
            _tabs.Dock = DockStyle.Fill;
            content.Controls.Add(_tabs, 0, 1);

            BuildTabs();
        }

        private Control MakeKpi(Panel host, string title, string value, Aggregations.AlertLevel level)
        {
            host.Padding = new Padding(10);
            host.Margin = new Padding(8, 6, 8, 6);
            host.BackColor = UIStyle.CardBg;
            host.BorderStyle = BorderStyle.FixedSingle;
            host.Width = 250; host.Height = 98;

            var t = new Label { Text = title, AutoSize = false, Dock = DockStyle.Top, Height = 22, ForeColor = UIStyle.TextDark };
            var v = new Label { Text = value, AutoSize = false, Dock = DockStyle.Fill, Font = new Font(Font, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft };
            v.ForeColor = level switch
            {
                Aggregations.AlertLevel.Green => UIStyle.Ok,
                Aggregations.AlertLevel.Yellow => UIStyle.Warn,
                Aggregations.AlertLevel.Red => UIStyle.Danger,
                _ => UIStyle.TextDark
            };

            host.Controls.Clear();
            host.Controls.Add(v);
            host.Controls.Add(t);
            return host;
        }

        private void BuildTabs()
        {
            // 概览页
            var pageOverview = new TabPage("概览") { BackColor = Color.White };
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5, Padding = new Padding(12) };
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 44)); // 趋势切换
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 20));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 20));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 20));

            // 趋势切换
            _trendSwitch.Dock = DockStyle.Fill;
            _trendSwitch.FlowDirection = FlowDirection.LeftToRight;
            _trendSwitch.WrapContents = false;
            _trendSwitch.Padding = new Padding(6, 8, 6, 0);
            _trendSwitch.Margin = new Padding(0);
            var windows = _cfg.ui?.trendWindows ?? new[] { 7, 14, 30 };
            foreach (var w in windows.Distinct().OrderBy(x => x))
            {
                var rb = new RadioButton { Text = $"{w} 日", AutoSize = true, Tag = w, Margin = new Padding(0, 6, 18, 0) };
                if (w == _trendWindow) rb.Checked = true;
                rb.CheckedChanged += async (s, e) =>
                {
                    var me = (RadioButton)s;
                    if (me.Checked) { _trendWindow = (int)me.Tag; await ReloadAsync(); }
                };
                _trendSwitch.Controls.Add(rb);
            }
            panel.Controls.Add(_trendSwitch, 0, 0);

            // 四张图
            _plot7d.Dock = DockStyle.Fill;
            _plotSize.Dock = DockStyle.Fill;
            _plotColor.Dock = DockStyle.Fill;
            _plotWarehouse.Dock = DockStyle.Fill;
            panel.Controls.Add(_plot7d, 0, 1);
            panel.Controls.Add(_plotSize, 0, 2);
            panel.Controls.Add(_plotColor, 0, 3);
            panel.Controls.Add(_plotWarehouse, 0, 4);
            pageOverview.Controls.Add(panel);
            _tabs.TabPages.Add(pageOverview);

            // 销售明细
            var pageDetail = new TabPage("销售明细") { BackColor = Color.White };
            var panelDetail = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(12) };
            panelDetail.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            panelDetail.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _boxSearch.Dock = DockStyle.Fill; _boxSearch.Font = UIStyle.Body; _boxSearch.PlaceholderText = "搜索（日期/款式/尺码/颜色/数量）";
            _boxSearch.TextChanged += (s, e) => ApplyFilter(_boxSearch.Text);
            panelDetail.Controls.Add(_boxSearch, 0, 0);
            _grid.Dock = DockStyle.Fill;
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.AllowUserToDeleteRows = false;
            _grid.RowHeadersVisible = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            _grid.DataSource = _binding;
            panelDetail.Controls.Add(_grid, 0, 1);
            pageDetail.Controls.Add(panelDetail);
            _tabs.TabPages.Add(pageDetail);

            // 库存页
            _invPage = new InventoryTabPage(_cfg);
            _invPage.SummaryUpdated += OnInventorySummaryUpdated;
            _tabs.TabPages.Add(_invPage);
        }

        // ===== TrayApp 兼容方法 =====
        public void FocusInput(){ try{ if(WindowState==FormWindowState.Minimized) WindowState=FormWindowState.Normal; _input.Focus(); _input.SelectAll(); }catch{} }
        public void ShowNoActivateAtCursor(){ try{ StartPosition=FormStartPosition.Manual; var pt=Cursor.Position; Location=new Point(Math.max(0,pt.X-Width/2),Math.max(0,pt.Y-Height/2)); Show(); }catch{ Show(); } }
        public void ShowAndFocusCentered(){ try{ StartPosition=FormStartPosition.CenterScreen; Show(); Activate(); FocusInput(); }catch{ Show(); } }
        public void ShowAndFocusCentered(bool alwaysOnTop){ TopMost=alwaysOnTop; ShowAndFocusCentered(); }
        public void SetLoading(string message){ /* 占位 */ }
        public void SetLoading(bool busy, string? message=null){ /* 占位 */ }
        public async void ApplyRawText(string selection, string parsed){ _input.Text=selection??string.Empty; await LoadTextAsync(parsed??string.Empty); }
        public void ApplyRawText(string text){ _input.Text=text??string.Empty; }

        // ===== 加载/绑定 =====
        public async System.Threading.Tasks.Task LoadTextAsync(string raw)=>await ReloadAsync(raw);
        private async System.Threading.Tasks.Task ReloadAsync()=>await ReloadAsync(_input.Text);

        private async System.Threading.Tasks.Task ReloadAsync(string displayText)
        {
            await System.Threading.Tasks.Task.Yield();
            var parsed = StyleWatcherWin.PayloadParser.Parse(displayText);

            // KPI（销售侧）
            var salesItems = parsed.Records.Select(r => new Aggregations.SalesItem
            {
                Date = r.Date, Size = r.Size ?? "", Color = r.Color ?? "", Qty = r.Qty
            }).ToList();
            var sales7 = salesItems.Where(x => x.Date >= DateTime.Today.AddDays(-6)).Sum(x => x.Qty);
            MakeKpi(_kpiSales7, "近7日销量", sales7.ToString(), Aggregations.AlertLevel.Unknown);

            // 库存 KPI 等待库存页触发 SummaryUpdated 来刷新
            MakeKpi(_kpiInv, "可用库存(总)", "—", Aggregations.AlertLevel.Unknown);
            MakeKpi(_kpiDoc, "库存天数(DoC)", "—", Aggregations.AlertLevel.Unknown);
            MakeKpi(_kpiMissing, "缺尺码数", "—", Aggregations.AlertLevel.Unknown);

            RenderCharts(salesItems);

            _binding.DataSource = parsed.Records.Select(r => new
            {
                日期 = r.Date.ToString("yyyy-MM-dd"),
                款式 = r.Name,
                尺码 = r.Size,
                颜色 = r.Color,
                数量 = r.Qty
            }).ToList();
            _grid.ClearSelection();
        }

        private void RenderCharts(List<Aggregations.SalesItem> salesItems)
        {
            // 趋势
            var series = Aggregations.BuildDateSeries(salesItems, _trendWindow);
            var modelTrend = new PlotModel { Title = $"近 {_trendWindow} 日总销量趋势", PlotMargins = new OxyThickness(50, 10, 10, 40) };
            var xAxis = new DateTimeAxis
            {
                Position = AxisPosition.Bottom,
                StringFormat = "MM-dd",
                IntervalType = DateTimeIntervalType.Days,
                MinorIntervalType = DateTimeIntervalType.Days,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.None,
                Angle = 0
            };
            var yAxis = new LinearAxis { Position = AxisPosition.Left, MinimumPadding = 0, AbsoluteMinimum = 0, MajorGridlineStyle = LineStyle.Solid };
            modelTrend.Axes.Add(xAxis); modelTrend.Axes.Add(yAxis);
            var line = new LineSeries { MarkerType = MarkerType.Circle };
            foreach (var (day, qty) in series) line.Points.Add(new DataPoint(DateTimeAxis.ToDouble(day), qty));
            modelTrend.Series.Add(line);

            if (_cfg.ui?.showMovingAverage ?? true)
            {
                var ma = Aggregations.MovingAverage(series.Select(x => x.qty).ToList(), 7);
                var maSeries = new LineSeries { LineStyle = LineStyle.Dash, Title = "MA7" };
                for (int i = 0; i < series.Count; i++)
                    maSeries.Points.Add(new DataPoint(DateTimeAxis.ToDouble(series[i].day), ma[i]));
                modelTrend.Series.Add(maSeries);
            }
            _plot7d.Model = modelTrend;

            // 尺码（全量降序）
            var sizeAgg = Aggregations.BySize(salesItems);
            var modelSize = new PlotModel { Title = "各尺码销量（降序，全部）", PlotMargins = new OxyThickness(80, 6, 6, 6) };
            modelSize.Axes.Add(new CategoryAxis { Position = AxisPosition.Left, ItemsSource = sizeAgg, LabelField = "Key", GapWidth = 0.4 });
            modelSize.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, MinimumPadding = 0, AbsoluteMinimum = 0 });
            modelSize.Series.Add(new BarSeries { ItemsSource = sizeAgg.Select(x => new BarItem { Value = x.Qty }) });
            _plotSize.Model = modelSize;

            // 颜色（全量降序）
            var colorAgg = Aggregations.ByColor(salesItems);
            var modelColor = new PlotModel { Title = "各颜色销量（降序，全部）", PlotMargins = new OxyThickness(80, 6, 6, 6) };
            modelColor.Axes.Add(new CategoryAxis { Position = AxisPosition.Left, ItemsSource = colorAgg, LabelField = "Key", GapWidth = 0.4 });
            modelColor.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, MinimumPadding = 0, AbsoluteMinimum = 0 });
            modelColor.Series.Add(new BarSeries { ItemsSource = colorAgg.Select(x => new BarItem { Value = x.Qty }) });
            _plotColor.Model = modelColor;
        }

        private void OnInventorySummaryUpdated()
        {
            var snap = _invPage.GetSummary();
            if (snap == null) return;

            MakeKpi(_kpiInv, "可用库存(总)", snap.总可用.ToString(), Aggregations.AlertLevel.Unknown);
            var sales7Text = _kpiSales7.Controls.OfType<Label>().LastOrDefault()?.Text ?? "0";
            if (!double.TryParse(sales7Text, out var sales7)) sales7 = 0;
            var avg7 = Math.Max(0.01, sales7 / 7.0);
            var doc = Aggregations.DaysOfCover(snap.总可用, avg7);
            MakeKpi(_kpiDoc, "库存天数(DoC)", (doc == int.MaxValue ? "—" : doc.ToString()), Aggregations.AlertLevel.Unknown);

            RenderWarehousePie(snap);
        }

        private void RenderWarehousePie(Aggregations.InventorySnapshot snap)
        {
            var model = new PlotModel { Title = "分仓库存占比（可用）", PlotMargins = new OxyThickness(10, 10, 10, 10) };
            var pie = new PieSeries { StrokeThickness = 0.5, InsideLabelPosition = 0.6, AngleSpan = 360, StartAngle = 0 };
            foreach (var kv in snap.分仓可用.OrderByDescending(k => k.Value))
                pie.Slices.Add(new PieSlice(kv.Key, kv.Value));
            model.Series.Add(pie);
            _plotWarehouse.Model = model;
        }

        private void ApplyFilter(string q)
        {
            q = (q ?? "").Trim();
            var current = _binding.DataSource as IEnumerable<object>;
            if (current == null) return;
            if (string.IsNullOrWhiteSpace(q)) { _binding.ResetBindings(false); return; }

            var filtered = current.Where(x =>
            {
                var t = x.GetType();
                string Get(string name) => t.GetProperty(name)?.GetValue(x)?.ToString() ?? "";
                return Get("日期").Contains(q, StringComparison.OrdinalIgnoreCase)
                    || Get("款式").Contains(q, StringComparison.OrdinalIgnoreCase)
                    || Get("尺码").Contains(q, StringComparison.OrdinalIgnoreCase)
                    || Get("颜色").Contains(q, StringComparison.OrdinalIgnoreCase)
                    || Get("数量").Contains(q, StringComparison.OrdinalIgnoreCase);
            }).ToList();

            _binding.DataSource = filtered;
            _grid.ClearSelection();
        }

        private void ExportExcel()
        {
            var saveDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "exports");
            Directory.CreateDirectory(saveDir);
            var path = Path.Combine(saveDir, $"StyleWatcher_{DateTime.Now:yyyyMMdd_HHmm}.xlsx");

            using var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("销售明细");
            ws1.Cell(1, 1).Value = "日期";
            ws1.Cell(1, 2).Value = "款式";
            ws1.Cell(1, 3).Value = "尺码";
            ws1.Cell(1, 4).Value = "颜色";
            ws1.Cell(1, 5).Value = "数量";

            var list = _binding.DataSource as IEnumerable<object>;
            if (list != null)
            {
                int r = 2;
                foreach (var it in list)
                {
                    var t = it.GetType();
                    ws1.Cell(r, 1).Value = t.GetProperty("日期")?.GetValue(it)?.ToString();
                    ws1.Cell(r, 2).Value = t.GetProperty("款式")?.GetValue(it)?.ToString();
                    ws1.Cell(r, 3).Value = t.GetProperty("尺码")?.GetValue(it)?.ToString();
                    ws1.Cell(r, 4).Value = t.GetProperty("颜色")?.GetValue(it)?.ToString();
                    ws1.Cell(r, 5).Value = t.GetProperty("数量")?.GetValue(it)?.ToString();
                    r++;
                }
            }
            ws1.Columns().AdjustToContents();

            var ws2 = wb.AddWorksheet("趋势7天");
            ws2.Cell(1, 1).Value = "日期";
            ws2.Cell(1, 2).Value = "数量";

            var detail = _binding.DataSource as IEnumerable<object>;
            if (detail != null)
            {
                var trend = detail
                    .Select(x =>
                    {
                        var t = x.GetType();
                        var d = DateTime.Parse(t.GetProperty("日期")?.GetValue(x)?.ToString() ?? DateTime.Now.ToString("yyyy-MM-dd"));
                        var q = int.Parse(t.GetProperty("数量")?.GetValue(x)?.ToString() ?? "0");
                        return new { 日期 = d.Date, 数量 = q };
                    })
                    .GroupBy(x => x.日期)
                    .OrderBy(x => x.Key)
                    .Select(x => new { Day = x.Key.ToString("yyyy-MM-dd"), Qty = x.Sum(y => y.数量) })
                    .ToList();

                int r = 2;
                foreach (var it in trend)
                {
                    ws2.Cell(r, 1).Value = it.Day;
                    ws2.Cell(r, 2).Value = it.Qty;
                    r++;
                }
            }
            ws2.Columns().AdjustToContents();

            wb.SaveAs(path);
            try { System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{path}\""); } catch { }
        }
    }
}
