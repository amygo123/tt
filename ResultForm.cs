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
    }

    public class ResultForm : Form
    {
        private readonly AppConfig _cfg;
        private readonly TextBox _input = new TextBox();
        private readonly Button _btnQuery = new Button();
        private readonly Button _btnExport = new Button();
        // 仍保留字段以兼容 TrayApp 的 SetLoading/Bind 调用，但不再添加到 UI
        private readonly Label _lblTitle = new Label();
        private readonly Label _lblSummary = new Label();
        private readonly TabControl _tabs = new TabControl();

        private PlotView _plot7d = new PlotView();
        private PlotView _plotSize = new PlotView();
        private PlotView _plotColor = new PlotView();

        private DataGridView _grid = new DataGridView();
        private BindingSource _binding = new BindingSource();
        private TextBox _boxSearch = new TextBox();

        public ResultForm(AppConfig cfg)
        {
            _cfg = cfg;
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
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 64)); // 顶栏
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // 内容
            Controls.Add(root);

            // 顶栏
            var top = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, Padding = new Padding(12,10,12,6), BackColor = UIStyle.HeaderBack };
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _input.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            _input.MinimumSize = new Size(300, 28);
            _btnQuery.Text = "重新查询";
            _btnQuery.AutoSize = true; _btnQuery.AutoSizeMode = AutoSizeMode.GrowAndShrink; _btnQuery.Padding = new Padding(10,4,10,4);
            _btnQuery.Click += async (s, e) => { _btnQuery.Enabled = false; try { await ReloadAsync(); } finally { _btnQuery.Enabled = true; } };
            _btnExport.Text = "导出 Excel";
            _btnExport.AutoSize = true; _btnExport.AutoSizeMode = AutoSizeMode.GrowAndShrink; _btnExport.Padding = new Padding(10,4,10,4);
            _btnExport.Click += (s, e) => ExportExcel();
            top.Controls.Add(_input, 0, 0);
            top.Controls.Add(_btnQuery, 1, 0);
            top.Controls.Add(_btnExport, 2, 0);
            root.Controls.Add(top, 0, 0);

            // Tabs
            _tabs.Dock = DockStyle.Fill;
            root.Controls.Add(_tabs, 0, 1);

            BuildTabs();
        }

        private void BuildTabs()
        {
            // 概览页
            var pageOverview = new TabPage("概览") { BackColor = Color.White };
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(12) };
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 34));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 33));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 33));
            _plot7d.Dock = DockStyle.Fill;
            _plotSize.Dock = DockStyle.Fill;
            _plotColor.Dock = DockStyle.Fill;
            panel.Controls.Add(_plot7d, 0, 0);
            panel.Controls.Add(_plotSize, 0, 1);
            panel.Controls.Add(_plotColor, 0, 2);
            pageOverview.Controls.Add(panel);
            _tabs.TabPages.Add(pageOverview);

            // 销售明细
            var pageDetail = new TabPage("销售明细") { BackColor = Color.White };
            var panelDetail = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(12) };
            panelDetail.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            panelDetail.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _boxSearch.Dock = DockStyle.Fill; _boxSearch.Font = UIStyle.Body; _boxSearch.PlaceholderText = "搜索（日期/款式/尺码/颜色/数量）";
            _boxSearch.TextChanged += (s, e) => ApplyFilter(_boxSearch.Text);
            panelDetail.Controls.Add(_boxSearch, 0, 0);
            _grid.Dock = DockStyle.Fill;
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.AllowUserToDeleteRows = false;
            _grid.RowHeadersVisible = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _grid.DataSource = _binding;
            panelDetail.Controls.Add(_grid, 0, 1);
            pageDetail.Controls.Add(panelDetail);
            _tabs.TabPages.Add(pageDetail);

            // 库存页
            _tabs.TabPages.Add(new InventoryTabPage(_cfg));
        }

        // 公开方法（TrayApp 调用保持兼容）
        public void FocusInput(){ try{ if(WindowState==FormWindowState.Minimized) WindowState=FormWindowState.Normal; _input.Focus(); _input.SelectAll(); }catch{} }
        public void ShowNoActivateAtCursor(){ try{ StartPosition=FormStartPosition.Manual; var pt=Cursor.Position; Location=new Point(Math.Max(0,pt.X-Width/2),Math.Max(0,pt.Y-Height/2)); Show(); }catch{ Show(); } }
        public void ShowAndFocusCentered(){ try{ StartPosition=FormStartPosition.CenterScreen; Show(); Activate(); FocusInput(); }catch{ Show(); } }
        public void ShowAndFocusCentered(bool alwaysOnTop){ TopMost=alwaysOnTop; ShowAndFocusCentered(); }
        public void SetLoading(string message){ _lblTitle.Text=message??""; _lblSummary.Text=""; _tabs.Enabled=false; Refresh(); }
        public void SetLoading(bool busy, string? message=null){ _lblTitle.Text=message??(busy?"加载中…":""); _lblSummary.Text=""; _tabs.Enabled=!busy; Refresh(); }
        public async void ApplyRawText(string selection, string parsed){ _input.Text=selection??string.Empty; await LoadTextAsync(parsed??string.Empty); }
        public void ApplyRawText(string text){ _input.Text=text??string.Empty; }

        // 数据绑定
        public async System.Threading.Tasks.Task LoadTextAsync(string raw)=>await ReloadAsync(raw);
        private async System.Threading.Tasks.Task ReloadAsync()=>await ReloadAsync(_input.Text);
        private async System.Threading.Tasks.Task ReloadAsync(string displayText)
        {
            _lblTitle.Text = "";
            _lblSummary.Text = "";
            await System.Threading.Tasks.Task.Yield();
            var parsed = StyleWatcherWin.PayloadParser.Parse(displayText);
            Bind(parsed);
            _tabs.Enabled = true;
        }

        private void Bind(ParsedPayload p)
        {
            // 图表
            RenderCharts(p);

            // 明细绑定
            _binding.DataSource = p.Records.Select(r => new
            {
                日期 = r.Date.ToString("yyyy-MM-dd"),
                款式 = r.Name,
                尺码 = r.Size,
                颜色 = r.Color,
                数量 = r.Qty
            }).ToList();

            _grid.ClearSelection();
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

        private void RenderCharts(ParsedPayload p)
        {
            // 近7天趋势
            var trend = p.Records
                .GroupBy(r => r.Date.Date)
                .OrderBy(g => g.Key)
                .Select(g => new { Day = g.Key, Qty = g.Sum(x => x.Qty) })
                .ToList();

            var modelTrend = new PlotModel { Title = "近 7 天总销量趋势" };
            var xAxis = new DateTimeAxis { Position = AxisPosition.Bottom, StringFormat = "MM-dd", IntervalType = DateTimeIntervalType.Days, MinorIntervalType = DateTimeIntervalType.Days };
            var yAxis = new LinearAxis { Position = AxisPosition.Left, MinimumPadding = 0, AbsoluteMinimum = 0 };
            modelTrend.Axes.Add(xAxis); modelTrend.Axes.Add(yAxis);
            var s1 = new LineSeries { MarkerType = MarkerType.Circle };
            foreach (var pt in trend) s1.Points.Add(new DataPoint(DateTimeAxis.ToDouble(pt.Day), pt.Qty));
            modelTrend.Series.Add(s1);
            _plot7d.Model = modelTrend;

            // 尺码：展示全量，按销量降序
            var sizeAgg = p.Records.GroupBy(r => r.Size).Select(g => new { Key = g.Key, Qty = g.Sum(x => x.Qty) })
                .OrderByDescending(x => x.Qty).ToList();
            var modelSize = new PlotModel { Title = "各尺码销量（降序，全部）" };
            var sizeAxis = new CategoryAxis { Position = AxisPosition.Left, ItemsSource = sizeAgg, LabelField = "Key" };
            var sizeVal = new LinearAxis { Position = AxisPosition.Bottom, MinimumPadding = 0, AbsoluteMinimum = 0 };
            modelSize.Axes.Add(sizeAxis); modelSize.Axes.Add(sizeVal);
            modelSize.Series.Add(new BarSeries { ItemsSource = sizeAgg.Select(x => new BarItem { Value = x.Qty }) });
            _plotSize.Model = modelSize;

            // 颜色：展示全量，按销量降序
            var colorAgg = p.Records.GroupBy(r => r.Color).Select(g => new { Key = g.Key, Qty = g.Sum(x => x.Qty) })
                .OrderByDescending(x => x.Qty).ToList();
            var modelColor = new PlotModel { Title = "各颜色销量（降序，全部）" };
            var colorAxis = new CategoryAxis { Position = AxisPosition.Left, ItemsSource = colorAgg, LabelField = "Key" };
            var colorVal = new LinearAxis { Position = AxisPosition.Bottom, MinimumPadding = 0, AbsoluteMinimum = 0 };
            modelColor.Axes.Add(colorAxis); modelColor.Axes.Add(colorVal);
            modelColor.Series.Add(new BarSeries { ItemsSource = colorAgg.Select(x => new BarItem { Value = x.Qty }) });
            _plotColor.Model = modelColor;
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
