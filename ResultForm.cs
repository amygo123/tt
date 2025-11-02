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
    }

    public class ResultForm : Form
    {
        private readonly AppConfig _cfg;
        private readonly TextBox _input = new TextBox();
        private readonly Button _btnQuery = new Button();
        private readonly Button _btnExport = new Button();
        private readonly Label _lblTitle = new Label();
        private readonly Label _lblSummary = new Label();
        private readonly TabControl _tabs = new TabControl();

        // 概览页
        private PlotView _plot7d;
        private PlotView _plotSizeTop;
        private PlotView _plotColorTop;

        // 明细页
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

            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, BackColor = Color.White };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Controls.Add(root);

            // 顶部栏：输入 + 按钮
            var bar = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, Padding = new Padding(12, 10, 12, 6), BackColor = UIStyle.HeaderBack };
            _input.Width = 460;
            _btnQuery.Text = "重新查询";
            _btnQuery.Click += async (s, e) =>
            {
                _btnQuery.Enabled = false;
                try { await ReloadAsync(); }
                finally { _btnQuery.Enabled = true; }
            };
            _btnExport.Text = "导出 Excel";
            _btnExport.Click += (s, e) => ExportExcel();
            bar.Controls.Add(_input);
            bar.Controls.Add(_btnQuery);
            bar.Controls.Add(_btnExport);
            root.Controls.Add(bar, 0, 0);

            // 信息区
            var info = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(12) };
            _lblTitle.Font = UIStyle.H1;
            _lblTitle.ForeColor = UIStyle.TextDark;
            _lblSummary.Font = UIStyle.Body;
            info.Controls.Add(_lblTitle, 0, 0);
            info.Controls.Add(_lblSummary, 0, 1);
            root.Controls.Add(info, 0, 1);

            // Tabs
            _tabs.Dock = DockStyle.Fill;
            root.Controls.Add(_tabs, 0, 2);

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
            _plot7d = new PlotView { Dock = DockStyle.Fill };
            _plotSizeTop = new PlotView { Dock = DockStyle.Fill };
            _plotColorTop = new PlotView { Dock = DockStyle.Fill };
            panel.Controls.Add(_plot7d, 0, 0);
            panel.Controls.Add(_plotSizeTop, 0, 1);
            panel.Controls.Add(_plotColorTop, 0, 2);
            pageOverview.Controls.Add(panel);
            _tabs.TabPages.Add(pageOverview);

            // 明细页
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
            _grid.DataSource = _binding;

            panelDetail.Controls.Add(bar, 0, 0);
            panelDetail.Controls.Add(_grid, 0, 1);
            pageDetail.Controls.Add(panelDetail);
            _tabs.TabPages.Add(pageDetail);

            // 库存页（新增）
            _tabs.TabPages.Add(new InventoryTabPage(_cfg));
        }

        // ========== TrayApp 需要的几个公共方法（补齐） ==========

        /// <summary>把焦点放到输入框</summary>
        public void FocusInput()
        {
            try
            {
                if (WindowState == FormWindowState.Minimized) WindowState = FormWindowState.Normal;
                _input.Focus();
                _input.SelectAll();
            }
            catch { }
        }

        /// <summary>在鼠标附近显示，不抢焦点（尽量接近原语义）</summary>
        public void ShowNoActivateAtCursor()
        {
            try
            {
                StartPosition = FormStartPosition.Manual;
                var pt = Cursor.Position;
                Location = new Point(Math.Max(0, pt.X - Width / 2), Math.Max(0, pt.Y - Height / 2));
                Show();          // WinForms 无直接 "不激活" 显示；若有自定义 Win32，可再优化
            }
            catch { Show(); }
        }

        /// <summary>居中显示并把焦点置于输入框</summary>
        public void ShowAndFocusCentered()
        {
            try
            {
                StartPosition = FormStartPosition.CenterScreen;
                Show();
                Activate();
                FocusInput();
            }
            catch { Show(); }
        }

        /// <summary>设置“加载中/状态提示”文字（TrayApp 在请求前后会调用）</summary>
        public void SetLoading(string message)
        {
            _lblTitle.Text = message ?? "";
            _lblSummary.Text = "";
            _tabs.Enabled = false;
            Refresh();
        }

        /// <summary>兼容另一种签名：SetLoading(bool busy, string? message = null)</summary>
        public void SetLoading(bool busy, string? message = null)
        {
            _lblTitle.Text = message ?? (busy ? "加载中…" : "");
            _lblSummary.Text = "";
            _tabs.Enabled = !busy;
            Refresh();
        }

        /// <summary>把文本写入输入框（TrayApp 捕获选区后调用）</summary>
        public void ApplyRawText(string text)
        {
            _input.Text = text ?? string.Empty;
        }

        // ==========================================================

        // 外部调用：载入文本并渲染
        public async System.Threading.Tasks.Task LoadTextAsync(string raw)
        {
            _input.Text = raw;
            await ReloadAsync();
        }

        private async System.Threading.Tasks.Task ReloadAsync()
        {
            var text = _input.Text;
            _lblTitle.Text = "…加载中";
            _lblSummary.Text = "";
            await System.Threading.Tasks.Task.Yield();

            var parsed = StyleWatcherWin.Parser.Parse(text); // 显式命名空间，避免解析不到
            Bind(parsed);
            _tabs.Enabled = true;
        }

        private void Bind(ParsedPayload p)
        {
            _lblTitle.Text = p.Title;
            _lblSummary.Text = $"{p.Yesterday}；近7天合计：{p.Sum7d}";

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

            if (string.IsNullOrWhiteSpace(q))
            {
                _binding.ResetBindings(false);
                return;
            }

            // 反射读取匿名类型属性
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
            modelTrend.Axes.Add(new CategoryAxis { Position = AxisPosition.Bottom, ItemsSource = trend, LabelField = "Day" });
            modelTrend.Axes.Add(new LinearAxis { Position = AxisPosition.Left, MinimumPadding = 0, AbsoluteMinimum = 0 });
            var s1 = new LineSeries { ItemsSource = trend, DataFieldX = "Day", DataFieldY = "Qty" };
            modelTrend.Series.Add(s1);
            _plot7d.Model = modelTrend;

            // 尺码 TOP
            var sizeTop = p.Records.GroupBy(r => r.Size).Select(g => new { Key = g.Key, Qty = g.Sum(x => x.Qty) })
                .OrderByDescending(x => x.Qty).Take(10).ToList();
            var modelSize = new PlotModel { Title = "尺码 Top 10" };
            modelSize.Axes.Add(new CategoryAxis { Position = AxisPosition.Left, ItemsSource = sizeTop, LabelField = "Key" });
            modelSize.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, MinimumPadding = 0, AbsoluteMinimum = 0 });
            var s2 = new BarSeries { ItemsSource = sizeTop.Select(x => new BarItem { Value = x.Qty }) };
            modelSize.Series.Add(s2);
            _plotSizeTop.Model = modelSize;

            // 颜色 TOP
            var colorTop = p.Records.GroupBy(r => r.Color).Select(g => new { Key = g.Key, Qty = g.Sum(x => x.Qty) })
                .OrderByDescending(x => x.Qty).Take(10).ToList();
            var modelColor = new PlotModel { Title = "颜色 Top 10" };
            modelColor.Axes.Add(new CategoryAxis { Position = AxisPosition.Left, ItemsSource = colorTop, LabelField = "Key" });
            modelColor.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, MinimumPadding = 0, AbsoluteMinimum = 0 });
            var s3 = new BarSeries { ItemsSource = colorTop.Select(x => new BarItem { Value = x.Qty }) };
            modelColor.Series.Add(s3);
            _plotColorTop.Model = modelColor;
        }

        private void ExportExcel()
        {
            var saveDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "exports");
            Directory.CreateDirectory(saveDir);
            var path = Path.Combine(saveDir, $"StyleWatcher_{DateTime.Now:yyyyMMdd_HHmm}.xlsx");

            using var wb = new XLWorkbook();
            // 明细
            var ws1 = wb.AddWorksheet("明细");
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

            // 趋势7天
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
