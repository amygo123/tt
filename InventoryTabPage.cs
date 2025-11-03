
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;
using OxyPlot;
using OxyPlot.Series;
using OxyPlot.Axes;
using OxyPlot.WindowsForms;

namespace StyleWatcherWin
{
    /// <summary>
    /// A2：库存页实现
    /// - 顶部：搜索框（品名/颜色/尺码/仓库），可防抖
    /// - 汇总 Tab + 动态分仓 Tab（按可用合计降序）
    /// - 每个子页：左右两列
    ///   左：颜色×尺码 热力图（自绘格子），可点击，支持 Tooltip
    ///   右上：按颜色可用量条形图（降序）
    ///   右下：按尺码可用量条形图（降序）
    /// </summary>
    public class InventoryTabPage : TabPage
    {
        private readonly AppConfig _cfg;

        private readonly TextBox _boxSearch = new();
        private readonly System.Windows.Forms.Timer _debounce = new() { Interval = 250 };
        private readonly TabControl _tabs = new();
        private readonly PlotView _pie = new();

        private List<Aggregations.InventoryRow> _allRows = new();
        private Aggregations.InventorySnapshot? _summary;

        public InventoryTabPage(AppConfig cfg)
        {
            _cfg = cfg;
            Text = "库存";
            BackColor = Color.White;
            Padding = new Padding(8);

            var root = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1 };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Controls.Add(root);

            // 顶部工具行：搜索 + 分仓饼图
            var tools = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            tools.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60));
            tools.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            root.Controls.Add(tools, 0, 0);

            _boxSearch.PlaceholderText = "搜索（品名/颜色/尺码/仓库）";
            _boxSearch.Dock = DockStyle.Fill;
            _boxSearch.TextChanged += (s, e) => { _debounce.Stop(); _debounce.Start(); };
            tools.Controls.Add(_boxSearch, 0, 0);

            _debounce.Tick += (s, e) => { _debounce.Stop(); RefreshTabs(); };

            _pie.Dock = DockStyle.Fill;
            tools.Controls.Add(_pie, 1, 0);

            // 主体：子 Tab
            _tabs.Dock = DockStyle.Fill;
            root.Controls.Add(_tabs, 0, 1);
        }

        public async System.Threading.Tasks.Task LoadInventoryAsync(string styleName)
        {
            // 读取后端文本（JSON 数组或纯文本）并解析
            var raw = await ApiHelper.QueryInventoryAsync(_cfg, styleName);
            _allRows = ParseInventory(raw);
            BuildSummaryAndPie();
            BuildWarehouseTabs();
        }

        private static List<Aggregations.InventoryRow> ParseInventory(string raw)
        {
            var list = new List<Aggregations.InventoryRow>();
            if (string.IsNullOrWhiteSpace(raw)) return list;

            try
            {
                // 尝试按 JSON 数组解析
                var arr = JsonSerializer.Deserialize<List<string>>(raw);
                if (arr != null)
                {
                    foreach (var line in arr)
                        TryParseLine(line, list);
                    return list;
                }
            }
            catch { /* 不是 JSON，走按行解析 */ }

            foreach (var line in raw.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries))
                TryParseLine(line, list);
            return list;
        }

        private static void TryParseLine(string line, List<Aggregations.InventoryRow> acc)
        {
            // 预期格式：品名，颜色，尺码，仓库，可用，现有 （中文逗号）
            if (string.IsNullOrWhiteSpace(line)) return;
            var parts = line.Trim().Trim('\"').Split('，');
            if (parts.Length < 6) return;

            var row = new Aggregations.InventoryRow
            {
                品名 = (parts[0] ?? "").Trim(),
                颜色 = (parts[1] ?? "").Trim(),
                尺码 = (parts[2] ?? "").Trim(),
                仓库 = (parts[3] ?? "").Trim(),
                可用 = SafeToInt(parts[4]),
                现有 = SafeToInt(parts[5])
            };
            // 清洗：品名异常/空直接丢弃；颜色/尺码空或合计为 0 的项可在图表维度上过滤
            if (!string.IsNullOrWhiteSpace(row.品名))
                acc.Add(row);
        }

        private static int SafeToInt(string s)
            => int.TryParse((s ?? "").Trim(), out var v) ? v : 0;

        private void BuildSummaryAndPie()
        {
            _summary = Aggregations.BuildSnapshot(_allRows);
            // 构造饼图（合并小扇区到“其他”）
            var total = Math.Max(1, _summary.分仓可用.Values.Sum());
            var data = _summary.分仓可用
                                .OrderByDescending(kv => kv.Value)
                                .Select(kv => new { Name = kv.Key, Val = kv.Value, Ratio = kv.Value / (double)total })
                                .ToList();
            var major = data.Where(d => d.Ratio >= 0.05).ToList();
            var otherVal = data.Where(d => d.Ratio < 0.05).Sum(d => d.Val);

            var model = new PlotModel { Title = "分仓可用占比" };
            var ps = new PieSeries { StrokeThickness = 0.5, InsideLabelPosition = 0.6, AngleSpan = 360, StartAngle = 0 };
            foreach (var m in major) ps.Slices.Add(new PieSlice(m.Name, m.Val));
            if (otherVal > 0) ps.Slices.Add(new PieSlice("其他", otherVal));
            model.Series.Add(ps);
            _pie.Model = model;
        }

        private void BuildWarehouseTabs()
        {
            _tabs.SuspendLayout();
            _tabs.TabPages.Clear();

            var subsets = Aggregations.GroupByWarehouse(_allRows); // (仓库, 可用合计, 行列表)
            // 汇总页
            _tabs.TabPages.Add(BuildSingleWarehousePage("汇总", _allRows));

            // 分仓页（按可用合计降序）
            foreach (var (仓库, _, 行) in subsets)
                _tabs.TabPages.Add(BuildSingleWarehousePage(仓库, 行));

            _tabs.ResumeLayout();
        }

        private TabPage BuildSingleWarehousePage(string name, List<Aggregations.InventoryRow> rows)
        {
            var page = new TabPage(name) { BackColor = Color.White };

            // 搜索过滤
            var q = (_boxSearch.Text ?? "").Trim();
            if (!string.IsNullOrEmpty(q))
            {
                rows = rows.Where(r =>
                            (r.品名 ?? "").Contains(q, StringComparison.OrdinalIgnoreCase) ||
                            (r.颜色 ?? "").Contains(q, StringComparison.OrdinalIgnoreCase) ||
                            (r.尺码 ?? "").Contains(q, StringComparison.OrdinalIgnoreCase) ||
                            (r.仓库 ?? "").Contains(q, StringComparison.OrdinalIgnoreCase)
                        ).ToList();
            }

            // 计算汇总（过滤掉颜色或尺码合计为 0 的类目）
            var sizeSum = rows.GroupBy(r => r.尺码 ?? "").ToDictionary(g => g.Key, g => g.Sum(x => x.可用));
            var colorSum = rows.GroupBy(r => r.颜色 ?? "").ToDictionary(g => g.Key, g => g.Sum(x => x.可用));
            var rowsFiltered = rows.Where(r => (colorSum.GetValueOrDefault(r.颜色 ?? "", 0) != 0) &&
                                               (sizeSum.GetValueOrDefault(r.尺码 ?? "", 0) != 0)).ToList();

            var snap = Aggregations.BuildSnapshot(rowsFiltered);

            // 左右两列布局
            var grid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2, Padding = new Padding(8) };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 55));
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 45));
            page.Controls.Add(grid);

            // 左：热力图（自绘 Label 网格）
            var heat = BuildHeatmapControl(snap);
            grid.Controls.Add(heat, 0, 0);
            grid.SetRowSpan(heat, 2);

            // 右上：颜色可用（降序）
            var plotColor = BuildBarPlot("颜色可用", snap.颜色List.Select(c => new { Key = c, Qty = rowsFiltered.Where(r => r.颜色 == c).Sum(r => r.可用) })
                                .Where(x => x.Qty != 0).OrderByDescending(x => x.Qty).ToList());
            grid.Controls.Add(plotColor, 1, 0);

            // 右下：尺码可用（降序）
            var plotSize = BuildBarPlot("尺码可用", snap.尺码List.Select(s => new { Key = s, Qty = rowsFiltered.Where(r => r.尺码 == s).Sum(r => r.可用) })
                                .Where(x => x.Qty != 0).OrderByDescending(x => x.Qty).ToList());
            grid.Controls.Add(plotSize, 1, 1);

            return page;
        }

        private Control BuildBarPlot(string title, List<dynamic> items)
        {
            var model = new PlotModel { Title = title, PlotMargins = new OxyThickness(80, 6, 6, 6) };
            var cat = new CategoryAxis { Position = AxisPosition.Left, StartPosition = 1, EndPosition = 0, GapWidth = 0.4 };
            foreach (var it in items) cat.Labels.Add((string)it.Key);
            model.Axes.Add(cat);
            model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, MinimumPadding = 0, AbsoluteMinimum = 0 });

            var bs = new BarSeries();
            foreach (var it in items) bs.Items.Add(new BarItem { Value = (double)it.Qty });
            model.Series.Add(bs);

            var pv = new PlotView { Dock = DockStyle.Fill, Model = model };
            return pv;
        }

        private Control BuildHeatmapControl(Aggregations.InventorySnapshot snap)
        {
            // 使用 TableLayoutPanel 渲染分类热力图：第一行/列是头
            var tl = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                ColumnCount = snap.尺码List.Count + 1,
                RowCount = snap.颜色List.Count + 1,
                Padding = new Padding(0),
                Margin = new Padding(0),
                AutoScroll = true
            };

            tl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80)); // 左上角 & 行头
            for (int i = 0; i < snap.尺码List.Count; i++)
                tl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 64));
            tl.RowStyles.Add(new RowStyle(SizeType.Absolute, 28)); // 列头
            for (int i = 0; i < snap.颜色List.Count; i++)
                tl.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));

            // 左上角空白
            tl.Controls.Add(new Label { Text = "", Dock = DockStyle.Fill }, 0, 0);

            // 列头（尺码）
            for (int x = 0; x < snap.尺码List.Count; x++)
            {
                var lbl = new Label
                {
                    Text = snap.尺码List[x],
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold)
                };
                tl.Controls.Add(lbl, x + 1, 0);
            }

            // 行头（颜色） + 单元格
            var max = snap.汇总矩阵.Count == 0 ? 1 : snap.汇总矩阵.Max(kv => kv.Value);
            for (int y = 0; y < snap.颜色List.Count; y++)
            {
                var colorName = snap.颜色List[y];
                var head = new Label
                {
                    Text = colorName,
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Padding = new Padding(8, 0, 0, 0),
                    Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold)
                };
                tl.Controls.Add(head, 0, y + 1);

                for (int x = 0; x < snap.尺码List.Count; x++)
                {
                    var size = snap.尺码List[x];
                    var val = snap.汇总矩阵.GetValueOrDefault((colorName, size), 0);
                    var bg = Aggregations.ColorScale(val, max);
                    var cell = new Label
                    {
                        Dock = DockStyle.Fill,
                        TextAlign = ContentAlignment.MiddleCenter,
                        BackColor = bg,
                        ForeColor = Color.Black,
                        BorderStyle = BorderStyle.FixedSingle,
                        Text = val.ToString()
                    };
                    var tip = new ToolTip();
                    tip.SetToolTip(cell, $"颜色：{colorName}，尺码：{size}，可用：{val}");
                    cell.Click += (s, e) =>
                    {
                        MessageBox.Show($"颜色：{colorName}\n尺码：{size}\n可用：{val}", "明细", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    };
                    tl.Controls.Add(cell, x + 1, y + 1);
                }
            }
            return tl;
        }

        private void RefreshTabs()
        {
            if (_allRows == null || _allRows.Count == 0) return;
            BuildWarehouseTabs();
        }
    }
}
