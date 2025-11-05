using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using OxyPlot.WindowsForms;

namespace StyleWatcherWin
{
    public class InventoryTabPage : TabPage
    {
        private readonly AppConfig _cfg;

        // UI
        private readonly TabControl _tabs = new TabControl();
        private readonly PlotView _pvHeat = new PlotView();
        private readonly PlotView _pvSize = new PlotView();
        private readonly PlotView _pvColor = new PlotView();

        // 明细
        private readonly TextBox _txtSearch = new TextBox();
        private readonly DataGridView _grid = new DataGridView();
        private readonly BindingSource _bs = new BindingSource();

        // data
        private List<InvRow> _rows = new List<InvRow>();
        private Dictionary<string, int> _aggAvailByWarehouse = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        public event Action<int,int,Dictionary<string,int>>? SummaryUpdated;

        public InventoryTabPage(AppConfig cfg)
        {
            _cfg = cfg;
            Text = "库存";
            BackColor = Color.White;

            _tabs.Dock = DockStyle.Fill;
            Controls.Add(_tabs);

            BuildSummaryTab();
            BuildDetailTab();

            // 默认隐藏任何 Tab 级别之外的搜索框 —— 只保留子 Tab 里的（满足你的“去掉库存页面的搜索框”要求）
        }

        private void BuildSummaryTab()
        {
            var tab = new TabPage("汇总") { BackColor = Color.White };
            var grid = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 2, Padding = new Padding(12) };
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 60));
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            _pvHeat.Dock = DockStyle.Fill;
            _pvSize.Dock = DockStyle.Fill;
            _pvColor.Dock = DockStyle.Fill;

            grid.Controls.Add(_pvHeat, 0, 0); grid.SetColumnSpan(_pvHeat, 2); // 热力图占上方整行
            grid.Controls.Add(_pvSize, 0, 1); // 左：尺码可用
            grid.Controls.Add(_pvColor, 1, 1); // 右：颜色可用

            _tabs.TabPages.Add(tab);
            tab.Controls.Add(grid);
        }

        private void BuildDetailTab()
        {
            var tab = new TabPage("明细") { BackColor = Color.White };
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, Padding = new Padding(12) };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            // 仅在“明细”子Tab中保留搜索框
            var bar = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight };
            _txtSearch.Width = 360;
            _txtSearch.PlaceholderText = "搜索（颜色/尺码/仓库/数量）";
            _txtSearch.TextChanged += (s, e) => ApplyFilter(_txtSearch.Text);
            bar.Controls.Add(_txtSearch);
            root.Controls.Add(bar, 0, 0);

            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.AllowUserToDeleteRows = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            _grid.RowHeadersVisible = false;
            _grid.Dock = DockStyle.Fill;
            _grid.DataSource = _bs;
            root.Controls.Add(_grid, 0, 1);

            _tabs.TabPages.Add(tab);
            tab.Controls.Add(root);
        }

        private class InvRow
        {
            public string Name { get; set; } = "";
            public string Color { get; set; } = "";
            public string Size { get; set; } = "";
            public string Warehouse { get; set; } = "";
            public int Avail { get; set; }
            public int OnHand { get; set; }
        }

        public async Task LoadInventoryAsync(string styleName)
        {
            try
            {
                var raw = await ApiHelper.QueryInventoryAsync(_cfg, styleName);
                // 约定返回：JSON 字符串数组，每项一个 CSV：品名,颜色,尺码,仓库,可用,现有
                var arr = JsonSerializer.Deserialize<List<string>>(raw, new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? new List<string>();
                var list = new List<InvRow>();
                foreach (var line in arr)
                {
                    var parts = (line ?? "").Split(',', StringSplitOptions.None);
                    if (parts.Length < 6) continue;
                    int avail = 0, onhand = 0;
                    int.TryParse(parts[4], out avail);
                    int.TryParse(parts[5], out onhand);
                    list.Add(new InvRow
                    {
                        Name = parts[0]?.Trim() ?? "",
                        Color = parts[1]?.Trim() ?? "",
                        Size = parts[2]?.Trim() ?? "",
                        Warehouse = parts[3]?.Trim() ?? "",
                        Avail = Math.Max(0, avail),
                        OnHand = Math.Max(0, onhand)
                    });
                }
                _rows = list;

                // 明细绑定
                _bs.DataSource = new BindingList<object>(_rows.Select(r => (object)new
                {
                    品名 = r.Name, 颜色 = r.Color, 尺码 = r.Size, 仓库 = r.Warehouse, 可用 = r.Avail, 现有 = r.OnHand
                }).ToList());

                // 汇总并通知
                var totalAvail = _rows.Sum(r => r.Avail);
                var totalOnHand = _rows.Sum(r => r.OnHand);
                _aggAvailByWarehouse = _rows.GroupBy(r => r.Warehouse ?? "")
                                            .ToDictionary(g => g.Key, g => g.Sum(x => x.Avail), StringComparer.OrdinalIgnoreCase);
                SummaryUpdated?.Invoke(totalAvail, totalOnHand, _aggAvailByWarehouse);

                // 渲染
                RenderHeat();
                RenderBarsBySize();
                RenderBarsByColor();
            }
            catch (Exception ex)
            {
                // 简单兜底
                _bs.DataSource = new BindingList<object>(new List<object> { new { 错误 = ex.Message } });
            }
        }

        private void ApplyFilter(string q)
        {
            q = (q ?? "").Trim();
            if (string.IsNullOrEmpty(q))
            {
                _bs.DataSource = new BindingList<object>(_rows.Select(r => (object)new
                {
                    品名 = r.Name, 颜色 = r.Color, 尺码 = r.Size, 仓库 = r.Warehouse, 可用 = r.Avail, 现有 = r.OnHand
                }).ToList());
                return;
            }
            var filtered = _rows.Where(r =>
                   r.Name.Contains(q, StringComparison.OrdinalIgnoreCase)
                || r.Color.Contains(q, StringComparison.OrdinalIgnoreCase)
                || r.Size.Contains(q, StringComparison.OrdinalIgnoreCase)
                || r.Warehouse.Contains(q, StringComparison.OrdinalIgnoreCase)
                || r.Avail.ToString().Contains(q, StringComparison.OrdinalIgnoreCase)
                || r.OnHand.ToString().Contains(q, StringComparison.OrdinalIgnoreCase)
            ).Select(r => (object)new
            {
                品名 = r.Name, 颜色 = r.Color, 尺码 = r.Size, 仓库 = r.Warehouse, 可用 = r.Avail, 现有 = r.OnHand
            }).ToList();
            _bs.DataSource = new BindingList<object>(filtered);
        }

        // --------- 渲染：热力图 ----------
        private void RenderHeat()
        {
            var colors = _rows.Select(r => r.Color).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().OrderBy(s => s).ToList();
            var sizes = _rows.Select(r => r.Size).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct()
                             .OrderBy(s => SizeSortKey(s)).ToList();

            var matrix = new Dictionary<(int x, int y), int>();
            for (int xi = 0; xi < colors.Count; xi++)
                for (int yi = 0; yi < sizes.Count; yi++)
                    matrix[(xi, yi)] = 0;

            foreach (var g in _rows.GroupBy(r => (Color: r.Color, Size: r.Size)))
            {
                var xi = colors.IndexOf(g.Key.Color);
                var yi = sizes.IndexOf(g.Key.Size);
                if (xi >= 0 && yi >= 0)
                    matrix[(xi, yi)] = g.Sum(x => x.Avail);
            }

            var model = new PlotModel { Title = "颜色 × 尺码 可用数热力图" };
            var xAxis = new CategoryAxis { Position = AxisPosition.Bottom };
            xAxis.Labels.AddRange(colors);
            var yAxis = new CategoryAxis { Position = AxisPosition.Left, StartPosition = 1, EndPosition = 0 };
            yAxis.Labels.AddRange(sizes);
            model.Axes.Add(xAxis);
            model.Axes.Add(yAxis);

            // 用矩形柱模拟热力格
            var rects = new RectangleBarSeries { StrokeThickness = 0.5 };
            foreach (var kv in matrix)
            {
                var (x, y) = kv.Key;
                var val = kv.Value;
                var item = new RectangleBarItem(x - 0.5, y - 0.5, x + 0.5, y + 0.5);
                // 简易颜色映射：按可用值深浅
                var norm = val <= 0 ? 0 : 1;
                item.Color = OxyColor.FromAColor(200, OxyColors.Orange);
                rects.Items.Add(item);
            }
            model.Series.Add(rects);

            // 覆盖一层“不可见”的散点用于点击/Tracker（只有一层，确保不会出现“两组数据”的重复）
            var pts = new ScatterSeries { MarkerType = MarkerType.Circle, MarkerSize = 0, TrackerFormatString = "{Tag}" };
            foreach (var kv in matrix)
            {
                var (x, y) = kv.Key;
                var val = kv.Value;
                var tag = $"颜色: {colors[x]}\n尺码: {sizes[y]}\n可用: {val}";
                pts.Points.Add(new ScatterPoint(x, y) { Tag = tag });
            }
            model.Series.Add(pts);

            _pvHeat.Model = model;
        }

        // 左侧：尺码可用（默认只显示前 10）
        private void RenderBarsBySize()
        {
            var list = _rows.GroupBy(r => r.Size).Select(g => new { Key = g.Key, Qty = g.Sum(x => x.Avail) })
                            .Where(a => !string.IsNullOrWhiteSpace(a.Key)).OrderByDescending(a => a.Qty).ToList();

            var model = new PlotModel { Title = "尺码可用（前 10 默认可视）" };
            var cat = new CategoryAxis { Position = AxisPosition.Left, StartPosition = 1, EndPosition = 0 };
            foreach (var a in list) cat.Labels.Add(a.Key);
            model.Axes.Add(cat);
            model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, MinimumPadding = 0, AbsoluteMinimum = 0 });

            var bs = new BarSeries();
            foreach (var a in list) bs.Items.Add(new BarItem { Value = a.Qty });
            model.Series.Add(bs);

            // 只显示前10（可滚轮/拖拽查看更多）
            if (cat.Labels.Count > 10)
            {
                cat.Minimum = -0.5;
                cat.Maximum = 9.5;
            }

            _pvSize.Model = model;
        }

        // 右侧：颜色可用（默认只显示前 10）
        private void RenderBarsByColor()
        {
            var list = _rows.GroupBy(r => r.Color).Select(g => new { Key = g.Key, Qty = g.Sum(x => x.Avail) })
                            .Where(a => !string.IsNullOrWhiteSpace(a.Key)).OrderByDescending(a => a.Qty).ToList();

            var model = new PlotModel { Title = "颜色可用（前 10 默认可视）" };
            var cat = new CategoryAxis { Position = AxisPosition.Left, StartPosition = 1, EndPosition = 0 };
            foreach (var a in list) cat.Labels.Add(a.Key);
            model.Axes.Add(cat);
            model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, MinimumPadding = 0, AbsoluteMinimum = 0 });

            var bs = new BarSeries();
            foreach (var a in list) bs.Items.Add(new BarItem { Value = a.Qty });
            model.Series.Add(bs);

            // 只显示前10
            if (cat.Labels.Count > 10)
            {
                cat.Minimum = -0.5;
                cat.Maximum = 9.5;
            }

            _pvColor.Model = model;
        }

        // 简单的尺码排序 key
        private static int SizeSortKey(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return int.MaxValue;
            // 常见顺序：XS < S < M < L < XL < 2XL < 3XL ... < 6XL < KXL/K2XL...
            var order = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["XS"]=1,["S"]=2,["M"]=3,["L"]=4,["XL"]=5,
                ["2XL"]=6,["3XL"]=7,["4XL"]=8,["5XL"]=9,["6XL"]=10,
                ["KXL"]=11,["K2XL"]=12,["K3XL"]=13,["K4XL"]=14
            };
            if (order.TryGetValue(s, out var v)) return v;
            return 1000 + s.GetHashCode();
        }
    }
}
