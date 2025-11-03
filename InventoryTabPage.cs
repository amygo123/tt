
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using OxyPlot.WindowsForms;

namespace StyleWatcherWin
{
    /// <summary>
    /// A2：库存页 —— 自包含的数据模型与聚合，不依赖 Aggregations.*（避免全局耦合）
    /// </summary>
    public class InventoryTabPage : TabPage
    {
        // —— 内部数据模型 —— //
        private sealed class InvRow
        {
            public string Name { get; set; } = "";
            public string Color { get; set; } = "";
            public string Size  { get; set; } = "";
            public string Warehouse { get; set; } = "";
            public int Available { get; set; }
            public int OnHand    { get; set; }
        }

        private sealed class InvSnapshot
        {
            public List<InvRow> Rows { get; } = new();
            public int TotalAvailable => Rows.Sum(r => r.Available);
            public int TotalOnHand    => Rows.Sum(r => r.OnHand);

            public IEnumerable<string> ColorsNonZero() =>
                Rows.GroupBy(r => r.Color)
                    .Select(g => new { c = g.Key, v = g.Sum(x => x.Available) })
                    .Where(x => !string.IsNullOrWhiteSpace(x.c) && x.v != 0)
                    .OrderByDescending(x => x.v).Select(x => x.c);

            public IEnumerable<string> SizesNonZero() =>
                Rows.GroupBy(r => r.Size)
                    .Select(g => new { s = g.Key, v = g.Sum(x => x.Available) })
                    .Where(x => !string.IsNullOrWhiteSpace(x.s) && x.v != 0)
                    .OrderByDescending(x => x.v).Select(x => x.s);

            public Dictionary<string,int> ByColor() =>
                Rows.GroupBy(r => r.Color).ToDictionary(g => g.Key, g => g.Sum(x => x.Available));

            public Dictionary<string,int> BySize() =>
                Rows.GroupBy(r => r.Size).ToDictionary(g => g.Key, g => g.Sum(x => x.Available));

            public Dictionary<string,int> ByWarehouse() =>
                Rows.GroupBy(r => r.Warehouse).ToDictionary(g => g.Key, g => g.Sum(x => x.Available));

            public IEnumerable<string> WarehousesSorted() =>
                ByWarehouse().OrderByDescending(kv => kv.Value).Select(kv => kv.Key);

            public InvSnapshot Filter(Func<InvRow,bool> pred)
            {
                var s = new InvSnapshot();
                foreach (var r in Rows.Where(pred)) s.Rows.Add(r);
                return s;
            }
        }

        private static readonly HttpClient _http = new HttpClient();

        private readonly AppConfig _cfg;
        private InvSnapshot _all = new InvSnapshot();
        private string _styleName = "";

        // —— 顶部工具区 —— //
        private readonly TextBox _search = new() { PlaceholderText = "搜索（颜色/尺码/仓库）" };
        private readonly Label _lblAvail = new() { AutoSize = true, Font = new Font("Microsoft YaHei UI", 10, FontStyle.Bold) };
        private readonly Label _lblOnHand = new() { AutoSize = true, Font = new Font("Microsoft YaHei UI", 10, FontStyle.Bold), Margin = new Padding(16,0,0,0) };
        private readonly System.Windows.Forms.Timer _debounce = new() { Interval = 220 };

        // —— 主区：四图 —— //
        private readonly PlotView _pvHeat = new() { Dock = DockStyle.Fill };
        private readonly PlotView _pvColor = new() { Dock = DockStyle.Fill };
        private readonly PlotView _pvSize = new() { Dock = DockStyle.Fill };
        private readonly PlotView _pvPie = new() { Dock = DockStyle.Fill };

        // —— 子 Tab（按仓库） —— //
        private readonly TabControl _subTabs = new() { Dock = DockStyle.Fill };

        public InventoryTabPage(AppConfig cfg)
        {
            _cfg = cfg;
            Text = "库存";
            BackColor = Color.White;

            // 布局：顶部工具行 + 主四图 + 子 tabs
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(12) };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 60));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
            Controls.Add(root);

            var tools = BuildTopTools();
            root.Controls.Add(tools, 0, 0);

            var four = BuildFourPlots();
            root.Controls.Add(four, 0, 1);

            root.Controls.Add(_subTabs, 0, 2);

            _debounce.Tick += (s,e)=> { _debounce.Stop(); ApplySearchAndRender(); };
        }

        private Control BuildTopTools()
        {
            var p = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            _search.Width = 300;
            _search.TextChanged += (s,e)=> { _debounce.Stop(); _debounce.Start(); };
            p.Controls.Add(_search);
            p.Controls.Add(_lblAvail);
            p.Controls.Add(_lblOnHand);
            return p;
        }

        private Control BuildFourPlots()
        {
            var grid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2 };
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 60));
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            grid.Controls.Add(_pvHeat, 0, 0);
            grid.Controls.Add(_pvPie,  1, 0);
            grid.Controls.Add(_pvColor,0, 1);
            grid.Controls.Add(_pvSize, 1, 1);
            return grid;
        }

        // —— 外部对接：载入库存 —— //
        public async Task LoadInventoryAsync(string styleName)
        {
            _styleName = styleName ?? "";
            var (ok, rows) = await FetchAndParseAsync(_styleName);
            if (!ok) { rows = new List<InvRow>(); }

            _all = new InvSnapshot();
            foreach (var r in rows) _all.Rows.Add(r);

            RenderAll();
        }

        private string GetInventoryBaseUrl()
        {
            // 通过反射读取 _cfg.inventory.url_base，如不存在则用默认值
            try
            {
                var inv = _cfg.GetType().GetProperty("inventory")?.GetValue(_cfg);
                if (inv != null)
                {
                    var p = inv.GetType().GetProperty("url_base")?.GetValue(inv)?.ToString();
                    if (!string.IsNullOrWhiteSpace(p)) return p!;
                }
            }
            catch {}
            return "http://192.168.40.97:8000/inventory?style_name=";
        }

        private async Task<(bool ok, List<InvRow> rows)> FetchAndParseAsync(string styleName)
        {
            try
            {
                var urlBase = GetInventoryBaseUrl();
                var url = urlBase + Uri.EscapeDataString(styleName ?? "");
                var json = await _http.GetStringAsync(url);

                // 期望是 JSON 数组，每个元素是一行文本
                var options = new JsonSerializerOptions { ReadCommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true };
                var lines = JsonSerializer.Deserialize<List<string>>(json, options) ?? new List<string>();

                var rows = new List<InvRow>();
                foreach (var raw in lines)
                {
                    if (string.IsNullOrWhiteSpace(raw)) continue;
                    // 兼容全角逗号/空格
                    var line = raw.Replace("，", ",").Trim();
                    var parts = line.Split(',', StringSplitOptions.TrimEntries);
                    if (parts.Length < 6) continue;

                    var name = parts[0];
                    if (string.IsNullOrWhiteSpace(name)) continue; // 品名为空，清洗掉

                    var color = parts[1];
                    var size  = parts[2];
                    var wh    = parts[3];
                    int.TryParse(parts[4], out var avail);
                    int.TryParse(parts[5], out var onhand);

                    rows.Add(new InvRow { Name = name, Color = color, Size = size, Warehouse = wh, Available = avail, OnHand = onhand });
                }

                return (true, rows);
            }
            catch
            {
                return (false, new List<InvRow>());
            }
        }

        private void ApplySearchAndRender()
        {
            var q = (_search.Text ?? "").Trim();
            if (string.IsNullOrEmpty(q))
            {
                RenderAll();
                return;
            }
            var filtered = _all.Filter(r =>
                (r.Color?.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0) ||
                (r.Size?.IndexOf(q,  StringComparison.OrdinalIgnoreCase) >= 0) ||
                (r.Warehouse?.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0));
            Render(filtered);
        }

        private void RenderAll() => Render(_all);

        private void Render(InvSnapshot snap)
        {
            // 顶部合计
            _lblAvail.Text = $"可用合计：{snap.TotalAvailable}";
            _lblOnHand.Text = $"现有合计：{snap.TotalOnHand}";

            // 汇总热力图 & 柱状图 & 饼图
            RenderHeatmap(snap, _pvHeat, "颜色×尺码 可用数热力图");
            RenderBarsByColor(snap, _pvColor, "颜色可用（降序）");
            RenderBarsBySize (snap, _pvSize,  "尺码可用（降序）");
            RenderWarehousePie(snap, _pvPie,  "分仓占比");

            // 子 Tab：按仓库降序创建
            BuildWarehouseTabs(snap);
        }

        private void BuildWarehouseTabs(InvSnapshot snap)
        {
            _subTabs.SuspendLayout();
            _subTabs.TabPages.Clear();

            var first = new TabPage("汇总") { BackColor = Color.White };
            first.Controls.Add(BuildWarehousePanel(snap));
            _subTabs.TabPages.Add(first);

            foreach (var wh in snap.WarehousesSorted())
            {
                var sub = new TabPage(wh) { BackColor = Color.White };
                var subSnap = snap.Filter(r => r.Warehouse == wh);
                sub.Controls.Add(BuildWarehousePanel(subSnap));
                _subTabs.TabPages.Add(sub);
            }

            _subTabs.ResumeLayout();
        }

        private Control BuildWarehousePanel(InvSnapshot snap)
        {
            var p = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2, Padding = new Padding(6) };
            p.RowStyles.Add(new RowStyle(SizeType.Percent, 60));
            p.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
            p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            var pvHeat = new PlotView { Dock = DockStyle.Fill };
            var pvColor = new PlotView { Dock = DockStyle.Fill };
            var pvSize = new PlotView { Dock = DockStyle.Fill };
            var grid = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells };

            RenderHeatmap(snap, pvHeat, "颜色×尺码（仓）");
            RenderBarsByColor(snap, pvColor, "颜色可用（降序）");
            RenderBarsBySize (snap, pvSize,  "尺码可用（降序）");

            // 明细：按 仓库/品名/颜色/尺码 顺序展示
            grid.DataSource = snap.Rows
                .OrderBy(r => r.Warehouse).ThenBy(r => r.Name).ThenBy(r => r.Color).ThenBy(r => r.Size)
                .Select(r => new { 仓库 = r.Warehouse, 品名 = r.Name, 颜色 = r.Color, 尺码 = r.Size, 可用 = r.Available, 现有 = r.OnHand })
                .ToList();

            p.Controls.Add(pvHeat, 0, 0);
            p.Controls.Add(pvColor,1, 0);
            p.Controls.Add(pvSize, 0, 1);
            p.Controls.Add(grid,   1, 1);

            return p;
        }

        // —— 可视化 —— //

        private static OxyColor HeatColor(double v, double vmin, double vmax)
        {
            // 简洁的白→绿渐变；v<=0 近白
            if (vmax <= vmin) return OxyColors.LightGray;
            var t = Math.Max(0, Math.Min(1, (v - vmin) / (vmax - vmin)));
            byte r = (byte)(255 * (1 - t));
            byte g = (byte)(255 * (1 - 0.37 * t)); // 255→160
            byte b = (byte)(255 * (1 - t));
            return OxyColor.FromRgb(r, g, b);
        }

        private static void RenderHeatmap(InvSnapshot snap, PlotView pv, string title)
        {
            // 仅使用非零类目构造坐标轴（按需求）
            var colors = snap.ColorsNonZero().ToList();
            var sizes  = snap.SizesNonZero().ToList();

            var model = new PlotModel { Title = title, PlotMargins = new OxyThickness(80, 6, 8, 40) };
            var x = new CategoryAxis { Position = AxisPosition.Bottom };
            var y = new CategoryAxis { Position = AxisPosition.Left };

            foreach (var s in sizes)  y.Labels.Add(s);
            foreach (var c in colors) x.Labels.Add(c);

            model.Axes.Add(x);
            model.Axes.Add(y);

            var rbs = new RectangleBarSeries { StrokeThickness = 0.5, StrokeColor = OxyColors.Gray, IsHitTestVisible = false };

            var map = snap.Rows.GroupBy(r => (r.Color, r.Size))
                .ToDictionary(g => g.Key, g => g.Sum(z => z.Available));

            double vmin = 0;
            double vmax = map.Count == 0 ? 1 : Math.Max(1, map.Max(kv => kv.Value));

            for (int ix = 0; ix < colors.Count; ix++)
            {
                for (int iy = 0; iy < sizes.Count; iy++)
                {
                    var key = (colors[ix], sizes[iy]);
                    map.TryGetValue(key, out var v);
                    var item = new RectangleBarItem(ix - 0.5, iy - 0.5, ix + 0.5, iy + 0.5)
                    {
                        Color = HeatColor(v, vmin, vmax)
                    };
                    rbs.Items.Add(item);
                }
            }

            var scatter = new ScatterSeries { MarkerType = MarkerType.None };
            for (int ix = 0; ix < colors.Count; ix++)
            {
                for (int iy = 0; iy < sizes.Count; iy++)
                {
                    var key = (colors[ix], sizes[iy]);
                    map.TryGetValue(key, out var v);
                    var pt = new ScatterPoint(ix, iy) { Tag = $"颜色: {key.Item1}\n尺码: {key.Item2}\n可用: {v}" };
                    scatter.Points.Add(pt);
                }
            }
            scatter.TrackerFormatString = "{Tag}";

            model.Series.Add(rbs);
            model.Series.Add(scatter);
            pv.Model = model;
        }

        private static void RenderBarsByColor(InvSnapshot snap, PlotView pv, string title)
        {
            var agg = snap.Rows.GroupBy(r => r.Color)
                .Select(g => new { Key = g.Key, Qty = g.Sum(x => x.Available) })
                .Where(a => !string.IsNullOrWhiteSpace(a.Key) && a.Qty != 0)
                .OrderByDescending(a => a.Qty).ToList();

            var model = new PlotModel { Title = title, PlotMargins = new OxyThickness(80,6,8,8) };
            var cat = new CategoryAxis { Position = AxisPosition.Left, StartPosition = 1, EndPosition = 0 };
            foreach (var a in agg) cat.Labels.Add(a.Key);
            model.Axes.Add(cat);
            model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, MinimumPadding = 0, AbsoluteMinimum = 0 });

            var bars = new BarSeries();
            foreach (var a in agg) bars.Items.Add(new BarItem { Value = a.Qty });
            model.Series.Add(bars);
            pv.Model = model;
        }

        private static void RenderBarsBySize(InvSnapshot snap, PlotView pv, string title)
        {
            var agg = snap.Rows.GroupBy(r => r.Size)
                .Select(g => new { Key = g.Key, Qty = g.Sum(x => x.Available) })
                .Where(a => !string.IsNullOrWhiteSpace(a.Key) && a.Qty != 0)
                .OrderByDescending(a => a.Qty).ToList();

            var model = new PlotModel { Title = title, PlotMargins = new OxyThickness(80,6,8,8) };
            var cat = new CategoryAxis { Position = AxisPosition.Left, StartPosition = 1, EndPosition = 0 };
            foreach (var a in agg) cat.Labels.Add(a.Key);
            model.Axes.Add(cat);
            model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, MinimumPadding = 0, AbsoluteMinimum = 0 });

            var bars = new BarSeries();
            foreach (var a in agg) bars.Items.Add(new BarItem { Value = a.Qty });
            model.Series.Add(bars);
            pv.Model = model;
        }

        private static void RenderWarehousePie(InvSnapshot snap, PlotView pv, string title)
        {
            var agg = snap.Rows.GroupBy(r => r.Warehouse).Select(g => new { Key = g.Key, Qty = g.Sum(x => x.Available) })
                .Where(x => !string.IsNullOrWhiteSpace(x.Key) && x.Qty > 0)
                .OrderByDescending(x => x.Qty).ToList();

            var total = agg.Sum(x => x.Qty);
            var model = new PlotModel { Title = title };
            var pie = new PieSeries { AngleSpan = 360, StartAngle = 0, StrokeThickness = 0.5, InsideLabelPosition = 0.6 };

            if (total <= 0)
            {
                pie.Slices.Add(new PieSlice("无数据", 1));
            }
            else
            {
                // 小于 5% 合并为“其他”
                double other = 0;
                foreach (var a in agg)
                {
                    var ratio = (double)a.Qty / total;
                    if (ratio < 0.05) other += a.Qty;
                    else pie.Slices.Add(new PieSlice(a.Key, a.Qty));
                }
                if (other > 0) pie.Slices.Add(new PieSlice("其他", other));
            }
            model.Series.Add(pie);
            pv.Model = model;
        }
    }
}
