using System;
using System.Collections.Generic;
using System.Data;
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
    /// 库存页：动态“汇总 + 分仓”子Tab；每个子Tab含 KPI / 热力图 / 条形图 + 表格
    /// 对外暴露：ShowWarehouse(name)、GetSummary()、SummaryUpdated 事件
    /// </summary>
    public class InventoryTabPage : TabPage
    {
        private readonly AppConfig _cfg;
        private readonly TabControl _tabs = new TabControl();

        // 最近一次汇总快照（概览页读取）
        private Aggregations.InventorySnapshot? _snapshot;

        public event Action? SummaryUpdated;

        public InventoryTabPage(AppConfig cfg)
        {
            _cfg = cfg;
            Text = "库存";
            BackColor = Color.White;

            BuildUI();
            _ = RefreshAsync();
        }

        private void BuildUI()
        {
            _tabs.Dock = DockStyle.Fill;
            Controls.Add(_tabs);
        }

        public Aggregations.InventorySnapshot? GetSummary() => _snapshot;

        public void ShowWarehouse(string name)
        {
            foreach (TabPage p in _tabs.TabPages)
            {
                if (string.Equals(p.Text, name, StringComparison.OrdinalIgnoreCase))
                {
                    _tabs.SelectedTab = p;
                    return;
                }
            }
        }

        private async Task RefreshAsync()
        {
            try
            {
                var style = _cfg.inventory?.default_style ?? string.Empty;
                if (string.IsNullOrWhiteSpace(style)) return;

                var baseUrl = _cfg.inventory?.url_base ?? "";
                var url = baseUrl + Uri.EscapeDataString(style);

                using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(Math.Max(3, _cfg.timeout_seconds)) };
                using var req = new HttpRequestMessage(HttpMethod.Get, url);
                foreach (var kv in _cfg.headers) req.Headers.TryAddWithoutValidation(kv.Key, kv.Value);

                var resp = await http.SendAsync(req);
                resp.EnsureSuccessStatusCode();
                var raw = await resp.Content.ReadAsStringAsync();

                var rows = ParseRows(CleanLines(raw));

                // 生成快照，通知外部
                _snapshot = Aggregations.BuildSnapshot(rows.Select(r => new Aggregations.InventoryRow
                {
                    品名 = r.品名, 颜色 = r.颜色, 尺码 = r.尺码, 仓库 = r.仓库, 可用 = r.可用, 现有 = r.现有
                }));
                SummaryUpdated?.Invoke();

                // 构建动态子 Tab
                BuildDynamicTabs(rows);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"库存请求失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BuildDynamicTabs(List<Row> rows)
        {
            _tabs.TabPages.Clear();

            // 汇总
            var summary = rows
                .GroupBy(r => new { r.颜色, r.尺码 })
                .Select(g => new Row { 品名 = "汇总", 颜色 = g.Key.颜色, 尺码 = g.Key.尺码, 仓库 = "ALL", 可用 = g.Sum(x => x.可用), 现有 = g.Sum(x => x.现有) })
                .OrderByDescending(r => r.可用).ToList();
            _tabs.TabPages.Add(CreateWarehousePage("汇总", summary));

            // 分仓（按可用合计降序）
            var byStore = rows.GroupBy(r => r.仓库)
                              .Select(g => new { 仓库 = g.Key, 行 = g.ToList(), 总可用 = g.Sum(x => x.可用) })
                              .OrderByDescending(x => x.总可用)
                              .ToList();
            foreach (var st in byStore)
            {
                var sorted = st.行.OrderByDescending(x => x.可用).ToList();
                _tabs.TabPages.Add(CreateWarehousePage(st.仓库, sorted));
            }
        }

        private TabPage CreateWarehousePage(string title, List<Row> data)
        {
            var page = new TabPage(title) { BackColor = Color.White };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(8) };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 64)); // KPI
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 55));  // 图
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 45));  // 表
            page.Controls.Add(layout);

            // KPI
            var kpi = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = true };
            var totalAvail = data.Sum(x => x.可用);
            var totalOnhand = data.Sum(x => x.现有);
            kpi.Controls.Add(MakeKpi("可用(合计)", totalAvail.ToString()));
            kpi.Controls.Add(MakeKpi("现有(合计)", totalOnhand.ToString()));
            layout.Controls.Add(kpi, 0, 0);

            // 图：左热力图 + 右条形图
            var charts = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            charts.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            charts.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            var heat = BuildHeatmap(data);
            var bar = BuildBar(data);
            charts.Controls.Add(heat, 0, 0);
            charts.Controls.Add(bar, 1, 0);
            layout.Controls.Add(charts, 0, 1);

            // 表
            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                DataSource = data
            };
            layout.Controls.Add(grid, 0, 2);

            return page;
        }

        private Control MakeKpi(string title, string value)
        {
            var p = new Panel { Width = 200, Height = 48, BorderStyle = BorderStyle.FixedSingle, Padding = new Padding(8) };
            var t = new Label { Text = title, Dock = DockStyle.Top, Height = 18 };
            var v = new Label { Text = value, Dock = DockStyle.Fill, Font = new Font("Microsoft YaHei UI", 10, FontStyle.Bold) };
            p.Controls.Add(v); p.Controls.Add(t);
            return p;
        }

        private PlotView BuildBar(List<Row> data)
        {
            var agg = data.GroupBy(x => x.颜色).Select(g => new { Key = g.Key, Qty = g.Sum(x => x.可用) }).OrderByDescending(x => x.Qty).ToList();
            var model = new PlotModel { Title = "按颜色可用（降序）" };
            model.Axes.Add(new CategoryAxis { Position = AxisPosition.Left, ItemsSource = agg, LabelField = "Key" });
            model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, MinimumPadding = 0, AbsoluteMinimum = 0 });
            model.Series.Add(new BarSeries { ItemsSource = agg.Select(x => new BarItem { Value = x.Qty }) });
            return new PlotView { Model = model, Dock = DockStyle.Fill };
        }

        private PlotView BuildHeatmap(List<Row> data)
        {
            // 用 RectangleBarSeries 实现类别热力图（X=颜色, Y=尺码）
            var colors = data.Select(x => x.颜色 ?? "").Distinct().OrderBy(x => x).ToList();
            var sizes = data.Select(x => x.尺码 ?? "").Distinct().OrderBy(x => x).ToList();
            var dict = data.GroupBy(x => (x.颜色 ?? "", x.尺码 ?? ""))
                           .ToDictionary(g => g.Key, g => g.Sum(x => x.可用));
            int vmax = Math.Max(1, dict.Values.DefaultIfEmpty(0).Max());

            var model = new PlotModel { Title = "库存热力图（颜色×尺码，可用）" };
            var xAxis = new CategoryAxis { Position = AxisPosition.Bottom, ItemsSource = colors };
            var yAxis = new CategoryAxis { Position = AxisPosition.Left, ItemsSource = sizes };
            model.Axes.Add(xAxis); model.Axes.Add(yAxis);

            var series = new RectangleBarSeries { StrokeThickness = 0.5 };
            for (int xi = 0; xi < colors.Count; xi++)
            {
                for (int yi = 0; yi < sizes.Count; yi++)
                {
                    var key = (colors[xi], sizes[yi]);
                    dict.TryGetValue(key, out var v);
                    var item = new RectangleBarItem(xi - 0.5, yi - 0.5, xi + 0.5, yi + 0.5)
                    {
                        Color = OxyColor.FromArgb(Aggregations.ColorScale(v, vmax).A,
                                                  Aggregations.ColorScale(v, vmax).R,
                                                  Aggregations.ColorScale(v, vmax).G,
                                                  Aggregations.ColorScale(v, vmax).B),
                        Value = v
                    };
                    series.Items.Add(item);
                }
            }
            model.Series.Add(series);
            return new PlotView { Model = model, Dock = DockStyle.Fill };
        }

        private static IEnumerable<string> CleanLines(string raw)
        {
            var items = new List<string>();
            try
            {
                var arr = JsonSerializer.Deserialize<string[]>(raw);
                if (arr != null) items.AddRange(arr);
            }
            catch
            {
                raw = raw.Replace("\r\n", "\n").Replace("\r", "\n");
                foreach (var ln in raw.Split('\n'))
                {
                    var s = ln.Trim();
                    if (!string.IsNullOrEmpty(s)) items.Add(s);
                }
            }

            return items
                .Select(s => s.Trim().Trim('"')
                    .Replace(", ", "，").Replace(",", " ，"))
                .Select(s => s.Replace(" ，", "，"))
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct()
                .OrderBy(s => s);
        }

        private static List<Row> ParseRows(IEnumerable<string> lines)
        {
            var result = new List<Row>();
            foreach (var ln in lines)
            {
                var parts = ln.Split('，').Select(x => x.Trim()).Where(x => x.Length > 0).ToArray();
                if (parts.Length < 4) continue;
                int available = 0, onhand = 0;
                if (parts.Length >= 6)
                {
                    int.TryParse(parts[^2], out available);
                    int.TryParse(parts[^1], out onhand);
                }
                result.Add(new Row
                {
                    品名 = parts[0],
                    颜色 = parts.Length > 1 ? parts[1] : "",
                    尺码 = parts.Length > 2 ? parts[2] : "",
                    仓库 = parts.Length > 3 ? parts[3] : "",
                    可用 = available,
                    现有 = onhand
                });
            }
            return result;
        }

        private class Row
        {
            public string 品名 { get; set; } = "";
            public string 颜色 { get; set; } = "";
            public string 尺码 { get; set; } = "";
            public string 仓库 { get; set; } = "";
            public int 可用 { get; set; }
            public int 现有 { get; set; }
        }
    }
}
