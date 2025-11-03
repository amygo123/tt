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
    public class InventoryTabPage : TabPage
    {
        private readonly AppConfig _cfg;
        private readonly TabControl _tabs = new();
        private Aggregations.InventorySnapshot? _snapshot;
        public event Action? SummaryUpdated;

        public InventoryTabPage(AppConfig cfg)
        {
            _cfg = cfg;
            Text = "库存";
            BackColor = Color.White;
            _tabs.Dock = DockStyle.Fill;
            Controls.Add(_tabs);
            _ = RefreshAsync();
        }

        public Aggregations.InventorySnapshot? GetSummary() => _snapshot;

        public void ShowWarehouse(string name)
        {
            foreach(TabPage p in _tabs.TabPages)
                if (string.Equals(p.Text,name,StringComparison.OrdinalIgnoreCase)){ _tabs.SelectedTab=p; return; }
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
                var rows = ParseRows(CleanLines(raw)).ToList();

                // —— 快照（用于概览 KPI 等），不做“0 类别剔除”
                _snapshot = Aggregations.BuildSnapshot(rows.Select(r => new Aggregations.InventoryRow
                {
                    品名 = r.品名, 颜色 = r.颜色, 尺码 = r.尺码, 仓库 = r.仓库, 可用 = r.可用, 现有 = r.现有
                }));
                SummaryUpdated?.Invoke();

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

            // —— 可视化用清洗：品名空/异常剔除；按照“颜色汇总=0或空”剔除，“尺码汇总=0或空”剔除（负值不剔）
            List<Row> CleanForVisuals(IEnumerable<Row> src)
            {
                var list = src.Where(r => !string.IsNullOrWhiteSpace(r.品名)
                                        && !string.IsNullOrWhiteSpace(r.颜色)
                                        && !string.IsNullOrWhiteSpace(r.尺码))
                              .ToList();

                var byColor = list.GroupBy(x=>x.颜色).ToDictionary(g=>g.Key, g=>g.Sum(z=>z.可用));
                var bySize  = list.GroupBy(x=>x.尺码).ToDictionary(g=>g.Key, g=>g.Sum(z=>z.可用));
                list = list.Where(x => byColor.GetValueOrDefault(x.颜色,0) != 0
                                     && bySize.GetValueOrDefault(x.尺码,0) != 0).ToList();
                return list;
            }

            // 汇总（ALL）
            var summary = rows
                .GroupBy(r => new { r.颜色, r.尺码 })
                .Select(g => new Row { 品名 = "汇总", 颜色 = g.Key.颜色, 尺码 = g.Key.尺码, 仓库 = "ALL", 可用 = g.Sum(x => x.可用), 现有 = g.Sum(x => x.现有) })
                .ToList();
            _tabs.TabPages.Add(CreateWarehousePage("汇总", summary, CleanForVisuals(summary)));

            // 分仓（按可用合计降序）
            var byStore = rows.GroupBy(r => r.仓库)
                              .Select(g => new { 仓库 = g.Key, 行 = g.ToList(), 总可用 = g.Sum(x => x.可用) })
                              .OrderByDescending(x => x.总可用)
                              .ToList();
            foreach (var st in byStore)
            {
                _tabs.TabPages.Add(CreateWarehousePage(st.仓库, st.行, CleanForVisuals(st.行)));
            }
        }

        private TabPage CreateWarehousePage(string title, List<Row> fullData, List<Row> visualData)
        {
            // 明细排序：仓库->品名->颜色->尺码；列顺序照此（表格展示用“fullData”，可视化用“visualData”）
            var ordered = fullData.OrderBy(x=>x.仓库).ThenBy(x=>x.品名).ThenBy(x=>x.颜色).ThenBy(x=>x.尺码).ToList();

            var page = new TabPage(title) { BackColor = Color.White };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(8) };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 64)); // KPI
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 55));  // 图
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 45));  // 表
            page.Controls.Add(layout);

            // KPI
            var kpi = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = true };
            var totalAvail = fullData.Sum(x => x.可用);
            var totalOnhand = fullData.Sum(x => x.现有);
            kpi.Controls.Add(MakeKpi("可用合计", totalAvail.ToString()));
            kpi.Controls.Add(MakeKpi("现有合计", totalOnhand.ToString()));
            layout.Controls.Add(kpi, 0, 0);

            // 图：左热力图 + 右条形图（坐标动态，仅用“visualData”的子集）
            var charts = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            charts.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            charts.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            var heat = BuildHeatmap(visualData);
            var bar = BuildBar(visualData);
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
                DataSource = ordered
            };
            layout.Controls.Add(grid, 0, 2);
            // 列顺序：仓库/品名/颜色/尺码/可用/现有
            if (grid.Columns.Contains("仓库")) grid.Columns["仓库"].DisplayIndex = 0;
            if (grid.Columns.Contains("品名")) grid.Columns["品名"].DisplayIndex = 1;
            if (grid.Columns.Contains("颜色")) grid.Columns["颜色"].DisplayIndex = 2;
            if (grid.Columns.Contains("尺码")) grid.Columns["尺码"].DisplayIndex = 3;
            if (grid.Columns.Contains("可用")) grid.Columns["可用"].DisplayIndex = 4;
            if (grid.Columns.Contains("现有")) grid.Columns["现有"].DisplayIndex = 5;

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

        // 颜色可用柱状图（标题精简；坐标轴基于“清洗后子集”）
        private PlotView BuildBar(List<Row> data)
        {
            var agg = data.GroupBy(x => x.颜色 ?? "")
                          .Select(g => new { Key = g.Key, Qty = g.Sum(z => z.可用) })
                          .Where(a => !string.IsNullOrWhiteSpace(a.Key) && a.Qty != 0) // 子集：去掉空/合计为0
                          .OrderByDescending(a => a.Qty)
                          .ToList();

            var model = new PlotModel { Title = "按颜色可用", PlotMargins = new OxyThickness(80, 6, 6, 6) };
            var cat = new CategoryAxis { Position = AxisPosition.Left, GapWidth = 0.4 };
            foreach (var a in agg) cat.Labels.Add(a.Key);
            model.Axes.Add(cat);
            model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, MinimumPadding = 0, AbsoluteMinimum = 0 });
            var series = new BarSeries();
            foreach (var a in agg) series.Items.Add(new BarItem { Value = a.Qty });
            model.Series.Add(series);
            return new PlotView { Model = model, Dock = DockStyle.Fill };
        }

        // 热力图（标题精简；坐标轴基于“清洗后子集”）
        private PlotView BuildHeatmap(List<Row> data)
        {
            var colors = data.GroupBy(x => x.颜色).Select(g => g.Key!).Where(k=>!string.IsNullOrWhiteSpace(k))
                             .Where(k => data.Where(r => r.颜色 == k).Sum(z => z.可用) != 0)
                             .OrderBy(k => k).ToList();
            var sizes  = data.GroupBy(x => x.尺码 ).Select(g => g.Key!).Where(k=>!string.IsNullOrWhiteSpace(k))
                             .Where(k => data.Where(r => r.尺码 == k).Sum(z => z.可用) != 0)
                             .OrderBy(k => k).ToList();

            var values = new double[Math.Max(1,colors.Count), Math.Max(1,sizes.Count)];
            foreach (var g in data.GroupBy(x => (x.颜色 ?? "", x.尺码 ?? "")))
            {
                var xi = colors.IndexOf(g.Key.Item1);
                var yi = sizes.IndexOf(g.Key.Item2);
                if (xi >= 0 && yi >= 0) values[xi, yi] = Math.Max(values[xi, yi], g.Sum(z => z.可用));
            }

            var vmax = Math.Max(1.0, values.Cast<double>().DefaultIfEmpty(0).Max());
            var model = new PlotModel { Title = "库存热力图" };

            var xAxis = new LinearAxis
            {
                Position = AxisPosition.Bottom,
                Minimum = -0.5,
                Maximum = colors.Count - 0.5,
                MajorStep = 1,
                MinorStep = 1,
                Angle = 45,
                LabelFormatter = v => {
                    int i = (int)Math.Round(v);
                    return (i >= 0 && i < colors.Count) ? colors[i] : "";
                }
            };
            var yAxis = new LinearAxis
            {
                Position = AxisPosition.Left,
                Minimum = -0.5,
                Maximum = sizes.Count - 0.5,
                MajorStep = 1,
                MinorStep = 1,
                StartPosition = 1, EndPosition = 0,
                LabelFormatter = v => {
                    int i = (int)Math.Round(v);
                    return (i >= 0 && i < sizes.Count) ? sizes[i] : "";
                }
            };
            var colorAxis = new LinearColorAxis
            {
                Position = AxisPosition.Right,
                Minimum = 0,
                Maximum = vmax,
                Palette = OxyPalettes.Jet(200),
                HighColor = OxyColors.Undefined,
                LowColor = OxyColors.Undefined
            };

            model.Axes.Add(xAxis);
            model.Axes.Add(yAxis);
            model.Axes.Add(colorAxis);

            var hm = new HeatMapSeries
            {
                X0 = 0, X1 = Math.Max(0, colors.Count - 1),
                Y0 = 0, Y1 = Math.Max(0, sizes.Count - 1),
                Interpolate = false,
                RenderMethod = HeatMapRenderMethod.Rectangles,
                Data = values,
                TrackerFormatString = "颜色: {2:0}\n尺码: {3:0}\n可用: {4:0}"
            };

            model.Series.Add(hm);
            model.PlotMargins = new OxyThickness(80, 10, 60, 50);

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

                // 颜色/尺码/品名若为空则视为“异常”，后续清洗时会剔除（这里仍保留原始行，供表格/KPI使用）
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
