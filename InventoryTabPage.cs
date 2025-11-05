using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
#pragma warning disable 0618
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using OxyPlot.WindowsForms;

namespace StyleWatcherWin
{
    /// <summary>
    /// 库存页：热力图（颜色×尺码）+ 颜色/尺码柱状 + 分仓 + 明细
    /// 本次修复/迭代：
    /// 1) 修复 HeatMap 缺少 ColorAxis 的异常（新增 LinearColorAxis）
    /// 2) 悬浮提示（颜色/尺码/库存）+ 点击格子联动右侧明细筛选
    /// 3) 兼容 ResultForm 对 LoadInventoryAsync 的调用（新增该方法包装到 LoadAsync）
    /// 4) 保持柱状图不做 TOP-N 限制（不锁前10）
    /// 5) 分仓子页面升级：左侧“分仓热力图”，右侧“该仓明细”，点击联动筛选
    /// </summary>
    public class InventoryTabPage : TabPage
    {
        public event Action<int, int, Dictionary<string, int>>? SummaryUpdated;

        #region 数据结构
        private sealed class InvRow
        {
            public string Name { get; set; } = "";
            public string Color { get; set; } = "";
            public string Size { get; set; } = "";
            public string Warehouse { get; set; } = "";
            public int Available { get; set; }
            public int OnHand { get; set; }
        }

        private sealed class InvSnapshot
        {
            public List<InvRow> Rows { get; } = new();

            public int TotalAvailable => Rows.Sum(r => r.Available);
            public int TotalOnHand => Rows.Sum(r => r.OnHand);

            public IEnumerable<string> ColorsNonZero() =>
                Rows.GroupBy(r => r.Color)
                    .Select(g => new { c = g.Key, v = g.Sum(x => x.Available) })
                    .Where(x => !string.IsNullOrWhiteSpace(x.c) && x.v != 0)
                    .OrderByDescending(x => x.v)
                    .Select(x => x.c);

            public IEnumerable<string> SizesNonZero() =>
                Rows.GroupBy(r => r.Size)
                    .Select(g => new { s = g.Key, v = g.Sum(x => x.Available) })
                    .Where(x => !string.IsNullOrWhiteSpace(x.s) && x.v != 0)
                    .OrderByDescending(x => x.v)
                    .Select(x => x.s);

            public Dictionary<string, int> ByWarehouse() =>
                Rows.GroupBy(r => r.Warehouse)
                    .ToDictionary(g => g.Key, g => g.Sum(x => x.Available));

            public InvSnapshot Filter(Func<InvRow, bool> pred)
            {
                var s = new InvSnapshot();
                foreach (var r in Rows.Where(pred)) s.Rows.Add(r);
                return s;
            }
        }
        #endregion

        private static readonly HttpClient _http = new();

        private readonly AppConfig _cfg;
        private InvSnapshot _all = new();
        private string _styleName = "";

        // 顶部工具区
        private readonly TextBox _search = new() { PlaceholderText = "搜索（颜色/尺码/仓库）" };
        private readonly Label _lblAvail = new() { AutoSize = true, Font = new Font("Microsoft YaHei UI", 10, FontStyle.Bold) };
        private readonly Label _lblOnHand = new() { AutoSize = true, Font = new Font("Microsoft YaHei UI", 10, FontStyle.Bold), Margin = new Padding(16, 0, 0, 0) };
        private readonly System.Windows.Forms.Timer _debounce = new() { Interval = 220 };

        // 主要图表
        private readonly PlotView _pvHeat = new() { Dock = DockStyle.Fill, BackColor = Color.White };
        private readonly PlotView _pvColor = new() { Dock = DockStyle.Fill, BackColor = Color.White };
        private readonly PlotView _pvSize = new() { Dock = DockStyle.Fill, BackColor = Color.White };

        // 明细
        private readonly DataGridView _grid = new() { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells };
        private BindingList<InvRow> _gridData = new();

        // 子 Tab（按仓库）
        private readonly TabControl _subTabs = new() { Dock = DockStyle.Fill };

        // 悬浮提示
        private readonly ToolTip _tip = new() { InitialDelay = 0, ReshowDelay = 0, AutoPopDelay = 8000, ShowAlways = true };

        // 当前点击筛选（主热力图）
        private (string? color, string? size)? _activeCell = null;

        public InventoryTabPage(AppConfig cfg)
        {
            _cfg = cfg;
            Text = "库存";

            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(12) };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 60));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
            Controls.Add(root);

            var tools = BuildTopTools();
            root.Controls.Add(tools, 0, 0);

            var four = BuildFourArea();
            root.Controls.Add(four, 0, 1);

            root.Controls.Add(_subTabs, 0, 2);

            _debounce.Tick += (s, e) => { _debounce.Stop(); ApplySearchFilter(); };
        }

        private Control BuildTopTools()
        {
            var p = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 1 };
            p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60));
            p.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            p.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            p.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            _search.Dock = DockStyle.Fill;
            _search.TextChanged += (s, e) => _debounce.Start();

            p.Controls.Add(_search, 0, 0);
            p.Controls.Add(_lblAvail, 1, 0);
            p.Controls.Add(_lblOnHand, 2, 0);

            var btnReload = new Button { Text = "刷新", AutoSize = true, Margin = new Padding(16, 0, 0, 0) };
            btnReload.Click += async (s, e) => await ReloadAsync(_styleName);
            p.Controls.Add(btnReload, 3, 0);

            return p;
        }

        private Control BuildFourArea()
        {
            var grid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2 };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 50));

            _grid.DataSource = _gridData;

            grid.Controls.Add(_pvSize, 0, 0);
            grid.Controls.Add(_pvColor, 1, 0);
            grid.Controls.Add(_pvHeat, 0, 1);
            grid.Controls.Add(_grid, 1, 1);

            AttachHeatmapInteractions(_pvHeat, sel =>
            {
                _activeCell = sel;
                ApplySearchFilter();
            });

            return grid;
        }

        #region 对外入口（兼容旧方法名）
        public async Task LoadAsync(string styleName)
        {
            _styleName = styleName ?? "";
            await ReloadAsync(_styleName);
        }

        public Task LoadInventoryAsync(string styleName) => LoadAsync(styleName); // 兼容 ResultForm 旧调用
        #endregion

        private async Task ReloadAsync(string styleName)
        {
            _activeCell = null; // 清筛选
            _search.Text = "";

            _all = await FetchInventoryAsync(styleName);
            RenderAll(_all);
        }

        private void RenderAll(InvSnapshot snap)
        {
            _lblAvail.Text = $"可用合计：{snap.TotalAvailable}";
            _lblOnHand.Text = $"现有合计：{snap.TotalOnHand}";

            RenderHeatmap(snap, _pvHeat, "颜色×尺码 可用数热力图");
            RenderBarsByColor(snap, _pvColor, "颜色可用（降序）");
            RenderBarsBySize(snap, _pvSize, "尺码可用（降序）");
            RenderWarehouseTabs(snap);

            try { SummaryUpdated?.Invoke(snap.TotalAvailable, snap.TotalOnHand, snap.ByWarehouse()); } catch { }
            FillGrid(_grid, snap.Rows);
        }

        #region 数据获取/解析
        private async Task<InvSnapshot> FetchInventoryAsync(string styleName)
        {
            var s = new InvSnapshot();
            if (string.IsNullOrWhiteSpace(styleName)) return s;

            try
            {
                var baseUrl = (_cfg?.inventory?.url_base ?? "");
                var url = baseUrl.Contains("style_name=")
                    ? baseUrl + Uri.EscapeDataString(styleName)
                    : baseUrl.TrimEnd('/') + "?style_name=" + Uri.EscapeDataString(styleName);

                using var req = new HttpRequestMessage(HttpMethod.Get, url);
                var resp = await _http.SendAsync(req);
                var raw = await resp.Content.ReadAsStringAsync();

                // 兼容：JSON 数组字符串 或 纯文本行
                List<string>? lines = null;
                try { lines = JsonSerializer.Deserialize<List<string>>(raw); } catch { }
                if (lines == null) lines = raw.Replace("\r\n", "\n").Split('\n').ToList();

                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    var seg = line.Replace('，', ',').Split(',');
                    if (seg.Length < 6) continue;
                    s.Rows.Add(new InvRow
                    {
                        Name = seg[0].Trim(),
                        Color = seg[1].Trim(),
                        Size = seg[2].Trim(),
                        Warehouse = seg[3].Trim(),
                        Available = int.TryParse(seg[4].Trim(), out var a) ? a : 0,
                        OnHand = int.TryParse(seg[5].Trim(), out var h) ? h : 0
                    });
                }
            }
            catch
            {
                // ignore，返回空快照
            }
            return s;
        }
        #endregion

        #region 绘图
        private void RenderBarsByColor(InvSnapshot snap, PlotView pv, string title)
        {
            var model = new PlotModel { Title = title };
            var data = snap.Rows.GroupBy(r => r.Color)
                                .Select(g => new { Key = g.Key, V = g.Sum(x => x.Available) })
                                .OrderByDescending(x => x.V)
                                .ToList();

            var cat = new CategoryAxis { Position = AxisPosition.Left }; // BarSeries => 类别轴在左边
            foreach (var d in data) cat.Labels.Add(d.Key);

            var val = new LinearAxis { Position = AxisPosition.Bottom, MinorGridlineStyle = LineStyle.Dot, MajorGridlineStyle = LineStyle.Solid };
            var series = new BarSeries();
            foreach (var d in data) series.Items.Add(new BarItem(d.V));

            model.Axes.Add(cat);
            model.Axes.Add(val);
            model.Series.Add(series);
            pv.Model = model;
        }

        private void RenderBarsBySize(InvSnapshot snap, PlotView pv, string title)
        {
            var model = new PlotModel { Title = title };
            var data = snap.Rows.GroupBy(r => r.Size)
                                .Select(g => new { Key = g.Key, V = g.Sum(x => x.Available) })
                                .OrderByDescending(x => x.V)
                                .ToList();

            var cat = new CategoryAxis { Position = AxisPosition.Left };
            foreach (var d in data) cat.Labels.Add(d.Key);

            var val = new LinearAxis { Position = AxisPosition.Bottom, MinorGridlineStyle = LineStyle.Dot, MajorGridlineStyle = LineStyle.Solid };
            var series = new BarSeries();
            foreach (var d in data) series.Items.Add(new BarItem(d.V));

            model.Axes.Add(cat);
            model.Axes.Add(val);
            model.Series.Add(series);
            pv.Model = model;
        }

        private sealed class HeatmapContext
        {
            public List<string> Colors = new();
            public List<string> Sizes = new();
            public double[,] Data = new double[0, 0];
        }

        private HeatmapContext BuildHeatmap(InvSnapshot snap, PlotView pv, string title)
        {
            var colors = snap.ColorsNonZero().ToList();
            var sizes = snap.SizesNonZero().ToList();

            var ci = colors.Select((c, i) => (c, i)).ToDictionary(x => x.c, x => x.i);
            var si = sizes.Select((s, i) => (s, i)).ToDictionary(x => x.s, x => x.i);

            var data = new double[colors.Count, sizes.Count];
            foreach (var g in snap.Rows.GroupBy(r => new { r.Color, r.Size }))
            {
                if (!ci.ContainsKey(g.Key.Color) || !si.ContainsKey(g.Key.Size)) continue;
                data[ci[g.Key.Color], si[g.Key.Size]] = g.Sum(x => x.Available);
            }

            var model = new PlotModel { Title = title };

            // 颜色轴（修复异常）
            var caxis = new LinearColorAxis { Position = AxisPosition.Right, Palette = OxyPalettes.Jet(256) };
            model.Axes.Add(caxis);

            // 类目映射轴
            var axX = new LinearAxis
            {
                Position = AxisPosition.Bottom,
                Minimum = -0.5, Maximum = Math.Max(colors.Count - 0.5, 0.5),
                MajorStep = 1, MinorStep = 1,
                IsZoomEnabled = true, IsPanEnabled = true,
                LabelFormatter = d =>
                {
                    var k = (int)Math.Round(d);
                    return (k >= 0 && k < colors.Count) ? colors[k] : "";
                }
            };

            var axY = new LinearAxis
            {
                Position = AxisPosition.Left,
                Minimum = -0.5, Maximum = Math.Max(sizes.Count - 0.5, 0.5),
                MajorStep = 1, MinorStep = 1,
                IsZoomEnabled = true, IsPanEnabled = true,
                LabelFormatter = d =>
                {
                    var k = (int)Math.Round(d);
                    return (k >= 0 && k < sizes.Count) ? sizes[k] : "";
                }
            };

            model.Axes.Add(axX);
            model.Axes.Add(axY);

            var hm = new HeatMapSeries
            {
                X0 = -0.5,
                X1 = colors.Count - 0.5,
                Y0 = -0.5,
                Y1 = sizes.Count - 0.5,
                Interpolate = false,
                RenderMethod = HeatMapRenderMethod.Rectangles,
                Data = data
            };

            model.Series.Add(hm);
            pv.Model = model;

            var ctx = new HeatmapContext { Colors = colors, Sizes = sizes, Data = data };
            pv.Tag = ctx;
            return ctx;
        }

        private void RenderHeatmap(InvSnapshot snap, PlotView pv, string title)
        {
            BuildHeatmap(snap, pv, title);
        }
        #endregion

        private void RenderWarehouseTabs(InvSnapshot snap)
        {
            _subTabs.SuspendLayout();
            _subTabs.TabPages.Clear();

            foreach (var g in snap.Rows.GroupBy(r => r.Warehouse).OrderByDescending(x => x.Sum(y => y.Available)))
            {
                var page = new TabPage($"{g.Key}（{g.Sum(x => x.Available)}）");

                var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
                panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
                panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

                var pv = new PlotView { Dock = DockStyle.Fill, BackColor = Color.White };
                var grid = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells };

                panel.Controls.Add(pv, 0, 0);
                panel.Controls.Add(grid, 1, 0);
                page.Controls.Add(panel);

                var subSnap = new InvSnapshot();
                foreach (var r in g) subSnap.Rows.Add(r);
                BuildHeatmap(subSnap, pv, $"{g.Key} 颜色×尺码");

                // 初始填充该仓明细
                FillGrid(grid, subSnap.Rows);

                // 联动：点击该仓热力图筛选右侧明细
                (string? color, string? size)? sel = null;
                AttachHeatmapInteractions(pv, newSel =>
                {
                    sel = newSel;
                    IEnumerable<InvRow> q = subSnap.Rows;
                    if (sel is { } ac)
                    {
                        if (!string.IsNullOrEmpty(ac.color)) q = q.Where(r => r.Color == ac.color);
                        if (!string.IsNullOrEmpty(ac.size)) q = q.Where(r => r.Size == ac.size);
                    }
                    FillGrid(grid, q);
                });

                _subTabs.TabPages.Add(page);
            }
            _subTabs.ResumeLayout();
        }

        private void FillGrid(DataGridView grid, IEnumerable<InvRow> rows)
        {
            grid.DataSource = new BindingList<InvRow>(rows.ToList());
        }

        private void ApplySearchFilter()
        {
            var key = (_search.Text ?? "").Trim();
            IEnumerable<InvRow> q = _all.Rows;

            if (!string.IsNullOrEmpty(key))
                q = q.Where(r => (r.Color?.IndexOf(key, StringComparison.OrdinalIgnoreCase) ?? -1) >= 0
                              || (r.Size?.IndexOf(key, StringComparison.OrdinalIgnoreCase) ?? -1) >= 0
                              || (r.Warehouse?.IndexOf(key, StringComparison.OrdinalIgnoreCase) ?? -1) >= 0);

            if (_activeCell is { } ac)
            {
                if (!string.IsNullOrEmpty(ac.color))
                    q = q.Where(r => r.Color == ac.color);
                if (!string.IsNullOrEmpty(ac.size))
                    q = q.Where(r => r.Size == ac.size);
            }

            FillGrid(_grid, q);
        }

        #region 交互：取消默认点击 → 自定义悬浮与联动筛选
        private void AttachHeatmapInteractions(PlotView pv, Action<(string? color, string? size)?> onSelectionChanged)
        {
            // 取消默认左键行为（包括默认 Tracker）
            try
            {
                var ctl = new PlotController();
                ctl.UnbindMouseDown(OxyMouseButton.Left);
                // 保留中键平移/滚轮缩放（可选）
                ctl.BindMouseWheel(PlotCommands.ZoomWheel);
                ctl.BindMouseDown(OxyMouseButton.Middle, PlotCommands.PanAt);
                pv.Controller = ctl;
            }
            catch { /* ignore */ }

            pv.MouseMove += (s, e) =>
            {
                var model = pv.Model;
                if (model == null) return;

                var hm = model.Series.OfType<HeatMapSeries>().FirstOrDefault();
                var ctx = pv.Tag as HeatmapContext;
                if (hm == null || ctx == null || ctx.Colors.Count == 0 || ctx.Sizes.Count == 0) return;

                var sp = new ScreenPoint(e.Location.X, e.Location.Y);
                var hit = hm.GetNearestPoint(sp, false);
                if (hit == null) { _tip.Hide(pv); return; }

                var xi = (int)Math.Round(hit.DataPoint.X);
                var yi = (int)Math.Round(hit.DataPoint.Y);

                if (xi >= 0 && xi < ctx.Colors.Count && yi >= 0 && yi < ctx.Sizes.Count)
                {
                    var color = ctx.Colors[xi];
                    var size = ctx.Sizes[yi];
                    var val = ctx.Data[xi, yi];
                    _tip.Show($"颜色：{color}  尺码：{size}  库存：{val:0}", pv, e.Location.X + 12, e.Location.Y + 12);
                }
                else
                {
                    _tip.Hide(pv);
                }
            };

            pv.MouseLeave += (s, e) => _tip.Hide(pv);

            pv.MouseDown += (s, e) =>
            {
                if (e.Button != MouseButtons.Left) return;

                var model = pv.Model;
                if (model == null) return;

                var hm = model.Series.OfType<HeatMapSeries>().FirstOrDefault();
                var ctx = pv.Tag as HeatmapContext;
                if (hm == null || ctx == null || ctx.Colors.Count == 0 || ctx.Sizes.Count == 0) return;

                var sp = new ScreenPoint(e.Location.X, e.Location.Y);
                var hit = hm.GetNearestPoint(sp, false);

                if (hit != null)
                {
                    var xi = (int)Math.Round(hit.DataPoint.X);
                    var yi = (int)Math.Round(hit.DataPoint.Y);
                    if (xi >= 0 && xi < ctx.Colors.Count && yi >= 0 && yi < ctx.Sizes.Count)
                    {
                        var color = ctx.Colors[xi];
                        var size = ctx.Sizes[yi];

                        // 二次点击取消
                        var current = (color, size);
                        // 读取之前的选择（通过外部闭包保存），这里只能通知外部决定
                        onSelectionChanged(current);
                        return;
                    }
                }

                // 点击空白取消
                onSelectionChanged(null);
            };
        }
        #endregion
    }
}
#pragma warning restore 0618
