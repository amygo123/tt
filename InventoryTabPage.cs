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
    public class InventoryTabPage : TabPage
    {
        public event Action<int, int, Dictionary<string, int>>? SummaryUpdated;

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
                Rows.GroupBy(r => r.Color).Select(g => new{ k=g.Key, v=g.Sum(x=>x.Available)})
                    .Where(x=>!string.IsNullOrWhiteSpace(x.k) && x.v!=0)
                    .OrderByDescending(x=>x.v).Select(x=>x.k);
            public IEnumerable<string> SizesNonZero() =>
                Rows.GroupBy(r => r.Size).Select(g => new{ k=g.Key, v=g.Sum(x=>x.Available)})
                    .Where(x=>!string.IsNullOrWhiteSpace(x.k) && x.v!=0)
                    .OrderByDescending(x=>x.v).Select(x=>x.k);
            public Dictionary<string,int> ByWarehouse() => Rows.GroupBy(r=>r.Warehouse).ToDictionary(g=>g.Key, g=>g.Sum(x=>x.Available));
        }

        private sealed class LookupDto
        {
            public string? style_name { get; set; }
            public string? grade { get; set; }
            public double? min_price_one { get; set; }
            public double? breakeven_one { get; set; }
        }

        private static readonly HttpClient _http = new();
        private readonly AppConfig _cfg;
        private InvSnapshot _all = new();
        private string _styleName = string.Empty;

        // UI
        private readonly PlotView _pvHeat = new() { Dock = DockStyle.Fill, BackColor = Color.White };
        private readonly PlotView _pvColor = new() { Dock = DockStyle.Fill, BackColor = Color.White };
        private readonly PlotView _pvSize  = new() { Dock = DockStyle.Fill, BackColor = Color.White };
        private readonly DataGridView _grid = new() { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells };
        private readonly TabControl _subTabs = new() { Dock = DockStyle.Fill };
        private readonly ToolTip _tip = new() { InitialDelay = 0, ReshowDelay = 0, AutoPopDelay = 8000, ShowAlways = true };

        private Label _lblAvail = new();
        private Label _lblOnHand = new();

        // KPI controls
        private readonly Label _kpiGradeVal = new() { AutoSize = true, Font = new Font("Microsoft YaHei UI", 14, FontStyle.Bold) };
        private readonly Label _kpiMinPriceVal = new() { AutoSize = true, Font = new Font("Microsoft YaHei UI", 14, FontStyle.Bold) };
        private readonly Label _kpiBreakevenVal = new() { AutoSize = true, Font = new Font("Microsoft YaHei UI", 14, FontStyle.Bold) };

        // selection on overview heatmap
        private (string? color,string? size)? _activeCell = null;

        public InventoryTabPage(AppConfig cfg)
        {
            _cfg = cfg;
            Text = "库存";
            BackColor = Color.White;

            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(12) };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 60));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
            Controls.Add(root);

            var tools = BuildTopTools();
            root.Controls.Add(tools, 0, 0);

            var four = BuildFourArea();
            root.Controls.Add(four, 0, 1);

            root.Controls.Add(_subTabs, 0, 2);
        }

        private Control BuildTopTools()
        {
            var p = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
            p.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            p.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            // row1: totals + refresh
            var row1 = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 1 };
            row1.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            row1.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            row1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            row1.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            _lblAvail = new Label { AutoSize = true, Font = new Font("Microsoft YaHei UI", 10, FontStyle.Bold) };
            _lblOnHand = new Label { AutoSize = true, Font = new Font("Microsoft YaHei UI", 10, FontStyle.Bold), Margin = new Padding(16,0,0,0) };
            row1.Controls.Add(_lblAvail, 0, 0);
            row1.Controls.Add(_lblOnHand, 1, 0);
            row1.Controls.Add(new Panel{Dock=DockStyle.Fill}, 2, 0);
            var btnReload = new Button { Text = "刷新", AutoSize = true };
            btnReload.Click += async (s,e)=> await ReloadAsync(_styleName);
            row1.Controls.Add(btnReload, 3, 0);

            // row2: KPI
            var row2 = BuildKpiRow();

            p.Controls.Add(row1, 0, 0);
            p.Controls.Add(row2, 0, 1);
            return p;
        }

        private Control BuildKpiRow()
        {
            var wrap = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false, AutoScroll = true, Margin = new Padding(0,6,0,0) };
            wrap.Controls.Add(BuildKpiCard("定级", _kpiGradeVal));
            wrap.Controls.Add(BuildKpiCard("最低价", _kpiMinPriceVal));
            wrap.Controls.Add(BuildKpiCard("保本价", _kpiBreakevenVal));
            return wrap;
        }

        private Control BuildKpiCard(string title, Label valueLabel)
        {
            var card = new Panel { AutoSize = true, Padding = new Padding(12), Margin = new Padding(0,0,12,0), BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle };
            var titleLbl = new Label { Text = title, AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft YaHei UI", 9, FontStyle.Regular) };
            valueLabel.Text = "-";
            var stack = new FlowLayoutPanel{ FlowDirection = FlowDirection.TopDown, AutoSize = true };
            stack.Controls.Add(titleLbl);
            stack.Controls.Add(valueLabel);
            card.Controls.Add(stack);
            return card;
        }

        private Control BuildFourArea()
        {
            var grid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2 };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 50));

            PrepareGridColumns(_grid);

            grid.Controls.Add(_pvHeat, 0, 0);
            grid.Controls.Add(_pvColor, 1, 0);
            grid.Controls.Add(_pvSize,  0, 1);
            grid.Controls.Add(_grid,    1, 1);

            AttachHeatmapInteractions(_pvHeat, sel => { _activeCell = sel; ApplyOverviewFilter(); });
            return grid;
        }

        public async Task LoadInventoryAsync(string styleName) => await LoadAsync(styleName);

        public async Task LoadAsync(string styleName)
        {
            _styleName = styleName ?? string.Empty;
            await ReloadAsync(_styleName);
        }

        private async Task ReloadAsync(string styleName)
        {
            _activeCell = null;
            _all = await FetchInventoryAsync(styleName);
            await UpdateKpisAsync(styleName);
            RenderAll(_all);
        }

        private void RenderAll(InvSnapshot snap)
        {
            _lblAvail.Text = $"可用合计：{snap.TotalAvailable}";
            _lblOnHand.Text = $"现有合计：{snap.TotalOnHand}";

            RenderHeatmap(snap, _pvHeat, "颜色×尺码 可用数热力图");
            RenderBarsByColor(snap, _pvColor, "颜色可用（降序，滚轮/右键查看更多）");
            RenderBarsBySize(snap, _pvSize, "尺码可用（降序，滚轮/右键查看更多）");
            RenderWarehouseTabs(snap);

            try { SummaryUpdated?.Invoke(snap.TotalAvailable, snap.TotalOnHand, snap.ByWarehouse()); } catch {}
            BindGrid(_grid, snap.Rows);
        }

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

                List<string>? lines = null;
                try { lines = JsonSerializer.Deserialize<List<string>>(raw); } catch {}
                if (lines == null) lines = raw.Replace("\r\n","\n").Split('\n').ToList();

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
            catch { }
            return s;
        }

        private async Task UpdateKpisAsync(string styleName)
        {
            try
            {
                var info = await FetchLookupAsync(styleName);
                if (info != null)
                {
                    _kpiGradeVal.Text = info.grade ?? "-";
                    _kpiMinPriceVal.Text = (info.min_price_one ?? 0).ToString("0.##");
                    _kpiBreakevenVal.Text = (info.breakeven_one ?? 0).ToString("0.##");
                }
                else
                {
                    _kpiGradeVal.Text = "-";
                    _kpiMinPriceVal.Text = "-";
                    _kpiBreakevenVal.Text = "-";
                }
            }
            catch
            {
                _kpiGradeVal.Text = "-";
                _kpiMinPriceVal.Text = "-";
                _kpiBreakevenVal.Text = "-";
            }
        }

        private async Task<LookupDto?> FetchLookupAsync(string styleName)
        {
            if (string.IsNullOrWhiteSpace(styleName)) return null;
            var url = "http://192.168.40.97:8002/lookup?name=" + Uri.EscapeDataString(styleName);
            try
            {
                using var req = new HttpRequestMessage(HttpMethod.Get, url);
                var resp = await _http.SendAsync(req);
                var raw = await resp.Content.ReadAsStringAsync();
                var list = JsonSerializer.Deserialize<List<LookupDto>>(raw);
                return (list != null && list.Count > 0) ? list[0] : null;
            }
            catch { return null; }
        }

        private void RenderBarsByColor(InvSnapshot snap, PlotView pv, string title)
        {
            var model = new PlotModel { Title = title };
            var data = snap.Rows.GroupBy(r => r.Color).Select(g => new{ Key=g.Key, V=g.Sum(x=>x.Available)})
                                .OrderByDescending(x=>x.V).ToList();

            var cat = new CategoryAxis { Position = AxisPosition.Left, StartPosition = 1, EndPosition = 0, IsZoomEnabled = true, IsPanEnabled = true };
            foreach (var d in data) cat.Labels.Add(d.Key);
            var val = new LinearAxis { Position = AxisPosition.Bottom, MinorGridlineStyle = LineStyle.Dot, MajorGridlineStyle = LineStyle.Solid, IsZoomEnabled = true, IsPanEnabled = true };
            var series = new BarSeries();
            foreach (var d in data) series.Items.Add(new BarItem(d.V));

            model.Axes.Add(cat); model.Axes.Add(val); model.Series.Add(series);
            pv.Model = model;
            ApplyTopNZoom(cat, data.Count, 10);
            BindPanZoom(pv);
        }

        private void RenderBarsBySize(InvSnapshot snap, PlotView pv, string title)
        {
            var model = new PlotModel { Title = title };
            var data = snap.Rows.GroupBy(r => r.Size).Select(g => new{ Key=g.Key, V=g.Sum(x=>x.Available)})
                                .OrderByDescending(x=>x.V).ToList();

            var cat = new CategoryAxis { Position = AxisPosition.Left, StartPosition = 1, EndPosition = 0, IsZoomEnabled = true, IsPanEnabled = true };
            foreach (var d in data) cat.Labels.Add(d.Key);
            var val = new LinearAxis { Position = AxisPosition.Bottom, MinorGridlineStyle = LineStyle.Dot, MajorGridlineStyle = LineStyle.Solid, IsZoomEnabled = true, IsPanEnabled = true };
            var series = new BarSeries();
            foreach (var d in data) series.Items.Add(new BarItem(d.V));

            model.Axes.Add(cat); model.Axes.Add(val); model.Series.Add(series);
            pv.Model = model;
            ApplyTopNZoom(cat, data.Count, 10);
            BindPanZoom(pv);
        }

        private void ApplyTopNZoom(CategoryAxis cat, int total, int n)
        {
            if (total <= 0) return;
            var maxIndex = Math.Min(n - 1, total - 1);
            cat.Minimum = -0.5;
            cat.Maximum = maxIndex + 0.5;
        }

        private sealed class HeatmapContext { public List<string> Colors=new(); public List<string> Sizes=new(); public double[,] Data=new double[0,0]; }

        private HeatmapContext BuildHeatmap(InvSnapshot snap, PlotView pv, string title)
        {
            var colors = snap.ColorsNonZero().ToList();
            var sizes  = snap.SizesNonZero().ToList();
            var ci = colors.Select((c,i)=>(c,i)).ToDictionary(x=>x.c, x=>x.i);
            var si = sizes.Select((s,i)=>(s,i)).ToDictionary(x=>x.s, x=>x.i);

            var data = new double[colors.Count, sizes.Count];
            foreach (var g in snap.Rows.GroupBy(r=>new{r.Color,r.Size}))
            {
                if (!ci.ContainsKey(g.Key.Color) || !si.ContainsKey(g.Key.Size)) continue;
                data[ci[g.Key.Color], si[g.Key.Size]] = g.Sum(x=>x.Available);
            }

            var model = new PlotModel { Title = title };

            // stats
            var vals = new List<double>();
            foreach (var v in data) if (v > 0) vals.Add(v);
            vals.Sort();
            double minPos = vals.Count>0 ? vals.First() : 1.0;
            double p95 = vals.Count>0 ? Percentile(vals, 0.95) : 1.0;
            if (p95 <= 0) p95 = minPos;

            // gradient: light green -> yellow -> orange -> red
            var palette = OxyPalette.Interpolate(256,
                OxyColor.FromRgb(229,245,224),
                OxyColor.FromRgb(161,217,155),
                OxyColor.FromRgb(255,224,102),
                OxyColor.FromRgb(253,174,97),
                OxyColor.FromRgb(244,109,67),
                OxyColor.FromRgb(215,48,39));

            var caxis = new LinearColorAxis
            {
                Position = AxisPosition.Right,
                Palette = palette,
                Minimum = minPos,
                Maximum = p95,
                LowColor = OxyColor.FromRgb(242,242,242), // zero values
                HighColor = OxyColor.FromRgb(153,0,0)
            };
            model.Axes.Add(caxis);

            var axX = new LinearAxis
            {
                Position = AxisPosition.Bottom,
                Minimum = -0.5, Maximum = Math.Max(colors.Count - 0.5, 0.5),
                MajorStep = 1, MinorStep = 1,
                IsZoomEnabled = true, IsPanEnabled = true,
                LabelFormatter = d => { var k=(int)Math.Round(d); return (k>=0 && k<colors.Count) ? colors[k] : ""; }
            };
            var axY = new LinearAxis
            {
                Position = AxisPosition.Left,
                Minimum = -0.5, Maximum = Math.Max(sizes.Count - 0.5, 0.5),
                MajorStep = 1, MinorStep = 1,
                IsZoomEnabled = true, IsPanEnabled = true,
                LabelFormatter = d => { var k=(int)Math.Round(d); return (k>=0 && k<sizes.Count) ? sizes[k] : ""; }
            };
            model.Axes.Add(axX); model.Axes.Add(axY);

            var hm = new HeatMapSeries
            {
                X0 = -0.5, X1 = colors.Count - 0.5,
                Y0 = -0.5, Y1 = sizes.Count - 0.5,
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

        private static double Percentile(List<double> sorted, double p)
        {
            if (sorted.Count == 0) return 0;
            if (p <= 0) return sorted.First();
            if (p >= 1) return sorted.Last();
            var idx = (sorted.Count - 1) * p;
            var lo = (int)Math.Floor(idx);
            var hi = (int)Math.Ceiling(idx);
            if (lo == hi) return sorted[lo];
            var frac = idx - lo;
            return sorted[lo] * (1 - frac) + sorted[hi] * frac;
        }

        private void RenderHeatmap(InvSnapshot snap, PlotView pv, string title)
        {
            BuildHeatmap(snap, pv, title);
            BindPanZoom(pv);
        }

        private void RenderWarehouseTabs(InvSnapshot snap)
        {
            _subTabs.SuspendLayout();
            _subTabs.TabPages.Clear();

            foreach (var g in snap.Rows.GroupBy(r => r.Warehouse).OrderByDescending(x => x.Sum(y => y.Available)))
            {
                var page = new TabPage($"{g.Key}（{g.Sum(x => x.Available)}）");

                var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
                root.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
                root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

                var search = new TextBox { Dock = DockStyle.Fill, PlaceholderText = "搜索本仓（颜色/尺码）" };
                root.Controls.Add(search, 0, 0);

                var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
                panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
                panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

                var pv = new PlotView { Dock = DockStyle.Fill, BackColor = Color.White };
                var grid = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells };
                PrepareGridColumns(grid);

                panel.Controls.Add(pv, 0, 0);
                panel.Controls.Add(grid, 1, 0);

                root.Controls.Add(panel, 0, 1);
                page.Controls.Add(root);

                var subSnap = new InvSnapshot();
                foreach (var r in g) subSnap.Rows.Add(r);
                BuildHeatmap(subSnap, pv, $"{g.Key} 颜色×尺码");

                BindGrid(grid, subSnap.Rows);

                (string? color, string? size)? sel = null;
                AttachHeatmapInteractions(pv, newSel =>
                {
                    sel = newSel;
                    ApplyWarehouseFilter(grid, subSnap, search.Text, sel);
                });

                var t = new System.Windows.Forms.Timer { Interval = 220 };
                search.TextChanged += (s2, e2) => { t.Stop(); t.Start(); };
                t.Tick += (s3, e3) => { t.Stop(); ApplyWarehouseFilter(grid, subSnap, search.Text, sel); };

                _subTabs.TabPages.Add(page);
            }
            _subTabs.ResumeLayout();
        }

        private void BindGrid(DataGridView grid, IEnumerable<InvRow> rows)
        {
            PrepareGridColumns(grid);
            grid.DataSource = new BindingList<InvRow>(rows.ToList());
        }

        private void PrepareGridColumns(DataGridView grid)
        {
            grid.AutoGenerateColumns = false;
            grid.Columns.Clear();
            grid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Name", HeaderText = "品名" });
            grid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Color", HeaderText = "颜色" });
            grid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Size", HeaderText = "尺码" });
            grid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Warehouse", HeaderText = "仓库" });
            grid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Available", HeaderText = "可用" });
            grid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "OnHand", HeaderText = "现有" });
        }

        private void ApplyOverviewFilter()
        {
            IEnumerable<InvRow> q = _all.Rows;
            if (_activeCell is { } ac)
            {
                if (!string.IsNullOrEmpty(ac.color)) q = q.Where(r => r.Color == ac.color);
                if (!string.IsNullOrEmpty(ac.size)) q = q.Where(r => r.Size == ac.size);
            }
            BindGrid(_grid, q);
        }

        private void ApplyWarehouseFilter(DataGridView grid, InvSnapshot snap, string key, (string? color, string? size)? sel)
        {
            IEnumerable<InvRow> q = snap.Rows;
            var k = (key ?? "").Trim();
            if (k.Length > 0)
            {
                q = q.Where(r => (r.Color?.IndexOf(k, StringComparison.OrdinalIgnoreCase) ?? -1) >= 0
                              || (r.Size?.IndexOf(k, StringComparison.OrdinalIgnoreCase) ?? -1) >= 0);
            }
            if (sel is { } ac)
            {
                if (!string.IsNullOrEmpty(ac.color)) q = q.Where(r => r.Color == ac.color);
                if (!string.IsNullOrEmpty(ac.size)) q = q.Where(r => r.Size == ac.size);
            }
            BindGrid(grid, q);
        }

        private void BindPanZoom(PlotView pv)
        {
            try
            {
                var ctl = pv.Controller ?? new PlotController();
                ctl.BindMouseWheel(PlotCommands.ZoomWheel);
                ctl.BindMouseDown(OxyMouseButton.Right, PlotCommands.PanAt);
                pv.Controller = ctl;
            }
            catch { }
        }

        private void AttachHeatmapInteractions(PlotView pv, Action<(string? color, string? size)?> onSelectionChanged)
        {
            try
            {
                var ctl = pv.Controller ?? new PlotController();
                ctl.UnbindMouseDown(OxyMouseButton.Left);
                ctl.BindMouseDown(OxyMouseButton.Middle, PlotCommands.PanAt);
                ctl.BindMouseDown(OxyMouseButton.Right, PlotCommands.PanAt);
                ctl.BindMouseWheel(PlotCommands.ZoomWheel);
                pv.Controller = ctl;
            }
            catch { }

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
                else _tip.Hide(pv);
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
                        onSelectionChanged((color, size));
                        return;
                    }
                }
                onSelectionChanged(null);
            };
        }

        // External API: allow selecting a warehouse tab by name from outside
        public void ActivateWarehouse(string warehouse)
        {
            if (string.IsNullOrWhiteSpace(warehouse)) return;
            foreach (TabPage tp in _subTabs.TabPages)
            {
                var name = tp.Text.Split('（')[0];
                if (string.Equals(name, warehouse, StringComparison.OrdinalIgnoreCase))
                {
                    _subTabs.SelectedTab = tp;
                    return;
                }
            }
        }
    }
}
#pragma warning restore 0618
