using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using OxyPlot.WindowsForms;

namespace StyleWatcherWin
{
    static class UI
    {
        public static readonly Font Title = new("Microsoft YaHei UI", 12, FontStyle.Bold);
        public static readonly Font Body  = new("Microsoft YaHei UI", 10);
        public static readonly Color HeaderBack = Color.FromArgb(245,247,250);
        public static readonly Color CardBack   = Color.FromArgb(250,250,250);
        public static readonly Color Text       = Color.FromArgb(47,47,47);
    }

    public class ResultForm : Form
    {
        private readonly AppConfig _cfg;

        // Header
        private readonly TextBox _input = new();
        private readonly Button _btnQuery = new();
        private readonly Button _btnExport = new();

        // KPI
        private readonly FlowLayoutPanel _kpi = new();
        private readonly Panel _kpiSales7 = new();
        private readonly Panel _kpiInv = new();
        private readonly Panel _kpiDoc = new();
        private readonly Panel _kpiMissing = new();

        // Tabs
        private readonly TabControl _tabs = new();

        // Overview
        private readonly FlowLayoutPanel _trendSwitch = new();
        private int _trendWindow = 7;
        private readonly PlotView _plotTrend = new();
        private readonly PlotView _plotSize = new();
        private readonly PlotView _plotColor = new();
        private readonly PlotView _plotWarehouse = new();
        private readonly CheckBox _chkTopN = new() { Text = "只看 Top 10", Checked = true, AutoSize = true, Margin = new Padding(0,0,10,0) };
        private const int DEFAULT_TOPN = 10;

        // Detail
        private readonly DataGridView _grid = new();
        private readonly BindingSource _binding = new();
        private readonly TextBox _boxSearch = new();
        private readonly FlowLayoutPanel _filterChips = new(); // 展示从图表点选带来的过滤标签
        private readonly Timer _searchDebounce = new() { Interval = 200 };

        // Inventory page（A1 不依赖，保持占位）
        private InventoryTabPage? _invPage;

        // Caches
        private string _lastDisplayText = string.Empty;
        private List<Aggregations.SalesItem> _sales = new();
        private List<object> _gridMaster = new();

        public ResultForm(AppConfig cfg)
        {
            _cfg = cfg;
            _trendWindow = (_cfg.ui?.trendWindows?.FirstOrDefault() ?? 7);

            Text = "StyleWatcher";
            Font = new Font("Microsoft YaHei UI", _cfg.window.fontSize);
            Width = Math.Max(1200, _cfg.window.width);
            Height = Math.Max(800, _cfg.window.height);
            StartPosition = FormStartPosition.CenterScreen;
            TopMost = _cfg.window.alwaysOnTop;
            BackColor = Color.White;
            KeyPreview = true;
            KeyDown += (s,e)=>{ if(e.KeyCode==Keys.Escape) Hide(); };

            var root = new TableLayoutPanel{Dock=DockStyle.Fill,RowCount=2,ColumnCount=1};
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 76));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Controls.Add(root);

            var header = BuildHeader();
            root.Controls.Add(header,0,0);

            var content = new TableLayoutPanel{Dock=DockStyle.Fill,RowCount=2,ColumnCount=1};
            content.RowStyles.Add(new RowStyle(SizeType.Absolute, 130));
            content.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.Controls.Add(content,0,1);

            _kpi.Dock = DockStyle.Fill;
            _kpi.FlowDirection = FlowDirection.LeftToRight;
            _kpi.WrapContents = true;
            _kpi.Padding = new Padding(12,8,12,8);
            _kpi.Controls.Add(MakeKpi(_kpiSales7,"近7日销量","—"));
            _kpi.Controls.Add(MakeKpi(_kpiInv,"可用库存总量","—"));
            _kpi.Controls.Add(MakeKpi(_kpiDoc,"库存天数","—"));
            _kpi.Controls.Add(MakeKpi(_kpiMissing,"缺失尺码","—"));
            content.Controls.Add(_kpi,0,0);

            _tabs.Dock = DockStyle.Fill;
            BuildTabs();
            content.Controls.Add(_tabs,0,1);

            _searchDebounce.Tick += (s,e)=> { _searchDebounce.Stop(); ApplyFilter(_boxSearch.Text); };
        }

        private Control BuildHeader()
        {
            var head = new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=3,RowCount=1,Padding=new(12,10,12,6),BackColor=UI.HeaderBack};
            head.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
            head.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            head.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            _input.MinimumSize = new Size(420,32);

            _btnQuery.Text="重新查询";
            _btnQuery.AutoSize=true; _btnQuery.Padding=new Padding(10,6,10,6);
            _btnQuery.Click += async (s,e)=>{ _btnQuery.Enabled=false; try{ await ReloadAsync(""); } finally{ _btnQuery.Enabled=true; } };

            _btnExport.Text="导出Excel";
            _btnExport.AutoSize=true; _btnExport.Padding=new Padding(10,6,10,6);
            _btnExport.Click += (s,e)=> ExportExcel();

            head.Controls.Add(_input,0,0);
            head.Controls.Add(_btnQuery,1,0);
            head.Controls.Add(_btnExport,2,0);
            return head;
        }

        private Control MakeKpi(Panel host,string title,string value)
        {
            host.Width=260; host.Height=110; host.Padding=new Padding(10);
            host.BackColor=UI.CardBack; host.BorderStyle=BorderStyle.FixedSingle;
            host.Margin = new Padding(8,4,8,4);

            var inner = new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=2};
            inner.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));                 // 固定标题行高度，避免遮挡
            inner.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var t=new Label{Text=title,Dock=DockStyle.Fill,Height=28,ForeColor=UI.Text,Font=UI.Body,TextAlign=ContentAlignment.MiddleLeft};
            var v=new Label{Text=value,Dock=DockStyle.Fill,Font=new Font("Microsoft YaHei UI", 16, FontStyle.Bold),TextAlign=ContentAlignment.MiddleLeft,Padding=new Padding(0,2,0,0)};

            inner.Controls.Add(t,0,0);
            inner.Controls.Add(v,0,1);
            host.Controls.Clear();
            host.Controls.Add(inner);
            return host;
        }

        private void BuildTabs()
        {
            // 概览：两行两列 + 顶部工具区（趋势窗口切换/TopN开关）
            var overview = new TabPage("概览"){BackColor=Color.White};

            var container = new TableLayoutPanel{Dock=DockStyle.Fill,RowCount=2,ColumnCount=1};
            container.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));
            container.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            // 顶部工具区
            var tools = new FlowLayoutPanel{Dock=DockStyle.Fill, FlowDirection=FlowDirection.LeftToRight, WrapContents=false, AutoScroll=true, Padding=new Padding(12,10,12,0)};
            _trendSwitch.FlowDirection=FlowDirection.LeftToRight;
            _trendSwitch.WrapContents=false;
            _trendSwitch.AutoSize=true;
            _trendSwitch.Padding=new Padding(0,0,12,0);
            var wins=_cfg.ui?.trendWindows??new[]{7,14,30};
            foreach(var w in wins.Distinct().OrderBy(x=>x)){
                var rb=new RadioButton{Text=$"{w} 日",AutoSize=true,Tag=w,Margin=new Padding(0,2,18,0)};
                if(w==_trendWindow) rb.Checked=true;
                rb.CheckedChanged+=(s,e)=>{ var me=(RadioButton)s; if(me.Checked){ _trendWindow=(int)me.Tag; if(_sales.Count>0) RenderCharts(_sales); } };
                _trendSwitch.Controls.Add(rb);
            }
            tools.Controls.Add(_trendSwitch);
            _chkTopN.CheckedChanged += (s,e)=> { if(_sales.Count>0) RenderCharts(_sales); };
            tools.Controls.Add(_chkTopN);

            container.Controls.Add(tools,0,0);

            // 主图网格
            var grid = new TableLayoutPanel{Dock=DockStyle.Fill,RowCount=2,ColumnCount=2,Padding=new Padding(12)};
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 55));
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 45));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            _plotTrend.Dock=DockStyle.Fill;
            _plotWarehouse.Dock=DockStyle.Fill;
            _plotSize.Dock=DockStyle.Fill;
            _plotColor.Dock=DockStyle.Fill;

            grid.Controls.Add(_plotTrend,0,0);
            grid.Controls.Add(_plotWarehouse,1,0);
            grid.Controls.Add(_plotSize,0,1);
            grid.Controls.Add(_plotColor,1,1);

            container.Controls.Add(grid,0,1);
            overview.Controls.Add(container);
            _tabs.TabPages.Add(overview);

            // 销售明细
            var detail = new TabPage("销售明细"){BackColor=Color.White};
            var panel = new TableLayoutPanel{Dock=DockStyle.Fill,RowCount=3,ColumnCount=1,Padding=new Padding(12)};
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent,100));

            _boxSearch.Dock=DockStyle.Fill; _boxSearch.PlaceholderText="搜索（日期/款式/尺码/颜色/数量）";
            _boxSearch.TextChanged += (s,e)=> { _searchDebounce.Stop(); _searchDebounce.Start(); };
            panel.Controls.Add(_boxSearch,0,0);

            _filterChips.Dock = DockStyle.Fill; _filterChips.FlowDirection = FlowDirection.LeftToRight;
            panel.Controls.Add(_filterChips,0,1);

            _grid.Dock=DockStyle.Fill; _grid.ReadOnly=true; _grid.AllowUserToAddRows=false; _grid.AllowUserToDeleteRows=false;
            _grid.RowHeadersVisible=false; _grid.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.AllCells;
            _grid.DataSource=_binding;
            panel.Controls.Add(_grid,0,2);
            detail.Controls.Add(panel);
            _tabs.TabPages.Add(detail);

            // 库存页（A1 暂不改动，只保留 Tab 占位；A2 重构）
            _invPage = null; // A2 中恢复为 new InventoryTabPage(_cfg)
        }

        private void SetKpiValue(Panel p,string value)
        {
            var val = p.Controls.OfType<TableLayoutPanel>().First().Controls.OfType<Label>().LastOrDefault();
            if(val!=null) val.Text=value ?? "—";
        }

        // —— 和 TrayApp.cs 对齐的接口 —— //
        public void FocusInput(){ try{ if(WindowState==FormWindowState.Minimized) WindowState=FormWindowState.Normal; _input.Focus(); _input.SelectAll(); }catch{} }
        public void ShowNoActivateAtCursor(){ try{ StartPosition=FormStartPosition.Manual; var pt=Cursor.Position; Location=new Point(Math.Max(0,pt.X-Width/2),Math.Max(0,pt.Y-Height/2)); Show(); }catch{ Show(); } }
        public void ShowAndFocusCentered(){ ShowAndFocusCentered(_cfg.window.alwaysOnTop); }
        public void ShowAndFocusCentered(bool alwaysOnTop){ TopMost=alwaysOnTop; StartPosition=FormStartPosition.CenterScreen; Show(); Activate(); FocusInput(); }
        public void SetLoading(string message){ SetKpiValue(_kpiSales7,"—"); SetKpiValue(_kpiInv,"—"); SetKpiValue(_kpiDoc,"—"); SetKpiValue(_kpiMissing,"—"); }
        public async void ApplyRawText(string selection, string parsed){ _input.Text=selection??string.Empty; _lastDisplayText = parsed ?? string.Empty; await LoadTextAsync(parsed??string.Empty); }
        public void ApplyRawText(string text){ _input.Text=text??string.Empty; }

        public async Task LoadTextAsync(string raw)=>await ReloadAsync(raw);
        private async Task ReloadAsync()=>await ReloadAsync(_input.Text);

        private async Task ReloadAsync(string displayText)
        {
            await Task.Yield();
            if (string.IsNullOrWhiteSpace(displayText))
                displayText = _lastDisplayText;

            var parsed = Parser.Parse(displayText ?? string.Empty);

            var newGrid = parsed.Records
                .OrderBy(r=>r.Name).ThenBy(r=>r.Color).ThenBy(r=>r.Size).ThenByDescending(r=>r.Date)
                .Select(r => (object)new { 日期=r.Date.ToString("yyyy-MM-dd"), 款式=r.Name, 颜色=r.Color, 尺码=r.Size, 数量=r.Qty }).ToList();

            var newSales = parsed.Records.Select(r=> new Aggregations.SalesItem{
                Date=r.Date, Size=r.Size??"", Color=r.Color??"", Qty=r.Qty
            }).ToList();

            _lastDisplayText = displayText ?? string.Empty;
            _sales = newSales;
            _gridMaster = newGrid;

            var sales7 = _sales.Where(x=>x.Date>=DateTime.Today.AddDays(-6)).Sum(x=>x.Qty);
            SetKpiValue(_kpiSales7, sales7.ToString());

            // 缺失尺码：展示具体列表
            SetKpiValue(_kpiMissing, string.Join(" / ", MissingSizes(_sales.Select(s=>s.Size))));

            RenderCharts(_sales);

            _binding.DataSource = new BindingList<object>(_gridMaster);
            _grid.ClearSelection();
            if (_grid.Columns.Contains("款式")) _grid.Columns["款式"].DisplayIndex = 0;
            if (_grid.Columns.Contains("颜色")) _grid.Columns["颜色"].DisplayIndex = 1;
            if (_grid.Columns.Contains("尺码")) _grid.Columns["尺码"].DisplayIndex = 2;
            if (_grid.Columns.Contains("日期")) _grid.Columns["日期"].DisplayIndex = 3;
            if (_grid.Columns.Contains("数量")) _grid.Columns["数量"].DisplayIndex = 4;

            // A1 暂时用“占位”方式渲染分仓饼图（没有库存 Summary 时显示空饼图），A2 接 InventoryTabPage 的 SummaryUpdated
            RenderWarehousePiePlaceholder();
        }

        private void RenderWarehousePiePlaceholder()
        {
            var model = new PlotModel { Title = "分仓库存占比" };
            var pie = new PieSeries { AngleSpan = 360, StartAngle = 0, StrokeThickness = 0.5, InsideLabelPosition = 0.6 };
            // A1：没有库存页联动时，给个占位，避免空白
            pie.Slices.Add(new PieSlice("无数据", 1));
            model.Series.Add(pie);
            _plotWarehouse.Model = model;
        }

        private static IEnumerable<string> MissingSizes(IEnumerable<string> sizes)
        {
            var set = new HashSet<string>(sizes.Where(s=>!string.IsNullOrWhiteSpace(s)), StringComparer.OrdinalIgnoreCase);
            var baseline = new []{"XS","S","M","L","XL","2XL","3XL","4XL","5XL","6XL","KXL","K2XL","K3XL","K4XL"};
            foreach (var s in baseline)
                if (!set.Contains(s)) yield return s;
        }

        private static List<Aggregations.SalesItem> CleanSalesForVisuals(IEnumerable<Aggregations.SalesItem> src)
        {
            var list = src.Where(s => !string.IsNullOrWhiteSpace(s.Color) && !string.IsNullOrWhiteSpace(s.Size)).ToList();
            var byColor = list.GroupBy(x=>x.Color).ToDictionary(g=>g.Key, g=>g.Sum(z=>z.Qty));
            var bySize  = list.GroupBy(x=>x.Size ).ToDictionary(g=>g.Key, g=>g.Sum(z=>z.Qty));
            list = list.Where(x => byColor.GetValueOrDefault(x.Color,0) != 0 && bySize.GetValueOrDefault(x.Size,0) != 0).ToList();
            return list;
        }

        private void RenderCharts(List<Aggregations.SalesItem> salesItems)
        {
            var cleaned = CleanSalesForVisuals(salesItems);

            // 1) 趋势（补零，按日）
            var series = Aggregations.BuildDateSeries(cleaned, _trendWindow);
            var modelTrend = new PlotModel { Title = $"近{_trendWindow}日销量趋势", PlotMargins = new OxyThickness(50,10,10,40) };
            modelTrend.IsLegendVisible = true;
            modelTrend.LegendPosition = LegendPosition.TopRight;

            var xAxis = new DateTimeAxis{ Position=AxisPosition.Bottom, StringFormat="MM-dd", IntervalType=DateTimeIntervalType.Days, MajorStep=1, MinorStep=1, IntervalLength=60, IsZoomEnabled=false, IsPanEnabled=false, MajorGridlineStyle=LineStyle.Solid };
            var yAxis = new LinearAxis{ Position=AxisPosition.Left, MinimumPadding=0, AbsoluteMinimum=0, MajorGridlineStyle=LineStyle.Solid };
            modelTrend.Axes.Add(xAxis); modelTrend.Axes.Add(yAxis);

            var line = new LineSeries{ Title="销量", MarkerType=MarkerType.Circle };
            foreach(var (day,qty) in series) line.Points.Add(new DataPoint(DateTimeAxis.ToDouble(day), qty));
            modelTrend.Series.Add(line);

            if(_cfg.ui?.showMovingAverage ?? true){
                var ma = Aggregations.MovingAverage(series.Select(x=> (double)x.qty).ToList(), 7);
                var maSeries = new LineSeries{ LineStyle=LineStyle.Dash, Title="MA7" };
                for(int i=0;i<series.Count;i++) maSeries.Points.Add(new DataPoint(DateTimeAxis.ToDouble(series[i].day), ma[i]));
                modelTrend.Series.Add(maSeries);
            }
            _plotTrend.Model = modelTrend;

            // TopN/全量
            int? topN = _chkTopN.Checked ? DEFAULT_TOPN : null;

            // 2) 尺码销量（降序）
            var sizeAgg = cleaned.GroupBy(x=>x.Size).Select(g=> new { Key=g.Key, Qty=g.Sum(z=>z.Qty)})
                                 .Where(a=>!string.IsNullOrWhiteSpace(a.Key) && a.Qty!=0)
                                 .OrderByDescending(a=>a.Qty)
                                 .ToList();
            if(topN.HasValue) sizeAgg = sizeAgg.Take(topN.Value).ToList();

            var modelSize = new PlotModel { Title = "尺码销量", PlotMargins = new OxyThickness(80,6,6,6) };
            var sizeCat = new CategoryAxis{ Position=AxisPosition.Left, GapWidth=0.4, StartPosition=1, EndPosition=0 };
            foreach(var a in sizeAgg) sizeCat.Labels.Add(a.Key);
            modelSize.Axes.Add(sizeCat);
            modelSize.Axes.Add(new LinearAxis{ Position=AxisPosition.Bottom, MinimumPadding=0, AbsoluteMinimum=0 });
            var bsSize = new BarSeries();
            foreach(var a in sizeAgg) bsSize.Items.Add(new BarItem{ Value=a.Qty });
            modelSize.Series.Add(bsSize);
            _plotSize.Model = modelSize;

            // 3) 颜色销量（降序）
            var colorAgg = cleaned.GroupBy(x=>x.Color).Select(g=> new { Key=g.Key, Qty=g.Sum(z=>z.Qty)})
                                  .Where(a=>!string.IsNullOrWhiteSpace(a.Key) && a.Qty!=0)
                                  .OrderByDescending(a=>a.Qty)
                                  .ToList();
            if(topN.HasValue) colorAgg = colorAgg.Take(topN.Value).ToList();

            var modelColor = new PlotModel { Title = "颜色销量", PlotMargins = new OxyThickness(80,6,6,6) };
            var colorCat = new CategoryAxis{ Position=AxisPosition.Left, GapWidth=0.4, StartPosition=1, EndPosition=0 };
            foreach(var a in colorAgg) colorCat.Labels.Add(a.Key);
            modelColor.Axes.Add(colorCat);
            modelColor.Axes.Add(new LinearAxis{ Position=AxisPosition.Bottom, MinimumPadding=0, AbsoluteMinimum=0 });
            var bsColor = new BarSeries();
            foreach(var a in colorAgg) bsColor.Items.Add(new BarItem{ Value=a.Qty });
            modelColor.Series.Add(bsColor);
            _plotColor.Model = modelColor;

            // 4) 分仓占比（A1 占位 + 合并“其他”逻辑保留；A2 用库存快照替换）
            RenderWarehousePiePlaceholder();
        }

        private void ApplyFilter(string q)
        {
            q = (q ?? "").Trim();
            if (string.IsNullOrWhiteSpace(q))
            {
                _binding.DataSource = new BindingList<object>(_gridMaster);
                _grid.ClearSelection();
                return;
            }
            var filtered = _gridMaster.Where(x=>{
                var t=x.GetType();
                string Get(string n)=> t.GetProperty(n)?.GetValue(x)?.ToString() ?? "";
                return Get("日期").Contains(q, StringComparison.OrdinalIgnoreCase)
                    || Get("款式").Contains(q, StringComparison.OrdinalIgnoreCase)
                    || Get("尺码").Contains(q, StringComparison.OrdinalIgnoreCase)
                    || Get("颜色").Contains(q, StringComparison.OrdinalIgnoreCase)
                    || Get("数量").Contains(q, StringComparison.OrdinalIgnoreCase);
            }).ToList();
            _binding.DataSource = new BindingList<object>(filtered);
            _grid.ClearSelection();
        }

        // —— 从图表点击带入过滤标签（展示/清空） —— //
        private void SetFilterChip(string key, string value, Action onRemove)
        {
            // 如果已存在相同 key 的 chip，先移除
            foreach (var c in _filterChips.Controls.OfType<Panel>().ToList())
            {
                if (c.Tag?.ToString()==key) { _filterChips.Controls.Remove(c); c.Dispose(); }
            }
            var p = new Panel { Height=26, Padding=new Padding(8,4,8,4), BackColor=Color.FromArgb(240,240,240), Margin=new Padding(0,3,6,3), Tag=key, AutoSize=true };
            var lbl = new Label { AutoSize=true, Text=$"{key}: {value}" };
            var btn = new Button { Text="×", AutoSize=true, Margin=new Padding(6,0,0,0) };
            btn.Click += (s,e)=>{ onRemove?.Invoke(); _filterChips.Controls.Remove(p); p.Dispose(); ApplyFilter(_boxSearch.Text); };
            p.Controls.Add(lbl); p.Controls.Add(btn);
            _filterChips.Controls.Add(p);
        }

        private void ExportExcel()
        {
            var saveDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "exports");
            Directory.CreateDirectory(saveDir);
            var path = Path.Combine(saveDir, $"StyleWatcher_{DateTime.Now:yyyyMMdd_HHmm}.xlsx");

            using var wb = new XLWorkbook();
            // 明细
            var ws1 = wb.AddWorksheet("销售明细");
            ws1.Cell(1,1).Value="日期"; ws1.Cell(1,2).Value="款式"; ws1.Cell(1,3).Value="颜色"; ws1.Cell(1,4).Value="尺码"; ws1.Cell(1,5).Value="数量";
            var list=_gridMaster;
            int r=2;
            foreach(var it in list){
                var t=it.GetType();
                ws1.Cell(r,1).Value=t.GetProperty("日期")?.GetValue(it)?.ToString();
                ws1.Cell(r,2).Value=t.GetProperty("款式")?.GetValue(it)?.ToString();
                ws1.Cell(r,3).Value=t.GetProperty("颜色")?.GetValue(it)?.ToString();
                ws1.Cell(r,4).Value=t.GetProperty("尺码")?.GetValue(it)?.ToString();
                ws1.Cell(r,5).Value=t.GetProperty("数量")?.GetValue(it)?.ToString();
                r++;
            }
            ws1.Columns().AdjustToContents();

            // 趋势
            var ws2 = wb.AddWorksheet("趋势");
            ws2.Cell(1,1).Value="日期"; ws2.Cell(1,2).Value="数量"; ws2.Cell(1,3).Value="MA7";
            var series = Aggregations.BuildDateSeries(_sales,_trendWindow);
            var ma = Aggregations.MovingAverage(series.Select(x=> (double)x.qty).ToList(), 7);
            int rr=2;
            for(int i=0;i<series.Count;i++){
                ws2.Cell(rr,1).Value=series[i].day.ToString("yyyy-MM-dd");
                ws2.Cell(rr,2).Value=series[i].qty;
                ws2.Cell(rr,3).Value=ma[i];
                rr++;
            }
            ws2.Columns().AdjustToContents();

            // 口径说明
            var ws3 = wb.AddWorksheet("口径说明");
            ws3.Cell(1,1).Value="趋势窗口（天）"; ws3.Cell(1,2).Value=_trendWindow;
            ws3.Cell(2,1).Value="是否显示MA7"; ws3.Cell(2,2).Value=(_cfg.ui?.showMovingAverage ?? true) ? "是" : "否";
            ws3.Cell(3,1).Value="TopN是否启用"; ws3.Cell(3,2).Value=_chkTopN.Checked ? $"是（Top {DEFAULT_TOPN}）" : "否（全量）";
            ws3.Columns().AdjustToContents();

            wb.SaveAs(path);
            try{ System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{path}\""); }catch{}
        }
    }
}
