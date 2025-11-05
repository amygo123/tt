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
        public static readonly Color ChipBack   = Color.FromArgb(235, 238, 244);
        public static readonly Color ChipBorder = Color.FromArgb(210, 214, 222);
        public static readonly Color Red        = Color.FromArgb(215, 58, 73);
        public static readonly Color Yellow     = Color.FromArgb(216, 160, 18);
        public static readonly Color Green      = Color.FromArgb(26, 127, 55);
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
        private FlowLayoutPanel? _kpiMissingFlow;

        // Tabs
        private readonly TabControl _tabs = new();

        // Overview
        private readonly FlowLayoutPanel _trendSwitch = new();
        private int _trendWindow = 7;
        private readonly PlotView _plotTrend = new();
        private readonly PlotView _plotSize = new();
        private readonly PlotView _plotColor = new();
        private readonly PlotView _plotWarehouse = new();

        // Detail
        private readonly DataGridView _grid = new();
        private readonly BindingSource _binding = new();
        private readonly TextBox _boxSearch = new();
        private readonly FlowLayoutPanel _filterChips = new();
        private readonly System.Windows.Forms.Timer _searchDebounce = new System.Windows.Forms.Timer() { Interval = 200 };

        // Inventory page
        private InventoryTabPage? _invPage;

        // Caches
        private string _lastDisplayText = string.Empty;
        private List<Aggregations.SalesItem> _sales = new();
        private List<object> _gridMaster = new();

        // cached inventory totals from event
        private int _invAvailTotal = 0;
        private int _invOnHandTotal = 0;
        private Dictionary<string,int> _invWarehouse = new Dictionary<string,int>();

        public ResultForm(AppConfig cfg)
        {
            _cfg = cfg;

            Text = "StyleWatcher";
            Font = new Font("Microsoft YaHei UI", _cfg.window.fontSize);
            Width = Math.Max(1200, _cfg.window.width);
            Height = Math.Max(800, _cfg.window.height);
            StartPosition = FormStartPosition.CenterScreen;
            TopMost = _cfg.window.alwaysOnTop;
            BackColor = Color.White;
            KeyPreview = true;
            KeyDown += (s,e)=>{ if(e.KeyCode==Keys.Escape) Hide(); };
            // 设置窗口图标：优先 EXE 内置图标，其次 Resources\app.ico
            try
            {
                var exeIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                if (exeIcon != null) this.Icon = exeIcon;
                else
                {
                    var icoPath = System.IO.Path.Combine(AppContext.BaseDirectory, "Resources", "app.ico");
                    if (System.IO.File.Exists(icoPath)) this.Icon = new Icon(icoPath);
                }
            }
            catch { /* ignore */ }


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
            _kpi.Controls.Add(MakeKpiMissing(_kpiMissing,"缺货尺码"));
            content.Controls.Add(_kpi,0,0);

            _tabs.Dock = DockStyle.Fill;
            BuildTabs();
            content.Controls.Add(_tabs,0,1);

            _searchDebounce.Tick += (s,e)=> { _searchDebounce.Stop(); ApplyFilter(_boxSearch.Text); };
        }

        private Control BuildHeader()
        {
            var head = new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=3,RowCount=1,Padding=new(12,10,12,6),BackColor=Color.FromArgb(245,247,250)};
            head.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
            head.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            head.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            _input.MinimumSize = new Size(420,32);
            _input.Height = 30;

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
            host.BackColor=Color.FromArgb(250,250,250); host.BorderStyle=BorderStyle.FixedSingle;
            host.Margin = new Padding(8,4,8,4);

            var inner = new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=2};
            inner.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            inner.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var t=new Label{Text=title,Dock=DockStyle.Fill,Height=28,ForeColor=Color.FromArgb(47,47,47),Font=new Font("Microsoft YaHei UI", 10),TextAlign=ContentAlignment.MiddleLeft};
            var v=new Label{Text=value,Dock=DockStyle.Fill,Font=new Font("Microsoft YaHei UI", 16, FontStyle.Bold),TextAlign=ContentAlignment.MiddleLeft,Padding=new Padding(0,2,0,0)};
            v.Name = "ValueLabel";

            inner.Controls.Add(t,0,0);
            inner.Controls.Add(v,0,1);
            host.Controls.Clear();
            host.Controls.Add(inner);
            return host;
        }

        private Control MakeKpiMissing(Panel host, string title)
        {
            host.Width=260; host.Height=110; host.Padding=new Padding(10);
            host.BackColor=Color.FromArgb(250,250,250); host.BorderStyle=BorderStyle.FixedSingle;
            host.Margin = new Padding(8,4,8,4);

            var inner = new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=2};
            inner.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            inner.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var t=new Label{Text=title,Dock=DockStyle.Fill,Height=28,ForeColor=Color.FromArgb(47,47,47),Font=new Font("Microsoft YaHei UI", 10),TextAlign=ContentAlignment.MiddleLeft};

            var flow = new FlowLayoutPanel{
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                AutoScroll = true,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            _kpiMissingFlow = flow;

            inner.Controls.Add(t,0,0);
            inner.Controls.Add(flow,0,1);
            host.Controls.Clear();
            host.Controls.Add(inner);
            return host;
        }

        private void SetMissingSizes(IEnumerable<string> sizes)
        {
            if (_kpiMissingFlow == null) return;
            _kpiMissingFlow.SuspendLayout();
            _kpiMissingFlow.Controls.Clear();

            foreach (var s in sizes)
            {
                var chip = new Label
                {
                    AutoSize = true,
                    Text = s,
                    Font = new Font("Microsoft YaHei UI", 8.5f, FontStyle.Regular),
                    BackColor = Color.FromArgb(235, 238, 244),
                    ForeColor = Color.FromArgb(47,47,47),
                    Padding = new Padding(6, 2, 6, 2),
                    Margin = new Padding(4, 2, 0, 2),
                    BorderStyle = BorderStyle.FixedSingle
                };
                _kpiMissingFlow.Controls.Add(chip);
            }

            if (_kpiMissingFlow.Controls.Count == 0)
            {
                var none = new Label
                {
                    AutoSize = true,
                    Text = "无",
                    Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Regular),
                    ForeColor = Color.FromArgb(47,47,47)
                };
                _kpiMissingFlow.Controls.Add(none);
            }

            _kpiMissingFlow.ResumeLayout();
        }

        private void BuildTabs()
        {
            // 概览
            var overview = new TabPage("概览"){BackColor=Color.White};

            var container = new TableLayoutPanel{Dock=DockStyle.Fill,RowCount=2,ColumnCount=1};
            container.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));
            container.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            // 顶部工具区：从配置读取 trendWindows
            var tools = new FlowLayoutPanel{Dock=DockStyle.Fill, FlowDirection=FlowDirection.LeftToRight, WrapContents=false, AutoScroll=true, Padding=new Padding(12,10,12,0)};
            _trendSwitch.FlowDirection=FlowDirection.LeftToRight;
            _trendSwitch.WrapContents=false;
            _trendSwitch.AutoSize=true;
            _trendSwitch.Padding=new Padding(0,0,12,0);

            var wins = (_cfg.ui?.trendWindows ?? new int[]{7,14,30})
                        .Where(x=>x>0 && x<=90).Distinct().OrderBy(x=>x).ToArray(); // 基本的容错
            if (!wins.Contains(_trendWindow)) _trendWindow = wins.FirstOrDefault(7);

            foreach(var w in wins)
            {
                var rb=new RadioButton{Text=$"{w} 日",AutoSize=true,Tag=w,Margin=new Padding(0,2,18,0)};
                if(w==_trendWindow) rb.Checked=true;
                rb.CheckedChanged += (s, e) => { var rbCtrl = s as RadioButton; if (rbCtrl == null || !rbCtrl.Checked) return; if (rbCtrl.Tag is int w2) { _trendWindow = w2; if (_sales != null && _sales.Count > 0) RenderCharts(_sales); } };
                _trendSwitch.Controls.Add(rb);
            }
            tools.Controls.Add(_trendSwitch);

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

            // 库存页
            _invPage = new InventoryTabPage(_cfg);
            _invPage.SummaryUpdated += OnInventorySummary;
            _tabs.TabPages.Add(_invPage);
        }

        private static Label? ValueLabelOf(Panel p)
        {
            var table = p.Controls.OfType<TableLayoutPanel>().FirstOrDefault();
            if (table != null)
            {
                var labels = table.Controls.OfType<Label>().ToList();
                if (labels.Count > 0) return labels.Last();
            }
            return null;
        }

        private void SetKpiValue(Panel p,string value, Color? color = null)
        {
            var lbl = ValueLabelOf(p);
            if (lbl == null) return;
            lbl.Text = value ?? "—";
            if (color.HasValue) lbl.ForeColor = color.Value;
            else lbl.ForeColor = Color.FromArgb(47,47,47);
        }

        // —— 与 TrayApp 对齐的接口 —— //
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

            // KPI: 近 N 天销量（固定用近 7 天）
            var sales7 = _sales.Where(x=>x.Date>=DateTime.Today.AddDays(-6)).Sum(x=>x.Qty);
            SetKpiValue(_kpiSales7, sales7.ToString());

            // 缺失尺码 chips（按销售基线）
            SetMissingSizes(MissingSizes(_sales.Select(s=>s.Size)));

            RenderCharts(_sales);

            _binding.DataSource = new BindingList<object>(_gridMaster);
            _grid.ClearSelection();
            if (_grid.Columns.Contains("款式")) _grid.Columns["款式"].DisplayIndex = 0;
            if (_grid.Columns.Contains("颜色")) _grid.Columns["颜色"].DisplayIndex = 1;
            if (_grid.Columns.Contains("尺码")) _grid.Columns["尺码"].DisplayIndex = 2;
            if (_grid.Columns.Contains("日期")) _grid.Columns["日期"].DisplayIndex = 3;
            if (_grid.Columns.Contains("数量")) _grid.Columns["数量"].DisplayIndex = 4;

            // 推断 styleName，若无则使用 default_style 兜底
            var styleName = parsed.Records
                .Select(r => r.Name)
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .GroupBy(n => n)
                .OrderByDescending(g => g.Count())
                .FirstOrDefault()
                ?.Key;

            if (string.IsNullOrWhiteSpace(styleName))
                styleName = _cfg.inventory?.default_style ?? "";

            if (!string.IsNullOrWhiteSpace(styleName))
            {
                try { _ = _invPage?.LoadInventoryAsync(styleName); } catch {}
            }
        }

        private void OnInventorySummary(int totalAvail, int totalOnHand, Dictionary<string,int> warehouseAgg)
        {
            _invAvailTotal = totalAvail;
            _invOnHandTotal = totalOnHand;
            _invWarehouse = warehouseAgg ?? new Dictionary<string,int>();

            SetKpiValue(_kpiInv, totalAvail.ToString());

            // 库存天数：用最近 N 天平均销量（来自配置 inventoryAlert.minSalesWindowDays）
            var baseDays = Math.Max(1, _cfg.inventoryAlert?.minSalesWindowDays ?? 7);
            var lastN = _sales.Where(x=> x.Date >= DateTime.Today.AddDays(-(baseDays-1))).Sum(x=>x.Qty);
            var avg = lastN / (double)baseDays;

            string daysText = "—";
            System.Drawing.Color? daysColor = null;
            if (avg > 0)
            {
                var d = Math.Round(totalAvail / avg, 1);
                daysText = d.ToString("0.0");

                var red = _cfg.inventoryAlert?.docRed ?? 3;
                var yellow = _cfg.inventoryAlert?.docYellow ?? 7;
                daysColor = d < red ? Color.FromArgb(215, 58, 73) : (d < yellow ? Color.FromArgb(216, 160, 18) : Color.FromArgb(26, 127, 55));
            }
            SetKpiValue(_kpiDoc, daysText, daysColor);

            // 概览页：分仓占比
            RenderWarehousePieOverview(_invWarehouse);
        }

        private void RenderWarehousePieOverview(Dictionary<string,int> warehouseAgg)
        {
            var model = new PlotModel { Title = "分仓库存占比" };
            var pie = new PieSeries { AngleSpan = 360, StartAngle = 0, StrokeThickness = 0.5, InsideLabelPosition = 0.6, InsideLabelFormat = "{0}" };

            var list = warehouseAgg.Where(kv => !string.IsNullOrWhiteSpace(kv.Key) && kv.Value > 0)
                                   .OrderByDescending(kv => kv.Value).ToList();
            var total = list.Sum(x => (double)x.Value);

            if (total <= 0)
            {
                pie.Slices.Add(new PieSlice("无数据", 1));
            }
            else
            {
                var keep = list.Where(kv => kv.Value / total >= 0.03).ToList();
                if (keep.Count < 3)
                    keep = list.Take(3).ToList();

                var keepSet = new HashSet<string>(keep.Select(k => k.Key));
                double other = 0;
                foreach (var kv in list)
                {
                    if (keepSet.Contains(kv.Key))
                        pie.Slices.Add(new PieSlice(kv.Key, kv.Value));
                    else
                        other += kv.Value;
                }
                if (other > 0) pie.Slices.Add(new PieSlice("其他", other));
            }

            model.Series.Add(pie);
            _plotWarehouse.Model = model;
            _plotWarehouse.MouseUp -= OnWarehousePieMouseUp;
            _plotWarehouse.MouseUp += OnWarehousePieMouseUp;
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

            // 1) 趋势
            var series = Aggregations.BuildDateSeries(cleaned, _trendWindow);
            var modelTrend = new PlotModel { Title = $"近{_trendWindow}日销量趋势", PlotMargins = new OxyThickness(50,10,10,40) };

            var xAxis = new DateTimeAxis{ Position=AxisPosition.Bottom, StringFormat="MM-dd", IntervalType=DateTimeIntervalType.Days, MajorStep=1, MinorStep=1, IntervalLength=60, IsZoomEnabled=false, IsPanEnabled=false, MajorGridlineStyle=LineStyle.Solid };
            var yAxis = new LinearAxis{ Position=AxisPosition.Left, MinimumPadding=0, AbsoluteMinimum=0, MajorGridlineStyle=LineStyle.Solid };
            modelTrend.Axes.Add(xAxis); modelTrend.Axes.Add(yAxis);

            var line = new LineSeries{ Title="销量", MarkerType=MarkerType.Circle };
            foreach(var (day,qty) in series) line.Points.Add(new DataPoint(DateTimeAxis.ToDouble(day), qty));
            modelTrend.Series.Add(line);

            if (false) // disabled MA7 on overview
            {
                var ma = Aggregations.MovingAverage(series.Select(x=> (double)x.qty).ToList(), 7);
                var maSeries = new LineSeries{ LineStyle=LineStyle.Dash, Title="MA7" };
                for(int i=0;i<series.Count;i++) maSeries.Points.Add(new DataPoint(DateTimeAxis.ToDouble(series[i].day), ma[i]));
                modelTrend.Series.Add(maSeries);
            }
            _plotTrend.Model = modelTrend;

            // 2) 尺码销量（降序）
            var sizeAgg = cleaned.GroupBy(x=>x.Size).Select(g=> new { Key=g.Key, Qty=g.Sum(z=>z.Qty)})
                                 .Where(a=>!string.IsNullOrWhiteSpace(a.Key) && a.Qty!=0)
                                 .OrderByDescending(a=>a.Qty).ToList();

            var modelSize = new PlotModel { Title = "尺码销量", PlotMargins = new OxyThickness(80,6,6,6) };
            var sizeCat = new CategoryAxis{ Position=AxisPosition.Left, GapWidth=0.4, StartPosition=1, EndPosition=0 };
            foreach(var a in sizeAgg) sizeCat.Labels.Add(a.Key);
            modelSize.Axes.Add(sizeCat);
            modelSize.Axes.Add(new LinearAxis{ Position = AxisPosition.Bottom, MinimumPadding = 0, AbsoluteMinimum = 0 });
            var bsSize = new BarSeries();
            foreach(var a in sizeAgg) bsSize.Items.Add(new BarItem{ Value=a.Qty });
            modelSize.Series.Add(bsSize);
            _plotSize.Model = modelSize;

            // 3) 颜色销量（降序）
            var colorAgg = cleaned.GroupBy(x=>x.Color).Select(g=> new { Key=g.Key, Qty=g.Sum(z=>z.Qty)})
                                  .Where(a=>!string.IsNullOrWhiteSpace(a.Key) && a.Qty!=0)
                                  .OrderByDescending(a=>a.Qty).ToList();

            var modelColor = new PlotModel { Title = "颜色销量", PlotMargins = new OxyThickness(80,6,6,6) };
            var colorCat = new CategoryAxis{ Position=AxisPosition.Left, GapWidth=0.4, StartPosition=1, EndPosition=0 };
            foreach(var a in colorAgg) colorCat.Labels.Add(a.Key);
            modelColor.Axes.Add(colorCat);
            modelColor.Axes.Add(new LinearAxis{ Position = AxisPosition.Bottom, MinimumPadding=0, AbsoluteMinimum=0 });
            var bsColor = new BarSeries();
            foreach(var a in colorAgg) bsColor.Items.Add(new BarItem{ Value=a.Qty });
            modelColor.Series.Add(bsColor);
            _plotColor.Model = modelColor;
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
            ws2.Cell(1,1).Value="日期"; ws2.Cell(1,2).Value="数量"; ws2.Cell(1,3).Value="MA7(若显示)";
            var series = Aggregations.BuildDateSeries(_sales,_trendWindow);
            var ma = Aggregations.MovingAverage(series.Select(x=> (double)x.qty).ToList(), 7);
            int rr=2;
            for(int i=0;i<series.Count;i++){
                ws2.Cell(rr,1).Value=series[i].day.ToString("yyyy-MM-dd");
                ws2.Cell(rr,2).Value=series[i].qty;
                ws2.Cell(rr,3).Value=(_cfg.ui?.showMovingAverage ?? true) ? ma[i] : 0;
                rr++;
            }
            ws2.Columns().AdjustToContents();

            // 口径说明
            var ws3 = wb.AddWorksheet("口径说明");
            ws3.Cell(1,1).Value="趋势窗口（天）"; ws3.Cell(1,2).Value=_trendWindow;
            ws3.Cell(2,1).Value="是否显示MA7"; ws3.Cell(2,2).Value=(_cfg.ui?.showMovingAverage ?? true) ? "是" : "否";
            ws3.Cell(3,1).Value="库存天数阈值"; ws3.Cell(3,2).Value=$"红<{_cfg.inventoryAlert?.docRed ?? 3}，黄<{_cfg.inventoryAlert?.docYellow ?? 7}";
            ws3.Cell(4,1).Value="销量基线天数"; ws3.Cell(4,2).Value=_cfg.inventoryAlert?.minSalesWindowDays ?? 7;
            ws3.Columns().AdjustToContents();

            wb.SaveAs(path);
            try{ System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{path}\""); }catch{}
        }
        // 概览页“分仓库存占比”饼图：点击后跳转库存并定位仓库
        private void OnWarehousePieMouseUp(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            try
            {
                var model = _plotWarehouse?.Model;
                if (model == null) return;
                var pie = model.Series.OfType<OxyPlot.Series.PieSeries>().FirstOrDefault();
                if (pie == null) return;

                var hit = pie.GetNearestPoint(new OxyPlot.ScreenPoint(e.Location.X, e.Location.Y), false);
                if (hit?.Item is OxyPlot.Series.PieSlice slice)
                {
                    var name = slice.Label ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(name)) return;
                    if (name == "其他" || name == "无数据") return;

                    // 切换到库存页 Tab
                    try
                    {
                        var invTab = _invPage?.Parent as TabPage;
                        var tabs = invTab?.Parent as TabControl;
                        if (invTab != null && tabs != null)
                            tabs.SelectedTab = invTab;
                    }
                    catch { /* ignore */ }

                    // 激活对应仓库子页
                    try { _invPage?.ActivateWarehouse(name); } catch { }
                }
            }
            catch { /* ignore */ }
        }
    
    }
}
