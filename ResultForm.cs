using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
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
        public static readonly Font Title = new("Microsoft YaHei UI", 13, FontStyle.Bold);
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

        // Detail
        private readonly DataGridView _grid = new();
        private readonly BindingSource _binding = new();
        private readonly TextBox _boxSearch = new();

        // Inventory page
        private InventoryTabPage _invPage;

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
            Width = Math.Max(1100, _cfg.window.width);
            Height = Math.Max(720, _cfg.window.height);
            StartPosition = FormStartPosition.CenterScreen;
            TopMost = _cfg.window.alwaysOnTop;
            BackColor = Color.White;
            KeyPreview = true;
            KeyDown += (s,e)=>{ if(e.KeyCode==Keys.Escape) Hide(); };

            var root = new TableLayoutPanel{Dock=DockStyle.Fill,RowCount=2,ColumnCount=1};
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 72));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Controls.Add(root);

            var header = BuildHeader();
            root.Controls.Add(header,0,0);

            var content = new TableLayoutPanel{Dock=DockStyle.Fill,RowCount=2,ColumnCount=1};
            content.RowStyles.Add(new RowStyle(SizeType.Absolute, 118));
            content.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.Controls.Add(content,0,1);

            _kpi.Dock = DockStyle.Fill;
            _kpi.FlowDirection = FlowDirection.LeftToRight;
            _kpi.WrapContents = true;
            _kpi.Padding = new Padding(12,8,12,8);
            _kpi.Controls.Add(MakeKpi(_kpiSales7,"近7日销量","—"));
            _kpi.Controls.Add(MakeKpi(_kpiInv,"可用库存(总)","—"));
            _kpi.Controls.Add(MakeKpi(_kpiDoc,"库存天数(DoC)","—"));
            _kpi.Controls.Add(MakeKpi(_kpiMissing,"缺尺码数","—"));
            content.Controls.Add(_kpi,0,0);

            _tabs.Dock = DockStyle.Fill;
            BuildTabs();
            content.Controls.Add(_tabs,0,1);
        }

        private Control BuildHeader()
        {
            var head = new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=3,RowCount=1,Padding=new(12,10,12,6),BackColor=UI.HeaderBack};
            head.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
            head.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            head.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            _input.MinimumSize = new Size(360,30);
            _btnQuery.Text="重新查询";
            _btnQuery.AutoSize=true; _btnQuery.Padding=new Padding(10,6,10,6);
            _btnQuery.Click += async (s,e)=>{ _btnQuery.Enabled=false; try{ await ReloadAsync(); } finally{ _btnQuery.Enabled=true; } };

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
            host.Width=250; host.Height=98; host.Padding=new Padding(10);
            host.BackColor=UI.CardBack; host.BorderStyle=BorderStyle.FixedSingle;
            var t=new Label{Text=title,Dock=DockStyle.Top,Height=22,ForeColor=UI.Text};
            var v=new Label{Text=value,Dock=DockStyle.Fill,Font=new Font(Font,FontStyle.Bold)};
            host.Controls.Add(v); host.Controls.Add(t);
            return host;
        }

        private void BuildTabs()
        {
            // 概览
            var overview = new TabPage("概览"){BackColor=Color.White};
            var layout = new TableLayoutPanel{Dock=DockStyle.Fill,RowCount=5,ColumnCount=1,Padding=new Padding(12)};
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,44));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent,40));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent,20));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent,20));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent,20));

            _trendSwitch.Dock=DockStyle.Fill;
            _trendSwitch.FlowDirection=FlowDirection.LeftToRight;
            _trendSwitch.WrapContents=false;
            _trendSwitch.Padding=new Padding(6,8,6,0);
            var wins=_cfg.ui?.trendWindows??new[]{7,14,30};
            foreach(var w in wins.Distinct().OrderBy(x=>x)){
                var rb=new RadioButton{Text=$"{w} 日",AutoSize=true,Tag=w,Margin=new Padding(0,6,18,0)};
                if(w==_trendWindow) rb.Checked=true;
                rb.CheckedChanged+=(s,e)=>{ var me=(RadioButton)s; if(me.Checked){ _trendWindow=(int)me.Tag; if(_sales.Count>0) RenderCharts(_sales); } };
                _trendSwitch.Controls.Add(rb);
            }
            layout.Controls.Add(_trendSwitch,0,0);

            _plotTrend.Dock=DockStyle.Fill;
            _plotSize.Dock=DockStyle.Fill;
            _plotColor.Dock=DockStyle.Fill;
            _plotWarehouse.Dock=DockStyle.Fill;
            layout.Controls.Add(_plotTrend,0,1);
            layout.Controls.Add(_plotSize,0,2);
            layout.Controls.Add(_plotColor,0,3);
            layout.Controls.Add(_plotWarehouse,0,4);
            overview.Controls.Add(layout);
            _tabs.TabPages.Add(overview);

            // 销售明细
            var detail = new TabPage("销售明细"){BackColor=Color.White};
            var panel = new TableLayoutPanel{Dock=DockStyle.Fill,RowCount=2,ColumnCount=1,Padding=new Padding(12)};
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute,38));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent,100));
            _boxSearch.Dock=DockStyle.Fill; _boxSearch.PlaceholderText="搜索（日期/款式/尺码/颜色/数量）";
            _boxSearch.TextChanged += (s,e)=> ApplyFilter(_boxSearch.Text);
            panel.Controls.Add(_boxSearch,0,0);
            _grid.Dock=DockStyle.Fill; _grid.ReadOnly=true; _grid.AllowUserToAddRows=false; _grid.AllowUserToDeleteRows=false;
            _grid.RowHeadersVisible=false; _grid.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.AllCells;
            _grid.DataSource=_binding;
            panel.Controls.Add(_grid,0,1);
            detail.Controls.Add(panel);
            _tabs.TabPages.Add(detail);

            // 库存
            _invPage = new InventoryTabPage(_cfg);
            _invPage.SummaryUpdated += OnInventorySummaryUpdated;
            _tabs.TabPages.Add(_invPage);
        }

        private void OnInventorySummaryUpdated()
        {
            var snap=_invPage.GetSummary();
            if(snap==null) return;

            // 可用总库存
            SetKpiValue(_kpiInv, snap.总可用.ToString());
            // DoC
            var sales7 = _sales.Where(x=>x.Date>=DateTime.Today.AddDays(-6)).Sum(x=>x.Qty);
            var avg7 = Math.Max(0.01, sales7/7.0);
            var doc = Aggregations.DaysOfCover(snap.总可用, avg7);
            SetKpiValue(_kpiDoc, doc==int.MaxValue? "—" : doc.ToString());

            // 分仓占比
            var model=new PlotModel{Title="分仓库存占比（可用）"};
            var pie=new PieSeries{AngleSpan=360,StartAngle=0,StrokeThickness=0.5,InsideLabelPosition=0.6};
            foreach(var kv in snap.分仓可用.OrderByDescending(k=>k.Value))
                pie.Slices.Add(new PieSlice(kv.Key, kv.Value));
            model.Series.Add(pie);
            _plotWarehouse.Model=model;
        }

        private void SetKpiValue(Panel p,string value)
        {
            var val = p.Controls.OfType<Label>().FirstOrDefault(l=>l.Dock==DockStyle.Fill);
            if(val!=null) val.Text=value;
        }

        // ====== Public methods used by TrayApp ======
        public void FocusInput(){ try{ if(WindowState==FormWindowState.Minimized) WindowState=FormWindowState.Normal; _input.Focus(); _input.SelectAll(); }catch{} }
        public void ShowNoActivateAtCursor(){ try{ StartPosition=FormStartPosition.Manual; var pt=Cursor.Position; Location=new Point(Math.Max(0,pt.X-Width/2),Math.Max(0,pt.Y-Height/2)); Show(); }catch{ Show(); } }
        public void ShowAndFocusCentered(){ ShowAndFocusCentered(_cfg.window.alwaysOnTop); }
        public void ShowAndFocusCentered(bool alwaysOnTop){ TopMost=alwaysOnTop; StartPosition=FormStartPosition.CenterScreen; Show(); Activate(); FocusInput(); }
        public void SetLoading(string message){ SetKpiValue(_kpiSales7,"—"); SetKpiValue(_kpiInv,"—"); SetKpiValue(_kpiDoc,"—"); SetKpiValue(_kpiMissing,"—"); }

        public async void ApplyRawText(string selection, string parsed){ _input.Text=selection??string.Empty; await LoadTextAsync(parsed??string.Empty); }
        public void ApplyRawText(string text){ _input.Text=text??string.Empty; }

        // ====== Load / Render ======
        public async System.Threading.Tasks.Task LoadTextAsync(string raw)=>await ReloadAsync(raw);
        private async System.Threading.Tasks.Task ReloadAsync()=>await ReloadAsync(_input.Text);

        private async System.Threading.Tasks.Task ReloadAsync(string displayText)
        {
            await System.Threading.Tasks.Task.Yield();
            if (string.IsNullOrWhiteSpace(displayText))
                displayText = _lastDisplayText;

            // 解析
            var parsed = Parser.Parse(displayText ?? string.Empty); // 兼容层包装到 PayloadParser
            var newSales = parsed.Records.Select(r=> new Aggregations.SalesItem{ Date=r.Date, Size=r.Size??"", Color=r.Color??"", Qty=r.Qty }).ToList();
            var newGrid = parsed.Records.Select(r => (object)new { 日期=r.Date.ToString("yyyy-MM-dd"), 款式=r.Name, 尺码=r.Size, 颜色=r.Color, 数量=r.Qty }).ToList();

            // 成功后替换缓存
            _lastDisplayText = displayText;
            _sales = newSales;
            _gridMaster = newGrid;

            // KPI
            var sales7 = _sales.Where(x=>x.Date>=DateTime.Today.AddDays(-6)).Sum(x=>x.Qty);
            SetKpiValue(_kpiSales7, sales7.ToString());
            SetKpiValue(_kpiMissing, Aggregations.MissingSizes(_sales.Select(s=>s.Size)).ToString());

            // 渲染
            RenderCharts(_sales);

            _binding.DataSource = new BindingList<object>(_gridMaster);
            _grid.ClearSelection();
        }

        private void RenderCharts(List<Aggregations.SalesItem> salesItems)
        {
            // 趋势
            var series = Aggregations.BuildDateSeries(salesItems, _trendWindow);
            var modelTrend = new PlotModel { Title = $"近 {_trendWindow} 日总销量趋势", PlotMargins = new OxyThickness(50,10,10,40) };
            var xAxis = new DateTimeAxis{ Position=AxisPosition.Bottom, StringFormat="MM-dd", IntervalType=DateTimeIntervalType.Days, MajorStep=1, MinorStep=1, IntervalLength=60, IsZoomEnabled=false, IsPanEnabled=false, MajorGridlineStyle=LineStyle.Solid };
            var yAxis = new LinearAxis{ Position=AxisPosition.Left, MinimumPadding=0, AbsoluteMinimum=0, MajorGridlineStyle=LineStyle.Solid };
            modelTrend.Axes.Add(xAxis); modelTrend.Axes.Add(yAxis);
            var line = new LineSeries{ MarkerType=MarkerType.Circle };
            foreach(var (day,qty) in series) line.Points.Add(new DataPoint(DateTimeAxis.ToDouble(day), qty));
            modelTrend.Series.Add(line);
            if(_cfg.ui?.showMovingAverage ?? true){
                var ma = Aggregations.MovingAverage(series.Select(x=>x.qty).ToList(), 7);
                var maSeries = new LineSeries{ LineStyle=LineStyle.Dash, Title="MA7" };
                for(int i=0;i<series.Count;i++) maSeries.Points.Add(new DataPoint(DateTimeAxis.ToDouble(series[i].day), ma[i]));
                modelTrend.Series.Add(maSeries);
            }
            _plotTrend.Model = modelTrend;

            // 尺码（全量降序）
            var sizeAgg = Aggregations.BySize(salesItems);
            var modelSize = new PlotModel { Title = "各尺码销量（降序，全部）", PlotMargins = new OxyThickness(80,6,6,6) };
            modelSize.Axes.Add(new CategoryAxis{ Position=AxisPosition.Left, ItemsSource=sizeAgg, LabelField="Key", GapWidth=0.4 });
            modelSize.Axes.Add(new LinearAxis{ Position=AxisPosition.Bottom, MinimumPadding=0, AbsoluteMinimum=0 });
            modelSize.Series.Add(new BarSeries{ ItemsSource=sizeAgg.Select(x=> new BarItem{ Value=x.Qty }) });
            _plotSize.Model = modelSize;

            // 颜色（全量降序）
            var colorAgg = Aggregations.ByColor(salesItems);
            var modelColor = new PlotModel { Title = "各颜色销量（降序，全部）", PlotMargins = new OxyThickness(80,6,6,6) };
            modelColor.Axes.Add(new CategoryAxis{ Position=AxisPosition.Left, ItemsSource=colorAgg, LabelField="Key", GapWidth=0.4 });
            modelColor.Axes.Add(new LinearAxis{ Position=AxisPosition.Bottom, MinimumPadding=0, AbsoluteMinimum=0 });
            modelColor.Series.Add(new BarSeries{ ItemsSource=colorAgg.Select(x=> new BarItem{ Value=x.Qty }) });
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
            var ws1 = wb.AddWorksheet("销售明细");
            ws1.Cell(1,1).Value="日期"; ws1.Cell(1,2).Value="款式"; ws1.Cell(1,3).Value="尺码"; ws1.Cell(1,4).Value="颜色"; ws1.Cell(1,5).Value="数量";
            var list=_gridMaster;
            int r=2;
            foreach(var it in list){
                var t=it.GetType();
                ws1.Cell(r,1).Value=t.GetProperty("日期")?.GetValue(it)?.ToString();
                ws1.Cell(r,2).Value=t.GetProperty("款式")?.GetValue(it)?.ToString();
                ws1.Cell(r,3).Value=t.GetProperty("尺码")?.GetValue(it)?.ToString();
                ws1.Cell(r,4).Value=t.GetProperty("颜色")?.GetValue(it)?.ToString();
                ws1.Cell(r,5).Value=t.GetProperty("数量")?.GetValue(it)?.ToString();
                r++;
            }
            ws1.Columns().AdjustToContents();

            var ws2 = wb.AddWorksheet("趋势");
            ws2.Cell(1,1).Value="日期"; ws2.Cell(1,2).Value="数量";
            var series = Aggregations.BuildDateSeries(_sales,_trendWindow);
            int rr=2;
            foreach(var it in series){ ws2.Cell(rr,1).Value=it.day.ToString("yyyy-MM-dd"); ws2.Cell(rr,2).Value=it.qty; rr++; }
            ws2.Columns().AdjustToContents();

            wb.SaveAs(path);
            try{ System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{path}\""); }catch{}
        }
    }
}
