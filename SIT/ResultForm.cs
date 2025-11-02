using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OxyPlot;
using OxyPlot.Series;
using OxyPlot.WindowsForms;
using OxyPlot.Axes;

namespace StyleWatcherWin
{
    public class ResultForm : Form
    {
        private readonly TextBox _txt;
        private readonly PlotView _plot;
        private readonly Button _btnCopy;
        private readonly Button _btnClose;

        public ResultForm(ParseResult result, int fontSize = 13, bool topMost = true, int w = 640, int h = 420)
        {
            Text = "销量结果";
            StartPosition = FormStartPosition.CenterScreen;
            TopMost = topMost;
            Width = Math.Max(400, w);
            Height = Math.Max(300, h);

            _txt = new TextBox { Multiline = true, Dock = DockStyle.Top, Height = Height/2, ScrollBars = ScrollBars.Vertical, Font = new Font("Segoe UI", fontSize) };
            _plot = new PlotView { Dock = DockStyle.Fill };
            _btnCopy = new Button { Text = "复制文本", Dock = DockStyle.Left, Width = 100 };
            _btnClose = new Button { Text = "关闭", Dock = DockStyle.Right, Width = 100 };

            var panel = new Panel { Dock = DockStyle.Bottom, Height = 36 };
            panel.Controls.Add(_btnCopy);
            panel.Controls.Add(_btnClose);

            Controls.Add(_plot);
            Controls.Add(_txt);
            Controls.Add(panel);

            _btnCopy.Click += (s, e) => { try { Clipboard.SetText(_txt.Text); } catch { } };
            _btnClose.Click += (s, e) => Close();

            Render(result);
        }

        private void Render(ParseResult r)
        {
            var lines = new List<string>();
            lines.Add($"昨日销量：{r.Yesterday}");
            lines.Add($"近7天合计：{r.Sum7d}");
            lines.Add("");
            if (r.Records.Count > 0)
            {
                lines.Add("明细：");
                foreach (var it in r.Records)
                    lines.Add($"{it.Date:yyyy-MM-dd}: {it.Quantity}");
            }
            else
            {
                lines.Add("未解析到明细。");
            }
            lines.Add("");
            lines.Add("原始：");
            lines.Add(r.Raw);
            _txt.Text = string.Join(Environment.NewLine, lines);

            // Chart
            var model = new PlotModel { Title = "近7天趋势" };
            model.Axes.Add(new DateTimeAxis { Position = AxisPosition.Bottom, StringFormat = "MM-dd", IntervalType = DateTimeIntervalType.Days });
            model.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Minimum = 0 });
            var series = new LineSeries { MarkerType = MarkerType.Circle };
            foreach (var it in r.Records.OrderBy(x=>x.Date))
                series.Points.Add(DateTimeAxis.CreateDataPoint(it.Date, it.Quantity));
            model.Series.Add(series);
            _plot.Model = model;
        }
    }
}
