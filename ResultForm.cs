using System;
using System.Drawing;
using System.Windows.Forms;

namespace StyleWatcherWin
{
    public class ResultForm : Form
    {
        readonly AppConfig _cfg;
        readonly TextBox _input;
        readonly TextBox _output;
        readonly Label _status;

        public ResultForm(AppConfig cfg)
        {
            _cfg = cfg;
            Text = "StyleWatcher";
            StartPosition = FormStartPosition.Manual;
            DoubleBuffered = true;
            Width = Math.Max(cfg.window.width, 800);
            Height = Math.Max(cfg.window.height, 560);
            BackColor = Color.White;

            var table = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1 };
            _status = new Label { Text = "就绪", Dock = DockStyle.Top, Padding = new Padding(8) };
            _input = new TextBox { Multiline = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical, ReadOnly = true };
            _output = new TextBox { Multiline = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical, ReadOnly = true };

            table.Controls.Add(_status, 0, 0);
            table.Controls.Add(_input, 0, 1);
            table.Controls.Add(_output, 0, 2);
            Controls.Add(table);
        }

        public void SetStatus(string text) { if (InvokeRequired) { Invoke(new Action<string>(SetStatus), text); return; } _status.Text = text; }
        public void SetPayload(string input, string output)
        {
            if (InvokeRequired) { Invoke(new Action<string, string>(SetPayload), input, output); return; }
            _input.Text = input ?? "";
            _output.Text = output ?? "";
            _status.Text = "完成";
        }

        public void ShowAndFocusCentered(bool topMost)
        {
            TopMost = topMost;
            if (!Visible) Show();
            var screen = Screen.PrimaryScreen.WorkingArea;
            Left = screen.Left + (screen.Width - Width) / 2;
            Top = screen.Top + (screen.Height - Height) / 2;
            Activate();
            Focus();
        }
    }
}
