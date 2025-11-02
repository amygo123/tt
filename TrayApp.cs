using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.IO;

namespace StyleWatcherWin
{
    public static class Formatter
    {
        public static string Prettify(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return raw;
            var s = raw.Replace("\\n", "\n").Replace("\r\n", "\n");
            var lines = s.Split('\n');
            for (int i = 0; i < lines.Length; i++) lines[i] = lines[i].Trim();
            s = string.Join("\n", lines);
            s = System.Text.RegularExpressions.Regex.Replace(s, @"\n{3,}", "\n\n");
            return s.Trim();
        }
    }

    public class TrayApp : Form
    {
        [DllImport("user32.dll")] static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")] static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        [DllImport("user32.dll")] static extern short GetAsyncKeyState(int vKey);
        [DllImport("user32.dll")] static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

        const int WM_HOTKEY = 0x0312;
        const int KEYEVENTF_KEYUP = 0x0002;

        AppConfig _cfg;
        ResultForm _window;
        readonly SemaphoreSlim _queryLock = new SemaphoreSlim(1, 1);
        DateTime _lastHotkeyAt = DateTime.MinValue;

        int _hotkeyId = 1;
        uint _mod;
        uint _vk;
        bool _allowCloseAll = false;

        NotifyIcon _tray;

        public TrayApp()
        {
            _cfg = AppConfig.Load();
            ShowInTaskbar = false;
            WindowState = FormWindowState.Minimized;
            Visible = false;

            BuildTray();
            ParseHotkey(_cfg.hotkey ?? "Alt+S", out _mod, out _vk);
            RegisterHotKey(this.Handle, _hotkeyId, _mod, _vk);
        }

        void BuildTray()
        {
            _tray = new NotifyIcon
            {
                Visible = true,
                Icon = System.Drawing.SystemIcons.Information,
                Text = "StyleWatcher"
            };
            var menu = new ContextMenuStrip();
            var mInput = new ToolStripMenuItem("手动输入...");
            var mConfig = new ToolStripMenuItem("打开配置文件");
            var mExit = new ToolStripMenuItem("退出");
            menu.Items.AddRange(new ToolStripItem[] { mInput, new ToolStripSeparator(), mConfig, new ToolStripSeparator(), mExit });
            _tray.ContextMenuStrip = menu;

            mInput.Click += async (s, e) =>
            {
                using var dlg = new InputBox("输入要解析的文本：");
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    await HandleQueryAsync(dlg.TextValue);
                }
            };
            mConfig.Click += (s, e) => System.Diagnostics.Process.Start("notepad.exe", AppConfig.ConfigPath);
            mExit.Click += (s, e) => ExitApp();
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && (int)m.WParam == _hotkeyId)
            {
                _ = OnHotkeyAsync();
            }
            base.WndProc(ref m);
        }

        async Task OnHotkeyAsync()
        {
            if ((DateTime.Now - _lastHotkeyAt).TotalMilliseconds < 300) return;
            _lastHotkeyAt = DateTime.Now;

            try { ReleaseAlt(); } catch { }
            string text = await TryGetSelectedTextAsync();
            try { ReleaseAlt(); } catch { }

            if (string.IsNullOrWhiteSpace(text)) return;
            await HandleQueryAsync(text);
        }

        async Task HandleQueryAsync(string text)
        {
            await _queryLock.WaitAsync();
            try
            {
                if (_window == null || _window.IsDisposed) _window = new ResultForm(_cfg);
                _window.ShowAndFocusCentered(_cfg.window.alwaysOnTop);
                _window.SetStatus("查询中...");
                var result = await ApiHelper.QueryAsync(_cfg, text);
                _window.SetPayload(text, result);
                _window.ShowAndFocusCentered(_cfg.window.alwaysOnTop);
            }
            finally
            {
                _queryLock.Release();
            }
        }

        async Task<string> TryGetSelectedTextAsync()
        {
            string backup = null;
            try
            {
                if (Clipboard.ContainsText()) backup = Clipboard.GetText();

                keybd_event((byte)Keys.ControlKey, 0, 0, 0);
                keybd_event((byte)Keys.C, 0, 0, 0);
                keybd_event((byte)Keys.C, 0, KEYEVENTF_KEYUP, 0);
                keybd_event((byte)Keys.ControlKey, 0, KEYEVENTF_KEYUP, 0);

                await Task.Delay(120);
                if (Clipboard.ContainsText()) return Clipboard.GetText();
                return "";
            }
            catch
            {
                return "";
            }
            finally
            {
                try { if (backup != null) Clipboard.SetText(backup); } catch { }
            }
        }

        void ReleaseAlt()
        {
            keybd_event((byte)Keys.Menu, 0, KEYEVENTF_KEYUP, 0);
        }

        void ParseHotkey(string text, out uint mod, out uint vk)
        {
            mod = 0; vk = 0;
            if (string.IsNullOrWhiteSpace(text)) { mod = 1 << 0; vk = (uint)Keys.S; return; }
            var parts = text.Split('+', StringSplitOptions.RemoveEmptyEntries);
            foreach (var p in parts)
            {
                var t = p.Trim().ToUpperInvariant();
                if (t == "CTRL" || t == "CONTROL") mod |= 0x0002;
                else if (t == "SHIFT") mod |= 0x0004;
                else if (t == "ALT") mod |= 0x0001;
                else if (t == "WIN" || t == "WINDOWS") mod |= 0x0008;
                else vk = (uint)Enum.Parse(typeof(Keys), t, true);
            }
            if (vk == 0) vk = (uint)Keys.S;
        }

        void ExitApp()
        {
            try { UnregisterHotKey(Handle, _hotkeyId); } catch { }
            _allowCloseAll = true;
            try { _window?.Close(); } catch { }
            if (_tray != null) _tray.Visible = false;
            Application.Exit();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_allowCloseAll) return;
            e.Cancel = true;
        }
    }

    public class InputBox : Form
    {
        TextBox _tb;
        public string TextValue => _tb.Text.Trim();
        public InputBox(string title)
        {
            Text = title;
            Width = 600;
            Height = 220;
            StartPosition = FormStartPosition.CenterScreen;
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1 };
            _tb = new TextBox { Multiline = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical };
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
            var ok = new Button { Text = "确定", DialogResult = DialogResult.OK };
            var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel };
            buttons.Controls.Add(ok); buttons.Controls.Add(cancel);
            panel.Controls.Add(_tb, 0, 0);
            panel.SetRowSpan(_tb, 2);
            panel.Controls.Add(buttons, 0, 2);
            Controls.Add(panel);
            AcceptButton = ok; CancelButton = cancel;
        }
    }
}
