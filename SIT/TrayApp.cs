using System;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;

namespace StyleWatcherWin
{
    public class TrayApp : Form
    {
        // --- Win32 热键 ---
        [DllImport("user32.dll")] private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")] private static extern bool UnregisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        private const int HOTKEY_ID = 1001;
        private const uint MOD_ALT = 0x0001;
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;
        private const uint MOD_WIN = 0x0008;

        private readonly NotifyIcon _notify;
        private readonly ContextMenuStrip _menu;
        private readonly ToolStripMenuItem _miInput;
        private readonly ToolStripMenuItem _miExit;
        private readonly HttpClient _http;
        private AppConfig _cfg;

        public TrayApp()
        {
            _cfg = AppConfig.Load();
            _http = new HttpClient();
            _http.Timeout = TimeSpan.FromSeconds(Math.Max(1, _cfg.timeout_seconds));

            // 托盘
            _menu = new ContextMenuStrip();
            _miInput = new ToolStripMenuItem("手动输入查询(&I)", null, (s,e)=>ManualInput());
            _miExit = new ToolStripMenuItem("退出(&X)", null, (s,e)=>Close());

            _menu.Items.Add(_miInput);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(_miExit);

            _notify = new NotifyIcon
            {
                Text = "StyleWatcher",
                Icon = SystemIcons.Information,
                Visible = true,
                ContextMenuStrip = _menu
            };

            // 注册热键
            ParseHotkey(_cfg.hotkey, out var mod, out var vk);
            RegisterHotKey(this.Handle, HOTKEY_ID, mod, vk);

            // 隐藏窗体（仅托盘）
            ShowInTaskbar = false;
            Opacity = 0;
            Load += (s,e)=> { Hide(); };
            FormClosing += (s,e)=> { _notify.Visible = false; UnregisterHotKey(this.Handle, HOTKEY_ID, 0, 0); };
        }

        protected override void WndProc(ref Message m)
        {
            const int WM_HOTKEY = 0x0312;
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                _ = OnHotkeyAsync();
            }
            base.WndProc(ref m);
        }

        private async Task OnHotkeyAsync()
        {
            try
            {
                string text = await ClipboardGetAsync();
                if (string.IsNullOrWhiteSpace(text))
                {
                    _notify.ShowBalloonTip(2000, "未获取文本", "请先在目标应用中选中文本后再试，或右键托盘选择手动输入。", ToolTipIcon.Info);
                    return;
                }
                await RunQueryAsync(text);
            }
            catch (Exception ex)
            {
                _notify.ShowBalloonTip(3000, "错误", ex.Message, ToolTipIcon.Error);
            }
        }

        private async Task RunQueryAsync(string code)
        {
            string resultText = await CallApiAsync(code);
            var parsed = PayloadParser.Parse(resultText);
            var form = new ResultForm(parsed, _cfg.window.fontSize, _cfg.window.alwaysOnTop, _cfg.window.width, _cfg.window.height);
            form.Show();
            form.Activate();
        }

        private async Task<string> CallApiAsync(string code)
        {
            var url = _cfg.api_url ?? "";
            var method = (_cfg.method ?? "POST").ToUpperInvariant();
            using var req = new HttpRequestMessage(new HttpMethod(method), url);

            var payload = new JsonDocumentOptions();
            var json = JsonSerializer.Serialize(new { code = code });
            if (!string.IsNullOrEmpty(_cfg.json_key))
            {
                json = JsonSerializer.Serialize(new System.Collections.Generic.Dictionary<string, string> { { _cfg.json_key, code } });
            }

            req.Content = new StringContent(json, Encoding.UTF8, "application/json");
            if (!string.IsNullOrEmpty(_cfg.headers?.Content_Type))
            {
                req.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(_cfg.headers.Content_Type);
                req.Content.Headers.ContentType.CharSet = "utf-8";
            }

            try
            {
                var res = await _http.SendAsync(req);
                var raw = await res.Content.ReadAsStringAsync();
                if (!res.IsSuccessStatusCode)
                {
                    return $"HTTP { (int)res.StatusCode }: { raw }";
                }
                // 兼容 { msg: ... }
                try
                {
                    using var doc = JsonDocument.Parse(raw);
                    if (doc.RootElement.TryGetProperty("msg", out var msg)) return msg.GetString() ?? raw;
                }
                catch { }
                return raw;
            }
            catch (Exception ex)
            {
                return $"请求失败：{ex.Message}";
            }
        }

        private async Task<string> ClipboardGetAsync()
        {
            // 尝试读取剪贴板，多次重试
            for (int i = 0; i < 3; i++)
            {
                try
                {
                    if (Clipboard.ContainsText())
                    {
                        var s = Clipboard.GetText();
                        if (!string.IsNullOrWhiteSpace(s))
                            return s.Trim();
                    }
                }
                catch { }
                await Task.Delay(120);
            }
            return "";
        }

        private void ManualInput()
        {
            using var dlg = new InputBox("输入查询内容：");
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                _ = RunQueryAsync(dlg.Value ?? "");
            }
        }

        private void ParseHotkey(string s, out uint mod, out uint vk)
        {
            mod = 0; vk = 0;
            if (string.IsNullOrWhiteSpace(s)) { mod = MOD_ALT; vk = (uint)Keys.S; return; }
            foreach (var part in s.Split('+'))
            {
                var t = part.Trim().ToUpperInvariant();
                if (t is "CTRL" or "CONTROL") mod |= MOD_CONTROL;
                else if (t == "SHIFT") mod |= MOD_SHIFT;
                else if (t == "ALT") mod |= MOD_ALT;
                else if (t is "WIN" or "WINDOWS") mod |= MOD_WIN;
                else if (Enum.TryParse<Keys>(t, true, out var key)) vk = (uint)key;
            }
            if (vk == 0) { mod = MOD_ALT; vk = (uint)Keys.S; }
        }
    }

    // 简单输入框
    internal class InputBox : Form
    {
        private readonly TextBox _tb;
        private readonly Button _ok;
        private readonly Button _cancel;
        public string? Value => _tb.Text;

        public InputBox(string title)
        {
            Text = title;
            StartPosition = FormStartPosition.CenterParent;
            Width = 480; Height = 160;
            _tb = new TextBox { Dock = DockStyle.Top, Height = 32 };
            _ok = new Button { Text = "确定", Dock = DockStyle.Left, Width = 100 };
            _cancel = new Button { Text = "取消", Dock = DockStyle.Right, Width = 100 };
            var panel = new Panel { Dock = DockStyle.Bottom, Height = 36 };
            panel.Controls.Add(_ok); panel.Controls.Add(_cancel);
            Controls.Add(panel); Controls.Add(_tb);

            _ok.Click += (s,e)=> DialogResult = DialogResult.OK;
            _cancel.Click += (s,e)=> DialogResult = DialogResult.Cancel;
        }
    }
}
