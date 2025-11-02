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
        // --- Win32 ---
        [DllImport("user32.dll")] static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")] static extern bool UnregisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")] static extern short GetAsyncKeyState(int vKey);
        [DllImport("user32.dll")] static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] static extern IntPtr GetFocus(); // 必须有
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
        [DllImport("user32.dll")] static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
        [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
        [DllImport("user32.dll", CharSet = CharSet.Auto)] static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Auto)] static extern IntPtr SendMessage(IntPtr hWnd, int msg, ref int wParam, ref int lParam);
        [DllImport("user32.dll", CharSet = CharSet.Auto)] static extern IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, StringBuilder lParam);

        const int WM_GETTEXTLENGTH = 0x000E;
        const int WM_GETTEXT = 0x000D;
        const int EM_GETSEL = 0x00B0;
        const int WM_HOTKEY = 0x0312;

        const uint MOD_ALT = 0x0001, MOD_CONTROL = 0x0002, MOD_SHIFT = 0x0004, MOD_WIN = 0x0008;
        const int KEYEVENTF_KEYUP = 0x0002;
        const byte VK_MENU = 0x12; // Alt

        static void ReleaseAlt()
        {
            if ((GetAsyncKeyState(VK_MENU) & 0x8000) != 0)
                keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
        }

        // Tray & config
        readonly NotifyIcon _tray = new NotifyIcon();
        readonly ContextMenuStrip _menu = new ContextMenuStrip();
        readonly HttpClient _http = new HttpClient();
        readonly AppConfig _cfg;

        // Single window & throttling
        ResultForm _window;
        readonly SemaphoreSlim _queryLock = new SemaphoreSlim(1, 1);
        DateTime _lastHotkeyAt = DateTime.MinValue;

        int _hotkeyId = 1;
        uint _mod;
        uint _vk;
        bool _allowCloseAll = false;

        public TrayApp()
        {
            _cfg = AppConfig.Load();
            ShowInTaskbar = false;
            WindowState = FormWindowState.Minimized;
            Visible = false;

            // 托盘图标（使用 Resources\app.ico）
            // 托盘图标：直接提取 EXE 的图标，确保与应用图标一致
            _tray.Text = "款式信息";
            try
            {
                var exeIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                if (exeIcon != null) _tray.Icon = exeIcon;
                else
                {
                    // 兜底再尝试 Resources\app.ico
                    var icoPath = Path.Combine(AppContext.BaseDirectory, "Resources", "app.ico");
                    _tray.Icon = File.Exists(icoPath) ? new Icon(icoPath) : SystemIcons.Application;
                }
            }
            catch
            {
                _tray.Icon = SystemIcons.Application;
            }
            _tray.Visible = true; // 确保在设置好 Icon 后再显示

            _tray.DoubleClick += (s, e) => ToggleWindow(show: true);

            var itemToggle = new ToolStripMenuItem("显示/隐藏 窗口", null, (s, e) => ToggleWindow(toggle: true));
            var itemQuery = new ToolStripMenuItem("手动输入查询", null, (s, e) =>
            {
                EnsureWindow();
                _window.FocusInput();
                _window.ShowNoActivateAtCursor();
            });
            var itemConfig = new ToolStripMenuItem("打开配置文件", null, (s, e) =>
            {
                try { System.Diagnostics.Process.Start("notepad.exe", AppConfig.ConfigPath); } catch { }
            });
            var itemExit = new ToolStripMenuItem("退出", null, (s, e) => ExitApp());

            _menu.Items.Add(itemToggle);
            _menu.Items.Add(itemQuery);
            _menu.Items.Add(itemConfig);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(itemExit);
            _tray.ContextMenuStrip = _menu;

            _http.Timeout = TimeSpan.FromSeconds(_cfg.timeout_seconds);
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            ParseHotkey(_cfg.hotkey, out _mod, out _vk);
            if (!RegisterHotKey(Handle, _hotkeyId, _mod, _vk))
                MessageBox.Show($"热键 {_cfg.hotkey} 注册失败，可能被占用。", "StyleWatcher",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);

            _tray.BalloonTipTitle = "StyleWatcher 已启动";
            _tray.BalloonTipText = $"选中文本后按 {_cfg.hotkey} 查询；双击托盘可显示窗口。";
            _tray.ShowBalloonTip(2500);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY)
                _ = OnHotkeyAsync();
            base.WndProc(ref m);
        }

        void EnsureWindow()
        {
            if (_window == null || _window.IsDisposed)
            {
                _window = new ResultForm(_cfg);
                _window.FormClosing += (s, e) =>
                {
                    if (!_allowCloseAll)
                    {
                        e.Cancel = true;
                        _window.Hide();
                    }
                };
            }
        }

        void ToggleWindow(bool show = false, bool toggle = false)
        {
            EnsureWindow();
            if (toggle)
            {
                if (_window.Visible) _window.Hide();
                else _window.ShowAndFocusCentered(_cfg.window.alwaysOnTop);
            }
            else if (show)
            {
                _window.ShowAndFocusCentered(_cfg.window.alwaysOnTop);
            }
        }

        void ExitApp()
        {
            try { UnregisterHotKey(Handle, _hotkeyId, _mod, _vk); } catch { }
            _allowCloseAll = true;
            try { _window?.Close(); } catch { }
            _tray.Visible = false;
            Application.Exit();
        }

        // 选区（Win32）+ 剪贴板兜底
        private string TryGetSelectedTextUsingWin32()
        {
            try
            {
                var fg = GetForegroundWindow();
                if (fg == IntPtr.Zero) return null;

                uint fgThread = GetWindowThreadProcessId(fg, out _);
                uint curThread = GetCurrentThreadId();
                bool attached = false;

                try
                {
                    attached = AttachThreadInput(curThread, fgThread, true);
                    var hFocus = GetFocus();
                    if (hFocus == IntPtr.Zero) return null;

                    int start = 0, end = 0;
                    SendMessage(hFocus, EM_GETSEL, ref start, ref end);

                    int len = (int)SendMessage(hFocus, WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero);
                    if (len <= 0) return null;

                    var sb = new StringBuilder(len + 1);
                    SendMessage(hFocus, WM_GETTEXT, sb.Capacity, sb);
                    var full = sb.ToString();

                    if (start < 0 || end < 0 || start > full.Length) return null;
                    if (end > full.Length) end = full.Length;
                    if (end > start) return full.Substring(start, end - start).Trim();
                }
                finally
                {
                    if (attached) AttachThreadInput(curThread, fgThread, false);
                }
            }
            catch { }
            return null;
        }

        private async Task<string> GetSelectionByClipboardRoundTripAsync()
        {
            IDataObject backup = null;
            try { backup = Clipboard.GetDataObject(); } catch { }

            SendKeys.SendWait("^c");
            await Task.Delay(120);

            string txt = "";
            try { txt = Clipboard.GetText()?.Trim() ?? ""; } catch { }

            if (backup != null)
            {
                try { Clipboard.SetDataObject(backup, true); } catch { }
            }
            return txt;
        }

        // 热键（去抖 + 限流 + 复用同窗）
        private async Task OnHotkeyAsync()
        {
            var now = DateTime.UtcNow;
            if ((now - _lastHotkeyAt).TotalMilliseconds < 500) return;
            _lastHotkeyAt = now;

            if (!await _queryLock.WaitAsync(0))
            {
                ToggleWindow(show: true);
                return;
            }

            try
            {
                ReleaseAlt();

                string txt = TryGetSelectedTextUsingWin32();
                if (string.IsNullOrEmpty(txt)) txt = await GetSelectionByClipboardRoundTripAsync();

                EnsureWindow();
                _window.ShowAndFocusCentered(_cfg.window.alwaysOnTop);

                if (string.IsNullOrEmpty(txt))
                {
                    _window.SetLoading("未检测到选中文本，请先选中一段文字再按热键。");
                    return;
                }

                _window.SetLoading("查询中...");
                string result = await QueryAsync(txt);
                _window.ApplyRawText(txt, result);
            }
            catch (Exception ex)
            {
                EnsureWindow();
                _window.SetLoading($"错误：{ex.Message}");
            }
            finally
            {
                ReleaseAlt();
                _queryLock.Release();
            }
        }

        private async Task<string> QueryAsync(string text)
        {
            try
            {
                var method = (_cfg.method ?? "POST").ToUpperInvariant();
                HttpRequestMessage req;

                if (method == "GET")
                {
                    req = new HttpRequestMessage(HttpMethod.Get, $"{_cfg.api_url}?{_cfg.json_key}={Uri.EscapeDataString(text)}");
                }
                else
                {
                    req = new HttpRequestMessage(HttpMethod.Post, _cfg.api_url);
                    var payload = new Dictionary<string, string> { { _cfg.json_key, text } };
                    var json = JsonSerializer.Serialize(payload);
                    req.Content = new StringContent(json, Encoding.UTF8, "application/json");
                }

                var resp = await _http.SendAsync(req);
                var raw = await resp.Content.ReadAsStringAsync();

                try
                {
                    using var doc = JsonDocument.Parse(raw);
                    if (doc.RootElement.TryGetProperty("msg", out var msgEl))
                        return Formatter.Prettify(msgEl.ToString());
                    return Formatter.Prettify(raw);
                }
                catch { return Formatter.Prettify(raw); }
            }
            catch (Exception ex) { return $"请求失败：{ex.Message}"; }
        }

        private void ParseHotkey(string s, out uint mod, out uint vk)
        {
            mod = 0; vk = 0;
            if (string.IsNullOrWhiteSpace(s)) { mod = MOD_ALT; vk = (uint)Keys.S; return; }
            var parts = s.Split(new[] { '+' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var p in parts)
            {
                var t = p.Trim().ToUpperInvariant();
                if (t == "CTRL" || t == "CONTROL") mod |= MOD_CONTROL;
                else if (t == "SHIFT") mod |= MOD_SHIFT;
                else if (t == "ALT") mod |= MOD_ALT;
                else if (t == "WIN" || t == "WINDOWS") mod |= MOD_WIN;
                else if (Enum.TryParse(t, true, out Keys key)) vk = (uint)key;
            }
            if (vk == 0) { mod = MOD_ALT; vk = (uint)Keys.S; }
        }
    }
}
