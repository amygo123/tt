using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace StyleWatcherWin
{
    public class AppConfig
    {
        public string api_url { get; set; } = "http://47.111.189.27:8089/qrcode/saleVolumeParser";
        public string method { get; set; } = "POST";
        public string json_key { get; set; } = "code";
        public int timeout_seconds { get; set; } = 6;
        public string hotkey { get; set; } = "Alt+S";
        public WindowCfg window { get; set; } = new WindowCfg();
        public Headers headers { get; set; } = new Headers();
        public InventoryCfg inventory { get; set; } = new InventoryCfg();

        public class WindowCfg
        {
            public int width { get; set; } = 560;
            public int height { get; set; } = 380;
            public int fontSize { get; set; } = 13;
            public bool alwaysOnTop { get; set; } = true;
        }

        public class Headers : Dictionary<string, string> { }

        public class InventoryCfg
        {
            public string url_base { get; set; } = "http://192.168.40.97:8000/inventory?style_name=";
            public string default_style { get; set; } = "纯色/通纯棉圆领短T/黑/XL";
        }

        public static string ConfigPath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");

        public static AppConfig Load()
        {
            try
            {
                if (!File.Exists(ConfigPath))
                {
                    var def = new AppConfig();
                    File.WriteAllText(ConfigPath, JsonSerializer.Serialize(def, new JsonSerializerOptions { WriteIndented = true, Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping }));
                    return def;
                }
                var raw = File.ReadAllText(ConfigPath);
                return JsonSerializer.Deserialize<AppConfig>(raw) ?? new AppConfig();
            }
            catch
            {
                return new AppConfig();
            }
        }

        public void Save()
        {
            File.WriteAllText(ConfigPath, JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true, Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping }));
        }
    }

    public static class ApiHelper
    {
        public static async Task<string> QueryAsync(AppConfig cfg, string text)
        {
            try
            {
                using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(Math.Max(3, cfg.timeout_seconds)) };
                using var req = new HttpRequestMessage(
                    string.Equals(cfg.method, "GET", StringComparison.OrdinalIgnoreCase) ? HttpMethod.Get : HttpMethod.Post,
                    cfg.api_url);

                foreach (var kv in cfg.headers)
                {
                    if (!string.Equals(kv.Key, "Content-Type", StringComparison.OrdinalIgnoreCase))
                        req.Headers.TryAddWithoutValidation(kv.Key, kv.Value);
                }

                if (req.Method == HttpMethod.Get)
                {
                    var url = cfg.api_url.Contains("?") ? cfg.api_url + "&" : cfg.api_url + "?";
                    url += Uri.EscapeDataString(cfg.json_key) + "=" + Uri.EscapeDataString(text ?? "");
                    req.RequestUri = new Uri(url);
                }
                else
                {
                    var payload = JsonSerializer.Serialize(new Dictionary<string, string> { [cfg.json_key] = text ?? "" });
                    req.Content = new StringContent(payload, Encoding.UTF8, "application/json");
                }

                var resp = await http.SendAsync(req);
                resp.EnsureSuccessStatusCode();
                var raw = await resp.Content.ReadAsStringAsync();

                try
                {
                    var doc = JsonDocument.Parse(raw);
                    if (doc.RootElement.TryGetProperty("msg", out var msgEl))
                        return msgEl.ToString();
                    return raw;
                }
                catch { return raw; }
            }
            catch (Exception ex)
            {
                return $"请求失败：{ex.Message}";
            }
        }
    }
}
