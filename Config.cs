using System;
using System.IO;
using System.Text.Json;

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
        public class WindowCfg { public int width { get; set; } = 900; public int height { get; set; } = 600; public int fontSize { get; set; } = 12; public bool alwaysOnTop { get; set; } = true; }
        public class Headers { public string Content_Type { get; set; } = "application/json"; }

        public static string ConfigPath => Path.Combine(AppContext.BaseDirectory, "appsettings.json");

        public static AppConfig Load()
        {
            try
            {
                if (File.Exists(ConfigPath))
                {
                    var json = File.ReadAllText(ConfigPath);
                    var cfg = JsonSerializer.Deserialize<AppConfig>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                    return cfg ?? new AppConfig();
                }
            }
            catch { }
            var def = new AppConfig();
            Save(def);
            return def;
        }
        public static void Save(AppConfig cfg)
        {
            var json = JsonSerializer.Serialize(cfg, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(ConfigPath, json);
        }
    }

    public static class ApiHelper
    {
        public static async System.Threading.Tasks.Task<string> QueryAsync(AppConfig cfg, string text)
        {
            try
            {
                using var http = new System.Net.Http.HttpClient { Timeout = System.TimeSpan.FromSeconds(cfg.timeout_seconds) };
                var method = (cfg.method ?? "POST").ToUpperInvariant();
                System.Net.Http.HttpRequestMessage req;
                if (method == "GET")
                {
                    req = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, $"{cfg.api_url}?{cfg.json_key}={System.Uri.EscapeDataString(text)}");
                }
                else
                {
                    var json = System.Text.Json.JsonSerializer.Serialize(new System.Collections.Generic.Dictionary<string,string> { { cfg.json_key, text } });
                    req = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Post, cfg.api_url);
                    req.Content = new System.Net.Http.StringContent(json, System.Text.Encoding.UTF8, "application/json");
                }
                var resp = await http.SendAsync(req);
                var raw = await resp.Content.ReadAsStringAsync();
                try
                {
                    var doc = System.Text.Json.JsonDocument.Parse(raw);
                    if (doc.RootElement.TryGetProperty("msg", out var msgEl))
                        return msgEl.ToString();
                    return raw;
                }
                catch { return raw; }
            }
            catch (System.Exception ex)
            {
                return $"请求失败：{ex.Message}";
            }
        }
    }
}
