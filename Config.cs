using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

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

        // A2: 库存接口
        public InventoryCfg inventory { get; set; } = new InventoryCfg();
        public LookupCfg lookup { get; set; } = new LookupCfg();

        // 新增：UI 配置（趋势窗口、是否显示 MA）
        public UiCfg ui { get; set; } = new UiCfg();

        // 新增：库存天数与销量窗口的告警阈值
        public InventoryAlertCfg inventoryAlert { get; set; } = new InventoryAlertCfg();

        public class WindowCfg
        {
            public int width { get; set; } = 560;
            public int height { get; set; } = 380;
            public int fontSize { get; set; } = 13;
            public bool alwaysOnTop { get; set; } = true;
        }

        public class Headers
        {
            // 兼容 "Content-Type" 键名
            [JsonPropertyName("Content-Type")]
            public string Content_Type { get; set; } = "application/json";
        }

        public class InventoryCfg
        {
            public string url_base { get; set; } 

        public class LookupCfg
        {
            // 通过款名查询定级/最低价/保本价
            public string url_base { get; set; } = "http://192.168.40.97:8002/lookup?name=";
        }
= "http://192.168.40.97:8000/inventory?style_name=";
            public string default_style { get; set; } = "纯色/通纯棉圆领短T/黑/XL";
        }

        public class UiCfg
        {
            public int[] trendWindows { get; set; } = new[] { 7, 14, 30 };
            public bool showMovingAverage { get; set; } = true;
        }

        public class InventoryAlertCfg
        {
            public double docRed { get; set; } = 3;         // 库存天数 < 3 天：红
            public double docYellow { get; set; } = 7;      // 库存天数 < 7 天：黄
            public int minSalesWindowDays { get; set; } = 7;// 最近 N 天销量作为基线
        }

        public static string ConfigPath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");

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
                using var http = new System.Net.Http.HttpClient
                {
                    Timeout = System.TimeSpan.FromSeconds(Math.Max(3, cfg.timeout_seconds))
                };
                var req = new System.Net.Http.HttpRequestMessage(
                    new System.Net.Http.HttpMethod(cfg.method ?? "POST"),
                    cfg.api_url ?? "");

                req.Content = new System.Net.Http.StringContent(
                    $"{{\"{cfg.json_key}\":\"{text?.Replace("\\","\\\\").Replace("\"","\\\"")}\"}}",
                    System.Text.Encoding.UTF8,
                    "application/json");

                var resp = await http.SendAsync(req);
                var raw = await resp.Content.ReadAsStringAsync();

                // 若返回 JSON 带 msg 字段，则优先取之
                try
                {
                    var doc = System.Text.Json.JsonDocument.Parse(raw);
                    if (doc.RootElement.TryGetProperty("msg", out var msgEl))
                        return msgEl.ToString();
                    return raw;
                }
                catch
                {
                    return raw;
                }
            }
            catch (System.Exception ex)
            {
                return $"请求失败：{ex.Message}";
            }
        }

        // A2: 查询库存（GET）
        public static async System.Threading.Tasks.Task<string> QueryInventoryAsync(AppConfig cfg, string styleName)
        {
            var baseUrl = cfg.inventory?.url_base ?? "";
            if (string.IsNullOrWhiteSpace(baseUrl)) return "";
            var url = baseUrl + Uri.EscapeDataString(styleName ?? "");
            try
            {
                using var http = new System.Net.Http.HttpClient
                {
                    Timeout = System.TimeSpan.FromSeconds(Math.Max(3, cfg.timeout_seconds))
                };
                var raw = await http.GetStringAsync(url);
                return raw ?? "";
            }
            catch (System.Exception ex)
            {
                return $"[] // 请求失败：{ex.Message}";
            }
        }
    }
}


        public static async System.Threading.Tasks.Task<string> QueryLookupAsync(AppConfig cfg, string styleName)
        {
            try
            {
                var baseUrl = cfg.lookup?.url_base ?? "";
                if (string.IsNullOrWhiteSpace(baseUrl)) return "[]";
                var url = baseUrl + System.Uri.EscapeDataString(styleName ?? "");
                using var http = new System.Net.Http.HttpClient
                {
                    Timeout = System.TimeSpan.FromSeconds(Math.Max(3, cfg.timeout_seconds))
                };
                var raw = await http.GetStringAsync(url);
                return raw ?? "[]";
            }
            catch (System.Exception ex)
            {
                return $"[] // 请求失败：{ex.Message}";
            }
        }
