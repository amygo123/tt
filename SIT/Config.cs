using System;
using System.IO;
using System.Text.Json;

namespace StyleWatcherWin
{
    public class AppConfig
    {
        // 主查询
        public string api_url { get; set; } = "http://127.0.0.1:8080/api";
        public string method { get; set; } = "POST";
        public string json_key { get; set; } = "code";
        public int timeout_seconds { get; set; } = 6;
        public string hotkey { get; set; } = "Alt+S";
        public WindowCfg window { get; set; } = new WindowCfg();
        public Headers headers { get; set; } = new Headers();

        // 库存（可选）
        public string inventory_api_url { get; set; } = "";
        public int inventory_timeout_seconds { get; set; } = 4;
        public int inventory_low_threshold { get; set; } = 10;
        public int inventory_cache_ttl_seconds { get; set; } = 300;

        public class WindowCfg { public int width { get; set; } = 640; public int height { get; set; } = 420; public int fontSize { get; set; } = 13; public bool alwaysOnTop { get; set; } = true; }
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
                    if (cfg != null) return cfg;
                }
            }
            catch { }
            return new AppConfig();
        }

        public void Save()
        {
            var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(ConfigPath, json);
        }
    }
}
