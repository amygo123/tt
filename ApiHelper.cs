
using System;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace StyleWatcherWin
{
    public static class ApiHelper
    {
        /// <summary>
        /// 调用解析接口：读取 AppConfig（方法/URL/headers/json_key），健壮性增强：
        /// - 使用超时（来自 cfg.timeout_seconds，最小 3s）
        /// - 透传可选 Authorization 等 Header
        /// - 检查 HTTP 状态码（非 2xx 返回友好错误）
        /// </summary>
        public static async Task<string> QueryAsync(AppConfig cfg, string text, CancellationToken ct = default)
        {
            if (cfg == null) throw new ArgumentNullException(nameof(cfg));

            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(Math.Max(3, cfg.timeout_seconds)) };

            using var req = new HttpRequestMessage(new HttpMethod(cfg.method ?? "POST"), cfg.api_url ?? "");

            // Content-Type from config, default application/json
            var contentType = cfg.headers?.Content_Type;
            if (string.IsNullOrWhiteSpace(contentType)) contentType = "application/json";

            // Build JSON payload
            var payload = $"{{\"{cfg.json_key}\":\"{(text ?? string.Empty).Replace("\\", "\\\\").Replace("\"", "\\\"")}\"}}";
            req.Content = new StringContent(payload, Encoding.UTF8, contentType);

            // Optional Authorization header
            var auth = cfg.headers?.Authorization;
            if (!string.IsNullOrWhiteSpace(auth))
            {
                if (!req.Headers.Contains("Authorization"))
                    req.Headers.TryAddWithoutValidation("Authorization", auth);
            }

            // Send
            using var resp = await http.SendAsync(req, ct).ConfigureAwait(false);

            // Ensure success or return rich error
            if (!resp.IsSuccessStatusCode)
            {
                var reason = resp.ReasonPhrase ?? "Unknown";
                var status = (int)resp.StatusCode;
                string body = string.Empty;
                try { body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false); } catch { }
                return $"请求失败：{status} {reason}" + (string.IsNullOrWhiteSpace(body) ? "" : $" | 响应：{Trim(body, 200)}");
            }

            var raw = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
            // 优先 msg 字段
            try
            {
                using var doc = System.Text.Json.JsonDocument.Parse(raw);
                if (doc.RootElement.TryGetProperty("msg", out var msgEl))
                    return msgEl.ToString();
            }
            catch { /* fall back to raw */ }

            return raw ?? string.Empty;
        }

        private static string Trim(string s, int max)
            => string.IsNullOrEmpty(s) ? s : (s.Length <= max ? s : s.Substring(0, max) + "…");
    }
}
