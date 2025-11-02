using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace StyleWatcherWin
{
    public static class ApiHelper
    {
        static readonly HttpClient _client = new HttpClient();

        public static async Task<string> QueryAsync(AppConfig cfg, string text, CancellationToken ct = default)
        {
            try
            {
                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(Math.Max(1, cfg.timeout_seconds)));
                using var linked = CancellationTokenSource.CreateLinkedTokenSource(cts.Token, ct);

                var method = (cfg.method ?? "POST").ToUpperInvariant();
                HttpRequestMessage req;
                if (method == "GET")
                {
                    var url = $"{cfg.api_url}?{cfg.json_key}={Uri.EscapeDataString(text)}";
                    req = new HttpRequestMessage(HttpMethod.Get, url);
                }
                else
                {
                    req = new HttpRequestMessage(HttpMethod.Post, cfg.api_url);
                    var payload = new Dictionary<string, string> { { cfg.json_key, text } };
                    var json = JsonSerializer.Serialize(payload);
                    req.Content = new StringContent(json, Encoding.UTF8, "application/json");
                }

                if (cfg.headers != null)
                {
                    foreach (var kv in cfg.headers)
                    {
                        if (string.Equals(kv.Key, "Content-Type", StringComparison.OrdinalIgnoreCase))
                        {
                            if (req.Content != null)
                                req.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(kv.Value);
                        }
                        else
                        {
                            if (!req.Headers.TryAddWithoutValidation(kv.Key, kv.Value))
                            {
                                if (req.Content != null)
                                    req.Content.Headers.TryAddWithoutValidation(kv.Key, kv.Value);
                            }
                        }
                    }
                }

                var resp = await _client.SendAsync(req, linked.Token);
                var raw = await resp.Content.ReadAsStringAsync(linked.Token);

                try
                {
                    using var doc = JsonDocument.Parse(raw);
                    if (doc.RootElement.TryGetProperty("msg", out var msgEl))
                        return Formatter.Prettify(msgEl.ToString());
                    return Formatter.Prettify(raw);
                }
                catch
                {
                    return Formatter.Prettify(raw);
                }
            }
            catch (Exception ex)
            {
                return $"请求失败：{ex.Message}";
            }
        }
    }
}
