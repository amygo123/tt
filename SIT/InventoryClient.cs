using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace StyleWatcherWin
{
    public class InventoryClient
    {
        private readonly HttpClient _http;
        private readonly AppConfig _cfg;

        public InventoryClient(HttpClient http, AppConfig cfg)
        {
            _http = http; _cfg = cfg;
        }

        public async Task<List<InventoryItem>> FetchAsync(string code)
        {
            var url = _cfg.inventory_api_url ?? "";
            if (string.IsNullOrWhiteSpace(url)) return new List<InventoryItem>();
            try
            {
                var res = await _http.GetAsync(url + "?q=" + Uri.EscapeDataString(code));
                res.EnsureSuccessStatusCode();
                var raw = await res.Content.ReadAsStringAsync();
                var list = JsonSerializer.Deserialize<List<InventoryItem>>(raw, new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? new();
                return list;
            }
            catch
            {
                return new List<InventoryItem>();
            }
        }
    }
}
