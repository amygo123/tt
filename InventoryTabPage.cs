using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StyleWatcherWin
{
    /// <summary>
    /// 库存页：自动按配置拉取默认款式；顶部无输入/刷新；
    /// 中部是一个 TabControl：第一个“汇总”，其后为各分仓 Tab（按总可用量降序排列）。
    /// </summary>
    public class InventoryTabPage : TabPage
    {
        private readonly AppConfig _cfg;
        private readonly TabControl _tabs = new TabControl();

        public InventoryTabPage(AppConfig cfg)
        {
            _cfg = cfg;
            Text = "库存";
            BackColor = System.Drawing.Color.White;

            BuildUI();
            _ = RefreshAsync();
        }

        private void BuildUI()
        {
            _tabs.Dock = DockStyle.Fill;
            Controls.Add(_tabs);
        }

        private async Task RefreshAsync()
        {
            try
            {
                var style = _cfg.inventory?.default_style ?? string.Empty;
                if (string.IsNullOrWhiteSpace(style))
                {
                    MessageBox.Show("appsettings.json 未配置 inventory.default_style。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var baseUrl = _cfg.inventory?.url_base ?? "";
                var url = baseUrl + Uri.EscapeDataString(style);

                using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(Math.Max(3, _cfg.timeout_seconds)) };
                using var req = new HttpRequestMessage(HttpMethod.Get, url);
                foreach (var kv in _cfg.headers)
                    req.Headers.TryAddWithoutValidation(kv.Key, kv.Value);

                var resp = await http.SendAsync(req);
                resp.EnsureSuccessStatusCode();
                var raw = await resp.Content.ReadAsStringAsync();

                // 解析
                var lines = CleanLines(raw).ToList();
                var rows = ParseRows(lines);

                // 构建动态子 Tab
                BuildDynamicTabs(rows);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"库存请求失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BuildDynamicTabs(List<Row> rows)
        {
            _tabs.TabPages.Clear();

            // 汇总：按 颜色+尺码 聚合
            var summary = rows
                .GroupBy(r => new { r.颜色, r.尺码 })
                .Select(g => new Row
                {
                    品名 = "汇总",
                    颜色 = g.Key.颜色,
                    尺码 = g.Key.尺码,
                    仓库 = "ALL",
                    可用 = g.Sum(x => x.可用),
                    现有 = g.Sum(x => x.现有)
                })
                .OrderByDescending(r => r.可用)
                .ToList();

            _tabs.TabPages.Add(CreateGridTab("汇总", summary));

            // 分仓：按仓库拆分，并按各仓总可用量降序排序 Tab 顺序
            var byStore = rows.GroupBy(r => r.仓库)
                              .Select(g => new { 仓库 = g.Key, 行 = g.ToList(), 总可用 = g.Sum(x => x.可用) })
                              .OrderByDescending(x => x.总可用)
                              .ToList();
            foreach (var st in byStore)
            {
                // 每个仓库内按 可用 降序
                var sorted = st.行.OrderByDescending(x => x.可用).ToList();
                _tabs.TabPages.Add(CreateGridTab(st.仓库, sorted));
            }
        }

        private TabPage CreateGridTab(string title, List<Row> data)
        {
            var page = new TabPage(title) { BackColor = System.Drawing.Color.White };
            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                DataSource = data
            };
            page.Controls.Add(grid);
            return page;
        }

        private static IEnumerable<string> CleanLines(string raw)
        {
            var items = new List<string>();
            try
            {
                var arr = JsonSerializer.Deserialize<string[]>(raw);
                if (arr != null) items.AddRange(arr);
            }
            catch
            {
                raw = raw.Replace("\r\n", "\n").Replace("\r", "\n");
                foreach (var ln in raw.Split('\n'))
                {
                    var s = ln.Trim();
                    if (!string.IsNullOrEmpty(s)) items.Add(s);
                }
            }

            return items
                .Select(s => s.Trim().Trim('"')
                    .Replace(", ", "，").Replace(",", " ，")) // 暂用窄空格占位，避免重复替换
                .Select(s => s.Replace(" ，", "，"))
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct()
                .OrderBy(s => s);
        }

        private static List<Row> ParseRows(IEnumerable<string> lines)
        {
            var result = new List<Row>();
            foreach (var ln in lines)
            {
                var parts = ln.Split('，').Select(x => x.Trim()).Where(x => x.Length > 0).ToArray();
                if (parts.Length < 4) continue; // 至少 品名/颜色/尺码/仓库
                int available = 0, onhand = 0;
                if (parts.Length >= 6)
                {
                    int.TryParse(parts[^2], out available);
                    int.TryParse(parts[^1], out onhand);
                }
                result.Add(new Row
                {
                    品名 = parts[0],
                    颜色 = parts.Length > 1 ? parts[1] : "",
                    尺码 = parts.Length > 2 ? parts[2] : "",
                    仓库 = parts.Length > 3 ? parts[3] : "",
                    可用 = available,
                    现有 = onhand
                });
            }
            return result;
        }

        private class Row
        {
            public string 品名 { get; set; } = "";
            public string 颜色 { get; set; } = "";
            public string 尺码 { get; set; } = "";
            public string 仓库 { get; set; } = "";
            public int 可用 { get; set; }
            public int 现有 { get; set; }
        }
    }
}
