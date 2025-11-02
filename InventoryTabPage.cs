using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StyleWatcherWin
{
    /// <summary>
    /// 库存页：调用 GET {cfg.inventory.url_base}{style}
    /// 左侧：清洗后的原始文本；右侧：结构化表格（品名/颜色/尺码/仓库/可用/现有）
    /// </summary>
    public class InventoryTabPage : TabPage
    {
        private readonly TextBox _txtQuery = new TextBox();
        private readonly Button _btnRefresh = new Button();
        private readonly RichTextBox _rawBox = new RichTextBox();
        private readonly DataGridView _grid = new DataGridView();

        private readonly AppConfig _cfg;

        public InventoryTabPage(AppConfig cfg)
        {
            _cfg = cfg;
            Text = "库存";
            BackColor = Color.White;

            BuildUI();
            _txtQuery.Text = cfg.inventory?.default_style ?? "";
            _ = RefreshAsync();
        }

        private void BuildUI()
        {
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(12) };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Controls.Add(root);

            var bar = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, Padding = new Padding(0), Margin = new Padding(0) };
            var lbl = new Label { Text = "款式路径：", AutoSize = true, Margin = new Padding(0, 8, 6, 0) };
            _txtQuery.Width = 420;
            _btnRefresh.Text = "刷新";
            _btnRefresh.Width = 72;
            _btnRefresh.Click += async (s, e) => await RefreshAsync();
            bar.Controls.Add(lbl);
            bar.Controls.Add(_txtQuery);
            bar.Controls.Add(_btnRefresh);
            root.Controls.Add(bar, 0, 0);

            var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, SplitterDistance = 520 };
            root.Controls.Add(split, 0, 1);

            // 左：清洗后原文
            _rawBox.ReadOnly = true;
            _rawBox.Dock = DockStyle.Fill;
            _rawBox.Font = new Font("Consolas", 10);
            split.Panel1.Controls.Add(_rawBox);

            // 右：结构化表格
            _grid.Dock = DockStyle.Fill;
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.AllowUserToDeleteRows = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            _grid.DataSource = new List<Row>();
            split.Panel2.Controls.Add(_grid);
        }

        private async Task RefreshAsync()
        {
            try
            {
                var style = _txtQuery.Text?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(style))
                {
                    MessageBox.Show("请输入款式路径，如：纯色/通纯棉圆领短T/黑/XL", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                // 清洗文本
                var lines = CleanLines(raw).ToList();
                _rawBox.Text = string.Join(Environment.NewLine, lines);

                // 结构化
                var rows = ParseRows(lines);
                _grid.DataSource = rows;
            }
            catch (Exception ex)
            {
                _rawBox.Text = $"请求失败：{ex.Message}";
                _grid.DataSource = new List<Row>();
            }
        }

        /// <summary>
        /// 兼容 JSON 数组 / 纯文本，做初步清洗：去空行、去重、统一中文逗号、兼容半角空格、排序
        /// </summary>
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
                    .Replace(", ", "，").Replace(",", "，")
                    .Replace(" ,", "，"))
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct()
                .OrderBy(s => s);
        }

        /// <summary>
        /// 解析“品名，颜色，尺码，仓库，可用，现有”，并容错
        /// </summary>
        private static List<Row> ParseRows(IEnumerable<string> lines)
        {
            var result = new List<Row>();
            foreach (var ln in lines)
            {
                var parts = ln.Split('，').Select(x => x.Trim()).Where(x => x.Length > 0).ToArray();
                if (parts.Length < 4) continue; // 起码有 品名/颜色/尺码/仓库
                int available = 0, onhand = 0;
                if (parts.Length >= 6)
                {
                    int.TryParse(parts[parts.Length - 2], out available);
                    int.TryParse(parts[parts.Length - 1], out onhand);
                }
                var name = parts[0];
                var color = parts.Length > 1 ? parts[1] : "";
                var size = parts.Length > 2 ? parts[2] : "";
                var store = parts.Length > 3 ? parts[3] : "";
                result.Add(new Row { 品名 = name, 颜色 = color, 尺码 = size, 仓库 = store, 可用 = available, 现有 = onhand });
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
