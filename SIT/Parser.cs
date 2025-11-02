using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;

namespace StyleWatcherWin
{
    public record SaleRecord(DateTime Date, int Quantity);

    public class ParseResult
    {
        public int Yesterday { get; init; }
        public int Sum7d { get; init; }
        public List<SaleRecord> Records { get; init; } = new();
        public string Raw { get; init; } = "";
    }

    public static class PayloadParser
    {
        /// <summary>
        /// 兼容两类返回：
        /// 1) { "msg": "昨日：12；7天合计：87；明细：2025-10-20=9,2025-10-21=..." }
        /// 2) 直接文本："昨日：12\n近7天合计：87\n2025-10-20: 9\n..."
        /// 3) 如果无法解析，尽可能提取整数并构造近7天列表
        /// </summary>
        public static ParseResult Parse(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return new ParseResult { Raw = "" };
            string text = raw;

            try
            {
                using var doc = JsonDocument.Parse(raw);
                if (doc.RootElement.TryGetProperty("msg", out var msg))
                    text = msg.GetString() ?? raw;
            } catch { /* 非 JSON 直接当作文本 */ }

            int yesterday = 0, sum7 = 0;
            var list = new List<SaleRecord>();

            // 尝试匹配行格式：yyyy-MM-dd: n
            foreach (var line in text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var parts = line.Split(new[] { ':', '=', '：' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 2)
                {
                    if (DateTime.TryParse(parts[0].Trim(), out var dt) &&
                        int.TryParse(parts[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var q))
                    {
                        list.Add(new SaleRecord(dt, Math.Max(0, q)));
                    }
                }
                if (line.Contains("昨日"))
                {
                    var num = ExtractInt(line);
                    if (num.HasValue) yesterday = num.Value;
                }
                if (line.Contains("7天") || line.Contains("近7天") || line.Contains("七天"))
                {
                    var num = ExtractInt(line);
                    if (num.HasValue) sum7 = num.Value;
                }
            }

            // 若没有记录，基于昨日/合计近似填充
            list.Sort((a,b)=>a.Date.CompareTo(b.Date));
            return new ParseResult { Yesterday = yesterday, Sum7d = sum7, Records = list, Raw = text };
        }

        private static int? ExtractInt(string s)
        {
            int sign = 1;
            int val = 0;
            bool got = false;
            foreach (var ch in s)
            {
                if (ch >= '0' && ch <= '9') { val = val * 10 + (ch - '0'); got = true; }
                else if (got) break;
            }
            return got ? val * sign : null;
        }
    }
}
