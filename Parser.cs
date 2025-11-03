using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace StyleWatcherWin
{
    public class SaleRecord
    {
        public DateTime Date { get; set; }
        public string   Name { get; set; } = "";
        public string   Size { get; set; } = "";
        public string   Color{ get; set; } = "";
        public int      Qty  { get; set; }
    }

    public class ParsedPayload
    {
        public string Title { get; set; } = "";
        public string Yesterday { get; set; } = "";
        public int?   Sum7d { get; set; }
        public List<SaleRecord> Records { get; set; } = new List<SaleRecord>();
    }

    public static class PayloadParser
    {
        // “标题: 昨日...” 行
        static readonly Regex RxTitle = new Regex(
            @"^(?<title>.+?)(?:[:：]\s*)(?<yest>昨日[^\n]*)$",
            RegexOptions.Compiled);

        // “XXX 近7天销量汇总：12345”
        static readonly Regex RxSum = new Regex(
            @"(?<name>.+?)\s*近\s*7\s*天\s*销量\s*汇\s*总[:：]\s*(?<sum>\d+)$",
            RegexOptions.Compiled);

        // “yyyy-MM-dd 名称 尺码 颜色: 99件”
        static readonly Regex RxLine = new Regex(
            @"^(?<date>20\d{2}-\d{2}-\d{2})\s+(?<rest>.+?)\s*[:：]\s*(?<qty>\d+)\s*件$",
            RegexOptions.Compiled);

        public static ParsedPayload Parse(string raw)
        {
            var result = new ParsedPayload();
            if (string.IsNullOrWhiteSpace(raw)) return result;

            var text = raw.Replace("\\n", "\n").Replace("\r\n", "\n").Trim();

            // 清洗空白行
            var lines = new List<string>();
            foreach (var l in text.Split('\n'))
            {
                var t = l.Trim();
                if (t.Length > 0) lines.Add(t);
            }

            foreach (var line in lines)
            {
                // 标题 & 昨日
                var mTitle = RxTitle.Match(line);
                if (mTitle.Success && string.IsNullOrEmpty(result.Yesterday))
                {
                    result.Title = mTitle.Groups["title"].Value.Trim();
                    result.Yesterday = mTitle.Groups["yest"].Value.Trim();
                    continue;
                }

                // 7天汇总
                var mSum = RxSum.Match(line);
                if (mSum.Success && !result.Sum7d.HasValue)
                {
                    if (int.TryParse(mSum.Groups["sum"].Value, out var s))
                        result.Sum7d = s;
                    if (string.IsNullOrEmpty(result.Title))
                        result.Title = mSum.Groups["name"].Value.Trim();
                    continue;
                }

                // 明细
                var m = RxLine.Match(line);
                if (!m.Success) continue;

                var rest = m.Groups["rest"].Value.Trim();
                var tokens = rest.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();

                string name = rest, size = "", color = "";
                if (tokens.Count >= 2)
                {
                    color = tokens[^1];
                    size  = tokens[^2];

                    // 将末尾 “尺码 颜色” 从名称中剥离
                    var idx = rest.LastIndexOf(size + " " + color, StringComparison.Ordinal);
                    if (idx > 0) name = rest.Substring(0, idx).Trim();
                }

                // 统一空值
                if (string.Equals(size,  "null", StringComparison.OrdinalIgnoreCase))  size  = "";
                if (string.Equals(color, "null", StringComparison.OrdinalIgnoreCase)) color = "";

                if (!DateTime.TryParseExact(m.Groups["date"].Value, "yyyy-MM-dd",
                        CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                    continue;

                if (!int.TryParse(m.Groups["qty"].Value, out var qty)) qty = 0;

                result.Records.Add(new SaleRecord
                {
                    Date = dt, Name = name, Size = size, Color = color, Qty = qty
                });
            }
            return result;
        }
    }
    // Backward-compat wrapper for legacy references: StyleWatcherWin.Parser.Parse(...)
    public static class Parser
    {
        public static ParsedPayload Parse(string text) => PayloadParser.Parse(text);
    }
}
