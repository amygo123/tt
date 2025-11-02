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
        public string Name { get; set; } = "";
        public string Size { get; set; } = "";
        public string Color { get; set; } = "";
        public int Qty { get; set; }
    }

    public class ParseResult
    {
        public string Title { get; set; } = "";
        public string Yesterday { get; set; } = "";
        public int? Sum7d { get; set; }
        public List<SaleRecord> Records { get; set; } = new List<SaleRecord>();
    }

    public static class PayloadParser
    {
        static readonly Regex RxLine = new Regex(
            @"(?<date>\d{4}[-/]\d{1,2}[-/]\d{1,2})?\s*(?<name>[^,，\s]+)\s*(?<size>[SMLX\d]+)?\s*(?<color>[^,，\s]+)?\s*(?<qty>\d+)",
            RegexOptions.Compiled);

        public static ParseResult Parse(string text)
        {
            var result = new ParseResult();
            if (string.IsNullOrWhiteSpace(text)) return result;

            foreach (Match m in RxLine.Matches(text))
            {
                var rec = new SaleRecord();
                if (DateTime.TryParse(m.Groups["date"].Value, out var dt)) rec.Date = dt.Date;
                rec.Name = m.Groups["name"].Value;
                rec.Size = m.Groups["size"].Value;
                rec.Color = m.Groups["color"].Value;
                if (int.TryParse(m.Groups["qty"].Value, out var q)) rec.Qty = q;
                if (!string.IsNullOrWhiteSpace(rec.Name) && rec.Qty > 0)
                    result.Records.Add(rec);
            }
            return result;
        }
    }
}
