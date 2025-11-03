using System;
using System.Collections.Generic;
using System.Linq;

namespace StyleWatcherWin
{
    public static class Aggregations
    {
        // —— 数据模型 —— //
        public struct SalesItem
        {
            public DateTime Date;
            public string Size;
            public string Color;
            public int Qty;
        }

        // —— 趋势序列：补齐日期（含 0）并按日升序 —— //
        public static List<(DateTime day, int qty)> BuildDateSeries(IEnumerable<SalesItem> items, int windowDays)
        {
            var end = DateTime.Today;                   // 包含今日
            var start = end.AddDays(-(windowDays - 1)); // 向前 N-1 天
            var dict = items
                .Where(x => x.Date.Date >= start && x.Date.Date <= end)
                .GroupBy(x => x.Date.Date)
                .ToDictionary(g => g.Key, g => g.Sum(z => z.Qty));

            var res = new List<(DateTime, int)>();
            for (var d = start; d <= end; d = d.AddDays(1))
            {
                dict.TryGetValue(d.Date, out var q);
                res.Add((d.Date, q));
            }
            return res;
        }

        // —— 移动平均（不足长度或空时安全返回 0） —— //
        public static List<double> MovingAverage(IList<double> src, int n)
        {
            var result = new List<double>(src.Count);
            if (src.Count == 0 || n <= 1)
            {
                for (int i=0;i<src.Count;i++) result.Add(0);
                return result;
            }

            double sum = 0;
            for (int i=0;i<src.Count;i++)
            {
                sum += src[i];
                if (i >= n) sum -= src[i - n];
                if (i < n - 1) result.Add(0);
                else result.Add(sum / n);
            }
            return result;
        }

        // —— 数字格式化（K/M） —— //
        public static string FormatNumber(double v)
        {
            if (Math.Abs(v) >= 1_000_000) return (v/1_000_000d).ToString("0.##") + "M";
            if (Math.Abs(v) >= 1_000) return (v/1_000d).ToString("0.##") + "K";
            return v.ToString("0");
        }
    }
}
