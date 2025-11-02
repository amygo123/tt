using System;
using System.Collections.Generic;
using System.Linq;

namespace StyleWatcherWin
{
    /// <summary>
    /// 纯计算：销量/库存聚合、补零、移动平均、DoC、预警
    /// 与 UI 解耦，便于单测与复用
    /// </summary>
    public static class Aggregations
    {
        // ==== 销售 ====
        public sealed class SalesItem
        {
            public DateTime Date { get; set; }
            public string Size { get; set; } = "";
            public string Color { get; set; } = "";
            public int Qty { get; set; }
        }

        public static List<(DateTime day, int qty)> BuildDateSeries(IEnumerable<SalesItem> records, int windowDays, DateTime? endDate = null)
        {
            var end = (endDate ?? DateTime.Today).Date;
            var start = end.AddDays(1 - windowDays).Date;

            var dict = records
                .GroupBy(r => r.Date.Date)
                .ToDictionary(g => g.Key, g => g.Sum(x => x.Qty));

            var series = new List<(DateTime, int)>();
            for (var d = start; d <= end; d = d.AddDays(1))
            {
                dict.TryGetValue(d, out var qty);
                series.Add((d, qty));
            }
            return series;
        }

        public static List<double> MovingAverage(IReadOnlyList<int> values, int span = 7)
        {
            if (values.Count == 0 || span <= 1) return values.Select(v => (double)v).ToList();
            var res = new double[values.Count];
            double sum = 0;
            var q = new Queue<int>();
            for (int i = 0; i < values.Count; i++)
            {
                q.Enqueue(values[i]); sum += values[i];
                while (q.Count > span) sum -= q.Dequeue();
                res[i] = sum / q.Count;
            }
            return res.ToList();
        }

        public static List<(string key, int qty)> BySize(IEnumerable<SalesItem> records)
            => records.GroupBy(r => r.Size ?? "")
                      .Select(g => (g.Key, g.Sum(x => x.Qty)))
                      .OrderByDescending(x => x.Item2).ToList();

        public static List<(string key, int qty)> ByColor(IEnumerable<SalesItem> records)
            => records.GroupBy(r => r.Color ?? "")
                      .Select(g => (g.Key, g.Sum(x => x.Qty)))
                      .OrderByDescending(x => x.Item2).ToList();

        // ==== 库存 & 预警（骨架，提交2补全联动） ====
        public sealed class InventoryThresholds
        {
            public int DocRed { get; set; } = 3;
            public int DocYellow { get; set; } = 7;
            public int MinSalesWindowDays { get; set; } = 7;
        }

        public static double DailyAvgFromLast7(IEnumerable<SalesItem> records)
        {
            var last7 = records.Where(r => r.Date >= DateTime.Today.AddDays(-6)).Sum(r => r.Qty);
            return Math.Max(last7 / 7.0, 0.01); // 防除零
        }

        public static int DaysOfCover(int available, double dailyAvg)
        {
            if (dailyAvg <= 0) return int.MaxValue;
            return (int)Math.Ceiling(available / dailyAvg);
        }

        public enum AlertLevel { Green, Yellow, Red, Unknown }

        public static AlertLevel LevelFromDoc(int doc, InventoryThresholds t)
        {
            if (doc == int.MaxValue) return AlertLevel.Unknown;
            if (doc < t.DocRed) return AlertLevel.Red;
            if (doc < t.DocYellow) return AlertLevel.Yellow;
            return AlertLevel.Green;
        }
    }
}
