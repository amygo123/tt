using System;
using System.Collections.Generic;
using System.Linq;

namespace StyleWatcherWin
{
    public static class Aggregations
    {
        public sealed class SalesItem
        {
            public DateTime Date { get; set; }
            public string Size { get; set; } = "";
            public string Color { get; set; } = "";
            public int Qty { get; set; }
        }

        public sealed class CategoryAgg
        {
            public string Key { get; set; } = "";
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

        public static List<CategoryAgg> BySize(IEnumerable<SalesItem> records)
            => records.GroupBy(r => r.Size ?? "")
                      .Select(g => new CategoryAgg { Key = g.Key, Qty = g.Sum(x => x.Qty) })
                      .OrderByDescending(x => x.Qty).ToList();

        public static List<CategoryAgg> ByColor(IEnumerable<SalesItem> records)
            => records.GroupBy(r => r.Color ?? "")
                      .Select(g => new CategoryAgg { Key = g.Key, Qty = g.Sum(x => x.Qty) })
                      .OrderByDescending(x => x.Qty).ToList();

        public sealed class InventoryThresholds
        {
            public int DocRed { get; set; } = 3;
            public int DocYellow { get; set; } = 7;
            public int MinSalesWindowDays { get; set; } = 7;
        }

        public static double DailyAvgFromLast7(IEnumerable<SalesItem> records)
        {
            var last7 = records.Where(r => r.Date >= DateTime.Today.AddDays(-6)).Sum(r => r.Qty);
            return Math.Max(last7 / 7.0, 0.01);
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
