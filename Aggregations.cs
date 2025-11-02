using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;

namespace StyleWatcherWin
{
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

        // ==== 库存 & 预警 ====
        public sealed class InventoryRow
        {
            public string 品名 { get; set; } = "";
            public string 颜色 { get; set; } = "";
            public string 尺码 { get; set; } = "";
            public string 仓库 { get; set; } = "";
            public int 可用 { get; set; }
            public int 现有 { get; set; }
        }

        public sealed class InventorySnapshot
        {
            public int 总可用 { get; set; }
            public int 总现有 { get; set; }
            public List<string> 尺码List { get; set; } = new();
            public List<string> 颜色List { get; set; } = new();
            public Dictionary<string,int> 分仓可用 { get; set; } = new();
            public Dictionary<(string 颜色,string 尺码),int> 汇总矩阵 { get; set; } = new();
        }

        public static InventorySnapshot BuildSnapshot(IEnumerable<InventoryRow> rows)
        {
            var list = rows.ToList();
            var snap = new InventorySnapshot
            {
                总可用 = list.Sum(x => x.可用),
                总现有 = list.Sum(x => x.现有),
            };
            snap.尺码List = list.Select(x => x.尺码 ?? "").Distinct().OrderBy(x => x).ToList();
            snap.颜色List = list.Select(x => x.颜色 ?? "").Distinct().OrderBy(x => x).ToList();
            snap.分仓可用 = list.GroupBy(x => x.仓库 ?? "").ToDictionary(g => g.Key, g => g.Sum(x => x.可用));
            snap.汇总矩阵 = list.GroupBy(x => (x.颜色 ?? "", x.尺码 ?? ""))
                               .ToDictionary(g => g.Key, g => g.Sum(x => x.可用));
            return snap;
        }

        public static List<(string 仓库,int 可用合计,List<InventoryRow> 行)> GroupByWarehouse(IEnumerable<InventoryRow> rows)
            => rows.GroupBy(r => r.仓库 ?? "")
                   .Select(g => (g.Key, g.Sum(x => x.可用), g.ToList()))
                   .OrderByDescending(t => t.Item2).ToList();

        public static Dictionary<string,int> ByColorInventory(IEnumerable<InventoryRow> rows)
            => rows.GroupBy(r => r.颜色 ?? "").ToDictionary(g => g.Key, g => g.Sum(x => x.可用));

        public static Dictionary<string,int> BySizeInventory(IEnumerable<InventoryRow> rows)
            => rows.GroupBy(r => r.尺码 ?? "").ToDictionary(g => g.Key, g => g.Sum(x => x.可用));

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

        // 颜色映射：根据值线性插值 0 -> 白，max -> 绿色
        public static Color ColorScale(int value, int max)
        {
            max = Math.Max(1, max);
            var ratio = Math.Max(0.0, Math.Min(1.0, value / (double)max));
            int r = (int)(240 - 120 * ratio);
            int g = (int)(255 - 40 * (1 - ratio));
            int b = (int)(240 - 200 * ratio);
            return Color.FromArgb(255, r, g, b);
        }
    }
}
