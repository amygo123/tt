using System;
using System.Drawing;
using System.Windows.Forms;

namespace StyleWatcherWin
{
    /// <summary>
    /// A1 阶段的占位页：仅为保证工程可编译。
    /// A2 会在此文件中实现真实的库存页（动态子Tab、热力图、联动等）。
    /// 保留对外接口（事件、方法签名），便于 ResultForm 在 A2 无缝联动。
    /// </summary>
    public class InventoryTabPage : TabPage
    {
        private readonly AppConfig _cfg;

        /// <summary>
        /// 当库存汇总计算完成时触发（A2 实现时会真正触发；A1 占位不触发）。
        /// </summary>
        public event Action? SummaryUpdated;

        public InventoryTabPage(AppConfig cfg)
        {
            _cfg = cfg;
            Text = "库存（A2 开发中）";
            BackColor = Color.White;

            // A1 只放一个说明标签，防止空白与异常操作
            var lbl = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft YaHei UI", 10, FontStyle.Regular),
                Text = "库存页将在 A2 提供：\n• 动态分仓子页\n• 颜色×尺码热力图（可点击联动）\n• 按颜色/尺码的可用柱状图（降序）\n• 明细表与导出、搜索防抖\n\n当前为占位实现，用于确保 A1 构建通过。"
            };
            Controls.Add(lbl);
        }

        /// <summary>
        /// 供 ResultForm 在概览饼图点击时切换到对应仓库（A2 实现）。
        /// A1 占位不做任何事，保留签名避免编译错误。
        /// </summary>
        public void ShowWarehouse(string name)
        {
            // no-op in A1
        }

        /// <summary>
        /// 返回库存汇总快照（A2 实现）。A1 占位返回 null。
        /// 若你在 Aggregations 中已有 InventorySnapshot 类型，此签名与 A2 对齐。
        /// </summary>
        public Aggregations.InventorySnapshot? GetSummary() => null;

        // —— 如果你希望 A1 就显示一个“刷新”按钮以避免误操作，也可以加上 no-op —— //
        // private void DummyRefresh()
        // {
        //     // A1 不做网络/IO，避免引入失败路径
        //     // A2 中会执行 HTTP 请求与数据清洗，并在完成后触发 SummaryUpdated()
        // }
    }
}
