
using System;
using System.Drawing;
using System.Windows.Forms;

namespace StyleWatcherWin
{
    /// <summary>
    /// A1 阶段占位：不包含真实库存逻辑，仅保证工程可编译。
    /// A2 会在此文件实现动态分仓、热力图与联动；当前不依赖任何 Aggregations.InventorySnapshot 类型。
    /// </summary>
    public class InventoryTabPage : TabPage
    {
        private readonly AppConfig _cfg;

#pragma warning disable CS0067 // 事件未使用（A1 占位）
        /// <summary>库存汇总完成事件（A2 使用）。</summary>
        public event Action? SummaryUpdated;
#pragma warning restore CS0067

        public InventoryTabPage(AppConfig cfg)
        {
            _cfg = cfg;
            Text = "库存（A2 开发中）";
            BackColor = Color.White;

            var lbl = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft YaHei UI", 10, FontStyle.Regular),
                Text = "库存页将在 A2 提供：\n• 动态分仓子页\n• 颜色×尺码热力图（可点击联动）\n• 按颜色/尺码的可用柱状图（降序）\n• 明细表与导出、搜索防抖\n\n当前为占位实现，用于确保 A1 构建通过。"
            };
            Controls.Add(lbl);
        }

        /// <summary>概览饼图点击时切换仓库（A2 实现）。A1 占位不做任何事。</summary>
        public void ShowWarehouse(string name) { /* no-op in A1 */ }

        /// <summary>返回库存汇总快照（A2 实现）。A1 返回 null，占位类型为 object?，不依赖 Aggregations.InventorySnapshot。</summary>
        public object? GetSummary() => null;
    }
}
