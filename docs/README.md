# 文档索引

> 路径约定：本文档目录内的相对路径以**仓库根目录**为基准（例如 `scripts/`、`data/`、`analysis.xlsx`）。

## 建议阅读顺序

| 顺序 | 文件 | 读它想知道什么 |
|---|---|---|
| 1 | [`metrics.md`](metrics.md) | 指标怎么定义的，为什么这么定义（含断点后的假设口径） |
| 2 | [`data_check.md`](data_check.md) | 数据进来时是什么样，做了哪些清洗决定（不删行） |
| 3 | [`validation.md`](validation.md) | SQL / pandas / Excel 三方怎么对，对到第几位，哪一条没对上 |
| 4 | [`break_analysis.md`](break_analysis.md) | 2017-10 断点的成因，哪些指标还能比、为什么 |
| 5 | [`ops_report.md`](ops_report.md) | 面向业务的结论、建议与执行顺序 |

## 过程记录（`process/`）

| 文件 | 内容 |
|---|---|
| [`process/项目计划.md`](process/项目计划.md) | 目标岗位、4 个问题、指标最小集、产出物与完成标准 |
| [`process/具体流程.md`](process/具体流程.md) | 分阶段的执行步骤：谁做、做什么、怎么检查 |
| [`process/ai_log.md`](process/ai_log.md) | 11 条 AI 协作记录：问题、输出、判断、修正 |
| [`process/mcp_setup.md`](process/mcp_setup.md) | `sqlite-readonly` MCP 的注册信息与只读验证结果 |

## 图表

`images/` 下的 6 张图由 `analysis.xlsx` 中的真实数值重绘（不重算指标），用于 README 展示。

| 文件 | 对应结论 |
|---|---|
| `fig1_region_delay.png` | Q1 各地区延迟率 |
| `fig2_shipping_mode.png` | Q2 运输方式承诺 vs 实际 |
| `fig3_quarterly_orders.png` | Q3 季度订单数与断点 |
| `fig4_rows_per_order.png` | 断点证据：每单行数 |
| `fig5_category_delay.png` | Q4 品类延迟率 |
| `fig6_robustness.png` | Q1 跨断点稳健性 |
