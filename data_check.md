# 数据体检记录

> 状态：**已定稿**（2026-10-14，清洗决定已由用户确认）。阶段 1 的 CSV 体检和 Excel 核对均已通过。
> 数据来源：`data\raw\DataCoSupplyChainDataset.csv`（公开样本，DataCo Smart Supply Chain）。以 CSV 为准。
> 脚本：`scripts\01_profile.py`（CSV 体检）、`scripts\01_excel_check.py`（Excel 核对）。

---

## 1. CSV 体检（步骤 1.1，已完成）

运行：`python.exe -I -X utf8 scripts\01_profile.py "data\raw\DataCoSupplyChainDataset.csv"`

| 检查项 | 结果 | 与计划预期 |
|---|---|---|
| 行数 | 180,519 | 一致 |
| 列数 | 53 | 一致 |
| `Order Item Id` 唯一值 / 重复行 | 180,519 / 0 | 一致 |
| `Order Id` 去重订单数 | 65,752 | 一致 |
| 完全重复的行 | 0 | — |
| `Product Description` 空值 | 180,519（100%） | 一致，不用于分析 |
| `Order Zipcode` 空值 | 155,679（86.24%） | 一致，不用于分析 |
| `Customer Lname` 空值 | 8 | 新发现，不影响本项目 |
| `Customer Zipcode` 空值 | 3 | 新发现，不影响本项目 |
| `Late_delivery_risk` 与 `Delivery Status` 不一致行 | 0 | 一致 |
| `Late_delivery_risk=1` 行数 | 98,977（等于 `Late delivery` 行数） | 一致 |
| 订单级延迟订单数 | 36,048 / 65,752（54.8%） | 新增，供阶段 4 核对 |

**CSV 格式**：ISO-8859-1 编码，无 BOM，CRLF 换行，逗号分隔，字段不加引号。表头 53 列，180,519 行每行都正好 53 个字段。

**日期格式**：`order date` 与 `shipping date` 为 `M/D/YYYY H:MM`，月和日位数不固定（如 `1/31/2018 22:56`）。首段最大值为 12，确认为月在前。

**Late_delivery_risk × Delivery Status**：

| | Advance shipping | Late delivery | Shipping canceled | Shipping on time |
|---|---|---|---|---|
| 0 | 41,592 | 0 | 7,754 | 32,196 |
| 1 | 0 | 98,977 | 0 | 0 |

---

## 2. Excel 核对（步骤 1.2，已通过）

工作簿：`data\工作簿2.xlsx`（按 ISO-8859-1 正确导入）。原文件已备份为 `data\工作簿2_备份.xlsx`（`cmp` 确认一致）。
运行：`python.exe -I -X utf8 scripts\01_excel_check.py "data\工作簿2.xlsx"`
新增工作表：`体检核对`（3 项检查）、`去重临时`（Order Id 去重的中间表）。原有数据未改动。

| 检查项 | 预期（CSV） | Excel 结果 | 结论 |
|---|---|---|---|
| 订单项数 `=COUNTA(DataCoSupplyChainDataset[Order Item Id])` | 180,519 | 180,519 | 通过 |
| 去重订单数（`去重临时` 去重后计数） | 65,752 | 65,752 | 通过 |
| 标志不一致行数 `=COUNTIFS(...)` | 0 | 0 | 通过 |

**方法说明**：计划原定去重用 `=ROWS(UNIQUE(...))`。本机 Excel 16.0（build 19127）不支持 `UNIQUE`（已在空白工作簿中实测，`Formula2` 报错）。改为把 `Order Id` 复制到 `去重临时` 表，用 Excel 的"删除重复项"去重，再 `=COUNTA(...)-1` 计数。

**工作簿2 与 CSV 的逐格比对**（独立于 Excel 公式，用 Python 完成）：

- 表头、180,519 行、53 列全部一致。
- 文本列完全一致（包括 `Chiriquí` 和尾随空格），没有编码替换。
- 数值列在 1e-6 容差内一致。
- 两列日期与 CSV 时间全部一一对应。Excel 将其存为日期类型，格式不同，但时间值一致。

---

## 3. 历史记录：工作簿1 的问题（已删除）

早先使用的 `工作簿1.xlsx` 与 CSV 不一致：`Order Item Id` 有 5,268 个单元格被价格类小数覆盖，导致去重订单数为 63,663。该文件已删除。`工作簿1_备份.xlsx` 作为原始副本保留，**不作为数据源**。

原因未查明。CSV 的分隔和引号结构正常，因此不是 CSV 格式导致的。工作簿1 的错位可能发生在导入或后续编辑过程中。

---

## 4. 已知数据问题（来自计划，已由体检确认的标注）

- `Order Profit Per Order` 是订单级值，**按订单项求和会重复计算**。需要利润时，按订单取一行。本项目以销售额为主，利润不作主要指标。
- 2018 年只有 1 月（阶段 2 导入后用 `orders_v` 核实月份数，目前尚未验证）。
- `Product Description` 全空，不用于分析。
- `Order Zipcode` 86.24% 为空，不用于分析。
- 延迟标志 `Late_delivery_risk` 与 `Delivery Status` 完全对应，属于定义性关系，**不能用于推断因果**。

---

## 5. 清洗决定（已确认，2026-10-14）

- **本项目不删除任何行，不修改 `data\raw\` 中的原始数据。**
- 空值字段保留，不作为分析字段。
- 日期列在 Excel 中为日期类型，在 SQL 和 pandas 中按 `%m/%d/%Y %H:%M` 解析。

---

## 6. 阶段 1 结论

- 阶段 1 完成：CSV 体检、Excel 核对（工作簿2）、清洗决定均已完成并确认。
- 阶段 2 导入脚本直接读取 `data\raw\` 的 CSV（ISO-8859-1），不生成额外副本。
