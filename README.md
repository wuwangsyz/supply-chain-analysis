# 供应链与配送履约分析

> 目录位置：`D:\work\实习\SQL项目\supply-chain-analysis`
> 项目目标与岗位对齐见 `项目计划.md`，分步操作见 `具体流程.md`。本文件是对外说明：这个项目做了什么、结论是什么、局限在哪里、怎么复现。

---

## 1. 背景

物流运营里最常见的提问是：配送延迟出在哪个环节，影响多大，应该先改什么。本项目用一份公开的供应链订单数据，把这类提问拆成 4 个能用数据回答的问题，并用 SQL、pandas、Excel 三条独立路径算出同一批数字互相核对，最后给出可执行的建议和明确的局限。

目标不是做一个漂亮的报告，而是证明一件事：**每个结论都能追溯到某个查询的某个输出，且知道它在哪里不成立。**

---

## 2. 数据来源

| 项目 | 值 |
|---|---|
| 数据集 | DataCo Smart Supply Chain（公开样本，**不是企业真实数据**） |
| 原始文件 | `data\raw\DataCoSupplyChainDataset.csv`，约 96 MB，ISO-8859-1 编码，无 BOM，CRLF |
| 规模 | 180,519 行订单项 × 53 列，65,752 个订单，2015-01 至 2018-01 共 37 个月 |
| 行的含义 | 一行 = 一个订单项；同一订单可以有多行（断点前平均每单 3 行） |
| 分析库 | `data\orders.db`（表 `orders` + 视图 `orders_v`），只读访问 |
| 仓库内容 | 只含脚本、结论文档与 `analysis.xlsx`。原始 CSV、SQLite 库和 xlsx 备份体积过大未入库（见 `.gitignore`），复现时把原始 CSV 放回 `data\raw\` 再跑 `scripts\02_load.py` |

数据里有两个必须知道的特征：`Late_delivery_risk` 是由"实际天数 > 承诺天数"定义出来的标志（不是外部观测的迟到事件）；2017-10 起有一个记录结构断点。两者的影响见第 5 节。

---

## 3. 流程

```
CSV 体检（pandas + Excel 逐格比对）→ 清洗决定（不删行）
  → 导入 SQLite（断言：180,519 行 / 65,752 单 / 37 个月）
  → 定稿指标口径（metrics.md）
  → 注册 sqlite-readonly MCP，验证只读保护
  → 每个问题三条独立路径计算：agent SQL ／ pandas ／ Excel 公式
  → 三方比对记录（validation.md）
  → 记录结构断点分析（break_analysis.md）
  → 分析簿 analysis.xlsx（总览 + 每题一表一图）
  → 运营报告 ops_report.md
```

三方核对的通过标准：计数完全相等，比例差值不超过 0.01 个百分点，金额差值不超过 0.01。全部 4 个问题加断点后的补充核对都通过（`validation.md`）。

---

## 4. 四个结论

**Q1 地区配送延迟**：全期订单级延迟率 54.82%（36,048 / 65,752）。23 个地区的延迟率都在 49.5%–57.6% 之间，跨度不到 8 个百分点；单量最大的 Western Europe（10,010 单）只比总体高 0.97 个百分点，接近抽样波动。**没有哪个地区明显更差**，所以按延迟订单绝对量分配资源比按延迟率排序更有用（西欧 + 中美洲合计占全部延迟订单的 29.8%）。这个结论在断点后窗口重复成立。

**Q2 运输方式承诺时效**：承诺天数越短的方式延迟率越高——First Class 承诺 1 天、实际 2 天，延迟率 95.27%；Second Class 承诺 2 天、实际 4 天，76.72%；Standard Class 承诺 4 天、实际 4 天，38.13%。因为延迟标志就是由承诺与实际的差定义的，**这个对应关系是定义性的，不能读成"这些方式的运输能力差"**。真正可执行的动作是把 First Class 与 Second Class 的承诺天数按实际重设——承诺不改准，延迟率这个指标就没有信号。

**Q3 需求节奏**：断点前 11 个完整季度的订单数在 5,141–5,309 之间，延迟率 53.94%–55.51%，**没有季节性**。2017Q4 记录值 6,280（约 +19%）与记录结构变化同月发生，无法区分，主口径按"与前三季持平"（假设 A：5,195）处理。更重要的是，行数、订单数、件数、销售额四个需求代理在断点两侧方向都不一致（−59%、+21%、−81%、−39%），**这份数据不能用来做排产或备料参考**。

**Q4 品类需求与履约**：断点前窗口里销售额前 5 的品类延迟率 54.42%–55.09%，与总体（54.77%）差距都在 1 个百分点内，**不存在"量大、销售额高、履约又差"的品类**。延迟与品类无关，指向各品类共用的环节（配送时间承诺），与 Q2 落在同一个位置。延迟率最高的小品类（Golf Bags & Carts 68.85%、61 单）只是 2.4 个标准误的偏离，在 32 个品类同时比较下不构成异常。

**综合建议的顺序**：① 重设 Second Class 与 First Class 的承诺天数 → ② 按延迟订单数配置核查资源（西欧、中美洲）→ ③ 索取 2017-10 的记录方式说明后再谈需求分析 → ④ 品类层面不动作。详见 `ops_report.md`。

---

## 5. 局限

1. **公开样本数据**。实际天数在 2–6 天之间近似均匀分布、First Class 恒为 2 天，这个形态不像真实物流运营，更像生成数据。所有结论只说明这份数据集的结构，不能直接搬到真实业务。
2. **2018 年只有 1 月**，不做任何 2018 年的结论。
3. **2017-10 记录结构断点**：每单行数从 3.0 变 1.0，一个月内完成。它破坏了"量"的可比性（行数、订单数、金额、件数、品类构成），没有破坏"率"的可比性（延迟率、运输方式占比、折扣率）。原因是 `Late_delivery_risk` 实际按订单整体赋值（45,902 个多行订单中 0 个出现订单内标志不一致），所以每单行数怎么变都不影响延迟率。涉及量的结论一律限定在同一窗口内，判定过程见 `break_analysis.md`。
4. **没有因果结论**。延迟标志由承诺与实际天数之差定义，运输方式与延迟的关系是定义性的，不是因果。
5. **样本量限制解读**。品类层面只有 10 个品类达到 2,500 单（延迟率标准误约 1 个百分点）的量级；地区层面延迟率最高和最低的两个地区都不足 700 单。小样本的偏离本项目一律只作线索。
6. **口径特殊性不外推**：订单级与行级延迟率在本数据中只差 0.005 个百分点，这是"订单内标志完全一致"的结果，换一份数据不成立。

---

## 6. 使用的工具

| 工具 | 用途 |
|---|---|
| Claude Code / Claude 桌面端 + `sqlite-readonly` MCP | 写并执行 SQL（只读，服务端以 `mode=ro` + `PRAGMA query_only` 双层保护） |
| SQLite（`data\orders.db`） | 分析库，`orders` 表 + `orders_v` 视图 |
| Python（venv `skills`，pandas、openpyxl） | 体检、导入、独立复算、生成分析簿 |
| Excel（本机 16.0 build 19127，win32com 驱动） | 用 `COUNTA`／`COUNTIFS`／`SUMIFS`／去重向导独立复算，以及透视与图表 |
| Markdown | 口径、体检、核对、断点分析、报告 |

Excel 侧不使用动态数组函数：本机 Excel 16.0 不支持 `UNIQUE`（已实测），去重用"复制到临时表 + 删除重复项"。

---

## 7. 文件清单

| 文件 | 内容 |
|---|---|
| `项目计划.md` | 目标岗位、4 个问题、指标最小集、产出物、完成标准 |
| `具体流程.md` | 分阶段的操作步骤（谁做、做什么、怎么检查） |
| `data_check.md` | CSV 体检、Excel 逐格核对、清洗决定 |
| `metrics.md` | 指标口径（含 2017Q4 与断点后的假设口径） |
| `queries.sql` | 4 个问题的 SQL + 断点分析 1–8，注释写明口径 |
| `validation.md` | 三方核对记录：agent SQL vs pandas vs Excel |
| `break_analysis.md` | 2017-10 记录结构断点的成因分析与可比重判定 |
| `analysis.xlsx` | 总览 + Q1–Q4 每题一表一图 + 断点与断点后 |
| `ops_report.md` | 面向业务同事的结论与建议 |
| `ai_log.md` | 11 条 AI 协作记录：问题、输出、判断、修正 |
| `mcp_setup.md` | MCP 注册信息与只读验证结果 |
| `scripts\` | `01_profile`、`01_excel_check`、`02_load`、`03_validate`、`04_excel_q1..q4`、`05_build_analysis` |
| `data\` | 原始 CSV、`orders.db`、`工作簿2.xlsx`（及其备份，不对外上传） |

---

## 8. 如何复现

前置：Python venv 在 `C:\Users\31451\.claude\venvs\skills`，本机装有 Excel。以下命令都在本目录执行。

```bat
:: 1. 数据体检（只读 CSV）
C:\Users\31451\.claude\venvs\skills\Scripts\python.exe -I -X utf8 scripts\01_profile.py "data\raw\DataCoSupplyChainDataset.csv"

:: 2. 导入 SQLite（断言：180,519 行、65,752 单、37 个月，不符即报错退出）
C:\Users\31451\.claude\venvs\skills\Scripts\python.exe -I -X utf8 scripts\02_load.py "data\raw\DataCoSupplyChainDataset.csv" "data\orders.db"

:: 3. pandas 独立复算（Q1 Q2 Q2S Q3 Q3M Q3S Q4 Q4P Q4S BREAK Q1B）
C:\Users\31451\.claude\venvs\skills\Scripts\python.exe -I -X utf8 scripts\03_validate.py Q1

:: 4. Excel 侧复算（会向 工作簿2.xlsx 新增结果表，已存在则停止不覆盖）
C:\Users\31451\.claude\venvs\skills\Scripts\python.exe -I -X utf8 scripts\04_excel_q1.py "data\工作簿2.xlsx"
C:\Users\31451\.claude\venvs\skills\Scripts\python.exe -I -X utf8 scripts\04_excel_q4.py "data\工作簿2.xlsx" post

:: 5. 重建分析簿（覆盖 analysis.xlsx，不要在该文件里手工编辑）
C:\Users\31451\.claude\venvs\skills\Scripts\python.exe -I -X utf8 scripts\05_build_analysis.py
```

SQL 侧：注册 MCP 后直接执行 `queries.sql` 里的各段。注册命令（user 范围，所有项目可用）：

```bat
claude mcp add sqlite-readonly -s user -e SQLITE_DB_PATH=<本目录绝对路径>/data/orders.db -- C:\Users\31451\.claude\venvs\skills\Scripts\python.exe -m mcp_sqlite.server
```

约束：不修改 `data\raw\` 中的原始数据；不往 `D:\pycharm` 环境安装任何包；`analysis.xlsx` 由脚本重建。
