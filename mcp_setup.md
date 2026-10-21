# MCP 注册与只读验证

> 日期：2026-10-14
> 状态：**已通过**。阶段 3 完成。

---

## 1. 注册信息

由用户之前通过 Claude CLI 注册，本次核对（`claude mcp get sqlite-readonly`）：

| 项目 | 值 |
|---|---|
| 名称 | `sqlite-readonly` |
| 范围 | User（所有项目可用） |
| 类型 | stdio |
| 命令 | `C:/Users/31451/.claude/venvs/skills/Scripts/python.exe -m mcp_sqlite.server` |
| 环境变量 | `SQLITE_DB_PATH=D:/work/实习/SQL项目/supply-chain-analysis/data/orders.db` |
| 健康检查 | ✔ Connected |

---

## 2. 只读保护（服务端源码核对）

`mcp_sqlite/db.py` 有两层保护：

1. 连接以 `file:...?mode=ro` 打开（操作系统级只读），并设置 `PRAGMA query_only`。
2. `run_query` 只接受单条 SELECT 或 WITH 语句，其他语句在进入数据库前被拒绝。

---

## 3. 只读验证结果（步骤 3.2）

| 步骤 | 操作 | 预期 | 结果 |
|---|---|---|---|
| 1 | `list_tables` | 包含 `orders` | 通过，返回 `["orders"]` |
| 2 | `describe_table orders` 与 `PRAGMA table_info(orders)` 对比 | 一致 | 通过，均为 53 列，字段名与类型一致 |
| 3 | `SELECT COUNT(*) FROM orders` | 180,519 | 通过，返回 180,519 |
| 4 | `DELETE FROM orders` | 被拒绝 | 通过，返回 `only read-only SELECT / WITH queries are allowed` |
| 5 | 删除尝试后再次 `COUNT(*)` | 仍为 180,519 | 通过，返回 180,519 |

**补充检查**

- `list_tables` 只列出表，不列出视图。通过 MCP 直接查询视图 `orders_v` 正常，返回 `2018-01` 2,123 行、`2017-12` 2,124 行，视图可用。
- 视图 `orders_v` 在 `sqlite_master` 中为 `view` 类型，与 `list_tables` 的结果不同是因为它只列出表。

---

## 4. 其他说明

- 当前会话中 `sqlite-readonly` 工具已可直接调用，未重启 Claude Code 即生效。如果新会话中找不到工具，再重启并检查 `claude mcp list`。
- 同一次 `claude mcp list` 中，`github` 插件连接失败（HTTP 400，授权头格式错误），`tavily` 工具超时。两者与本项目无关，未处理。
