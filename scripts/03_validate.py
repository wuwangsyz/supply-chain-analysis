"""阶段 4 步骤 B：用 pandas 独立复算各题，只读 data\\orders.db 的 orders 表。

不读取 queries.sql，保证与 SQL 的计算方式相互独立。

用法：
    C:\\Users\\31451\\.claude\\venvs\\skills\\Scripts\\python.exe -I -X utf8 scripts\\03_validate.py Q1
"""

import sqlite3
import sys

import pandas as pd

DB = "data/orders.db"


def load_orders(columns: str) -> pd.DataFrame:
    con = sqlite3.connect(DB)
    try:
        return pd.read_sql(f"SELECT {columns} FROM orders", con)
    finally:
        con.close()


def q1() -> pd.DataFrame:
    """Q1：各地区订单数与订单级延迟率。订单级：同一订单内任一行延迟即记为延迟。"""
    df = load_orders('"Order Id", "Order Region", Late_delivery_risk')
    orders = df.groupby("Order Id").agg(
        region=("Order Region", "first"),
        late=("Late_delivery_risk", "max"),
        n_region=("Order Region", "nunique"),
    )
    # 同一订单的地区必须唯一，否则按地区分组会重复计数
    assert (orders["n_region"] == 1).all(), "存在多地区订单"

    result = orders.groupby("region").agg(
        orders=("late", "size"),
        late_orders=("late", "sum"),
    )
    result["late_rate_pct"] = 100.0 * result["late_orders"] / result["orders"]
    result.index = result.index.str.strip()
    return result.sort_values("orders", ascending=False)


def q2() -> pd.DataFrame:
    """Q2：各运输方式的订单数、平均计划/实际天数、订单级延迟率。每单取一行。"""
    df = load_orders(
        '"Order Id", "Shipping Mode", "Days for shipment (scheduled)", '
        '"Days for shipping (real)", Late_delivery_risk'
    )
    orders = df.groupby("Order Id").agg(
        mode=("Shipping Mode", "first"),
        sched=("Days for shipment (scheduled)", "first"),
        real=("Days for shipping (real)", "first"),
        late=("Late_delivery_risk", "max"),
    )
    result = orders.groupby("mode").agg(
        orders=("late", "size"),
        avg_sched_days=("sched", "mean"),
        avg_real_days=("real", "mean"),
        late_orders=("late", "sum"),
    )
    result["late_rate_pct"] = 100.0 * result["late_orders"] / result["orders"]
    return result.sort_values("orders", ascending=False)


def q2_sensitivity() -> pd.DataFrame:
    """Q2 敏感性检查：排除已取消订单（Shipping canceled）后的延迟率。"""
    df = load_orders('"Order Id", "Shipping Mode", "Delivery Status", Late_delivery_risk')
    orders = df.groupby("Order Id").agg(
        mode=("Shipping Mode", "first"),
        status=("Delivery Status", "first"),
        late=("Late_delivery_risk", "max"),
    )
    kept = orders[orders["status"] != "Shipping canceled"]
    result = kept.groupby("mode").agg(orders=("late", "size"), late_orders=("late", "sum"))
    result["late_rate_pct"] = 100.0 * result["late_orders"] / result["orders"]
    return result.sort_values("orders", ascending=False)


def q3_orders() -> pd.DataFrame:
    """Q3 基础表：每单一行，含下单月份与订单级延迟标志（同一订单下单月唯一，已验证）。

    月份由 pandas 从原始日期文本解析，不使用视图 orders_v，与 SQL 的解析互相独立。
    """
    df = load_orders('"Order Id", "order date (DateOrders)", Late_delivery_risk')
    df["ym"] = pd.to_datetime(df["order date (DateOrders)"], format="%m/%d/%Y %H:%M").dt.strftime("%Y-%m")
    return df.groupby("Order Id").agg(
        ym=("ym", "first"),
        late=("Late_delivery_risk", "max"),
    )


def q3() -> pd.DataFrame:
    """Q3 季度汇总：2015Q1–2017Q4 共 12 个完整季度，不含 2018-01。"""
    orders = q3_orders()
    orders = orders[(orders["ym"] >= "2015-01") & (orders["ym"] <= "2017-12")].copy()
    year = orders["ym"].str[:4]
    month = orders["ym"].str[5:7].astype(int)
    orders["quarter"] = year + "Q" + ((month - 1) // 3 + 1).astype(str)
    result = orders.groupby("quarter").agg(orders=("late", "size"), late_orders=("late", "sum"))
    result["late_rate_pct"] = 100.0 * result["late_orders"] / result["orders"]
    return result


def q3_monthly() -> pd.DataFrame:
    """Q3 补充：月度行数、订单数、每单行数。用于判断 2017-10 起的记录结构变化。"""
    df = load_orders('"Order Id", "order date (DateOrders)"')
    df["ym"] = pd.to_datetime(df["order date (DateOrders)"], format="%m/%d/%Y %H:%M").dt.strftime("%Y-%m")
    result = df.groupby("ym").agg(item_rows=("Order Id", "size"), orders=("Order Id", "nunique"))
    result["rows_per_order"] = result["item_rows"] / result["orders"]
    return result


def q3_scenarios() -> pd.DataFrame:
    """Q3 的 2017Q4 假设口径。这些都是假设，不是数据观测。

    - 记录值：数据中的 2017Q4 订单数（不做调整）。
    - 假设 A（需求持平）：2017Q4 订单数 = 2017Q1–Q3 季度平均。
    - 假设 B（按行数换算）：2017Q4 行数 ÷ 断点前每单平均行数（2015-01 至 2017-09）。
    """
    monthly = q3_monthly()
    q4 = ["2017-10", "2017-11", "2017-12"]
    pre = monthly.loc[:"2017-09"]
    rows_per_order_pre = pre["item_rows"].sum() / pre["orders"].sum()

    quarterly = q3()
    q4_recorded = int(quarterly.loc["2017Q4", "orders"])
    q4_rows = int(monthly.loc[q4, "item_rows"].sum())
    q123 = quarterly.loc[["2017Q1", "2017Q2", "2017Q3"], "orders"].mean()

    return pd.DataFrame(
        {
            "orders": [q4_recorded, round(q123), round(q4_rows / rows_per_order_pre)],
            "basis": [
                "记录值（不调整）",
                "假设 A：与 2017Q1–Q3 平均持平",
                f"假设 B：行数 {q4_rows:,} ÷ 前期每单行数 {rows_per_order_pre:.3f}",
            ],
        },
        index=["记录值", "假设A", "假设B"],
    )


PRE_WINDOW = ("2015-01", "2017-09")
POST_WINDOW = ("2017-10", "2018-01")


def q4(window: tuple[str, str] = PRE_WINDOW) -> pd.DataFrame:
    """Q4：各品类含该品类的订单数、销售额（按订单项求和）、订单级延迟率。

    默认窗口：2015-01 至 2017-09（断点前）。断点后用 POST_WINDOW。
    一单可含多个品类，品类订单数加总会大于窗口内订单数。品类名去掉尾随空格后再分组。
    """
    df = load_orders('"Order Id", "Category Name", Sales, Late_delivery_risk, "order date (DateOrders)"')
    df["ym"] = pd.to_datetime(df["order date (DateOrders)"], format="%m/%d/%Y %H:%M").dt.strftime("%Y-%m")
    df = df[(df["ym"] >= window[0]) & (df["ym"] <= window[1])].copy()
    df["category"] = df["Category Name"].str.strip()

    # 销售额：按订单项求和（每行只属于一个品类，不重复）
    sales = df.groupby("category")["Sales"].sum()

    # 订单数与延迟：按（订单，品类）去重，延迟取该单最大值
    pairs = df.groupby(["Order Id", "category"]).agg(late=("Late_delivery_risk", "max")).reset_index()
    result = pairs.groupby("category").agg(
        orders=("late", "size"),
        late_orders=("late", "sum"),
    )
    result["sales"] = sales
    result["late_rate_pct"] = 100.0 * result["late_orders"] / result["orders"]
    return result.sort_values("sales", ascending=False)


def q4_post_scenario() -> pd.DataFrame:
    """断点后品类的假设口径（与 Q3 同样的思路：记录值 + 假设）。

    假设 C：断点后的订单总数代表当前需求，但品类组成与断点前一致。
    - 品类订单数（假设）= 断点后总单数 × 断点前该品类占比（断点前含该品类订单数 ÷ 断点前总单数）。
    - 延迟率沿用断点前该品类的延迟率。
    只对断点前存在的品类做假设；断点后新出现的品类没有依据，不做假设。
    """
    pre = q4(PRE_WINDOW)
    post = q4(POST_WINDOW)
    df_pre =load_orders('"Order Id", "order date (DateOrders)"')
    df_pre["ym"] = pd.to_datetime(df_pre["order date (DateOrders)"], format="%m/%d/%Y %H:%M").dt.strftime("%Y-%m")
    pre_orders_total = int(df_pre[(df_pre["ym"] >= PRE_WINDOW[0]) & (df_pre["ym"] <= PRE_WINDOW[1])]["Order Id"].nunique())
    df_post = df_pre[(df_pre["ym"] >= POST_WINDOW[0]) & (df_pre["ym"] <= POST_WINDOW[1])]
    post_orders_total = int(df_post["Order Id"].nunique())

    result = pd.DataFrame(index=post.index)
    result["post_recorded"] = post["orders"]
    result["pre_orders"] = pre["orders"].reindex(result.index)
    result["pre_late_rate_pct"] = pre["late_rate_pct"].reindex(result.index)
    result["post_late_rate_pct"] = post["late_rate_pct"]
    result["assumed_post_orders"] = (
        post_orders_total * result["pre_orders"] / pre_orders_total
    ).round()
    result["basis"] = result["pre_orders"].notna().map({True: "断点前有依据", False: "断点后新品类，无依据"})
    return result.sort_values("post_recorded", ascending=False)


BREAK_MONTH = "2017-10"


def add_segment(df: pd.DataFrame) -> pd.DataFrame:
    """按下单月份加段落标记：2017-10 之前为 pre，起为 post。"""
    df = df.copy()
    df["ym"] = pd.to_datetime(df["order date (DateOrders)"], format="%m/%d/%Y %H:%M").dt.strftime("%Y-%m")
    df["seg"] = df["ym"].lt(BREAK_MONTH).map({True: "pre", False: "post"})
    return df


def break_compare() -> pd.DataFrame:
    """断点前后两段的规模、结构、履约对比。判定与结论见 break_analysis.md。

    月份在这里由原始日期文本解析，SQL 侧用的是视图 orders_v，两者相互独立。
    行为订单级指标时先按订单聚合，避免多行订单被重复计算。
    """
    df = add_segment(load_orders(
        '"Order Id", "Order Item Id", "Order Customer Id", "order date (DateOrders)", '
        '"Order Item Quantity", "Order Item Product Price", Sales, "Order Item Discount Rate", '
        '"Market", "Order Country", "Category Name", Late_delivery_risk, '
        '"Days for shipping (real)", "Days for shipment (scheduled)"'
    ))
    orders = df.groupby(["seg", "Order Id"]).agg(
        late=("Late_delivery_risk", "max"),
        real=("Days for shipping (real)", "first"),
        sched=("Days for shipment (scheduled)", "first"),
    )
    by_seg = df.groupby("seg")
    late_by_seg = orders.groupby("seg")
    result = pd.DataFrame(index=["pre", "post"])
    result["item_rows"] = by_seg.size()
    result["orders"] = late_by_seg.size()
    result["rows_per_order"] = result["item_rows"] / result["orders"]
    result["customers"] = by_seg["Order Customer Id"].nunique()
    result["orders_per_customer"] = result["orders"] / result["customers"]
    result["markets"] = by_seg["Market"].nunique()
    result["order_countries"] = by_seg["Order Country"].nunique()
    result["categories"] = by_seg["Category Name"].apply(lambda s: s.str.strip().nunique())
    result["avg_qty"] = by_seg["Order Item Quantity"].mean()
    result["avg_unit_price"] = by_seg["Order Item Product Price"].mean()
    result["avg_sales_per_row"] = by_seg["Sales"].mean()
    result["sales_per_order"] = by_seg["Sales"].sum() / result["orders"]
    result["avg_disc_rate"] = by_seg["Order Item Discount Rate"].mean()
    result["late_rate_pct"] = 100.0 * late_by_seg["late"].mean()
    result["avg_gap_days"] = late_by_seg["real"].mean() - late_by_seg["sched"].mean()
    result["min_item_id"] = by_seg["Order Item Id"].min()
    result["max_item_id"] = by_seg["Order Item Id"].max()
    result["min_order_id"] = by_seg["Order Id"].min()
    result["max_order_id"] = by_seg["Order Id"].max()
    result["id_gaps"] = result["max_order_id"] - result["min_order_id"] + 1 - result["orders"]
    return result.T


def q1_break() -> pd.DataFrame:
    """Q1 跨断点稳健性：两段都出现的地区，各自的订单数与订单级延迟率。"""
    df = add_segment(load_orders('"Order Id", "Order Region", Late_delivery_risk, "order date (DateOrders)"'))
    df["region"] = df["Order Region"].str.strip()
    orders = df.groupby(["seg", "Order Id"]).agg(region=("region", "first"), late=("Late_delivery_risk", "max"))
    pivot = orders.groupby(["region", "seg"]).agg(orders=("late", "size"), late_orders=("late", "sum"))
    pivot["late_rate_pct"] = 100.0 * pivot["late_orders"] / pivot["orders"]
    pre, post = pivot.xs("pre", level="seg"), pivot.xs("post", level="seg")
    result = pd.DataFrame({
        "pre_orders": pre["orders"],
        "pre_late_rate_pct": pre["late_rate_pct"],
        "post_orders": post["orders"],
        "post_late_rate_pct": post["late_rate_pct"],
    }).dropna()
    return result.sort_values("pre_orders", ascending=False)


QUESTIONS = {
    "Q1": q1,
    "Q4": q4,
    "Q4P": lambda: q4(POST_WINDOW),
    "Q4S": q4_post_scenario,
    "Q2": q2,
    "Q2S": q2_sensitivity,
    "Q3": q3,
    "Q3M": q3_monthly,
    "Q3S": q3_scenarios,
    "BREAK": break_compare,
    "Q1B": q1_break,
}

if __name__ == "__main__":
    pd.set_option("display.width", 200)
    table = QUESTIONS[sys.argv[1]]()
    print(table.to_string())
    if "orders" in table.columns:
        print(f"\n合计订单数: {int(table['orders'].sum()):,}")
    if "late_orders" in table.columns:
        print(f"合计延迟订单数: {int(table['late_orders'].sum()):,}")
