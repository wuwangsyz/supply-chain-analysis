"""阶段 1 步骤 1.1：只读体检 DataCo CSV，输出到标准输出。

用法：
    python.exe -I scripts\\01_profile.py "data\\raw\\DataCoSupplyChainDataset.csv"
"""

import re
import sys

import pandas as pd


def date_pattern(value: str) -> str:
    """把日期字符串替换成格式模板，例如 2/3/2015 0:00 -> D/D/DDDD D:DD。"""
    return re.sub(r"\d", "D", value)


def main(path: str) -> None:
    # 全部按字符串读入，避免 pandas 自动转换掩盖格式问题
    df = pd.read_csv(path, encoding="ISO-8859-1", dtype=str, low_memory=False)

    print(f"行数: {len(df):,}")
    print(f"列数: {df.shape[1]}")

    print("\n== 每列空值数 ==")
    nulls = df.isna().sum()
    for col, n in nulls.items():
        if n:
            print(f"{col}: {n:,} ({100.0 * n / len(df):.2f}%)")
    print(f"有空值的列数: {(nulls > 0).sum()}")

    print("\n== 主键 ==")
    print(f"Order Item Id 唯一值数: {df['Order Item Id'].nunique():,}")
    print(f"Order Item Id 重复行数: {df['Order Item Id'].duplicated().sum():,}")
    print(f"Order Id 去重订单数: {df['Order Id'].nunique():,}")
    dup_rows = df.duplicated().sum()
    print(f"完全重复的行数: {dup_rows:,}")

    print("\n== 日期格式分布 ==")
    for col in ["order date (DateOrders)", "shipping date (DateOrders)"]:
        patterns = df[col].dropna().map(date_pattern).value_counts()
        print(f"{col}:")
        for pat, n in patterns.items():
            print(f"  {pat}: {n:,}")
        # 首段（月或日）的最大值；若超过 12，说明不是 M/D 顺序
        first_field = df[col].dropna().str.split("/").str[0].astype(int)
        print(f"  首段最大值: {first_field.max()}（超过 12 则不是 M/D 顺序）")

    print("\n== Late_delivery_risk × Delivery Status ==")
    print(pd.crosstab(df["Late_delivery_risk"], df["Delivery Status"]).to_string())
    inconsistent = (df["Late_delivery_risk"] == "1") & (df["Delivery Status"] != "Late delivery")
    inconsistent |= (df["Late_delivery_risk"] == "0") & (df["Delivery Status"] == "Late delivery")
    print(f"标志与状态不一致的行数: {inconsistent.sum():,}")

    print("\n== 订单级延迟（按 Order Id 聚合）==")
    late_orders = df.assign(late=df["Late_delivery_risk"] == "1").groupby("Order Id")["late"].any()
    print(f"延迟订单数: {late_orders.sum():,} / 订单数: {len(late_orders):,}")


if __name__ == "__main__":
    main(sys.argv[1])
