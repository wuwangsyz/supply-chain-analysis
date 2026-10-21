"""阶段 2 步骤 2.1：把 DataCo CSV 导入 SQLite，并建立视图 orders_v。

规则：
- 编码 ISO-8859-1，只读 CSV，不修改 data\\raw\\。
- 非空值全部为数字的列用 REAL，其余用 TEXT；空字符串存为 NULL。
- 表名 orders，字段名保留原样（含空格）。
- 视图 orders_v 增加 year_month 列（YYYY-MM），由 order date 的 M/D/YYYY 文本解析而来。
- 导入后断言：行数 180,519、去重订单数 65,752、orders_v 中月份数 37。任一不符则退出并报错。

用法：
    C:\\Users\\31451\\.claude\\venvs\\skills\\Scripts\\python.exe -I -X utf8 scripts\\02_load.py "data\\raw\\DataCoSupplyChainDataset.csv" "data\\orders.db"
"""

import csv
import os
import re
import sqlite3
import sys

csv.field_size_limit(sys.maxsize)

NUMERIC = re.compile(r"^-?\d+(\.\d+)?$")
ORDER_DATE = "order date (DateOrders)"


def infer_types(path: str) -> tuple[list[str], dict[str, str]]:
    """第一遍扫描：每列是否全部为数字（忽略空值）。返回表头和列类型。"""
    with open(path, encoding="ISO-8859-1", newline="") as f:
        reader = csv.reader(f)
        header = next(reader)
        numeric = {col: True for col in header}
        seen = {col: False for col in header}
        for row in reader:
            for col, val in zip(header, row):
                if val == "":
                    continue
                seen[col] = True
                if numeric[col] and not NUMERIC.match(val):
                    numeric[col] = False
    types = {col: "REAL" if numeric[col] and seen[col] else "TEXT" for col in header}
    return header, types


def load(csv_path: str, db_path: str) -> None:
    header, types = infer_types(csv_path)

    if os.path.exists(db_path):
        os.remove(db_path)
    con = sqlite3.connect(db_path)
    cols_sql = ", ".join(f'"{c}" {types[c]}' for c in header)
    con.execute(f"CREATE TABLE orders ({cols_sql})")

    placeholders = ", ".join("?" for _ in header)
    with open(csv_path, encoding="ISO-8859-1", newline="") as f:
        reader = csv.reader(f)
        next(reader)
        batch = []
        for row in reader:
            values = []
            for col, val in zip(header, row):
                if val == "":
                    values.append(None)
                elif types[col] == "REAL":
                    values.append(float(val))
                else:
                    values.append(val)
            batch.append(values)
            if len(batch) == 10000:
                con.executemany(f"INSERT INTO orders VALUES ({placeholders})", batch)
                batch.clear()
        if batch:
            con.executemany(f"INSERT INTO orders VALUES ({placeholders})", batch)

    # M/D/YYYY H:MM -> YYYY-MM：年在第二个 / 之后，月在第一个 / 之前
    d = f'"{ORDER_DATE}"'
    con.execute(
        f"""
        CREATE VIEW orders_v AS
        SELECT *,
               substr({d}, instr({d}, '/') + instr(substr({d}, instr({d}, '/') + 1), '/') + 1, 4)
               || '-' || printf('%02d', CAST(substr({d}, 1, instr({d}, '/') - 1) AS INTEGER)) AS year_month
        FROM orders
        """
    )
    con.commit()
    con.close()


def check(db_path: str) -> None:
    con = sqlite3.connect(db_path)
    rows = con.execute("SELECT COUNT(*) FROM orders").fetchone()[0]
    orders = con.execute('SELECT COUNT(DISTINCT "Order Id") FROM orders').fetchone()[0]
    months = con.execute("SELECT COUNT(DISTINCT year_month) FROM orders_v").fetchone()[0]
    con.close()

    print(f"行数: {rows:,}（期望 180,519）")
    print(f"去重订单数: {orders:,}（期望 65,752）")
    print(f"月份数: {months}（期望 37）")
    if (rows, orders, months) != (180519, 65752, 37):
        raise SystemExit("断言失败，导入结果与预期不一致。")
    print("断言通过。")


if __name__ == "__main__":
    load(sys.argv[1], sys.argv[2])
    check(sys.argv[2])
