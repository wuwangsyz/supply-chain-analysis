"""阶段 4 步骤 E + 阶段 5 步骤 5.2：生成 analysis.xlsx。

每题一个工作表（结果表 + 图表），另加"断点与断点后"和"总览"两个工作表。
数据来自 scripts\\03_validate.py 的 pandas 结果（已与 SQL、Excel 三方核对）。
每次运行都会重新生成 analysis.xlsx，因此不要在该文件中手工编辑。

用法：
    C:\\Users\\31451\\.claude\\venvs\\skills\\Scripts\\python.exe -I -X utf8 scripts\\05_build_analysis.py
"""

import importlib.util
import os

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font

OUT = "analysis.xlsx"

BREAK_NOTE = "断点（2017-10）破坏了量的可比性，没有破坏率的可比性。判定依据见 break_analysis.md。"


def load_validate_module():
    """直接加载 03_validate.py 中的函数，保证图表与核对使用同一份计算。"""
    path = os.path.join(os.path.dirname(__file__), "03_validate.py")
    spec = importlib.util.spec_from_file_location("validate", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def add_bar_chart(ws, title: str, value_col: int, label_col: int, n: int, anchor: str) -> None:
    """横向柱状图：数值取自 value_col，分类标签取自 label_col，首行为表头。"""
    chart = BarChart()
    chart.type = "bar"
    chart.title = title
    chart.y_axis.title = "延迟率（%）"
    chart.legend = None
    chart.add_data(Reference(ws, min_col=value_col, min_row=1, max_row=n + 1), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=label_col, min_row=2, max_row=n + 1))
    chart.height = 14
    chart.width = 20
    ws.add_chart(chart, anchor)


def overview_chart(src, dst, title: str, value_col: int, label_col: int, n: int, anchor: str, kind: str) -> None:
    """总览页的小图：数据区引用各题工作表的前 n 行，不在总览页重复放数据。"""
    chart = LineChart() if kind == "line" else BarChart()
    if kind != "line":
        chart.type = "bar"
    chart.title = title
    chart.legend = None
    chart.add_data(Reference(src, min_col=value_col, min_row=1, max_row=n + 1), titles_from_data=True)
    chart.set_categories(Reference(src, min_col=label_col, min_row=2, max_row=n + 1))
    chart.height = 8.5
    chart.width = 16
    dst.add_chart(chart, anchor)


def fmt(v):
    """表格数值写入：ID 和计数取整，比率保留 4 位小数。"""
    if isinstance(v, float):
        return int(v) if abs(v - round(v)) < 1e-9 else round(v, 4)
    return v


def write_q1(wb: Workbook, table) -> None:
    ws = wb.create_sheet("Q1 地区延迟")
    ws.append(["地区", "订单数", "延迟订单数", "订单级延迟率（%）"])
    for region, row in table.iterrows():
        ws.append([region, int(row["orders"]), int(row["late_orders"]), round(float(row["late_rate_pct"]), 2)])

    n = len(table)
    ws.cell(1, 5).value = "图表标签"
    for r in range(2, n + 2):
        # 分类标签带订单数，满足"标注订单数"的要求
        ws.cell(r, 5).value = f"{ws.cell(r, 1).value}（{ws.cell(r, 2).value:,}）"

    add_bar_chart(ws, "Q1 各地区订单级延迟率（括号内为订单数）", value_col=4, label_col=5, n=n, anchor="G2")
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["E"].width = 34
    return ws


def write_q2(wb: Workbook, table, sensitivity) -> None:
    ws = wb.create_sheet("Q2 运输方式")
    ws.append(["运输方式", "订单数", "平均计划天数", "平均实际天数", "订单级延迟率（%）", "图表标签"])
    for mode, row in table.iterrows():
        ws.append([
            mode,
            int(row["orders"]),
            round(float(row["avg_sched_days"]), 2),
            round(float(row["avg_real_days"]), 2),
            round(float(row["late_rate_pct"]), 2),
            f"{mode}（{int(row['orders']):,}）",
        ])

    n = len(table)
    add_bar_chart(ws, "Q2 各运输方式订单级延迟率（括号内为订单数）", value_col=5, label_col=6, n=n, anchor="A10")

    # 敏感性：排除已取消订单，放在表下方
    start = n + 4
    ws.cell(start, 1).value = "敏感性检查：排除已取消订单"
    ws.cell(start + 1, 1).value = "运输方式"
    ws.cell(start + 1, 2).value = "订单数"
    ws.cell(start + 1, 3).value = "订单级延迟率（%）"
    for i, (mode, row) in enumerate(sensitivity.iterrows(), start=start + 2):
        ws.cell(i, 1).value = mode
        ws.cell(i, 2).value = int(row["orders"])
        ws.cell(i, 3).value = round(float(row["late_rate_pct"]), 2)

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["F"].width = 28
    return ws


def write_q3(wb: Workbook, quarterly, monthly, scenarios) -> None:
    ws = wb.create_sheet("Q3 月度需求")
    ws.append(["季度（12 个完整季度）", "订单数", "延迟订单数", "订单级延迟率（%）", "标注"])
    for i, (quarter, row) in enumerate(quarterly.iterrows(), start=2):
        note = "记录结构断点：2017-10 起每单行数由约 3 降为 1，不可解读为需求增长" if quarter == "2017Q4" else ""
        ws.append([quarter, int(row["orders"]), int(row["late_orders"]), round(float(row["late_rate_pct"]), 2), note])

    n = len(quarterly)
    chart = LineChart()
    chart.title = "Q3 季度订单数（2017Q4 为记录结构断点，见下方说明）"
    chart.y_axis.title = "订单数"
    chart.legend = None
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=n + 1), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=n + 1))
    chart.height = 12
    chart.width = 20
    ws.add_chart(chart, "G2")

    # 月度记录结构：证明 2017Q4 的变化来自记录方式
    start = n + 4
    ws.cell(start, 1).value = "月度记录结构（SQL 与 pandas 一致）"
    header = ["月份", "行数", "订单数", "每单行数"]
    for j, h in enumerate(header, start=1):
        ws.cell(start + 1, j).value = h
    for i, (ym, row) in enumerate(monthly.iterrows(), start=start + 2):
        ws.cell(i, 1).value = ym
        ws.cell(i, 2).value = int(row["item_rows"])
        ws.cell(i, 3).value = int(row["orders"])
        ws.cell(i, 4).value = round(float(row["rows_per_order"]), 2)

    # 2017Q4 假设口径：仅影响订单数，延迟率不变
    s = start + 2 + len(monthly) + 2
    ws.cell(s, 1).value = "2017Q4 订单数口径（假设，见 metrics.md 第 4 节）"
    ws.cell(s + 1, 1).value = "口径"
    ws.cell(s + 1, 2).value = "订单数"
    ws.cell(s + 1, 3).value = "说明"
    for i, (label, row) in enumerate(scenarios.iterrows(), start=s + 2):
        ws.cell(i, 1).value = label
        ws.cell(i, 2).value = int(row["orders"])
        ws.cell(i, 3).value = row["basis"]

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["E"].width = 60
    return ws


def write_q4(wb: Workbook, table) -> None:
    ws = wb.create_sheet("Q4 品类")
    ws.append(["品类（断点前 2015-01 至 2017-09）", "含该品类订单数", "销售额", "延迟订单数", "订单级延迟率（%）", "图表标签"])
    for cat, row in table.iterrows():
        ws.append([
            cat,
            int(row["orders"]),
            round(float(row["sales"]), 2),
            int(row["late_orders"]),
            round(float(row["late_rate_pct"]), 2),
            f"{cat}（{int(row['orders']):,}）",
        ])

    n = len(table)
    # 图表取销售额前 15 的品类，避免 32 个柱子过密
    top = 15
    add_bar_chart(ws, "Q4 销售额前 15 品类的订单级延迟率（括号内为订单数）", value_col=5, label_col=6, n=top, anchor="H2")

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["F"].width = 34
    ws.cell(n + 3, 1).value = "注：一单可含多个品类，品类订单数加总大于窗口内订单数（57,349）。"
    ws.cell(n + 4, 1).value = "总体订单级延迟率（窗口内）：54.77%（57,349 单，31,412 单延迟），作为对照。"
    return ws


def write_break(wb: Workbook, compare, q4p, q4s, q1b) -> None:
    """断点分析页：两段结构对比、断点后品类记录值、假设 C、Q1 稳健性。"""
    ws = wb.create_sheet("断点与断点后")
    bold = Font(bold=True)
    r = 1

    ws.cell(r, 1).value = "1. 两段对比（pandas 复算，与 SQL 一致）"
    ws.cell(r, 1).font = bold
    r += 1
    ws.cell(r, 1).value = "指标"
    ws.cell(r, 2).value = "断点前 2015-01–2017-09"
    ws.cell(r, 3).value = "断点后 2017-10–2018-01"
    for c in (1, 2, 3):
        ws.cell(r, c).font = bold
    r += 1
    for metric, row in compare.iterrows():
        ws.cell(r, 1).value = metric
        ws.cell(r, 2).value = fmt(row["pre"])
        ws.cell(r, 3).value = fmt(row["post"])
        r += 1
    ws.cell(r, 1).value = (
        "注：item_rows、avg_qty、sales_per_order、id_gaps 是断点特有；markets、order_countries、categories "
        "的两段差异是并集口径造成的假象，已撤回（break_analysis.md 第 3 节）。"
    )
    r += 2

    ws.cell(r, 1).value = "2. Q4P 断点后各品类记录值（口径同 Q4，只换窗口）"
    ws.cell(r, 1).font = bold
    r += 1
    header = ["品类", "含该品类订单数", "销售额", "延迟订单数", "订单级延迟率（%）"]
    for j, h in enumerate(header, start=1):
        ws.cell(r, j).value = h
        ws.cell(r, j).font = bold
    r += 1
    for cat, row in q4p.iterrows():
        ws.cell(r, 1).value = cat
        ws.cell(r, 2).value = int(row["orders"])
        ws.cell(r, 3).value = round(float(row["sales"]), 2)
        ws.cell(r, 4).value = int(row["late_orders"])
        ws.cell(r, 5).value = round(float(row["late_rate_pct"]), 2)
        r += 1
    ws.cell(r, 1).value = (
        f"合计：品类订单数 {int(q4p['orders'].sum()):,}，窗口内订单数 8,403，差 127 来自 63 个含 2–5 行的订单。"
        "各品类样本 3–30 单，标准误约 9–29 个百分点，不出图、不作品类层面结论。"
    )
    r += 2

    ws.cell(r, 1).value = "3. Q4S 假设 C（只有 13 个断点前后都存在的品类有依据）"
    ws.cell(r, 1).font = bold
    r += 1
    header = ["品类", "断点后记录值", "假设 C 订单数", "断点前延迟率（%）", "断点后记录延迟率（%）"]
    for j, h in enumerate(header, start=1):
        ws.cell(r, j).value = h
        ws.cell(r, j).font = bold
    r += 1
    grounded = q4s[q4s["basis"] == "断点前有依据"]
    for cat, row in grounded.iterrows():
        ws.cell(r, 1).value = cat
        ws.cell(r, 2).value = int(row["post_recorded"])
        ws.cell(r, 3).value = int(row["assumed_post_orders"])
        ws.cell(r, 4).value = round(float(row["pre_late_rate_pct"]), 2)
        ws.cell(r, 5).value = round(float(row["post_late_rate_pct"]), 2)
        r += 1
    ws.cell(r, 1).value = (
        f"合计：记录值 {int(grounded['post_recorded'].sum()):,}，假设 C {int(grounded['assumed_post_orders'].sum()):,}。"
        "两个数不能直接比较——每单品类数从约 2.6 降到约 1，分母含义变了。"
        "可比的是同一窗口内的占比：13 个共有品类占全部（订单，品类）对的比例，断点前 90.77%，断点后 2.45%。"
    )
    ws.cell(r, 1).font = bold
    r += 2

    ws.cell(r, 1).value = "4. Q1 跨断点稳健性（两段都出现的 7 个地区）"
    ws.cell(r, 1).font = bold
    r += 1
    header = ["地区", "断点前订单数", "断点前延迟率（%）", "断点后订单数", "断点后延迟率（%）", "差值（百分点）"]
    for j, h in enumerate(header, start=1):
        ws.cell(r, j).value = h
        ws.cell(r, j).font = bold
    r += 1
    for region, row in q1b.iterrows():
        pre_rate, post_rate = float(row["pre_late_rate_pct"]), float(row["post_late_rate_pct"])
        ws.cell(r, 1).value = region
        ws.cell(r, 2).value = int(row["pre_orders"])
        ws.cell(r, 3).value = round(pre_rate, 2)
        ws.cell(r, 4).value = int(row["post_orders"])
        ws.cell(r, 5).value = round(post_rate, 2)
        ws.cell(r, 6).value = round(post_rate - pre_rate, 2)
        r += 1
    ws.cell(r, 1).value = (
        "注：本表按段落分窗口，断点前（33 个月）+ 断点后（4 个月）= Q1 页的全期数字，"
        "例如 Western Europe 8,395 + 1,615 = 10,010。两处不能直接混用。"
    )

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 26
    ws.column_dimensions["D"].width = 24
    ws.column_dimensions["E"].width = 24
    ws.column_dimensions["F"].width = 20


def write_overview(wb: Workbook, facts, sheets) -> None:
    """总览页：4 个问题的关键数字与结论，加 4 张引用各题数据区的图。每个数字都能在对应工作表找到。"""
    ws = wb.create_sheet("总览", 0)
    bold = Font(bold=True)
    ws.cell(1, 1).value = "供应链与配送履约分析：总览"
    ws.cell(1, 1).font = Font(bold=True, size=14)
    ws.cell(2, 1).value = (
        "数据：DataCo 公开供应链数据集，180,519 行订单项、65,752 个订单、37 个月（2015-01 至 2018-01）。"
        "公开样本，不是企业真实数据。口径见 metrics.md，核对见 validation.md，断点分析见 break_analysis.md。"
    )

    header = ["题号", "业务问题", "关键数字", "结论（一句话）", "限制条件"]
    for j, h in enumerate(header, start=1):
        ws.cell(4, j).value = h
        ws.cell(4, j).font = bold
    for i, row in enumerate(facts, start=5):
        for j, v in enumerate(row, start=1):
            ws.cell(i, j).value = v
        ws.cell(i, 1).font = bold

    ws.cell(10, 1).value = BREAK_NOTE
    ws.cell(10, 1).font = bold

    overview_chart(sheets["Q1"], ws, "总览 Q1：各地区延迟率（按订单数前 10）", 4, 5, 10, "G4", "bar")
    overview_chart(sheets["Q2"], ws, "总览 Q2：各运输方式延迟率", 5, 6, 4, "G22", "bar")
    overview_chart(sheets["Q3"], ws, "总览 Q3：季度订单数（2017Q4 为断点）", 2, 1, 12, "S4", "line")
    overview_chart(sheets["Q4"], ws, "总览 Q4：销售额前 10 品类的延迟率", 5, 6, 10, "S22", "bar")

    for col, width in {"A": 8, "B": 30, "C": 34, "D": 44, "E": 34}.items():
        ws.column_dimensions[col].width = width


def main() -> None:
    validate = load_validate_module()
    wb = Workbook()
    wb.remove(wb.active)

    q1, q2, q3 = validate.q1(), validate.q2(), validate.q3()
    q3m, q3s = validate.q3_monthly(), validate.q3_scenarios()
    q4 = validate.q4()
    compare = validate.break_compare()

    ws_q1 = write_q1(wb, q1)
    ws_q2 = write_q2(wb, q2, validate.q2_sensitivity())
    ws_q3 = write_q3(wb, q3, q3m, q3s)
    ws_q4 = write_q4(wb, q4)
    write_break(wb, compare, validate.q4(validate.POST_WINDOW), validate.q4_post_scenario(), validate.q1_break())

    # 总览的关键数字一律从上面的结果表里取，保证与工作表同源
    total_orders, total_late = int(q1["orders"].sum()), int(q1["late_orders"].sum())
    overall_rate = 100.0 * total_late / total_orders
    by_rate = q1.sort_values("late_rate_pct")
    low, high = by_rate.iloc[0], by_rate.iloc[-1]
    biggest = q1.sort_values("orders").iloc[-1]
    pre_break = q3.drop(index="2017Q4")
    pre_rate = float(compare.loc["late_rate_pct", "pre"])
    top5 = q4.head(5)
    top5_lo, top5_hi = float(top5["late_rate_pct"].min()), float(top5["late_rate_pct"].max())

    def m2(mode, col):
        return round(float(q2.loc[mode, col]), 2)

    facts = [
        ["Q1", "哪些地区延迟率高且订单量大",
         f"全期订单级延迟率 {overall_rate:.2f}%（{total_late:,} / {total_orders:,}）。"
         f"延迟率最高：{high.name} {float(high['late_rate_pct']):.2f}%（{int(high['orders']):,} 单）；"
         f"最低：{low.name} {float(low['late_rate_pct']):.2f}%（{int(low['orders']):,} 单）。"
         f"单量最大：{biggest.name} {int(biggest['orders']):,} 单、{float(biggest['late_rate_pct']):.2f}%",
         "23 个地区的延迟率跨度不到 8 个百分点，单量最大的地区只比总体高约 1 个百分点，"
         "看不出哪个地区明显更差；要按影响面排序，先看延迟订单数最大的地区",
         "延迟率最高和最低的地区样本都不足 700 单，标准误约 2 个百分点以上；"
         "该结论在断点后窗口重复成立（见断点与断点后页第 4 节）"],
        ["Q2", "各运输方式的承诺时效是否匹配",
         f"First Class：承诺 {q2.loc['First Class', 'avg_sched_days']:.0f} 天、实际 {q2.loc['First Class', 'avg_real_days']:.0f} 天，"
         f"延迟率 {m2('First Class', 'late_rate_pct')}%；"
         f"Second Class：承诺 {q2.loc['Second Class', 'avg_sched_days']:.0f} 天、实际 {q2.loc['Second Class', 'avg_real_days']:.0f} 天，"
         f"延迟率 {m2('Second Class', 'late_rate_pct')}%；"
         f"Standard Class：承诺 {q2.loc['Standard Class', 'avg_sched_days']:.0f} 天、实际 {q2.loc['Standard Class', 'avg_real_days']:.0f} 天，"
         f"延迟率 {m2('Standard Class', 'late_rate_pct')}%",
         "承诺天数越短的方式延迟率越高；First Class 与 Second Class 的承诺天数与实际所需天数不匹配，"
         "这两类的承诺应改为按实际天数设定",
         "延迟标志由实际天数与承诺天数之差决定，是定义性关系，不能当作因果；"
         "实际天数分布近似均匀，不像真实物流数据（validation.md Q2 第 5 节）"],
        ["Q3", "订单量在各季度如何波动，哪些季度要提前备料",
         f"2015Q1–2017Q3 的 11 个季度订单数在 {int(pre_break['orders'].min()):,}–{int(pre_break['orders'].max()):,} 之间，"
         f"延迟率 {float(pre_break['late_rate_pct'].min()):.2f}%–{float(pre_break['late_rate_pct'].max()):.2f}%；"
         f"2017Q4 记录值 {int(q3.loc['2017Q4', 'orders']):,}（较前一季度约 +19%），同期每单行数由 3.0 降为 1.0",
         "前期需求平稳、没有明显季节性，看不出需要按季度提前备料的窗口；"
         "2017Q4 的订单数上升与记录方式变化同时发生，不能当作需求高峰",
         f"2017Q4 订单数按口径不同落在 {int(q3s.loc['假设B', 'orders']):,}–{int(q3s.loc['记录值', 'orders']):,} 之间，"
         f"主口径取假设 A {int(q3s.loc['假设A', 'orders']):,}（是假设，不是观测）；2018 年只有 1 月"],
        ["Q4", "哪些品类量大、销售额高但履约差",
         f"断点前窗口总体延迟率 {pre_rate:.2f}%；销售额前 5 品类"
         f"（{', '.join(top5.index)}）延迟率 {top5_lo:.2f}%–{top5_hi:.2f}%，与总体差距都在 1 个百分点内",
         "没有量大且履约差的品类；延迟在各品类之间分布均匀，问题更可能在整体配送流程，而不是某个品类",
         "32 个品类同时比较，个别小样本品类偏离属偶然（最小的 Basketball 只有 56 单）；"
         "断点后品类结构变化，不参与排名（Q4P 各品类 3–30 单，不足以下结论）"],
    ]
    write_overview(wb, facts, {"Q1": ws_q1, "Q2": ws_q2, "Q3": ws_q3, "Q4": ws_q4})
    wb.save(OUT)
    print(f"已写入 {OUT}，工作表：{wb.sheetnames}")


if __name__ == "__main__":
    main()
