"""阶段 4 步骤 C：用 Excel 核对 Q4 的一个品类（Fishing，断点前窗口 2015-01 至 2017-09）。

核对内容：含该品类的订单数、订单级延迟订单数、延迟率、销售额。

做法：
1. 把 Order Id、Category Name、order date、Late_delivery_risk 复制到"Q4临时"工作表。
2. 按 Order Id、Category Name 排序（延迟降序），再按（订单，品类）去重。每个（订单，品类）保留一行，延迟取最大值。
3. 在"Q4"工作表用 COUNTIFS 计算订单数与延迟；销售额直接用 SUMIFS 在原表上计算（按订单项求和，无需去重）。

只新增工作表，不修改原有数据。

用法：
    C:\\Users\\31451\\.claude\\venvs\\skills\\Scripts\\python.exe -I -X utf8 scripts\\04_excel_q4.py "data\\工作簿2.xlsx"
"""

import os
import sys

import win32com.client as win32

TABLE = "DataCoSupplyChainDataset"
DATA_SHEET = "DataCoSupplyChainDataset"
CATEGORY = "Fishing"
# 窗口：pre = 断点前 2015-01 至 2017-09（默认）；post = 断点后 2017-10 至 2018-01
WINDOWS = {
    "pre": {"sheets": ("Q4临时", "Q4"), "start": "DATE(2015,1,1)", "end": "DATE(2017,10,1)"},
    "post": {"sheets": ("Q4P临时", "Q4P"), "start": "DATE(2017,10,1)", "end": "DATE(2018,2,1)"},
}
COLUMNS = ["Order Id", "Category Name", "order date (DateOrders)", "Late_delivery_risk"]
XL_ASCENDING = 1
XL_DESCENDING = 2
XL_UP = -4162


def main(path: str, window: str = "pre") -> None:
    tmp_name, result_name = WINDOWS[window]["sheets"]
    win_start, win_end = WINDOWS[window]["start"], WINDOWS[window]["end"]
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(path))
        names = [s.Name for s in wb.Worksheets]
        if tmp_name in names or result_name in names:
            raise SystemExit("目标工作表已存在，停止以免覆盖。")

        src = wb.Worksheets(DATA_SHEET)
        table = src.ListObjects(TABLE)
        tmp = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        tmp.Name = tmp_name

        # 1. 复制四列（含表头）到临时表 A:D
        for col_idx, col_name in enumerate(COLUMNS, start=1):
            table.ListColumns(col_name).Range.Copy(tmp.Cells(1, col_idx))
        last_row = tmp.Cells(tmp.Rows.Count, 1).End(XL_UP).Row

        # 2. 按（订单，品类）去重：排序后保留每组第一行（延迟为 1 优先）
        data = tmp.Range(f"A1:D{last_row}")
        data.Sort(
            Key1=tmp.Range("A1"), Order1=XL_ASCENDING,
            Key2=tmp.Range("B1"), Order2=XL_ASCENDING,
            Key3=tmp.Range("D1"), Order3=XL_DESCENDING,
            Header=1,
        )
        data.RemoveDuplicates(Columns=(1, 2), Header=1)
        last = tmp.Cells(tmp.Rows.Count, 1).End(XL_UP).Row

        # 3. 结果表
        ws = wb.Worksheets.Add(After=tmp)
        ws.Name = result_name
        cat_rng = f"'{tmp_name}'!$B$2:$B${last}"
        date_rng = f"'{tmp_name}'!$C$2:$C${last}"
        late_rng = f"'{tmp_name}'!$D$2:$D${last}"
        win = f'{date_rng},">="&{win_start},{date_rng},"<"&{win_end}'
        tbl_cat = f"{TABLE}[Category Name]"
        tbl_date = f"{TABLE}[order date (DateOrders)]"
        tbl_sales = f"{TABLE}[Sales]"

        ws.Range("A1").Value = "检查项"
        ws.Range("B1").Value = "Excel 结果"
        ws.Range("A2").Value = f"{CATEGORY} 含该品类订单数"
        ws.Range("B2").Formula = f'=COUNTIFS({cat_rng},"{CATEGORY}",{win})'
        ws.Range("A3").Value = f"{CATEGORY} 延迟订单数"
        ws.Range("B3").Formula = f'=COUNTIFS({cat_rng},"{CATEGORY}",{late_rng},1,{win})'
        ws.Range("A4").Value = f"{CATEGORY} 订单级延迟率（%）"
        ws.Range("B4").Formula = "=100*B3/B2"
        ws.Range("A5").Value = f"{CATEGORY} 销售额"
        ws.Range("B5").Formula = (
            f'=SUMIFS({tbl_sales},{tbl_cat},"{CATEGORY}",'
            f'{tbl_date},">="&{win_start},{tbl_date},"<"&{win_end})'
        )

        xl.CalculateFull()
        for r in range(2, 6):
            print(f"{ws.Cells(r, 1).Value}: {ws.Cells(r, 2).Value}")

        wb.Save()
        wb.Close(SaveChanges=True)
    finally:
        xl.Quit()


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "pre")
