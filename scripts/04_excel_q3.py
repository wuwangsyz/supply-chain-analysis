"""阶段 4 步骤 C：用 Excel 核对 Q3 的一个数字（2017Q4 的订单数、延迟订单数、延迟率）。

选择 2017Q4 是因为它是季度订单数的最高值，也是数据记录方式发生变化的时期（见 validation.md）。

做法：
1. 把 Order Id、order date、Late_delivery_risk 复制到"Q3临时"工作表。
2. 按 Order Id 升序、延迟降序排序，再按 Order Id 去重，得到每单一行（同一订单的下单日期唯一，已验证）。
3. 在"Q3"工作表用 COUNTIFS 按日期区间（2017-10-01 至 2017-12-31）计算订单数、延迟订单数和延迟率。

只新增工作表，不修改原有数据。

用法：
    C:\\Users\\31451\\.claude\\venvs\\skills\\Scripts\\python.exe -I -X utf8 scripts\\04_excel_q3.py "data\\工作簿2.xlsx"
"""

import os
import sys

import win32com.client as win32

TABLE = "DataCoSupplyChainDataset"
DATA_SHEET = "DataCoSupplyChainDataset"
TMP_SHEET = "Q3临时"
RESULT_SHEET = "Q3"
COLUMNS = ["Order Id", "order date (DateOrders)", "Late_delivery_risk"]
XL_ASCENDING = 1
XL_DESCENDING = 2
XL_UP = -4162


def main(path: str) -> None:
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(path))
        names = [s.Name for s in wb.Worksheets]
        if TMP_SHEET in names or RESULT_SHEET in names:
            raise SystemExit("目标工作表已存在，停止以免覆盖。")

        table = wb.Worksheets(DATA_SHEET).ListObjects(TABLE)
        tmp = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        tmp.Name = TMP_SHEET

        # 1. 复制三列（含表头）到临时表 A:C
        for col_idx, col_name in enumerate(COLUMNS, start=1):
            table.ListColumns(col_name).Range.Copy(tmp.Cells(1, col_idx))
        last_row = tmp.Cells(tmp.Rows.Count, 1).End(XL_UP).Row

        # 2. 排序后按 Order Id 去重，每单保留延迟为 1 的行（如有）
        data = tmp.Range(f"A1:C{last_row}")
        data.Sort(
            Key1=tmp.Range("A1"), Order1=XL_ASCENDING,
            Key2=tmp.Range("C1"), Order2=XL_DESCENDING,
            Header=1,
        )
        data.RemoveDuplicates(Columns=1, Header=1)

        # 3. 结果表：2017Q4 = 2017-10-01 至 2017-12-31
        last = tmp.Cells(tmp.Rows.Count, 1).End(XL_UP).Row
        date_rng = f"'{TMP_SHEET}'!$B$2:$B${last}"
        late_rng = f"'{TMP_SHEET}'!$C$2:$C${last}"
        start = '">="&DATE(2017,10,1)'
        end = '"<"&DATE(2018,1,1)'

        ws = wb.Worksheets.Add(After=tmp)
        ws.Name = RESULT_SHEET
        ws.Range("A1").Value = "检查项"
        ws.Range("B1").Value = "Excel 结果"
        ws.Range("A2").Value = "2017Q4 订单数"
        ws.Range("B2").Formula = f"=COUNTIFS({date_rng},{start},{date_rng},{end})"
        ws.Range("A3").Value = "2017Q4 延迟订单数"
        ws.Range("B3").Formula = f"=COUNTIFS({date_rng},{start},{date_rng},{end},{late_rng},1)"
        ws.Range("A4").Value = "2017Q4 订单级延迟率（%）"
        ws.Range("B4").Formula = "=100*B3/B2"

        xl.CalculateFull()
        for r in range(2, 5):
            print(f"{ws.Cells(r, 1).Value}: {ws.Cells(r, 2).Value}")

        wb.Save()
        wb.Close(SaveChanges=True)
    finally:
        xl.Quit()


if __name__ == "__main__":
    main(sys.argv[1])
