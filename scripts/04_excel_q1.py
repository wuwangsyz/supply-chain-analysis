"""阶段 4 步骤 C：用 Excel 核对 Q1 的一个数字（西欧的订单数与延迟率）。

做法：
1. 把 Order Id、Order Region、Late_delivery_risk 复制到"Q1临时"工作表。
2. 按 Order Id 升序、延迟降序排序，再按 Order Id 去重。去重后每单保留延迟为 1 的那行（若有），
   与订单级口径一致。
3. 在"Q1"工作表用 COUNTIFS 计算订单数、延迟订单数和延迟率。

只新增工作表，不修改原有数据。工作簿已在阶段 1 前备份为 data\\工作簿2_备份.xlsx。

用法：
    C:\\Users\\31451\\.claude\\venvs\\skills\\Scripts\\python.exe -I -X utf8 scripts\\04_excel_q1.py "data\\工作簿2.xlsx"
"""

import os
import sys

import win32com.client as win32

TABLE = "DataCoSupplyChainDataset"
DATA_SHEET = "DataCoSupplyChainDataset"
TMP_SHEET = "Q1临时"
RESULT_SHEET = "Q1"
REGION = "Western Europe"
XL_ASCENDING = 1
XL_DESCENDING = 2


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
        for col_idx, col_name in enumerate(["Order Id", "Order Region", "Late_delivery_risk"], start=1):
            table.ListColumns(col_name).Range.Copy(tmp.Cells(1, col_idx))
        last_row = tmp.Cells(tmp.Rows.Count, 1).End(-4162).Row  # xlUp

        # 2. 排序：Order Id 升序，延迟降序；去重后每单保留延迟为 1 的行
        data = tmp.Range(f"A1:C{last_row}")
        data.Sort(
            Key1=tmp.Range("A1"), Order1=XL_ASCENDING,
            Key2=tmp.Range("C1"), Order2=XL_DESCENDING,
            Header=1,
        )
        data.RemoveDuplicates(Columns=1, Header=1)  # 按 Order Id 去重

        # 3. 结果表
        ws = wb.Worksheets.Add(After=tmp)
        ws.Name = RESULT_SHEET
        last = tmp.Cells(tmp.Rows.Count, 1).End(-4162).Row
        rng_region = f"'{TMP_SHEET}'!$B$2:$B${last}"
        rng_late = f"'{TMP_SHEET}'!$C$2:$C${last}"

        ws.Range("A1").Value = "检查项"
        ws.Range("B1").Value = "Excel 结果"
        ws.Range("A2").Value = "去重后订单数（全部）"
        ws.Range("B2").Formula = f"=COUNTA('{TMP_SHEET}'!$A$2:$A${last})"
        ws.Range("A3").Value = f"{REGION} 订单数"
        ws.Range("B3").Formula = f'=COUNTIFS({rng_region},"{REGION}")'
        ws.Range("A4").Value = f"{REGION} 延迟订单数"
        ws.Range("B4").Formula = f'=COUNTIFS({rng_region},"{REGION}",{rng_late},1)'
        ws.Range("A5").Value = f"{REGION} 订单级延迟率（%）"
        ws.Range("B5").Formula = "=100*B4/B3"

        xl.CalculateFull()
        for r in range(2, 6):
            print(f"{ws.Cells(r, 1).Value}: {ws.Cells(r, 2).Value}")

        wb.Save()
        wb.Close(SaveChanges=True)
    finally:
        xl.Quit()


if __name__ == "__main__":
    main(sys.argv[1])
