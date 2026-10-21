"""阶段 1 步骤 1.2：在 工作簿1.xlsx 中新增"体检核对"工作表，用 Excel 计算 3 项检查。

只新增工作表，不修改已有数据和"体检"工作表。改动前请先保留 data\\工作簿1_备份.xlsx。

去重订单数不用 UNIQUE：本机 Excel 16.0（build 19127）不支持动态数组函数（已实测）。
改为：把 Order Id 整列复制到新工作表"去重临时"，用 Excel 自带的"删除重复项"去重，再计数。

用法：
    C:\\Users\\31451\\.claude\\venvs\\skills\\Scripts\\python.exe -I -X utf8 scripts\\01_excel_check.py "data\\工作簿2.xlsx"
"""

import os
import sys

import win32com.client as win32

CHECK_SHEET = "体检核对"
TMP_SHEET = "去重临时"
TABLE = "DataCoSupplyChainDataset"
DATA_SHEET = "DataCoSupplyChainDataset"

CHECKS = [
    ("订单项数", f"=COUNTA({TABLE}[Order Item Id])"),
    ("去重订单数", f"=COUNTA({TMP_SHEET}!A:A)-1"),
    (
        "标志不一致行数",
        f'=COUNTIFS({TABLE}[Late_delivery_risk],1,{TABLE}[Delivery Status],"<>Late delivery")',
    ),
]


def main(path: str) -> None:
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(path))
        names = [s.Name for s in wb.Worksheets]
        if CHECK_SHEET in names or TMP_SHEET in names:
            raise SystemExit("目标工作表已存在，停止以免覆盖。")

        # 1. 复制 Order Id 列（含表头）到临时表，并去重
        src = wb.Worksheets(DATA_SHEET)
        order_id_col = src.ListObjects(TABLE).ListColumns("Order Id").Range
        tmp = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        tmp.Name = TMP_SHEET
        order_id_col.Copy(tmp.Range("A1"))
        last_row = tmp.Cells(tmp.Rows.Count, 1).End(-4162).Row  # xlUp
        tmp.Range(f"A1:A{last_row}").RemoveDuplicates(Columns=1, Header=1)  # Header=xlYes

        # 2. 写入检查表
        ws = wb.Worksheets.Add(After=tmp)
        ws.Name = CHECK_SHEET
        ws.Range("A1").Value = "检查项"
        ws.Range("B1").Value = "公式"
        ws.Range("C1").Value = "结果"
        for i, (label, formula) in enumerate(CHECKS, start=2):
            ws.Cells(i, 1).Value = label
            # 前导 ' 让公式以文本形式显示在 B 列，而不是被计算
            ws.Cells(i, 2).Value = "'" + formula
            ws.Cells(i, 3).Formula = formula

        xl.CalculateFull()
        for i, (label, _) in enumerate(CHECKS, start=2):
            print(f"{label}: {ws.Cells(i, 3).Value}")
        print(f"Excel 版本: {xl.Version}")

        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()


if __name__ == "__main__":
    main(sys.argv[1])
