import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
from io import BytesIO
from datetime import datetime

st.title("Excel 做賬模板處理工具")

uploaded_file = st.file_uploader("請上傳 Excel 檔案（.xlsx）", type="xlsx")

if uploaded_file:
    wb = openpyxl.load_workbook(uploaded_file)

    sheetnames = wb.sheetnames

    # ✅ 針對第 1 分頁：根據 A 欄排序（index 0）
    if len(sheetnames) >= 1:
        ws1 = wb[sheetnames[0]]
        rows = list(ws1.iter_rows(values_only=True))
        header = rows[0]
        data_rows = sorted(rows[1:], key=lambda x: (x[0] if x[0] is not None else ""))
        ws1.delete_rows(1, ws1.max_row)
        for i, row in enumerate([header] + data_rows, 1):
            ws1.append(row)

    # ✅ 第 2 分頁：刪除 A 欄，再依新 A 欄排序
    if len(sheetnames) >= 2:
        ws2 = wb[sheetnames[1]]
        ws2.delete_cols(1)  # 刪除原 A 欄
        rows = list(ws2.iter_rows(values_only=True))
        header = rows[0]
        data_rows = sorted(rows[1:], key=lambda x: (x[0] if x[0] is not None else ""))
        ws2.delete_rows(1, ws2.max_row)
        for i, row in enumerate([header] + data_rows, 1):
            ws2.append(row)

    # 🗑️ 刪除第 3 分頁（index=2）
    if len(sheetnames) >= 3:
        del wb[sheetnames[2]]

    # 重新取得工作表名稱（因已刪除）
    sheetnames = wb.sheetnames

    # ✅ 第 4 分頁開始（index 3）清空內容（保留排序和表頭）
    for sheet_name in sheetnames[3:]:
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row,
                                min_col=1, max_col=ws.max_column):
            for cell in row:
                if cell.value not in (None, ""):
                    cell.value = ""

    # 🕐 從檔名取月份
    month = datetime.now().month
    try:
        file_name = uploaded_file.name
        dt = datetime.strptime(file_name[:10], "%Y-%m-%d")
        month = dt.month
    except:
        pass
