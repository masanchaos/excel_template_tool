import streamlit as st
import openpyxl
from io import BytesIO
from datetime import datetime

st.title("Excel 做賬模板處理工具")

uploaded_file = st.file_uploader("請上傳 Excel 檔案（.xlsx）", type="xlsx")

if uploaded_file:
    wb = openpyxl.load_workbook(uploaded_file)

    sheetnames = wb.sheetnames

    # ✅ 第 1 分頁：根據 A 欄升序排序
    if len(sheetnames) >= 1:
        ws1 = wb[sheetnames[0]]
        rows = list(ws1.iter_rows(values_only=True))
        if rows:
            header = rows[0]
            data_rows = sorted(
                [r for r in rows[1:] if r and len(r) > 0],
                key=lambda x: str(x[0]) if x[0] is not None else ""
            )
            ws1.delete_rows(1, ws1.max_row)
            for row in [header] + data_rows:
                ws1.append(row)

    # ✅ 第 2 分頁：刪除 A 欄，然後根據新的 A 欄排序
    if len(sheetnames) >= 2:
        ws2 = wb[sheetnames[1]]
        ws2.delete_cols(1)
        rows = list(ws2.iter_rows(values_only=True))
        if rows:
            header = rows[0]
            data_rows = sorted(
                [r for r in rows[1:] if r and len(r) > 0],
                key=lambda x: str(x[0]) if x[0] is not None else ""
            )
            ws2.delete_rows(1, ws2.max_row)
            for row in [header] + data_rows:
                ws2.append(row)

    # 🗑️ 刪除第 3 分頁（index=2）
    if len(sheetnames) >= 3:
        del wb[sheetnames[2]]
    sheetnames = wb.sheetnames  # 更新分頁名稱

    # ✅ 清空第 4 分頁起的內容（從 index 3 開始）
    for sheet_name in sheetnames[3:]:
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row,
                                min_col=1, max_col=ws.max_column):
            for cell in row:
                if cell.value not in (None, ""):
                    cell.value = ""

    # ⏰ 根據檔名或系統時間取月份
    month = datetime.now().month
    try:
        dt = datetime.strptime(uploaded_file.name[:10], "%Y-%m-%d")
        month = dt.month
    except:
        pass

    # 📥 準備下載
    result_filename = f"{month}月做賬模板.xlsx"
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.success("✅ 處理完成，請下載：")
    st.download_button(
        label="📥 下載做賬模板",
        data=output,
        file_name=result_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
