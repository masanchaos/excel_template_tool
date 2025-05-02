import streamlit as st
import openpyxl
from io import BytesIO
from datetime import datetime

st.title("Excel 做賬模板處理工具")

uploaded_file = st.file_uploader("請上傳 Excel 檔案（.xlsx）", type="xlsx")

if uploaded_file:
    wb = openpyxl.load_workbook(uploaded_file)

    sheetnames = wb.sheetnames

    # 🗑️ 刪除第 3 個分頁（index=2）
    if len(sheetnames) >= 3:
        del wb[sheetnames[2]]

    # 重新取得刪除後的 sheet 名稱
    sheetnames = wb.sheetnames

    # 🆕 刪除第 2 分頁的第 1 欄（A欄）
    if len(sheetnames) >= 2:
        ws = wb[sheetnames[1]]
        ws.delete_cols(1)

    # ✅ 清空第 4 分頁（index=3）開始的內容
    for sheet_name in sheetnames[3:]:
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row,
                                min_col=1, max_col=ws.max_column):
            for cell in row:
                if cell.value not in (None, ""):
                    cell.value = ""

    # ⏰ 取得月份（檔名 or 系統時間）
    month = datetime.now().month
    try:
        file_name = uploaded_file.name
        dt = datetime.strptime(file_name[:10], "%Y-%m-%d")
        month = dt.month
    except:
        pass

    # 💾 儲存成下載檔
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
