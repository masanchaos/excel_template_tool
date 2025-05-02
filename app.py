import streamlit as st
import openpyxl
from io import BytesIO
from datetime import datetime

st.title("Excel 做賬模板處理工具")

uploaded_file = st.file_uploader("請上傳 Excel 檔案（.xlsx）", type="xlsx")

if uploaded_file:
    wb = openpyxl.load_workbook(uploaded_file)

    # 原始分頁列表
    sheetnames = wb.sheetnames

    # 刪除第三分頁（索引為 2）
    if len(sheetnames) >= 3:
        del wb[sheetnames[2]]
        sheetnames = wb.sheetnames  # 更新清單

    # 清空第 3 張工作表（索引2）以後的內容，但不改變行數或儲存格
    for sheet_name in sheetnames[2:]:
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row,
                                min_col=1, max_col=ws.max_column):
            for cell in row:
                if cell.value not in (None, ""):
                    cell.value = ""

    # 嘗試從檔名擷取月份
    month = datetime.now().month
    try:
        file_name = uploaded_file.name
        dt = datetime.strptime(file_name[:10], "%Y-%m-%d")
        month = dt.month
    except:
        pass

    # 輸出 Excel
    result_filename = f"{month}月做賬模板.xlsx"
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.success("處理完成！請下載檔案：")
    st.download_button(
        label="📥 下載做賬模板",
        data=output,
        file_name=result_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
