import streamlit as st
import openpyxl
from io import BytesIO
from datetime import datetime

st.title("Excel 做賬模板處理工具")

uploaded_file = st.file_uploader("請上傳 Excel 檔案（.xlsx）", type="xlsx")

if uploaded_file:
    # 讀取 Excel
    wb = openpyxl.load_workbook(uploaded_file)

    # 清除第 3 張工作表開始的內容
    sheetnames = wb.sheetnames
    for sheet in sheetnames[2:]:
        ws = wb[sheet]
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row,
                                min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.value = None

    # 從檔案時間推斷月份
    month = datetime.now().month
    try:
        file_name = uploaded_file.name
        dt = datetime.strptime(file_name[:10], "%Y-%m-%d")
        month = dt.month
    except:
        pass

    # 產生檔案名稱
    result_filename = f"{month}月做賬模板.xlsx"

    # 儲存到記憶體
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    # 提供下載
    st.success("處理完成！請下載檔案：")
    st.download_button(
        label="📥 下載做賬模板",
        data=output,
        file_name=result_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
