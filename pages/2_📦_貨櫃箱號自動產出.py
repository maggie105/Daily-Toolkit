import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill

st.set_page_config(page_title="貨櫃箱號產出", page_icon="📦", layout="wide")
st.markdown("<h2 style='color: #2b5c8f; font-weight: 700;'>📦 貨櫃箱號自動產出系統</h2>", unsafe_allow_html=True)

ctn_file = st.file_uploader("📁 請上傳拆櫃明細原始檔案 (.xlsx)", type=["xlsx"])

if ctn_file is not None and st.button("🚀 啟動貨櫃箱號自動產出", type="primary", use_container_width=True):
    with st.spinner("正在執行產出中..."):
        try:
            df = pd.read_excel(ctn_file, skiprows=4, header=None)
            header_names = ['col_A', 'col_B', 'col_C', 'col_D', 'col_E', 'col_F', 'col_G', 'col_H', 'col_I']
            df.columns = header_names[:len(df.columns)]
            df = df[df['col_A'].notna()]
            df = df[~df['col_A'].astype(str).str.contains('合计|总计|CTN|SKU|品项|合計|總計|品項', case=False, na=False)]
            
            df['is_empty_g'] = df['col_G'].isna()
            df['temp_g'] = pd.to_numeric(df['col_G'], errors='coerce').fillna(1).astype(int)
            df_expanded = df.loc[df.index.repeat(df['temp_g'])].copy()
            
            df_expanded['col_H'] = [f"{int(r['col_G'])}箱-{c+1}" if not r['is_empty_g'] else "" for r, c in zip(df_expanded.to_dict('records'), df_expanded.groupby(level=0).cumcount())]
            
            is_empty_list = df_expanded['is_empty_g'].tolist()
            output_df = df_expanded.iloc[:, :9].copy()
            final_headers = ['商品編號', '商品名稱', '樣式', '品項條碼', '廠商批價', '叫貨數量', '箱數', '箱數', '拆櫃日期']
            output_df.columns = final_headers

            wb = Workbook(); ws = wb.active; ws.title = "拆櫃明細"
            thin = Side(border_style="thin", color="000000")
            full_border = Border(top=thin, left=thin, right=thin, bottom=thin)
            center_align = Alignment(horizontal='center', vertical='center')
            red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

            ws.append(final_headers)
            for idx in range(1, 10):
                cell = ws.cell(row=1, column=idx)
                cell.border = full_border; cell.alignment = center_align; cell.font = Font(name='微軟正黑體', size=11, bold=True)

            for row_data in output_df.fillna("").values.tolist(): ws.append(row_data)
            ws.auto_filter.ref = f"A1:I{ws.max_row}"

            for r_idx, is_empty in enumerate(is_empty_list, start=2):
                for c_idx in range(1, 10):
                    cell = ws.cell(row=r_idx, column=c_idx)
                    cell.border = full_border; cell.alignment = center_align; cell.font = Font(name='微軟正黑體', size=11)
                    if is_empty: cell.fill = red_fill

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = max_len + 5

            excel_data = BytesIO(); wb.save(excel_data); excel_data.seek(0)
            st.success("✨ 貨櫃箱號整理&產出完畢！")
            st.download_button(label="📥 點此一鍵下載全新拆櫃 Excel 檔案", data=excel_data, file_name="#義烏櫃 拆櫃明細-福北路-箱數.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        except Exception as e: st.error(f"❌ 錯誤: {e}")
