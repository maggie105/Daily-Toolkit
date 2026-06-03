import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill

st.set_page_config(page_title="庫存條碼產出", page_icon="🏷️", layout="wide")
st.markdown("<h2 style='color: #2b5c8f; font-weight: 700;'>🏷️ 庫存資料與一維條碼整合系統</h2>", unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)
with col1: file_main = st.file_uploader("📦 ① 拆櫃明細檔案", type=["xlsx"])
with col2: file_shelf = st.file_uploader("🗺️ ② 貨架位檔案", type=["xlsx"])
with col3: file_purchase = st.file_uploader("📋 ③ 採購建議檔案", type=["xlsx"])

if st.button("🚀 啟動多表整合&自動產出", type="primary", use_container_width=True):
    if not (file_main and file_shelf and file_purchase): st.error("🚨 請上傳完整三份報表。")
    else:
        with st.spinner("整合中..."):
            try:
                df_main = pd.read_excel(file_main, skiprows=3)
                new_headers = ["商品編號", "商品名稱", "商品規格", "品項條碼", "箱裝數", "叫貨數量", "件數", "福北總庫存", "15日銷售", "福撿儲位", "一維條碼"]
                df_main.columns = new_headers[:len(df_main.columns)]
                
                df_shelf_raw = pd.read_excel(file_shelf); df_shelf_raw.iloc[:, 1] = df_shelf_raw.iloc[:, 1].ffill()
                shelf_dict = {str(r.iloc[4]).strip(): str(r.iloc[1]).strip() for _, r in df_shelf_raw.iterrows() if not pd.isna(r.iloc[4])}
                
                df_pur_raw = pd.read_excel(file_purchase)
                p_stock = {str(r.iloc[0]).strip(): r.iloc[5] for _, r in df_pur_raw.iterrows() if not pd.isna(r.iloc[0])}
                p_sale = {str(r.iloc[0]).strip(): r.iloc[12] for _, r in df_pur_raw.iterrows() if not pd.isna(r.iloc[0])}

                yellow_rows = set()
                for idx, row in df_main.iterrows():
                    key = str(row["品項條碼"]).strip()
                    if not key or key == "nan": continue
                    df_main.at[idx, "福北總庫存"] = p_stock.get(key, 0)
                    df_main.at[idx, "15日銷售"] = p_sale.get(key, 0)
                    s_val = shelf_dict.get(key, "")
                    df_main.at[idx, "福撿儲位"] = s_val
                    df_main.at[idx, "一維條碼"] = f'="*" & D{idx+5} & "*"'
                    if s_val and s_val not in ["W-兩倉暫存區", "W-線東品"] and (float(p_stock.get(key, 0)) - float(p_sale.get(key, 0)) <= 0):
                        yellow_rows.add(idx + 5)

                wb = Workbook(); ws = wb.active; ws.title = "庫存一維碼明細"
                thin = Side(border_style="thin", color="000000")
                full_border = Border(top=thin, left=thin, right=thin, bottom=thin)
                y_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")

                for _ in range(3): ws.append([])
                ws.append(new_headers)
                for c in range(1, 12):
                    cell = ws.cell(row=4, column=c)
                    cell.border = full_border; cell.alignment = Alignment(horizontal='center', vertical='center'); cell.font = Font(name='微軟正黑體', size=11, bold=True)

                for _, row_data in df_main.fillna("").iterrows(): ws.append(list(row_data))
                ws.auto_filter.ref = f"A4:K{ws.max_row}"

                for r in range(5, ws.max_row + 1):
                    for c in range(1, 12):
                        cell = ws.cell(row=r, column=c)
                        cell.border = full_border; cell.alignment = Alignment(horizontal='center', vertical='center'); cell.font = Font(name='Free 3 of 9 Extended', size=30) if c == 11 else Font(name='微軟正黑體', size=11)
                        if r in yellow_rows: cell.fill = y_fill

                for col in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    ws.column_dimensions[col[0].column_letter].width = max_len + 5

                excel_data = BytesIO(); wb.save(excel_data); excel_data.seek(0)
                st.success("✨ VBA 庫存大整合全自動演算法執行成功！")
                st.download_button(label="📥 點此一鍵下載全新庫存條碼整合 Excel 檔案", data=excel_data, file_name="#義烏櫃 拆櫃明細-福北路-庫存資料&一維條碼.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            except Exception as e: st.error(f"❌ 錯誤: {e}")
