import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill

# ================= 網頁頁面初始設定 =================
st.set_page_config(page_title="貨櫃箱號產出", page_icon="📦", layout="wide")

# 🛠️ 超強制 CSS 注入 (保持美美的北歐風)
st.markdown("""
    <style>
    .stApp { background-color: #fcfcfc !important; color: #2b303a !important; }
    [data-testid="stSidebar"] { background-color: #f3f4f6 !important; border-right: 1px solid #e5e7eb; }
    h1, h2, h3, h4, h5, h6, p, label, .stMarkdown { color: #2b303a !important; }
    .upload-card {
        background-color: #ffffff; padding: 24px; border-radius: 12px; 
        border: 1px solid #e5e7eb; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); margin-bottom: 20px;
    }
    [data-testid="stFileUploader"] { background-color: #fafafa !important; border: 1px dashed #cbd5e1 !important; border-radius: 8px !important; padding: 10px !important; }
    button[data-testid*="stBaseButton"] { background-color: #e2e8f0 !important; color: #475569 !important; border: 1px solid #cbd5e1 !important; }
    .custom-main-title { color: #1e293b !important; font-size: 1.8rem !important; font-weight: 700 !important; margin-bottom: 0.5rem; }
    .stButton > button { background-color: #475569 !important; border-color: #475569 !important; color: #ffffff !important; font-weight: 600 !important; font-size: 16px !important; border-radius: 8px !important; padding: 12px 0 !important; }
    .stButton > button:hover { background-color: #334155 !important; border-color: #334155 !important; }
    .stButton > button p { color: #ffffff !important; font-weight: 600 !important; }
    .custom-section-title { color: #2b5c8f !important; font-size: 18px !important; font-weight: 600 !important; margin-bottom: 12px !important; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="custom-main-title">📦 貨櫃箱號自動產出系統</div>', unsafe_allow_html=True)
st.markdown("<p style='color: #555; margin-bottom: 5px;'>請上傳原始拆櫃 Excel 報表，系統會全自動依箱數進行列數倍增、產出 H 欄序號，並將 G 欄空白者全自動塗上紅底標記。</p>", unsafe_allow_html=True)
st.markdown("---")

st.markdown('<div class="upload-card">', unsafe_allow_html=True)
st.markdown('<div class="custom-section-title">📁 請上傳拆櫃明細原始檔案 (.xlsx)</div>', unsafe_allow_html=True)
ctn_file = st.file_uploader("將檔案拖放到此處", type=["xlsx"], key="uctn", label_visibility="collapsed")
st.markdown("</div>", unsafe_allow_html=True)

if ctn_file is not None:
    if st.button("🚀 啟動貨櫃箱號自動產出", type="primary", use_container_width=True):
        with st.spinner("正在執行產出中..."):
            try:
                # ─── 100% 恢復您最初本機最成功的完全原始處理 ───
                df = pd.read_excel(ctn_file, skiprows=4, header=None)
                
                header_names = ['col_A', 'col_B', 'col_C', 'col_D', 'col_E', 'col_F', 'col_G', 'col_H', 'col_I']
                actual_col_count = len(df.columns)
                df.columns = header_names[:actual_col_count]
                for col in header_names[actual_col_count:]:
                    df[col] = None

                # 完美過濾
                df = df[df['col_A'].notna()]
                exclude_keywords = '合計|總計|CTN|SKU|品項'
                df = df[~df['~df['col_A'].astype(str).str.contains(exclude_keywords, case=False, na=False)] if 'col_A' in df.columns else df] # 安全防護
                df = df[~df['col_A'].astype(str).str.contains('合計|總計|CTN|SKU|品項', case=False, na=False)]

                # 箱數空白標記與倍增
                df['is_empty_g'] = df['col_G'].isna()
                df['temp_g'] = pd.to_numeric(df['col_G'], errors='coerce').fillna(1).astype(int)
                df_expanded = df.loc[df.index.repeat(df['temp_g'])].copy()

                is_empty_list = df_expanded['is_empty_g'].tolist()
                output_df = df_expanded.iloc[:, :9].copy()
                
                final_headers = ['商品編號', '商品名稱', '樣式', '品項條碼', '廠商批價', '叫貨數量', '箱數', '箱數', '拆櫃日期']
                output_df.columns = final_headers
                
                # ─── 進入 Openpyxl 最終強制硬塗儲存格階段 ───
                wb = Workbook()
                ws = wb.active
                ws.title = "拆櫃明細"

                thin_side = Side(border_style="thin", color="000000")
                full_border = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)
                center_align = Alignment(horizontal='center', vertical='center')
                ms_font = Font(name='微軟正黑體', size=11)
                red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

                # 表頭建立
                ws.append(final_headers)
                for col_idx in range(1, 10):
                    header_cell = ws.cell(row=1, column=col_idx)
                    header_cell.border = full_border
                    header_cell.alignment = center_align
                    header_cell.font = Font(name='微軟正黑體', size=11, bold=True)

                # 丟入基礎資料
                for row_data in output_df.fillna("").values.tolist(): 
                    ws.append(row_data)

                # 🛠️ 暴力硬改，直接覆蓋 openpyxl 儲存格，避開所有快取和 Pandas 錯位
                for r_idx in range(2, ws.max_row + 1):
                    # 第 7 欄是「箱數」(G欄)
                    g_val = ws.cell(row=r_idx, column=7).value
                    
                    # 強制把第 8 欄 (H欄) 刷成：[G欄數字]-1 
                    if g_val is not None and str(g_val).strip() != "":
                        try:
                            clean_g = str(int(float(g_val)))
                            ws.cell(row=r_idx, column=8).value = f"{clean_g}-1"
                        except:
                            ws.cell(row=r_idx, column=8).value = f"{str(g_val).strip()}-1"
                    else:
                        ws.cell(row=r_idx, column=8).value = ""

                    # 強制把第 9 欄 (I欄) 從第二列開始，全部清洗成空字串
                    ws.cell(row=r_idx, column=9).value = ""

                # 篩選器
                ws.auto_filter.ref = f"A1:I{ws.max_row}"

                # 紅底與美化
                for row_idx, is_empty in enumerate(is_empty_list, start=2):
                    for col_idx in range(1, 10):
                        cell = ws.cell(row=row_idx, column=col_idx)
                        cell.border = full_border
                        cell.alignment = center_align
                        cell.font = ms_font
                        if is_empty:
                            cell.fill = red_fill

                # 寬度自動調整
                for col in ws.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        try:
                            if cell.value:
                                val_str = str(cell.value)
                                byte_len = len(val_str.encode('big5')) if val_str else 0
                                if byte_len > max_length:
                                    max_length = byte_len
                        except: pass
                    ws.column_dimensions[column].width = max_length + 4

                excel_data = BytesIO()
                wb.save(excel_data)
                excel_data.seek(0)
                
                st.success("✨ 貨櫃箱號整理&產出完畢！請點擊下方按鈕下載成果：")
                st.download_button(
                    label="📥 點此一鍵下載全新拆櫃 Excel 檔案", 
                    data=excel_data, 
                    file_name="#義烏櫃 拆櫃明細-福北路-箱數.xlsx", 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                    use_container_width=True
                )
            except Exception as e: 
                st.error(f"❌ 錯誤: {e}")
