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

st.markdown('<div class="custom-main-title">📦 貨櫃箱號自動產出系統 (精密除錯版)</div>', unsafe_allow_html=True)
st.markdown("<p style='color: #555; margin-bottom: 5px;'>系統會自動依箱數進行列數倍增、產出 H 欄序號，並在下方即時顯示資料處理 Log。</p>", unsafe_allow_html=True)
st.markdown("---")

st.markdown('<div class="upload-card">', unsafe_allow_html=True)
st.markdown('<div class="custom-section-title">📁 請上傳拆櫃明細原始檔案 (.xlsx)</div>', unsafe_allow_html=True)
ctn_file = st.file_uploader("將檔案拖放到此處", type=["xlsx"], key="uctn", label_visibility="collapsed")
st.markdown("</div>", unsafe_allow_html=True)

if ctn_file is not None:
    if st.button("🚀 啟動貨櫃箱號自動產出", type="primary", use_container_width=True):
        
        # 建立一個除錯日誌容器，呈現在前台
        log_container = st.container()
        with log_container:
            st.info("🔍 [LOG] 開始解析上傳檔案...")
            
        with st.spinner("正在執行產出中..."):
            try:
                # 1. 讀取原始檔案 (跳過前 4 行，不設定表頭名稱，直接用數字 0~8 索引，防止兩個「箱數」衝突)
                df = pd.read_excel(ctn_file, skiprows=4, header=None)
                
                # 保證至少有 9 欄
                for i in range(df.shape[1], 9):
                    df[i] = None
                    
                # 只取前 9 欄
                df = df.iloc[:, :9]
                
                # 過濾雜項 (0 代表第一欄 col_A)
                df = df[df[0].notna()]
                exclude_keywords = '合計|總計|CTN|SKU|品項'
                df = df[~df[0].astype(str).str.contains(exclude_keywords, case=False, na=False)]

                with log_container:
                    st.write("📋 [LOG] 原始資料過濾完成，前 3 筆的 G 欄 (索引6) 箱數內容為：")
                    st.code(df[[0, 1, 6]].head(3).to_string())

                # 2. 處理 G 欄 (索引 6) 是否為空以及轉換為數字
                df['is_empty_g'] = df[6].isna()
                df['temp_g'] = pd.to_numeric(df[6], errors='coerce').fillna(1).astype(int)
                
                # 3. 執行倍增展開
                df_expanded = df.loc[df.index.repeat(df['temp_g'])].copy()

                with log_container:
                    st.write(f"📈 [LOG] 資料展開完成！列數從 {len(df)} 列倍增至 {len(df_expanded)} 列。")

                # 4. 生成傳統的「10箱-1, 10箱-2」邏輯（完全回歸你本機版成功的邏輯）
                def generate_h_native(row, group_count):
                    if row['is_empty_g']:
                        return ""
                    return f"{int(row[6])}箱-{group_count + 1}"

                df_expanded[7] = [
                    generate_h_native(row, count) 
                    for row, count in zip(df_expanded.to_dict('records'), df_expanded.groupby(level=0).cumcount())
                ]
                
                # 5. 強制清空第 9 欄 (索引 8，拆櫃日期)
                df_expanded[8] = ""
                
                # 記錄紅底清單
                is_empty_list = df_expanded['is_empty_g'].tolist()
                
                # 取出最終準備寫入 Excel 的資料
                final_data_df = df_expanded[[0, 1, 2, 3, 4, 5, 6, 7, 8]].copy()
                
                with log_container:
                    st.write("🎯 [LOG] H 欄序號計算完畢，即時預覽產出結果前 5 筆：")
                    st.code(final_data_df.head(5).to_string())

                # ─── 進入 Openpyxl 導出與美化 ───
                wb = Workbook()
                ws = wb.active
                ws.title = "拆櫃明細"

                thin_side = Side(border_style="thin", color="000000")
                full_border = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)
                center_align = Alignment(horizontal='center', vertical='center')
                ms_font = Font(name='微軟正黑體', size=11)
                red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

                # 寫入最終期望的標題 (包含兩個一模一樣的「箱數」)
                final_headers = ['商品編號', '商品名稱', '樣式', '品項條碼', '廠商批價', '叫貨數量', '箱數', '箱數', '拆櫃日期']
                ws.append(final_headers)
                
                for col_idx in range(1, 10):
                    header_cell = ws.cell(row=1, column=col_idx)
                    header_cell.border = full_border
                    header_cell.alignment = center_align
                    header_cell.font = Font(name='微軟正黑體', size=11, bold=True)

                # 寫入過濾並計算後的每行資料
                for row_data in final_data_df.fillna("").values.tolist(): 
                    ws.append(row_data)

                # 加入篩選器
                ws.auto_filter.ref = f"A1:I{ws.max_row}"

                # 刷入紅底與邊框格式
                for row_idx, is_empty in enumerate(is_empty_list, start=2):
                    for col_idx in range(1, 10):
                        cell = ws.cell(row=row_idx, column=col_idx)
                        cell.border = full_border
                        cell.alignment = center_align
                        cell.font = ms_font
                        if is_empty:
                            cell.fill = red_fill

                # 自動調整欄位寬度
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

                # 轉為二進位流供瀏覽器下載
                excel_data = BytesIO()
                wb.save(excel_data)
                excel_data.seek(0)
                
                with log_container:
                    st.success("🎉 [SUCCESS] 檔案成功產生完畢，未發生任何異常！")
                    
                st.download_button(
                    label="📥 點此一鍵下載全新拆櫃 Excel 檔案", 
                    data=excel_data, 
                    file_name="#義烏櫃 拆櫃明細-福北路-箱數.xlsx", 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                    use_container_width=True
                )
            except Exception as e: 
                st.error(f"❌ 執行中發生未預期錯誤: {e}")
