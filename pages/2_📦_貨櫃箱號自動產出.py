import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill

# ================= 網頁頁面初始設定 =================
st.set_page_config(page_title="貨櫃箱號產出", page_icon="📦", layout="wide")

# 🛠️ 終極北歐風裝潢加強版 (超強制 CSS 注入)
st.markdown("""
    <style>
    /* 全站北歐風淺色背景與冷灰色調文字 */
    .stApp {
        background-color: #fcfcfc !important;
        color: #2b303a !important;
    }
    /* 側邊欄改為溫潤的淺米灰 */
    [data-testid="stSidebar"] {
        background-color: #f3f4f6 !important;
        border-right: 1px solid #e5e7eb;
    }
    /* 標題與內文颜色統一 */
    h1, h2, h3, h4, h5, h6, p, label, .stMarkdown {
        color: #2b303a !important;
    }
    /* 卡片外框 */
    .upload-card {
        background-color: #ffffff;
        padding: 24px; 
        border-radius: 12px; 
        border: 1px solid #e5e7eb;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        margin-bottom: 24px;
    }
    /* 按鈕樣式調和 */
    div.stButton > button:first-child {
        background-color: #4f46e5 !important;
        color: white !important;
        border-radius: 6px !important;
        border: none !important;
        padding: 8px 16px !important;
    }
    div.stButton > button:first-child:hover {
        background-color: #4338ca !important;
    }
    </style>
""", unsafe_allow_html=True)


st.title("📦 貨櫃箱號自動對應系統")
st.caption("誠信、精準、高效 ── 讓跨境物流化繁為簡")

st.markdown('<div class="upload-card">', unsafe_allow_html=True)
col1, col2 = st.columns(2)

with col1:
    purchase_file = st.file_uploader("1️⃣ 上傳【採購系統檔】(Excel/CSV)", type=["xlsx", "csv"])
with col2:
    mapping_file = st.file_uploader("2️⃣ 上傳【箱號對應檔】(Excel/CSV)", type=["xlsx", "csv"])
st.markdown('</div>', unsafe_allow_html=True)

if purchase_file and mapping_file:
    # 讀取採購系統檔
    if purchase_file.name.endswith('.csv'):
        df_purchase = pd.read_csv(purchase_file)
    else:
        df_purchase = pd.read_excel(purchase_file)

    # 讀取箱號對應檔
    if mapping_file.name.endswith('.csv'):
        df_mapping = pd.read_csv(mapping_file)
    else:
        df_mapping = pd.read_excel(mapping_file)

    st.success("🎉 雙方報表皆上傳成功！已就緒。")

    if st.button("🚀 開始執行自動化對應與格式化"):
        with st.spinner("系統正在處理大數據對應與格式標準化中，請稍候..."):
            
            # 1. 整理對應檔：建立字典 (品項條碼 -> 起始箱號)
            box_mapping = {}
            for _, row in df_mapping.iterrows():
                sku = str(row.get('品項條碼', '')).strip()
                start_box = row.get('起始箱號', '')
                if sku:
                    box_mapping[sku] = {'start_box': start_box}

            output_data = []
            is_empty_list = []

            # 2. 逐行遍歷採購系統，進行比對與防呆
            for _, row in df_purchase.iterrows():
                cabinet = str(row.get('櫃號', '')).strip()
                box_range = str(row.get('箱麥', '')).strip()
                po_no = str(row.get('採購單號', '')).strip()
                note = str(row.get('備註', '')).strip()
                order_1688 = str(row.get('1688訂單', '')).strip()
                supplier = str(row.get('供應商', '')).strip()
                sku = str(row.get('SKU 編碼', '')).strip()
                
                # ==================== 🛠️ 核心修正邏輯 ====================
                raw_start = box_mapping.get(sku, {}).get('start_box', None)
                
                # 判斷是否為有效數值或非空字串
                if pd.notna(raw_start) and str(raw_start).strip() != "":
                    try:
                        # 先轉 float 再轉 int，完美消滅 351.0 這類浮點數小數點
                        start_box_val = int(float(raw_start))
                        start_box_str = str(start_box_val)
                        actual_start_box = f"{start_box_str}-1"  # 達成 =TEXT(G2,"0") & "-1" 的邏輯
                        is_empty = False
                    except:
                        # 轉換失敗則保持原始字串型態
                        start_box_str = str(raw_start).strip()
                        actual_start_box = f"{start_box_str}-1" if start_box_str else ""
                        is_empty = True
                else:
                    # 當對應不到或原始欄位為空時，H與I欄皆保持完全空白，不顯示為 -1
                    start_box_str = ""
                    actual_start_box = ""
                    is_empty = True
                # ==========================================================

                is_empty_list.append(is_empty)

                output_data.append([
                    cabinet,
                    box_range,
                    po_no,
                    note,
                    order_1688,
                    supplier,
                    sku,
                    start_box_str,
                    actual_start_box
                ])

            # 建立輸出的 DataFrame
            output_cols = ["櫃號", "箱麥", "採購單號", "備註", "1688訂單", "供應商", "SKU 編碼", "起始箱號", "實際起始箱號"]
            output_df = pd.DataFrame(output_data, columns=output_cols)

            # ================= openpyxl 導出與美化 =================
            wb = Workbook()
            ws = wb.active
            ws.title = "拆櫃明細"

            # 樣式定義
            thin_side = Side(border_style="thin", color="CCCCCC")
            full_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
            center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
            header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

            # 寫入標題
            ws.append(output_cols)
            for cell in ws[1]:
                cell.fill = header_fill
                cell.alignment = center_align
                cell.font = Font(name='微軟正黑體', size=11, bold=True)

            # 寫入資料
            for row_data in output_df.fillna("").values.tolist(): 
                ws.append(row_data)
            ws.auto_filter.ref = f"A1:I{ws.max_row}"

            # 刷入邊框與防呆標記顏色
            for r_idx, is_empty in enumerate(is_empty_list, start=2):
                for c_idx in range(1, 10):
                    cell = ws.cell(row=r_idx, column=c_idx)
                    cell.border = full_border
                    cell.alignment = center_align
                    cell.font = Font(name='微軟正黑體', size=11)
                    if is_empty: 
                        cell.fill = red_fill

            # 自動調整欄寬
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = max_len + 5

            # 寫入二進位流並提供下載
            excel_data = BytesIO()
            wb.save(excel_data)
            excel_data.seek(0)
            
            st.success("✨ 貨櫃箱號整理&產出完畢！請點擊下方按鈕下載成果：")
            st.download_button(
                label="📥 點此一鍵下載全新拆櫃 Excel 檔案", 
                data=excel_data, 
                file_name="#義烏櫃 拆櫃明細-福北路-箱數.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
