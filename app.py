import streamlit as st
import pandas as pd
import os
import time
import gspread
import warnings
import re
import json
from io import BytesIO
from datetime import datetime  # 這是之前補上的
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build         
from googleapiclient.http import MediaIoBaseUpload  
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill

# 1. 隱藏 openpyxl 的格式警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ================= 網頁頁面初始設定 =================
st.set_page_config(page_title="自動化工具", page_icon="🔧", layout="wide")

# 🔥 終極北歐簡約風裝潢 (超強制 CSS 注入)
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
    /* 左右兩大區塊卡片外框 */
    .upload-card {
        background-color: #ffffff;
        padding: 24px; 
        border-radius: 12px; 
        border: 1px solid #e5e7eb;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        margin-bottom: 20px;
    }
    /* 將上傳檔案的大背景改為北歐簡約淡灰 + 細虛線外框 */
    [data-testid="stFileUploader"] {
        background-color: #fafafa !important;
        border: 1px dashed #cbd5e1 !important;
        border-radius: 8px !important;
        padding: 10px !important;
    }
    [data-testid="stFileUploader"] section {
        background-color: transparent !important;
    }
    div[data-testid="stFileUploadDropzone"] {
        background-color: #fafafa !important;
    }
    
    /* 強制將所有 Upload 小按鈕改為「常態北歐灰底黑字」 */
    button[data-testid*="stBaseButton"] {
        background-color: #e2e8f0 !important;
        color: #475569 !important;
        border: 1px solid #cbd5e1 !important;
        transition: all 0.2s ease;
    }
    button[data-testid*="stBaseButton"]:hover {
        background-color: #cbd5e1 !important;
        color: #334155 !important;
    }
    
    /* 強制將大按鈕改為深藍底、白字、粗體 */
    .stButton > button {
        background-color: #475569 !important;
        border-color: #475569 !important;
        color: #ffffff !important;
        font-weight: 600 !important;
        font-size: 16px !important;
        border-radius: 8px !important;
        padding: 12px 0 !important;
        transition: all 0.2s ease !important;
    }
    .stButton > button:hover {
        background-color: #334155 !important;
        border-color: #334155 !important;
    }
    .stButton > button p {
        color: #ffffff !important;
        font-weight: 600 !important;
    }
    
    /* 自訂精緻的中等標題字體樣式 */
    .custom-section-title {
        color: #2b5c8f !important;
        font-size: 18px !important;
        font-weight: 600 !important;
        margin-bottom: 12px !important;
        margin-top: 5px !important;
        display: flex;
        align-items: center;
    }

    /* 右上角 Toast 視窗優化 */
    div[data-testid="stToast"] {
        background-color: #334155 !important; 
        border: 1px solid #1e293b !important;  
        border-radius: 8px !important;         
        box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04) !important; 
    }
  /* 強制將 Toast 內的所有層級文字都改為純白與粗體 */
    div[data-testid="stToast"], 
    div[data-testid="stToast"] div, 
    div[data-testid="stToast"] span, 
    div[data-testid="stToast"] p {
        color: #ffffff !important;            
        font-weight: 900 !important;          
        font-size: 15px !important;
        letter-spacing: 0.5px !important;
    }
    div[data-testid="stToast"] button {
        color: #94a3b8 !important;
    }
    </style>
""", unsafe_allow_html=True)

# ================= Google Sheets 核心設定 =================
SPREADSHEET_ID = '1FLfAbqq1TmQnXFR3rHxGkrXELrlKSMpiXGYXk7hVZm8'
SERVICE_ACCOUNT_FILE = 'service_account.json'

SHEET_CONFIGS = {
    "01": {"gid": "1735114343", "end_col": "AB", "features": ["SKU编号", "现有库存"]}, 
    "02": {"gid": "830730835", "end_col": "S", "features": ["店铺昵称", "Item ID", "SKU"]},  
    "03": {"gid": "150603716", "end_col": "Q", "features": ["商品名称", "有效商品销售额"]},  
    "04": {"gid": "790740822", "end_col": "W", "features": ["商品SKU", "利润", "利润率"]},  
    "target_main": {"gid": "1324377276"}         
}

# ================= 標準表頭定義 =================
HEADERS_01 = ["Image URL", "SKU编号", "名称", "重量(g)", "长(cm)", "宽(cm)", "高(cm)", "仓库", "库区", "货架位", "现有库存", "订单已锁", "整仓可用", "整仓未上架", "活动预留", "在途中", "在途总成本", "警戒库存", "预测日销量", "预计可售天数", "加權成本價", "總成本價", "銷售狀態", "一級分類", "二級分類", "三級分類", "SKU類型", "備註"]
HEADERS_02 = ["店铺昵称", "分类_L1", "分类_L2", "分类_L3", "产品名称", "Item ID", "销量", "收藏", "浏览量", "Parent SKU", "变种", "变种ID", "SKU", "库存", "价格", "促销价", "限购", "平台创建時間", "发货期"]
HEADERS_03 = ["商品名称", "店铺", "商品SKU", "有效商品销售额", "有效订单量", "有效商品销量", "商品平均价格", "商品销售额", "商品销量", "订单总量", "包裹总量", "退款商品金额", "退款订单数", "退款商品数", "取消商品金額", "取消訂單數", "取消商品數"]
HEADERS_04 = ["商品SKU", "店铺", "商品名称", "分类", "商品收入", "总成本", "利润", "单个商品利润", "利润率", "订单数量", "销售数量", "退货数量", "退货率", "商品总销售额", "折扣&优惠补贴", "买家支付运费", "卖家支付運費", "佣金", "交易費", "服務費", "營銷費用", "退款金額", "平台其他費用"]

# ================= 工具函數 =================
def normalize_header(header_name):
    s = str(header_name).strip()
    trad_to_simp = {
        "量": "量", "銷": "销", "額": "额", "庫": "库", "存": "存", 
        "單": "单", "價": "价", "利": "利", "潤": "润", "貨": "货", 
        "類": "类", "幣": "币", "應": "应", "實": "实", "項": "项",
        "權": "权", "總": "总", "狀": "状", "態": "态", "級": "级", 
        "備": "备", "註": "注"
    }
    for trad, simp in trad_to_simp.items():
        s = s.replace(trad, simp)
    return s

def get_worksheet_by_gid(sh, gid):
    for sheet in sh.worksheets():
        if str(sheet.id) == str(gid):
            return sheet
    return None

def safe_read_and_align_uploaded(uploaded_file, target_headers, task_key, header_row=0):
    try:
        df = pd.read_excel(uploaded_file, header=header_row)
    except Exception as e:
        st.error(f"❌ 無法讀取檔案 `{uploaded_file.name}`。錯誤資訊: {e}")
        return None
        
    original_cols = df.columns.tolist()
    norm_original_cols = [normalize_header(col) for col in original_cols]
    norm_to_orig = {normalize_header(col): col for col in original_cols}
    
    features = SHEET_CONFIGS[task_key]["features"]
    for f in features:
        if normalize_header(f) not in norm_original_cols:
            st.error(f"🚨 檔案防呆攔截：您上傳的 `{uploaded_file.name}` 找不到關鍵欄位 `{f}`！")
            return None
            
    df_aligned = pd.DataFrame()
    for target in target_headers:
        target_norm = normalize_header(target)
        if target_norm in norm_to_orig:
            df_aligned[target] = df[norm_to_orig[target_norm]]
        else:
            df_aligned[target] = ""
    return df_aligned[target_headers]

def get_gspread_client():
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    if len(st.secrets) > 0:
        try:
            if "gcp_service_account" in st.secrets:
                secret_data = st.secrets["gcp_service_account"]
                creds_dict = json.loads(secret_data) if isinstance(secret_data, str) else secret_data
                creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
                return gspread.authorize(creds)
            elif "type" in st.secrets and st.secrets["type"] == "service_account":
                creds_dict = dict(st.secrets)
                creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
                return gspread.authorize(creds)
        except Exception as e:
            st.error(f"❌ 嘗試解析雲端 Secrets 金鑰時發生錯誤: {e}")
            
    if os.path.exists(SERVICE_ACCOUNT_FILE):
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope)
            return gspread.authorize(creds)
        except Exception as e:
            st.error(f"❌ 讀取本地憑證 `{SERVICE_ACCOUNT_FILE}` 失敗。錯誤: {e}")
            return None
    return None

def upload_to_google_sheets(df, task_key, title_list=None):
    if df is None or df.empty:
        return
    try:
        client = get_gspread_client()
        if not client: return
        sh = client.open_by_key(SPREADSHEET_ID)
        worksheet = get_worksheet_by_gid(sh, SHEET_CONFIGS[task_key]["gid"])
        if not worksheet: return
            
        end_col = SHEET_CONFIGS[task_key]["end_col"]
        worksheet.batch_clear([f"A:{end_col}"])
        
        data_to_upload = []
        if title_list:
            data_to_upload.append([("" if pd.isna(x) else x) for x in title_list])
            data_to_upload.append(df.columns.tolist())
        else:
            data_to_upload.append(df.columns.tolist())

        clean_df = df.fillna("")
        for row in clean_df.values.tolist():
            data_to_upload.append([("" if pd.isna(cell) else cell) for cell in row])
        
        worksheet.update(values=data_to_upload, range_name='A1')
        st.write(f"🟢 {worksheet.title} 數據已成功同步。")
    except Exception as e:
        st.error(f"❌ 雲端同步失敗 (Task {task_key})。錯誤: {e}")

def post_process_steps():
    try:
        client = get_gspread_client()
        if not client: return
        sh = client.open_by_key(SPREADSHEET_ID)
        sheet_source = get_worksheet_by_gid(sh, SHEET_CONFIGS["03"]["gid"])
        sheet_target = get_worksheet_by_gid(sh, SHEET_CONFIGS["target_main"]["gid"])
        if not sheet_source or not sheet_target: return

        raw_date_text = sheet_source.acell('B1').value 
        found_dates = re.findall(r'\d{4}-\d{2}-\d{2}', str(raw_date_text))
        if len(found_dates) >= 2:
            d1 = found_dates[0].replace("-", "")[4:]
            d2 = found_dates[1].replace("-", "")[4:]
            final_date_str = f"{d1}-{d2}"
            sheet_target.update(values=[[final_date_str]], range_name='B1')
            st.write(f"📅 雲端主表 B1 日期同步完成: `{final_date_str}`")

        v_values = sheet_source.col_values(22)[4:] 
        filtered_v = [[v] for v in v_values if v and str(v).strip() not in ["ZZ000-04", "總和"]]
        sheet_target.batch_clear(["A3:A"]) 
        if filtered_v:
            sheet_target.update(values=filtered_v, range_name='A3')
            st.write("🏷️ 雲端 A 欄品項條碼二次同步成功。")
    except Exception as e:
        st.error(f"❌ 後處理二次同步出錯: {e}")

# ==================== 🛠️ 側邊欄：方案 A (極簡無框文字卡片) ====================
with st.sidebar:
    st.markdown("<h3 style='color: #2b5c8f; font-weight: 700;'>🔧 Daily Toolkit</h3>", unsafe_allow_html=True)
    st.write("數據處理中心")
    st.markdown("---")
    
    # 🌟 方案 A 的核心：將所有子項目直接攤平，做成高對比的極簡大按鈕清單
    st.markdown("<p style='font-weight: 600; color: #475569; font-size: 14px; margin-bottom: 5px;'>🛠️ 請選擇欲執行的工具動作：</p>", unsafe_allow_html=True)
    app_mode = st.radio(
        label="隱藏標籤",
        options=[
            "📊 BS銷售更新", 
            "📦 貨櫃箱號 自動產出", 
            "🏷️ 庫存資料+一維碼 一鍵產出", 
            "🧾 正隆帳單核對"
        ],
        label_visibility="collapsed" # 隱藏原生討厭的 radio 小黑點標題
    )
    st.markdown("---")
    st.caption("✨ 目前版本: V3.2 (方案 A 清晰無框版)")

# ==================== 🖥️ 右側主畫面：功能一 (BS銷售更新) ====================
if app_mode == "📊 BS銷售更新":
    st.markdown("<h2 style='color: #2b5c8f; font-weight: 700;'>📊 BigSeller 銷售數據更新系統</h2>", unsafe_allow_html=True)
    st.markdown("""
        <p style='color: #555; margin-bottom: 5px;'>對應本地 01 至 04 資料夾。請上傳對應的 Excel 報表，系統將自動進行多檔清洗、字體校正並同步至雲端。</p>
        <p style='margin-top: 0;'>👉 <a href='https://docs.google.com/spreadsheets/d/1FLfAbqq1TmQnXFR3rHxGkrXELrlKSMpiXGYXk7hVZm8/edit?gid=1324377276#gid=1324377276' target='_blank' style='color: #2b5c8f; font-weight: bold; text-decoration: underline;'>📊 點此打開 Google Sheets 雲端主表</a></p>
    """, unsafe_allow_html=True)
    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="upload-card">', unsafe_allow_html=True)
        st.markdown('<div class="custom-section-title">📁 01. 庫存清單</div>', unsafe_allow_html=True)
        uploaded_files_01 = st.file_uploader("請上傳「库存清单*.xlsx」檔案 (可多選)", type=["xlsx"], accept_multiple_files=True, key="u01")
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="custom-section-title">📁 02. 在線產品 更新促銷價</div>', unsafe_allow_html=True)
        uploaded_files_lan = st.file_uploader("① 請上傳「懶餅乾*.xlsx」檔案 (可多選)", type=["xlsx"], accept_multiple_files=True, key="ulan")
        uploaded_files_onl = st.file_uploader("② 請上傳「Online_products*.xlsx」檔案 (可多選)", type=["xlsx"], accept_multiple_files=True, key="uonl")
        st.markdown("</div>", unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="upload-card">', unsafe_allow_html=True)
        st.markdown('<div class="custom-section-title">📁 03. 銷量報告</div>', unsafe_allow_html=True)
        uploaded_file_03 = st.file_uploader("請上傳「销量报告*.xlsx」檔案 (單選)", type=["xlsx"], key="u03")
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="custom-section-title">📁 04. 利潤報告</div>', unsafe_allow_html=True)
        uploaded_file_04 = st.file_uploader("請上傳「商品利润*.xlsx」檔案 (單選)", type=["xlsx"], key="u04")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("🔥 啟動自動化整合全流程", type="primary", use_container_width=True):
        if not (uploaded_files_01 or uploaded_files_lan or uploaded_files_onl or uploaded_file_03 or uploaded_file_04):
            st.error("🚨 流程終止：您尚未上傳任何 Excel 檔案。")
        else:
            start_time = time.time()
            has_error = False
            with st.spinner("系統正在全速同步數據，請稍候..."):
                if uploaded_files_01:
                    st.toast("⏳ 正在處理 01.庫存清單...")
                    list_dfs = []
                    for f in uploaded_files_01:
                        res_df = safe_read_and_align_uploaded(f, HEADERS_01, "01")
                        if res_df is not None: list_dfs.append(res_df)
                        else: has_error = True
                    if list_dfs:
                        df01 = pd.concat(list_dfs, ignore_index=True)
                        upload_to_google_sheets(df01, "01")

                if uploaded_files_lan or uploaded_files_onl:
                    st.toast("⏳ 正在處理 02.在線產品更新...")
                    lan_dfs, onl_dfs = [], []
                    for f in uploaded_files_lan:
                        res_df = safe_read_and_align_uploaded(f, HEADERS_02, "02")
                        if res_df is not None: lan_dfs.append(res_df)
                        else: has_error = True
                    for f in uploaded_files_onl:
                        res_df = safe_read_and_align_uploaded(f, HEADERS_02, "02")
                        if res_df is not None: onl_dfs.append(res_df)
                        else: has_error = True
                    df_lan = pd.concat(lan_dfs, ignore_index=True).drop_duplicates(subset=['SKU']) if lan_dfs else pd.DataFrame(columns=HEADERS_02)
                    df_onl = pd.concat(onl_dfs, ignore_index=True).drop_duplicates(subset=['SKU']) if onl_dfs else pd.DataFrame(columns=HEADERS_02)
                    df02 = pd.concat([df_lan, df_onl], ignore_index=True)
                    if not df02.empty:
                        df02['priority'] = df02['店铺昵称'].apply(lambda x: 0 if x == "02-懶餅乾家居" else 1)
                        df02 = df02.sort_values(by=['SKU', 'priority']).drop_duplicates(subset=['SKU'], keep='first').drop(columns=['priority'])
                        upload_to_google_sheets(df02, "02")

                if uploaded_file_03:
                    st.toast("⏳ 正在處理 03.銷量報告...")
                    try:
                        t03 = pd.read_excel(uploaded_file_03, nrows=1, header=None).values.tolist()[0]
                        df03 = safe_read_and_align_uploaded(uploaded_file_03, HEADERS_03, "03", header_row=1)
                        if df03 is not None: upload_to_google_sheets(df03, "03", title_list=t03)
                        else: has_error = True
                    except Exception as e: has_error = True

                if uploaded_file_04:
                    st.toast("⏳ 正在處理 04.利潤報告...")
                    try:
                        t04 = pd.read_excel(uploaded_file_04, nrows=1, header=None).values.tolist()[0]
                        df04 = safe_read_and_align_uploaded(uploaded_file_04, HEADERS_04, "04", header_row=1)
                        if df04 is not None: upload_to_google_sheets(df04, "04", title_list=t04)
                        else: has_error = True
                    except Exception as e: has_error = True

                if uploaded_file_03 and not has_error:
                    post_process_steps()
                    
            if has_error:
                st.warning("⚠️ 流程已結束，但過程中發生過上述錯誤攔截，請檢查修正。")
            else:
                st.success(f"✨ 全部自動化流程已執行完畢！總耗時：{round(time.time() - start_time, 2)} 秒")

# ==================== 🖥️ 功能二：貨櫃箱號自動產出 ====================
elif app_mode == "📦 貨櫃箱號 自動產出":
    st.markdown("<h2 style='color: #2b5c8f; font-weight: 700;'>📦 貨櫃箱號自動產出系統</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #555;'>請上傳原始拆櫃 Excel 報表，系統會全自動依箱數進行列數倍增、產出 H 欄序號，並將 G 欄空白者全自動塗上紅底標記。</p>", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown('<div class="upload-card">', unsafe_allow_html=True)
    st.markdown('<div class="custom-section-title">📁 請上傳拆櫃明細原始檔案 (.xlsx)</div>', unsafe_allow_html=True)
    ctn_file = st.file_uploader("將檔案拖放到此處", type=["xlsx"], key="uctn")
    st.markdown("</div>", unsafe_allow_html=True)

    if ctn_file is not None:
        if st.button("🚀 啟動貨櫃箱號自動產出", type="primary", use_container_width=True):
            with st.spinner("正在執行產出中..."):
                try:
                    df = pd.read_excel(ctn_file, skiprows=4, header=None)
                    header_names = ['col_A', 'col_B', 'col_C', 'col_D', 'col_E', 'col_F', 'col_G', 'col_H', 'col_I']
                    actual_col_count = len(df.columns)
                    df.columns = header_names[:actual_col_count]
                    for col in header_names[actual_col_count:]:
                        df[col] = None

                    df = df[df['col_A'].notna()]
                    exclude_keywords = '合计|总计|CTN|SKU|品项|合計|總計|品項'
                    df = df[~df['col_A'].astype(str).str.contains(exclude_keywords, case=False, na=False)]

                    df['is_empty_g'] = df['col_G'].isna()
                    df['temp_g'] = pd.to_numeric(df['col_G'], errors='coerce').fillna(1).astype(int)
                    df_expanded = df.loc[df.index.repeat(df['temp_g'])].copy()

                    def generate_h(row, group_count):
                        if row['is_empty_g']: return "" 
                        return f"{int(row['col_G'])}箱-{group_count + 1}"

                    df_expanded['col_H'] = [
                        generate_h(row, count) 
                        for row, count in zip(df_expanded.to_dict('records'), df_expanded.groupby(level=0).cumcount())
                    ]

                    is_empty_list = df_expanded['is_empty_g'].tolist()
                    output_df = df_expanded.iloc[:, :9].copy()
                    final_headers = ['商品編號', '商品名稱', '樣式', '品項條碼', '廠商批價', '叫貨數量', '箱數', '箱數', '拆櫃日期']
                    output_df.columns = final_headers

                    wb = Workbook()
                    ws = wb.active
                    ws.title = "拆櫃明細"

                    thin_side = Side(border_style="thin", color="000000")
                    full_border = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)
                    center_align = Alignment(horizontal='center', vertical='center')
                    ms_font = Font(name='微軟正黑體', size=11)
                    red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

                    ws.append(final_headers)
                    for col_idx in range(1, 10):
                        cell = ws.cell(row=1, column=col_idx)
                        cell.border = full_border
                        cell.alignment = center_align
                        cell.font = Font(name='微軟正黑體', size=11, bold=True)

                    for row_data in output_df.fillna("").values.tolist():
                        ws.append(row_data)

                    ws.auto_filter.ref = f"A1:I{ws.max_row}"

                    for row_idx, is_empty in enumerate(is_empty_list, start=2):
                        for col_idx in range(1, 10):
                            cell = ws.cell(row=row_idx, column=col_idx)
                            cell.border = full_border
                            cell.alignment = center_align
                            cell.font = ms_font
                            if is_empty:
                                cell.fill = red_fill

                    for col in ws.columns:
                        max_length = 0
                        column = col[0].column_letter
                        for cell in col:
                            if cell.value:
                                val_str = str(cell.value)
                                try: byte_len = len(val_str.encode('big5'))
                                except: byte_len = len(val_str)
                                if byte_len > max_length: max_length = byte_len
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
                    st.error(f"❌ 處理過程中發生非預期錯誤: {e}")

# ==================== 🖥️ 功能三：庫存資料 + 一維碼一鍵產出 ====================
elif app_mode == "🏷️ 庫存資料+一維碼 一鍵產出":
    st.markdown("<h2 style='color: #2b5c8f; font-weight: 700;'>🏷️ 庫存資料與一維條碼整合系統</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #555;'>對應原 VBA 巨集。請同時投入三份對應報表，系統會全自動跨表關聯、智能取消合併儲存格並向下填補、自動篩選清除中文字儲位，並針對「庫存減銷售小於等於 0」的列自動刷黃底標記！</p>", unsafe_allow_html=True)
    st.markdown("---")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="upload-card">', unsafe_allow_html=True)
        st.markdown('<div class="custom-section-title">📦 ① 拆櫃明細檔案</div>', unsafe_allow_html=True)
        file_main = st.file_uploader("請上傳 拆櫃明細 (.xlsx)", type=["xlsx"], key="vba_main")
        st.markdown("</div>", unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="upload-card">', unsafe_allow_html=True)
        st.markdown('<div class="custom-section-title">🗺️ ② 貨架位檔案</div>', unsafe_allow_html=True)
        file_shelf = st.file_uploader("請上傳 貨架位/儲位表 (.xlsx)", type=["xlsx"], key="vba_shelf")
        st.markdown("</div>", unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="upload-card">', unsafe_allow_html=True)
        st.markdown('<div class="custom-section-title">📋 ③ 採購建議檔案</div>', unsafe_allow_html=True)
        file_purchase = st.file_uploader("請上傳 採購建議/建議表 (.xlsx)", type=["xlsx"], key="vba_purchase")
        st.markdown("</div>", unsafe_allow_html=True)

    if st.button("🚀 啟動多表整合&自動產出", type="primary", use_container_width=True):
        if not (file_main and file_shelf and file_purchase):
            st.error("🚨 錯誤：您必須同時上傳「拆櫃明細」、「貨架位」與「採購建議」三份檔案才能啟動流程！")
        else:
            with st.spinner("正在執行中..."):
                try:
                    df_main = pd.read_excel(file_main, skiprows=3)
                    new_headers = ["商品編號", "商品名稱", "商品規格", "品項條碼", "箱裝數", "叫貨數量", "件數", "福北總庫存", "15日銷售", "福撿儲位", "一維條碼"]
                    df_main.columns = new_headers[:len(df_main.columns)]
                    
                    df_shelf_raw = pd.read_excel(file_shelf)
                    df_shelf_raw.iloc[:, 1] = df_shelf_raw.iloc[:, 1].ffill()
                    
                    shelf_dict = {}
                    for _, row in df_shelf_raw.iterrows():
                        barcode = str(row.iloc[4]).strip() if not pd.isna(row.iloc[4]) else ""
                        loc = str(row.iloc[1]).strip() if not pd.isna(row.iloc[1]) else ""
                        if barcode:
                            shelf_dict[barcode] = loc

                    df_pur_raw = pd.read_excel(file_purchase)
                    pur_stock_dict = {}
                    pur_sale_dict = {}
                    for _, row in df_pur_raw.iterrows():
                        barcode = str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else ""
                        stock = row.iloc[5] if not pd.isna(row.iloc[5]) else 0
                        sale = row.iloc[12] if not pd.isna(row.iloc[12]) else 0
                        if barcode:
                            pur_stock_dict[barcode] = stock
                            pur_sale_dict[barcode] = sale

                    yellow_rows = set()
                    
                    for idx, row in df_main.iterrows():
                        search_key = str(row["品項條碼"]).strip() if not pd.isna(row["品項條碼"]) else ""
                        if not search_key or search_key == "nan":
                            continue
                        
                        stock_val = pur_stock_dict.get(search_key, 0)
                        sale_val = pur_sale_dict.get(search_key, 0)
                        df_main.at[idx, "福北總庫存"] = stock_val
                        df_main.at[idx, "15日銷售"] = sale_val
                        
                        shelf_val = shelf_dict.get(search_key, "")
                        df_main.at[idx, "福撿儲位"] = shelf_val
                        df_main.at[idx, "一維條碼"] = f'="*" & D{idx+5} & "*"'
                        
                        if shelf_val and shelf_val not in ["W-兩倉暫存區", "W-線東品"]:
                            try:
                                diff = float(stock_val) - float(sale_val)
                                if diff <= 0:
                                    yellow_rows.add(idx + 5)
                            except:
                                pass

                    wb = Workbook()
                    ws = wb.active
                    ws.title = "庫存一維碼明細"

                    thin_side = Side(border_style="thin", color="000000")
                    full_border = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)
                    center_align = Alignment(horizontal='center', vertical='center')
                    ms_font = Font(name='微軟正黑體', size=11)
                    barcode_font = Font(name='Free 3 of 9 Extended', size=30)
                    yellow_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")

                    for _ in range(3):
                        ws.append([])
                    
                    ws.append(new_headers)
                    for col_idx in range(1, 12):
                        cell = ws.cell(row=4, column=col_idx)
                        cell.border = full_border
                        cell.alignment = center_align
                        cell.font = Font(name='微軟正黑體', size=11, bold=True)

                    for _, row_data in df_main.fillna("").iterrows():
                        ws.append(list(row_data))

                    ws.auto_filter.ref = f"A4:K{ws.max_row}"

                    for r_idx in range(5, ws.max_row + 1):
                        is_yellow = r_idx in yellow_rows
                        for c_idx in range(1, 12):
                            cell = ws.cell(row=r_idx, column=c_idx)
                            cell.border = full_border
                            cell.alignment = center_align
                            cell.font = ms_font
                            if c_idx == 11:
                                cell.font = barcode_font
                            if is_yellow:
                                cell.fill = yellow_fill

                    for col in ws.columns:
                        max_length = 0
                        column = col[0].column_letter
                        for cell in col:
                            if cell.value:
                                val_str = str(cell.value)
                                if val_str.startswith('="*'):
                                    byte_len = 15
                                else:
                                    try: byte_len = len(val_str.encode('big5'))
                                    except: byte_len = len(val_str)
                                if byte_len > max_length: max_length = byte_len
                        ws.column_dimensions[column].width = max_length + 4

                    excel_data = BytesIO()
                    wb.save(excel_data)
                    excel_data.seek(0)

                    st.success("✨ VBA 庫存大整合全自動演算法執行成功！中文字體儲位已自動排除。")
                    st.download_button(
                        label="📥 點此一鍵下載全新庫存條碼整合 Excel 檔案",
                        data=excel_data,
                        file_name="#義烏櫃 拆櫃明細-福北路-庫存資料&一維條碼.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"❌ 跨表運算時發生非預期錯誤: {e}")

# ==================== 🖥️ 功能四：正隆帳單核對 ====================
elif app_mode == "🧾 正隆帳單核對":
    st.markdown("<h2 style='color: #2b5c8f; font-weight: 700;'>🧾 正隆帳單自動化核對系統</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #555;'>本系統分為兩個階段：【第一階段】整合 LINE 對話叫貨紀錄，以及【第二階段】核對正隆帳單 PDF。請依序操作。</p>", unsafe_allow_html=True)
    st.markdown("---")

    # 建立兩個分頁，方便未來擴充第二階段
    tab1, tab2 = st.tabs(["📱 第一階段：整合 LINE 叫貨紀錄", "📄 第二階段：正隆帳單 PDF 核對"])

    # ------------------ 📱 TAB 1: LINE 紀錄整合 ------------------
    with tab1:
        st.markdown('<div class="upload-card">', unsafe_allow_html=True)
        st.markdown('<div class="custom-section-title">📅 步驟 1：選擇篩選日期區間</div>', unsafe_allow_html=True)
        
        # 使用 Streamlit 行事曆元件，免去手動輸入 MM/DD 的麻煩與格式錯誤風險
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            start_date = st.date_input("開始日期", value=pd.Timestamp.now().floor('D') - pd.Timedelta(days=30), key="line_start")
        with col_d2:
            end_date = st.date_input("結束日期", value=pd.Timestamp.now().floor('D'), key="line_end")
            
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="upload-card">', unsafe_allow_html=True)
        st.markdown('<div class="custom-section-title">📁 步驟 2：上傳 LINE 對話紀錄檔案 (.txt)</div>', unsafe_allow_html=True)
        line_file = st.file_uploader("請上傳 `[LINE]桃園正隆X沐樂.txt` 檔案", type=["txt"], key="uline_chat")
        st.markdown('</div>', unsafe_allow_html=True)

        if line_file is not None:
            # 讀取上傳的 TXT 文字，支援多種編碼格式防呆
            raw_chat_data = ""
            file_bytes = line_file.read()
            for enc in ['utf-8', 'utf-16', 'cp950']:
                try:
                    raw_chat_data = file_bytes.decode(enc)
                    break
                except:
                    continue

            if not raw_chat_data:
                st.error("🚨 無法解析該文字檔的編碼，請確認檔案是否損毀或使用標準 TXT 格式。")
            else:
                # ------------------ 核心解析演算法邏輯 ------------------
                current_year = start_date.year
                start_dt = datetime(start_date.year, start_date.month, start_date.day)
                end_dt = datetime(end_date.year, end_date.month, end_date.day)

                anchor_text = "07:38 善敏 01/05(一)排程報告："
                report_start = raw_chat_data.find(anchor_text)
                data_to_process = raw_chat_data[report_start + len(anchor_text):] if report_start != -1 else raw_chat_data

                order_marker_pattern = r"您好\s*[，,]\s*與您訂購紙箱"
                order_markers = list(re.finditer(order_marker_pattern, data_to_process))
                date_pattern = r"(?<![:\d\-])(?:\d{4}/)?(\d{1,2})/(\d{1,2})(?!\d)"
                all_dates = list(re.finditer(date_pattern, data_to_process))
                timestamp_pattern = r"(?m)^\d{1,2}:\d{2}\s"
                all_timestamps = list(re.finditer(timestamp_pattern, data_to_process))

                results = []
                for o_idx, o_match in enumerate(order_markers):
                    order_pos = o_match.start()
                    nearest_date = None
                    for d_match in all_dates:
                        if d_match.start() < order_pos: 
                            nearest_date = d_match
                        else: 
                            break
                    
                    if not nearest_date: 
                        continue
                    try:
                        msg_dt = datetime(current_year, int(nearest_date.group(1)), int(nearest_date.group(2)))
                    except: 
                        continue
                    
                    if not (start_dt <= msg_dt <= end_dt): 
                        continue
                    
                    next_timestamp_pos = len(data_to_process)
                    for t_match in all_timestamps:
                        if t_match.start() > order_pos:
                            next_timestamp_pos = t_match.start()
                            break
                    
                    next_order_pos = order_markers[o_idx+1].start() if o_idx + 1 < len(order_markers) else len(data_to_process)
                    block_end = min(next_timestamp_pos, next_order_pos)
                    order_block = data_to_process[o_match.start():block_end]
                    
                    name_m = re.search(r"收件人(?:姓名)?[:：]\s*(.*)", order_block)
                    addr_m = re.search(r"(?:收件人地址|收件地址|收貨地址|地址)[:：]\s*(.*)", order_block)
                    name = name_m.group(1).strip() if name_m else "未填"
                    addr = addr_m.group(1).strip() if addr_m else "未填"
                    
                    size_pattern = r"(\d+(?:\.\d+)?\s*\*\s*\d+(?:\.\d+)?\s*\*\s*\d+(?:\.\d+)?)"
                    qty_pattern = r"(\d+)\s*個"
                    all_sizes = list(re.finditer(size_pattern, order_block))
                    
                    for s_idx, s_match in enumerate(all_sizes):
                        size_val = s_match.group(1).replace(" ", "")
                        qty_m = re.search(qty_pattern, order_block[s_match.end():])
                        if qty_m:
                            results.append({
                                "訂購尺寸": size_val,
                                "叫貨日期": msg_dt.strftime("%Y/%m/%d"),
                                "數量": int(qty_m.group(1)),
                                "未稅單價": "", "未稅總價": "", "含稅總價": "",
                                "收件人姓名": name, "收件人地址": addr
                            })

                # ------------------ 畫面顯示與同步處理 ------------------
                if results:
                    results = sorted(results, key=lambda x: (x["叫貨日期"], x["訂購尺寸"]))
                    df_results = pd.DataFrame(results)
                    
                    st.markdown("### 📊 偵測並解析成功之訂單預覽")
                    st.dataframe(df_results, use_container_width=True)
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("☁️ 確認無誤，一鍵回填至雲端 Google Sheets", type="primary", use_container_width=True):
                        with st.spinner("正在連線並上傳至雲端試算表..."):
                            try:
                                client = get_gspread_client()
                                if client:
                                    # 使用本地 parse_order.py 設定之專屬 SHEET_ID 與工作表名稱
                                    LINE_SHEET_ID = "1hklqBQ_9Z0HZcgHF-3Kl1jdXqDa13NXOqwHZ1BZnxQg"
                                    LINE_SHEET_NAME = "叫貨紀錄"
                                    
                                    sh = client.open_by_key(LINE_SHEET_ID)
                                    worksheet = sh.worksheet(LINE_SHEET_NAME)
                                    
                                    first_row = worksheet.row_values(1)
                                    headers = ["訂購尺寸", "叫貨日期", "數量", "未稅單價", "未稅總價", "含稅總價", "收件人姓名", "收件人地址"]
                                    
                                    if not first_row:
                                        worksheet.insert_row(headers, 1)
                                    
                                    rows_to_append = []
                                    for d in results:
                                        rows_to_append.append([
                                            d.get("訂購尺寸", ""), d.get("叫貨日期", ""), d.get("數量", ""),
                                            d.get("未稅單價", ""), d.get("未稅總價", ""), d.get("含稅總價", ""),
                                            d.get("收件人姓名", ""), d.get("收件人地址", "")
                                        ])
                                    
                                    worksheet.append_rows(rows_to_append, value_input_option='USER_ENTERED')
                                    st.toast(f"✅ 成功同步 {len(rows_to_append)} 筆資料到雲端！")
                                    st.success(f"✨ 雲端同步完成！已成功回填 {len(rows_to_append)} 筆叫貨資料。")
                            except Exception as e:
                                st.error(f"❌ 雲端上傳失敗：{e}")
                else:
                    st.warning("⚠️ 在選擇的日期區間內，找不到符合格式的訂單數據。請檢查日期區間是否正確。")

  # ------------------ 📄 TAB 2: PDF 帳單核對 (高精度 OCR 交叉核對版) ------------------
    with tab2:
        st.markdown('<div class="upload-card">', unsafe_allow_html=True)
        st.markdown('<div class="custom-section-title">📄 步驟 1：請上傳正隆帳單混合 PDF 原始檔案 (.pdf)</div>', unsafe_allow_html=True)
        pdf_file = st.file_uploader("將正隆帳單 PDF 檔案拖放到此處", type=["pdf"], key="updf_billing_main")
        st.markdown("</div>", unsafe_allow_html=True)

        if pdf_file is not None:
            if st.button("🚀 啟動 Google 雲端高精度 OCR 提取與交叉對帳", type="primary", use_container_width=True):
                with st.spinner("正在透過 Google 雲端 AI 引擎進行 99% 精準度字元識別，請稍候..."):
                    try:
                        # 1. 建立 Google Drive API 連線 (直接複用現有的 secrets 金鑰)
                        scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
                        
                        # 優先讀取 Streamlit 雲端環境的 secrets，若無則讀取本地變數
                        if "gcp_service_account" in st.secrets:
                            secret_data = st.secrets["gcp_service_account"]
                            creds_dict = json.loads(secret_data) if isinstance(secret_data, str) else secret_data
                            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
                        else:
                            # 讀取本地端
                            SERVICE_ACCOUNT_FILE = "credentials.json" # 若檔名為 service_account.json 請自行修正
                            creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope)
                        
                        drive_service = build('drive', 'v3', credentials=creds)
                        st.info(f"🔍 目前網頁端正在使用的機器人信箱為：`{creds.service_account_email}`")

                        # 2. 上傳 PDF 並指定轉換為 Google Docs 格式 (此動做會觸發 Google 免費的高精度 OCR 引擎)
                    # 💡 修改後的完美寫法（指定丟到有空間的共用資料夾）
                     # 2. 上傳 PDF 並指定轉換為 Google Docs 格式
                        # 💡 關鍵修正：以下這幾行全部要向右縮排 4 個空格，保持在 try: 的肚子裡面！
                        file_metadata = {
                            'name': 'Temp_OCR_File',
                            'mimeType': 'application/vnd.google-apps.document',
                            'parents': ['1fR0E2oUqiRBkdR58VAJ8ooO2Im381zfE']
                        }
                        
                        pdf_bytes = pdf_file.getvalue()
                        media = MediaIoBaseUpload(BytesIO(pdf_bytes), mimetype='application/pdf', resumable=True)
                        
                        uploaded_file = drive_service.files().create(
                            body=file_metadata, media_body=media, fields='id'
                        ).execute()
                        
                        file_id = uploaded_file.get('id')

                        # 3. 導出下載 OCR 轉換完畢後的純文字檔
                        export_media = drive_service.files().export_media(
                            fileId=file_id, mimeType='text/plain'
                        ).execute()
                        
                        raw_ocr_text = export_media.decode('utf-8')

                        # 4. 立即刪除雲端暫存檔，保護營業隱私與空間安全
                        drive_service.files().delete(fileId=file_id).execute()
                        st.toast("⚡ OCR 文字辨識成功！暫存檔已安全刪除。開始解析核心欄位...")

                        # --------------------------------------------------------
                        # 5. 核心解析演算法：針對 Google OCR 文本結構進行提取
                        # --------------------------------------------------------
                        lines = [line.strip() for line in raw_ocr_text.split('\n') if line.strip()]
                        full_flat_text = "".join(lines).replace(" ", "")

                        statement_results = []
                        invoice_results = []
                        
                        # 發票號碼正規表達式 (2碼大寫英文+8碼數字)
                        inv_pattern = r"([A-Z]{2}\d{8})"

                        # 金額防呆清洗函數 (承襲自您的演算法精髓)
                        def clean_and_parse_amount(s):
                            s = s.replace(",", "").replace(" ", "")
                            if "." in s and len(s.split(".")[-1]) == 3: # 修正千分位被誤判為句點的問題
                                s = s.replace(".", "")
                            try: return float(s)
                            except: return 0.0

                        # ================= [解析樣式 1：對帳單區塊] =================
                        for idx, line in enumerate(lines):
                            # 判定如果該行含有發票號碼，且整份文件有對帳單特徵
                            if "對帳單" in full_flat_text and re.search(inv_pattern, line):
                                m_inv = re.search(inv_pattern, line)
                                inv_no = m_inv.group(1)
                                
                                # 抓取品名規格 (通常材質長度含有 * 乘號特徵)
                                size_m = re.search(r"([\u4e00-\u9fff\w\(\)]+.*?\d+\*\d+\*\d+.*?)(?:\s|$)", line)
                                if not size_m and idx + 1 < len(lines):
                                    # 嘗試在下一行尋找品名
                                    size_m = re.search(r"([\u4e00-\u9fff\w\(\)]+.*?\d+\*\d+\*\d+.*?)(?:\s|$)", lines[idx+1])
                                
                                size_val = size_m.group(1).strip() if size_m else "紙箱 (常規規格)"
                                
                                # 提取所有數值，過濾掉單純的日期等雜訊
                                all_nums = re.findall(r"[\d,]+\.?\d*", line)
                                if len(all_nums) < 4 and idx + 1 < len(lines):
                                    all_nums += re.findall(r"[\d,]+\.?\d*", lines[idx+1])
                                
                                clean_nums = [n for n in all_nums if len(n.replace(",", "").split(".")[0]) >= 1]
                                
                                if len(clean_nums) >= 4:
                                    statement_results.append({
                                        "發票號碼": inv_no,
                                        "品名規格": size_val,
                                        "數量": clean_nums[-4],
                                        "未稅金額": clean_and_parse_amount(clean_nums[-3]),
                                        "稅額": clean_and_parse_amount(clean_nums[-2]),
                                        "總金額": clean_and_parse_amount(clean_nums[-1])
                                    })

                        # ================= [解析樣式 2：電子發票證明聯區塊] =================
                        blocks = raw_ocr_text.split("電子發票證明聯")
                        for block in blocks[1:]:
                            m_inv = re.search(inv_pattern, block)
                            inv_no = m_inv.group(1) if m_inv else "未知號碼"
                            
                            sales_amt, tax_amt, total_amt = 0.0, 0.0, 0.0
                            
                            for bline in block.split("\n"):
                                b_clean = bline.replace(" ", "")
                                if "銷售額" in b_clean or "金額" in b_clean:
                                    nums = re.findall(r"[\d,\.]+", bline)
                                    if nums: sales_amt = clean_and_parse_amount(nums[-1])
                                if "營業稅" in b_clean or "稅額" in b_clean:
                                    nums = re.findall(r"[\d,\.]+", bline)
                                    if nums: tax_amt = clean_and_parse_amount(nums[-1])
                                if "總計" in b_clean and "新台幣" not in bline:
                                    nums = re.findall(r"[\d,\.]+", bline)
                                    if nums: total_amt = clean_and_parse_amount(nums[-1])
                            
                            # 萬一總計漏抓，啟動金額關聯公式自動計算
                            if total_amt == 0.0 and sales_amt > 0:
                                tax_amt = tax_amt if tax_amt > 0 else round(sales_amt * 0.05)
                                total_amt = sales_amt + tax_amt

                            invoice_results.append({
                                "發票號碼": inv_no,
                                "品名規格": "電子發票明細",
                                "數量": "1",
                                "單價": sales_amt,
                                "營業稅": tax_amt,
                                "總計": total_amt
                            })

                        # --------------------------------------------------------
                        # 6. 視覺化對帳呈現
                        # --------------------------------------------------------
                        st.success("✨ PDF 帳單數據全自動高精度提取完畢！")
                        
                        df_stmt = pd.DataFrame(statement_results)
                        df_inv = pd.DataFrame(invoice_results)

                        col_v1, col_v2 = st.columns(2)
                        with col_v1:
                            st.markdown("#### 📊 1. 對帳單（樣式一）明細提取")
                            if not df_stmt.empty:
                                st.dataframe(df_stmt, use_container_width=True)
                            else:
                                st.info("未偵測到對帳單格式頁面。")
                                
                        with col_v2:
                            st.markdown("#### 📄 2. 電子發票證明聯（樣式二）明細提取")
                            if not df_inv.empty:
                                st.dataframe(df_inv, use_container_width=True)
                            else:
                                st.info("未偵測到電子發票證明聯。")

                        # ================= [兩條線交叉對帳系統核心] =================
                        if not df_stmt.empty and not df_inv.empty:
                            st.markdown("---")
                            st.markdown("<h3 style='color: #2b5c8f;'>⚖️ 兩條線自動化交叉對帳結果 (以發票號碼為主鍵)</h3>", unsafe_allow_html=True)
                            
                            # 整合兩張表格進行一對一比對
                            df_recon = pd.merge(
                                df_stmt[['發票號碼', '總金額']], 
                                df_inv[['發票號碼', '總計']], 
                                on='發票號碼', how='outer'
                            ).fillna(0.0)
                            
                            df_recon.columns = ['發票號碼', '對帳單總金額 (A)', '發票證明聯總計 (B)']
                            df_recon['金額差異 (B-A)'] = df_recon['發票證明聯總計 (B)'] - df_recon['對帳單總金額 (A)']
                            
                            def get_recon_status(diff):
                                return "✓ 金額完全一致" if abs(diff) <= 1 else f"❌ 異常！金額不符 (差額 {diff:g})"
                            df_recon['核對結果狀態'] = df_recon['金額差異 (B-A)'].apply(get_recon_status)
                            
                            # 美化表格輸出
                            st.dataframe(df_recon, use_container_width=True)
                            
                            # 統計摘要摘要
                            bad_counts = df_recon['核對結果狀態'].str.contains("異常").sum()
                            if bad_counts == 0:
                                st.balloons()
                                st.success(f"🎉 終極大特報：本份帳單中所有發票與對帳單金額【完全吻合】，未發現任何異常！")
                            else:
                                st.error(f"⚠️ 注意：偵測到共有 {bad_counts} 筆發票存在金額不符或單邊漏開現象，請查看上方表格紅字！")
                                
                    except Exception as e:
                        st.error(f"❌ 雲端高精度 OCR 執行時發生非預期錯誤，請檢查專案權限。錯誤詳情: {e}")
