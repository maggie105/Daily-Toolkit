import streamlit as st
import pandas as pd
import os
import time
import gspread
import warnings
import re
import json
from oauth2client.service_account import ServiceAccountCredentials
from gspread.exceptions import APIError

# 1. 隱藏 openpyxl 的格式警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ================= 網頁頁面初始設定 =================
st.set_page_config(page_title="BS銷售更新", page_icon="📊", layout="wide")

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
    
    /* 🎨 1. 將大標題強制改為極簡深黑色 */
    .custom-main-title {
        color: #1e293b !important;  
        font-size: 1.8rem !important;
        font-weight: 700 !important;
        margin-bottom: 0.5rem;
    }
    
    /* 🎨 2. 將啟動按鈕改為北歐風冷岩灰藍色、白字 */
    .stButton > button {
        background-color: #475569 !important; 
        border-color: #475569 !important;
        color: #ffffff !important;
        font-weight: 600 !important;
        font-size: 16px !important;
        border-radius: 8px !important;
        padding: 12px 0 !important;
        transition: all 0.2s ease !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05) !important;
    }
    /* 按鈕滑鼠懸停效果 */
    .stButton > button:hover {
        background-color: #334155 !important; 
        border-color: #334155 !important;
        color: #ffffff !important;
    }
    /* 強制按鈕內部的文字維持純白與粗體 */
    .stButton > button p {
        color: #ffffff !important;
        font-weight: 600 !important;
    }
    
    /* 自訂精訊的中等標題字體樣式 */
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
    def retry_on_api_error(func, *args, max_retries=3, delay=5, **kwargs):
    """
    自動重試機制:遇到暫時性的 API 錯誤(503服務不可用、429太多請求)時，
    會自動等待後重試，最多重試 max_retries 次。
    """
    for attempt in range(1, max_retries + 1):
        try:
            return func(*args, **kwargs)
        except APIError as e:
            error_str = str(e)
            is_temporary = "503" in error_str or "429" in error_str or "500" in error_str
            if is_temporary and attempt < max_retries:
                st.toast(f"⏳ 雲端服務暫時忙碌，{delay}秒後自動重試 ({attempt}/{max_retries})...")
                time.sleep(delay)
                continue
            else:
                raise

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
        retry_on_api_error(worksheet.batch_clear, [f"A:{end_col}"])
        
        data_to_upload = []
        if title_list:
            data_to_upload.append([("" if pd.isna(x) else x) for x in title_list])
            data_to_upload.append(df.columns.tolist())
        else:
            data_to_upload.append(df.columns.tolist())

        clean_df = df.fillna("")
        for row in clean_df.values.tolist():
            data_to_upload.append([("" if pd.isna(cell) else cell) for cell in row])
        
        retry_on_api_error(worksheet.update, values=data_to_upload, range_name='A1')
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

        raw_date_text = retry_on_api_error(sheet_source.acell, 'B1').value 
        found_dates = re.findall(r'\d{4}-\d{2}-\d{2}', str(raw_date_text))
        if len(found_dates) >= 2:
            d1 = found_dates[0].replace("-", "")[4:]
            d2 = found_dates[1].replace("-", "")[4:]
            final_date_str = f"{d1}-{d2}"
            retry_on_api_error(sheet_target.update, values=[[final_date_str]], range_name='B1')
            st.write(f"📅 雲端主表 B1 日期同步完成: `{final_date_str}`")

        v_values = retry_on_api_error(sheet_source.col_values, 22)[4:] 
        filtered_v = [[v] for v in v_values if v and str(v).strip() not in ["ZZ000-04", "總和"]]
        retry_on_api_error(sheet_target.batch_clear, ["A3:A"]) 
        if filtered_v:
            retry_on_api_error(sheet_target.update, values=filtered_v, range_name='A3')
            st.write("🏷️ 雲端 A 欄品項條碼二次同步成功。")
    except Exception as e:
        st.error(f"❌ 後處理二次同步出錯: {e}")

# ==================== 主畫面內容 ====================
# 🎯 套用自訂深黑標題類別
st.markdown('<div class="custom-main-title">📊 BigSeller 銷售數據更新系統</div>', unsafe_allow_html=True)
st.markdown("""
    <p style='color: #555; margin-bottom: 5px;'>請上傳對應的 Excel 報表，系統將自動進行多檔清洗、字體校正並同步至雲端。</p>
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
    uploaded_files_onl = st.file_uploader("② 請上傳「其他店鋪」的檔案 (可多選)", type=["xlsx"], accept_multiple_files=True, key="uonl")
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

# 🎯 這裡的 Button 就會自動套用 CSS 變成優雅的冷岩灰藍底色囉！
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
