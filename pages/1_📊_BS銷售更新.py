import streamlit as st
import pandas as pd
import os
import time
import gspread
import warnings
import re
import json
from oauth2client.service_account import ServiceAccountCredentials

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
st.set_page_config(page_title="BS銷售更新", page_icon="📊", layout="wide")

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

HEADERS_01 = ["Image URL", "SKU编号", "名称", "重量(g)", "长(cm)", "宽(cm)", "高(cm)", "仓库", "库区", "货架位", "现有库存", "订单已锁", "整仓可用", "整仓未上架", "活动预留", "在途中", "在途总成本", "警戒库存", "预测日销量", "预计可售天数", "加權成本價", "總成本價", "銷售狀態", "一級分類", "二級分類", "三級分類", "SKU類型", "備註"]
HEADERS_02 = ["店铺昵称", "分类_L1", "分类_L2", "分类_L3", "产品名称", "Item ID", "销量", "收藏", "浏览量", "Parent SKU", "变种", "变种ID", "SKU", "库存", "价格", "促销价", "限购", "平台创建時間", "发货期"]
HEADERS_03 = ["商品名称", "店铺", "商品SKU", "有效商品销售额", "有效订单量", "有效商品销量", "商品平均价格", "商品销售额", "商品销量", "订单总量", "包裹总量", "退款商品金额", "退款订单数", "退款商品数", "取消商品金額", "取消訂單數", "取消商品數"]
HEADERS_04 = ["商品SKU", "店铺", "商品名称", "分类", "商品收入", "总成本", "利润", "单个商品利润", "利润率", "订单数量", "销售数量", "退货数量", "退货率", "商品总销售额", "折扣&优惠补贴", "买家支付运费", "卖家支付運費", "佣金", "交易費", "服務費", "營銷費用", "退款金額", "平台其他費用"]

def normalize_header(header_name):
    s = str(header_name).strip()
    trad_to_simp = {"量": "量", "銷": "销", "額": "额", "庫": "库", "存": "存", "單": "单", "價": "价", "利": "利", "潤": "润", "貨": "货", "類": "类", "幣": "币", "應": "应", "實": "实", "項": "项", "權": "权", "總": "总", "狀": "状", "態": "态", "級": "级", "備": "备", "註": "注"}
    for trad, simp in trad_to_simp.items(): s = s.replace(trad, simp)
    return s

def get_worksheet_by_gid(sh, gid):
    for sheet in sh.worksheets():
        if str(sheet.id) == str(gid): return sheet
    return None

def safe_read_and_align_uploaded(uploaded_file, target_headers, task_key, header_row=0):
    try: df = pd.read_excel(uploaded_file, header=header_row)
    except Exception as e: st.error(f"❌ 無法讀取檔案 `{uploaded_file.name}`。錯誤資訊: {e}"); return None
    original_cols = df.columns.tolist()
    norm_original_cols = [normalize_header(col) for col in original_cols]
    norm_to_orig = {normalize_header(col): col for col in original_cols}
    for f in SHEET_CONFIGS[task_key]["features"]:
        if normalize_header(f) not in norm_original_cols: st.error(f"🚨 檔案防呆攔截：找不到關鍵欄位 `{f}`！"); return None
    df_aligned = pd.DataFrame()
    for target in target_headers:
        target_norm = normalize_header(target)
        if target_norm in norm_to_orig: df_aligned[target] = df[norm_to_orig[target_norm]]
        else: df_aligned[target] = ""
    return df_aligned[target_headers]

def get_gspread_client():
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    if len(st.secrets) > 0 and "gcp_service_account" in st.secrets:
        return gspread.authorize(ServiceAccountCredentials.from_json_keyfile_dict(json.loads(st.secrets["gcp_service_account"]) if isinstance(st.secrets["gcp_service_account"], str) else st.secrets["gcp_service_account"], scope))
    if os.path.exists(SERVICE_ACCOUNT_FILE): return gspread.authorize(ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope))
    return None

def upload_to_google_sheets(df, task_key, title_list=None):
    if df is None or df.empty: return
    try:
        client = get_gspread_client()
        if not client: return
        sh = client.open_by_key(SPREADSHEET_ID)
        worksheet = get_worksheet_by_gid(sh, SHEET_CONFIGS[task_key]["gid"])
        if not worksheet: return
        worksheet.batch_clear([f"A:{SHEET_CONFIGS[task_key]['end_col']}"])
        data_to_upload = []
        if title_list: data_to_upload.append([("" if pd.isna(x) else x) for x in title_list])
        data_to_upload.append(df.columns.tolist())
        for row in df.fillna("").values.tolist(): data_to_upload.append([("" if pd.isna(cell) else cell) for cell in row])
        worksheet.update(values=data_to_upload, range_name='A1')
        st.write(f"🟢 {worksheet.title} 數據已成功同步。")
    except Exception as e: st.error(f"❌ 雲端同步失敗。錯誤: {e}")

def post_process_steps():
    try:
        client = get_gspread_client()
        sh = client.open_by_key(SPREADSHEET_ID)
        s_sheet = get_worksheet_by_gid(sh, SHEET_CONFIGS["03"]["gid"])
        t_sheet = get_worksheet_by_gid(sh, SHEET_CONFIGS["target_main"]["gid"])
        dates = re.findall(r'\d{4}-\d{2}-\d{2}', str(s_sheet.acell('B1').value))
        if len(dates) >= 2:
            f_date = f"{dates[0].replace('-', '')[4:]}-{dates[1].replace('-', '')[4:]}"
            t_sheet.update(values=[[f_date]], range_name='B1')
        v_vals = [[v] for v in s_sheet.col_values(22)[4:] if v and str(v).strip() not in ["ZZ000-04", "總和"]]
        t_sheet.batch_clear(["A3:A"])
        if v_vals: t_sheet.update(values=v_vals, range_name='A3')
    except Exception as e: st.error(f"❌ 後處理出錯: {e}")

st.markdown("<h2 style='color: #2b5c8f; font-weight: 700;'>📊 BigSeller 銷售數據更新系統</h2>", unsafe_allow_html=True)
st.markdown("👉 <a href='https://docs.google.com/spreadsheets/d/1FLfAbqq1TmQnXFR3rHxGkrXELrlKSMpiXGYXk7hVZm8/edit?gid=1324377276#gid=1324377276' target='_blank' style='color: #2b5c8f; font-weight: bold; text-decoration: underline;'>點此打開 Google Sheets 雲端主表</a>", unsafe_allow_html=True)
st.markdown("---")

col1, col2 = st.columns(2)
with col1:
    uploaded_files_01 = st.file_uploader("📁 01. 庫存清單 (可多選)", type=["xlsx"], accept_multiple_files=True)
    uploaded_files_lan = st.file_uploader("📁 02. ① 懶餅乾檔案 (可多選)", type=["xlsx"], accept_multiple_files=True)
    uploaded_files_onl = st.file_uploader("📁 02. ② Online產品檔案 (可多選)", type=["xlsx"], accept_multiple_files=True)
with col2:
    uploaded_file_03 = st.file_uploader("📁 03. 銷量報告 (單選)", type=["xlsx"])
    uploaded_file_04 = st.file_uploader("📁 04. 利潤報告 (單選)", type=["xlsx"])

if st.button("🔥 啟動自動化整合全流程", type="primary", use_container_width=True):
    if not (uploaded_files_01 or uploaded_files_lan or uploaded_files_onl or uploaded_file_03 or uploaded_file_04):
        st.error("🚨 請先上傳檔案。")
    else:
        start_time = time.time(); has_error = False
        with st.spinner("同步中..."):
            if uploaded_files_01:
                list_dfs = [safe_read_and_align_uploaded(f, HEADERS_01, "01") for f in uploaded_files_01 if safe_read_and_align_uploaded(f, HEADERS_01, "01") is not None]
                if list_dfs: upload_to_google_sheets(pd.concat(list_dfs, ignore_index=True), "01")
            if uploaded_files_lan or uploaded_files_onl:
                l_dfs = [safe_read_and_align_uploaded(f, HEADERS_02, "02") for f in uploaded_files_lan if safe_read_and_align_uploaded(f, HEADERS_02, "02") is not None]
                o_dfs = [safe_read_and_align_uploaded(f, HEADERS_02, "02") for f in uploaded_files_onl if safe_read_and_align_uploaded(f, HEADERS_02, "02") is not None]
                df_l = pd.concat(l_dfs, ignore_index=True).drop_duplicates(subset=['SKU']) if l_dfs else pd.DataFrame(columns=HEADERS_02)
                df_o = pd.concat(o_dfs, ignore_index=True).drop_duplicates(subset=['SKU']) if o_dfs else pd.DataFrame(columns=HEADERS_02)
                df02 = pd.concat([df_l, df_o], ignore_index=True)
                if not df02.empty:
                    df02['priority'] = df02['店铺昵称'].apply(lambda x: 0 if x == "02-懶餅乾家居" else 1)
                    df02 = df02.sort_values(by=['SKU', 'priority']).drop_duplicates(subset=['SKU'], keep='first').drop(columns=['priority'])
                    upload_to_google_sheets(df02, "02")
            if uploaded_file_03:
                try: upload_to_google_sheets(safe_read_and_align_uploaded(uploaded_file_03, HEADERS_03, "03", header_row=1), "03", title_list=pd.read_excel(uploaded_file_03, nrows=1, header=None).values.tolist()[0])
                except: has_error = True
            if uploaded_file_04:
                try: upload_to_google_sheets(safe_read_and_align_uploaded(uploaded_file_04, HEADERS_04, "04", header_row=1), "04", title_list=pd.read_excel(uploaded_file_04, nrows=1, header=None).values.tolist()[0])
                except: has_error = True
            if uploaded_file_03 and not has_error: post_process_steps()
        st.success(f"✨ 全部自動化流程已執行完畢！總耗時：{round(time.time() - start_time, 2)} 秒")
