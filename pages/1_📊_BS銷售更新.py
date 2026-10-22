import json
import os
import re
import time
import unicodedata
import warnings
from datetime import date

import gspread
import pandas as pd
import streamlit as st
from gspread.exceptions import APIError
from oauth2client.service_account import ServiceAccountCredentials


warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


# ================= Google Sheets 核心設定 =================
SPREADSHEET_ID = "1FLfAbqq1TmQnXFR3rHxGkrXELrlKSMpiXGYXk7hVZm8"
SERVICE_ACCOUNT_FILE = "service_account.json"

SHEET_CONFIGS = {
    "01": {
        "gid": "1735114343", "end_col": "AB", "key_col": "B",
        "features": ["SKU编号", "现有库存"],
    },
    "02": {
        "gid": "830730835", "end_col": "S", "key_col": "M",
        "features": ["店铺昵称", "Item ID", "SKU"],
    },
    "03": {
        "gid": "150603716", "end_col": "Q", "key_col": "C",
        "features": ["商品名称", "有效商品销售额"],
    },
    "04": {
        "gid": "790740822", "end_col": "W", "key_col": "A",
        "features": ["商品SKU", "利润", "利润率"],
    },
    "target_main": {"gid": "1324377276"},
}


# ================= 標準表頭定義 =================
HEADERS_01 = [
    "Image URL", "SKU编号", "名称", "重量(g)", "长(cm)", "宽(cm)", "高(cm)",
    "仓库", "库区", "货架位", "现有库存", "订单已锁", "整仓可用", "整仓未上架",
    "活动预留", "在途中", "在途总成本", "警戒库存", "预测日销量", "预计可售天数",
    "加權成本價", "總成本價", "銷售狀態", "一級分類", "二級分類", "三級分類",
    "SKU類型", "備註",
]

HEADERS_02 = [
    "店铺昵称", "分类_L1", "分类_L2", "分类_L3", "产品名称", "Item ID", "销量",
    "收藏", "浏览量", "Parent SKU", "变种", "变种ID", "SKU", "库存", "价格",
    "促销价", "限购", "平台创建時間", "发货期",
]

# 雲端 03 工作表維持舊版 A:Q 架構，避免既有透視表和公式大幅改寫。
HEADERS_03 = [
    "商品名称", "店铺", "商品SKU", "有效商品销售额", "有效订单量", "有效商品销量",
    "商品平均价格", "商品销售额", "商品销量", "订单总量", "包裹总量", "退款商品金额",
    "退款订单数", "退款商品数", "取消商品金額", "取消訂單數", "取消商品數",
]

HEADERS_04 = [
    "商品SKU", "店铺", "商品名称", "分类", "商品收入", "总成本", "利润",
    "单个商品利润", "利润率", "订单数量", "销售数量", "退货数量", "退货率",
    "商品总销售额", "折扣&优惠补贴", "买家支付运费", "卖家支付運費", "佣金",
    "交易費", "服務費", "營銷費用", "退款金額", "平台其他費用",
]

NEW_03_REQUIRED_HEADERS = [
    "产品名称", "SKU_ID", "店铺名称", "店铺SKU编号", "订单总量", "有效订单总量",
    "产品销量", "有效产品销量", "产品销售额", "有效产品销售额", "取消订单量",
    "取消产品销售额", "取消产品销量", "退款订单量", "退款产品销售额", "退款产品销量",
]


# ================= 表頭與資料處理 =================
TRAD_TO_SIMP = {
    "銷": "销", "額": "额", "庫": "库", "單": "单", "價": "价", "潤": "润",
    "貨": "货", "類": "类", "幣": "币", "應": "应", "實": "实", "項": "项",
    "權": "权", "總": "总", "狀": "状", "態": "态", "級": "级", "備": "备",
    "註": "注", "時": "时", "間": "间", "運": "运", "費": "费", "務": "务",
    "營": "营", "稱": "称", "號": "号", "產": "产", "訂": "订", "鋪": "铺",
    "變": "变", "長": "长", "寬": "宽", "預": "预", "測": "测", "創": "创",
}

HEADER_ALIASES = {
    "商品均价": "商品平均价格",
}


def normalize_header(header_name):
    """統一空白、全半形與常見繁簡字，避免同義表頭被誤判為缺欄。"""
    value = unicodedata.normalize("NFKC", str(header_name)).strip()
    value = re.sub(r"\s+", "", value)
    for trad, simp in TRAD_TO_SIMP.items():
        value = value.replace(trad, simp)
    return HEADER_ALIASES.get(value, value)


def normalize_identifier(value):
    """將 Excel 常見的識別碼型態轉成穩定字串，並保留 SKU 內部符號。"""
    if pd.isna(value):
        return ""
    text = unicodedata.normalize("NFKC", str(value)).strip()
    if text.lower() in {"nan", "none", "<na>"}:
        return ""
    if re.fullmatch(r"[+-]?\d+\.0+", text):
        text = text.split(".", 1)[0]
    return text


def normalize_identifier_series(series):
    return series.map(normalize_identifier)


def looks_like_sku(value):
    """只用於判斷 01 的 SKU／名稱是否明顯對調，不否定較特殊的合法 SKU。"""
    text = normalize_identifier(value)
    return bool(
        text
        and text != "0"
        and re.fullmatch(r"[A-Za-z0-9-]+", text)
        and re.search(r"[A-Za-z]", text)
        and re.search(r"\d", text)
    )


def repair_inventory_sku_name_columns(df, file_name):
    """修正新版庫存清單中少數 SKU编号／名称欄位互換的資料列。"""
    repaired = df.copy()
    repaired["SKU编号"] = normalize_identifier_series(repaired["SKU编号"])
    repaired["名称"] = repaired["名称"].map(
        lambda value: "" if pd.isna(value) else str(value).strip()
    )

    swap_mask = (
        ~repaired["SKU编号"].map(looks_like_sku)
        & repaired["名称"].map(looks_like_sku)
    )
    if swap_mask.any():
        old_sku = repaired.loc[swap_mask, "SKU编号"].copy()
        repaired.loc[swap_mask, "SKU编号"] = repaired.loc[swap_mask, "名称"].values
        repaired.loc[swap_mask, "名称"] = old_sku.values
        repaired["SKU编号"] = normalize_identifier_series(repaired["SKU编号"])
        st.info(
            f"ℹ️ 01 資料已自動整理完成（修正 {int(swap_mask.sum())} 筆欄位對調）。"
        )

    return repaired


def validate_sales_skus(series, report_label, excel_row_offset):
    """銷量報告的 SKU 是下游彙總鍵；空白或 0 必須中止，避免錯誤合併。"""
    normalized = normalize_identifier_series(series)
    invalid_mask = normalized.isin(["", "0"])
    if invalid_mask.any():
        excel_rows = [int(index) + excel_row_offset for index in normalized.index[invalid_mask][:10]]
        raise ValueError(
            f"{report_label} 有 {int(invalid_mask.sum())} 列店鋪SKU為空白或 0；"
            f"請回 BS 修正後重新匯出。Excel 列號示例：{excel_rows}。"
        )
    return normalized


def reset_uploaded_file(uploaded_file):
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)


def uploaded_file_name(uploaded_file):
    return getattr(uploaded_file, "name", os.path.basename(str(uploaded_file)))


def read_excel(uploaded_file, **kwargs):
    reset_uploaded_file(uploaded_file)
    return pd.read_excel(uploaded_file, **kwargs)


def to_numeric_series(series):
    cleaned = series.astype(str).str.replace(",", "", regex=False).str.strip()
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


def clean_cell_for_google(cell):
    if pd.isna(cell):
        return ""
    if hasattr(cell, "item"):
        cell = cell.item()
    if isinstance(cell, (pd.Timestamp, date)):
        return cell.isoformat()
    return cell


def safe_read_and_align_uploaded(uploaded_file, target_headers, task_key, header_row=0):
    """讀取 01／02／04 或舊版 03，依標準欄位名稱重新排列。"""
    try:
        df = read_excel(uploaded_file, header=header_row)
    except Exception as exc:
        st.error(f"❌ 無法讀取檔案 `{uploaded_file_name(uploaded_file)}`。錯誤資訊: {exc}")
        return None

    original_cols = df.columns.tolist()
    normalized_to_original = {}
    for col in original_cols:
        normalized_to_original[normalize_header(col)] = col

    for feature in SHEET_CONFIGS[task_key]["features"]:
        if normalize_header(feature) not in normalized_to_original:
            st.error(
                f"🚨 檔案防呆攔截：`{uploaded_file_name(uploaded_file)}` "
                f"找不到關鍵欄位 `{feature}`。"
            )
            return None

    aligned = pd.DataFrame(index=df.index)
    for target in target_headers:
        source = normalized_to_original.get(normalize_header(target))
        aligned[target] = df[source] if source is not None else ""

    if task_key == "01":
        aligned = repair_inventory_sku_name_columns(
            aligned, uploaded_file_name(uploaded_file)
        )
    elif task_key == "02":
        # 必須在後續 drop_duplicates 之前整理，否則同一 SKU 只因空白而被當成兩筆。
        aligned["SKU"] = normalize_identifier_series(aligned["SKU"])
        aligned["店铺昵称"] = aligned["店铺昵称"].map(
            lambda value: "" if pd.isna(value) else str(value).strip()
        )
    elif task_key == "03":
        aligned["商品SKU"] = validate_sales_skus(
            aligned["商品SKU"], "舊版銷量報告", 3
        )

    return aligned[target_headers]


def detect_sales_report_type(uploaded_file):
    """辨識新版經營分析或舊版銷量報告。"""
    first_row_header = read_excel(uploaded_file, header=0, nrows=5)
    normalized = {normalize_header(col) for col in first_row_header.columns}
    required = {normalize_header(col) for col in NEW_03_REQUIRED_HEADERS}
    if required.issubset(normalized):
        return "business_analysis"

    second_row_header = read_excel(uploaded_file, header=1, nrows=5)
    normalized_old = {normalize_header(col) for col in second_row_header.columns}
    old_required = {
        normalize_header("商品名称"),
        normalize_header("店铺"),
        normalize_header("商品SKU"),
        normalize_header("有效商品销售额"),
        normalize_header("有效商品销量"),
    }
    if old_required.issubset(normalized_old):
        return "legacy_sales_report"

    raise ValueError("無法辨識 03 報表：既不是新版經營分析，也不是舊版銷量報告。")


def validate_new_sales_relationships(df):
    checks = [
        ("有效订单总量", "订单总量", "取消订单量", 0.000001),
        ("有效产品销量", "产品销量", "取消产品销量", 0.000001),
        ("有效产品销售额", "产品销售额", "取消产品销售额", 0.011),
    ]
    problems = []
    for effective, gross, cancelled, tolerance in checks:
        difference = (
            to_numeric_series(df[effective])
            - (to_numeric_series(df[gross]) - to_numeric_series(df[cancelled]))
        ).abs()
        bad_count = int((difference > tolerance).sum())
        if bad_count:
            problems.append(f"{effective} 有 {bad_count} 列不符合 {gross}－{cancelled}")
    if problems:
        raise ValueError("；".join(problems))


def transform_new_sales_report(uploaded_file, report_start, report_end):
    if report_start is None or report_end is None:
        raise ValueError("新版經營分析檔案沒有日期資訊，請先選擇報表開始日與結束日。")
    if report_start > report_end:
        raise ValueError("報表開始日不可晚於結束日。")

    source = read_excel(uploaded_file, header=0)
    normalized_to_original = {normalize_header(col): col for col in source.columns}
    missing = [
        col for col in NEW_03_REQUIRED_HEADERS
        if normalize_header(col) not in normalized_to_original
    ]
    if missing:
        raise ValueError("新版經營分析缺少必要欄位：" + "、".join(missing))

    # 重新命名成固定的新版欄位名稱，避免繁簡體或空白差異。
    canonical = pd.DataFrame(index=source.index)
    for target in NEW_03_REQUIRED_HEADERS:
        canonical[target] = source[normalized_to_original[normalize_header(target)]]

    canonical = canonical.dropna(how="all").copy()
    canonical["店铺名称"] = canonical["店铺名称"].fillna("").astype(str).str.strip()
    canonical["店铺SKU编号"] = validate_sales_skus(
        canonical["店铺SKU编号"], "新版經營分析", 2
    )

    blank_key = (canonical["店铺名称"] == "") | (canonical["店铺SKU编号"] == "")
    if blank_key.any():
        raise ValueError(f"新版經營分析有 {int(blank_key.sum())} 列缺少店鋪名稱或店鋪SKU編號。")

    numeric_columns = [
        "订单总量", "有效订单总量", "产品销量", "有效产品销量", "产品销售额",
        "有效产品销售额", "取消订单量", "取消产品销售额", "取消产品销量",
        "退款订单量", "退款产品销售额", "退款产品销量",
    ]
    for col in numeric_columns:
        canonical[col] = to_numeric_series(canonical[col])

    validate_new_sales_relationships(canonical)

    gross_quantity = canonical["产品销量"]
    average_price = canonical["产品销售额"].div(gross_quantity.where(gross_quantity != 0)).fillna(0)

    output = pd.DataFrame({
        "商品名称": canonical["产品名称"],
        "店铺": canonical["店铺名称"],
        "商品SKU": canonical["店铺SKU编号"],
        "有效商品销售额": canonical["有效产品销售额"],
        "有效订单量": canonical["有效订单总量"],
        "有效商品销量": canonical["有效产品销量"],
        "商品平均价格": average_price,
        "商品销售额": canonical["产品销售额"],
        "商品销量": canonical["产品销量"],
        "订单总量": canonical["订单总量"],
        # 新版經營分析沒有包裹總量；目前下游沒有使用此欄，保留空白以維持 A:Q。
        "包裹总量": "",
        "退款商品金额": canonical["退款产品销售额"],
        "退款订单数": canonical["退款订单量"],
        "退款商品数": canonical["退款产品销量"],
        "取消商品金額": canonical["取消产品销售额"],
        "取消訂單數": canonical["取消订单量"],
        "取消商品數": canonical["取消产品销量"],
    })[HEADERS_03]

    title = ["经营分析", f"时间:{report_start.isoformat()}-{report_end.isoformat()}"]
    title.extend([""] * (len(HEADERS_03) - len(title)))

    metadata = {
        "report_type": "新版經營分析－店鋪SKU",
        "rows": len(output),
        "stores": canonical["店铺名称"].nunique(),
        "skus": canonical["店铺SKU编号"].nunique(),
        "duplicate_store_sku_rows": int(
            canonical.duplicated(["店铺名称", "店铺SKU编号"], keep=False).sum()
        ),
        "effective_quantity": float(canonical["有效产品销量"].sum()),
        "effective_sales": float(canonical["有效产品销售额"].sum()),
    }
    return output, title, metadata


def transform_legacy_sales_report(uploaded_file):
    title = read_excel(uploaded_file, nrows=1, header=None).values.tolist()[0]
    output = safe_read_and_align_uploaded(uploaded_file, HEADERS_03, "03", header_row=1)
    if output is None:
        raise ValueError("舊版銷量報告欄位驗證失敗。")

    metadata = {
        "report_type": "舊版銷量報告",
        "rows": len(output),
        "stores": output["店铺"].nunique(),
        "skus": output["商品SKU"].nunique(),
        "duplicate_store_sku_rows": int(output.duplicated(["店铺", "商品SKU"], keep=False).sum()),
        "effective_quantity": float(to_numeric_series(output["有效商品销量"]).sum()),
        "effective_sales": float(to_numeric_series(output["有效商品销售额"]).sum()),
    }
    return output, title, metadata


def prepare_sales_report(uploaded_file, report_start, report_end):
    report_type = detect_sales_report_type(uploaded_file)
    if report_type == "business_analysis":
        return transform_new_sales_report(uploaded_file, report_start, report_end)
    return transform_legacy_sales_report(uploaded_file)


def show_sales_preflight(metadata):
    st.info(f"ℹ️ 03 資料檢查通過，準備同步（{metadata['rows']:,} 筆）。")


# ================= Google Sheets 存取 =================
def build_worksheet_map(spreadsheet):
    """一次讀取工作表清單，避免每個任務重複呼叫 Google API。"""
    return {str(sheet.id): sheet for sheet in spreadsheet.worksheets()}


def get_worksheet_by_gid(spreadsheet, gid, worksheet_map=None):
    if worksheet_map is None:
        worksheet_map = build_worksheet_map(spreadsheet)
    return worksheet_map.get(str(gid))


def quote_sheet_title(title):
    """將工作表名稱安全地放入 A1 range。"""
    return "'" + str(title).replace("'", "''") + "'"


def retry_on_api_error(func, *args, max_retries=3, delay=5, **kwargs):
    for attempt in range(1, max_retries + 1):
        try:
            return func(*args, **kwargs)
        except APIError as exc:
            error_text = str(exc)
            temporary = any(code in error_text for code in ("429", "500", "502", "503", "504"))
            if temporary and attempt < max_retries:
                st.toast(f"⏳ 雲端服務暫時忙碌，{delay} 秒後自動重試 ({attempt}/{max_retries})...")
                time.sleep(delay)
                continue
            raise


def get_gspread_client():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    if len(st.secrets) > 0:
        try:
            if "gcp_service_account" in st.secrets:
                secret_data = st.secrets["gcp_service_account"]
                credentials_dict = json.loads(secret_data) if isinstance(secret_data, str) else secret_data
                credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
                return gspread.authorize(credentials)
        except Exception as exc:
            st.error(f"❌ 嘗試解析雲端 Secrets 金鑰時發生錯誤: {exc}")

    if os.path.exists(SERVICE_ACCOUNT_FILE):
        try:
            credentials = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope)
            return gspread.authorize(credentials)
        except Exception as exc:
            st.error(f"❌ 讀取本地憑證 `{SERVICE_ACCOUNT_FILE}` 失敗。錯誤: {exc}")
            return None

    st.error("❌ 找不到 Google Sheets 服務帳戶憑證。")
    return None


def dataframe_to_google_values(df, title_list=None):
    values = []
    if title_list:
        width = len(df.columns)
        clean_title = [clean_cell_for_google(cell) for cell in title_list[:width]]
        clean_title.extend([""] * (width - len(clean_title)))
        values.append(clean_title)
    values.append([clean_cell_for_google(cell) for cell in df.columns.tolist()])
    for row in df.itertuples(index=False, name=None):
        values.append([clean_cell_for_google(cell) for cell in row])
    return values


def prepare_cloud_upload(df, task_key, title_list=None):
    """把已驗證資料轉成待批次上傳內容；此步驟不會連線或修改雲端。"""
    if df is None or df.empty:
        raise ValueError(f"Task {task_key} 沒有可上傳的資料。")
    return {
        "task_key": task_key,
        "end_col": SHEET_CONFIGS[task_key]["end_col"],
        "key_col": SHEET_CONFIGS[task_key]["key_col"],
        "values": dataframe_to_google_values(df, title_list=title_list),
    }


def last_populated_row(value_range):
    """values.batchGet 會省略尾端空白；回傳關鍵欄最後一個有值的列號。"""
    columns = value_range.get("values", [])
    return len(columns[0]) if columns else 0


def batch_upload_to_google_sheets(prepared_uploads):
    """
    先一次讀取各表關鍵欄列數，再以單一 values.batchUpdate 寫入所有資料。

    寫入失敗時 Google 不會套用該批次；寫入成功後才清除舊資料多出的尾端列，
    因此不再出現「先清空工作表、上傳途中失敗」的空白視窗。
    """
    if not prepared_uploads:
        return False, None, None

    client = get_gspread_client()
    if not client:
        return False, None, None

    spreadsheet = None
    worksheet_map = None

    try:
        spreadsheet = retry_on_api_error(client.open_by_key, SPREADSHEET_ID)
        worksheet_map = retry_on_api_error(build_worksheet_map, spreadsheet)

        resolved_uploads = []
        key_ranges = []
        for upload in prepared_uploads:
            task_key = upload["task_key"]
            worksheet = get_worksheet_by_gid(
                spreadsheet,
                SHEET_CONFIGS[task_key]["gid"],
                worksheet_map,
            )
            if not worksheet:
                raise ValueError(f"找不到 Task {task_key} 對應的 Google Sheet。")

            sheet_name = quote_sheet_title(worksheet.title)
            resolved = dict(upload)
            resolved["worksheet"] = worksheet
            resolved["sheet_name"] = sheet_name
            resolved_uploads.append(resolved)
            key_ranges.append(
                f"{sheet_name}!{upload['key_col']}:{upload['key_col']}"
            )

        # 只讀每張表的一個關鍵欄，不再下載 A:末欄的完整舊資料。
        old_rows_response = retry_on_api_error(
            spreadsheet.values_batch_get,
            key_ranges,
            params={"majorDimension": "COLUMNS"},
        )
        old_value_ranges = old_rows_response.get("valueRanges", [])
        for index, upload in enumerate(resolved_uploads):
            value_range = old_value_ranges[index] if index < len(old_value_ranges) else {}
            upload["old_last_row"] = last_populated_row(value_range)

        update_data = []
        for upload in resolved_uploads:
            new_last_row = len(upload["values"])
            upload["new_last_row"] = new_last_row
            update_data.append({
                "range": (
                    f"{upload['sheet_name']}!A1:"
                    f"{upload['end_col']}{new_last_row}"
                ),
                "majorDimension": "ROWS",
                "values": upload["values"],
            })

        retry_on_api_error(
            spreadsheet.values_batch_update,
            body={
                "valueInputOption": "RAW",
                "includeValuesInResponse": False,
                "data": update_data,
            },
        )

        # 新資料確定寫入後，才一次清除舊資料較長時留下的尾端。
        stale_ranges = []
        for upload in resolved_uploads:
            if upload["old_last_row"] > upload["new_last_row"]:
                stale_ranges.append(
                    f"{upload['sheet_name']}!A{upload['new_last_row'] + 1}:"
                    f"{upload['end_col']}{upload['old_last_row']}"
                )
        if stale_ranges:
            try:
                retry_on_api_error(
                    spreadsheet.values_batch_clear,
                    body={"ranges": stale_ranges},
                )
            except Exception as exc:
                st.error(
                    "⚠️ 新資料已寫入，但舊資料尾端清理失敗；"
                    f"請重跑一次或檢查雲端表格。錯誤: {exc}"
                )
                return False, spreadsheet, worksheet_map

        for upload in resolved_uploads:
            st.write(f"🟢 {upload['worksheet'].title} 數據已成功同步。")
        return True, spreadsheet, worksheet_map
    except Exception as exc:
        st.error(f"❌ 雲端批次同步失敗，已停止後續處理。錯誤: {exc}")
        return False, spreadsheet, worksheet_map


def post_process_steps(spreadsheet=None, worksheet_map=None):
    try:
        if spreadsheet is None:
            client = get_gspread_client()
            if not client:
                return False
            spreadsheet = retry_on_api_error(client.open_by_key, SPREADSHEET_ID)
        if worksheet_map is None:
            worksheet_map = retry_on_api_error(build_worksheet_map, spreadsheet)
        source_sheet = get_worksheet_by_gid(
            spreadsheet, SHEET_CONFIGS["03"]["gid"], worksheet_map
        )
        target_sheet = get_worksheet_by_gid(
            spreadsheet, SHEET_CONFIGS["target_main"]["gid"], worksheet_map
        )
        if not source_sheet or not target_sheet:
            st.error("❌ 找不到 03 或採購主表工作表。")
            return False

        raw_date_text = retry_on_api_error(source_sheet.acell, "B1").value
        found_dates = re.findall(r"\d{4}-\d{2}-\d{2}", str(raw_date_text))
        if len(found_dates) < 2:
            st.error("❌ 03!B1 找不到完整的開始日與結束日，已停止後處理。")
            return False

        start_text = found_dates[0].replace("-", "")[4:]
        end_text = found_dates[1].replace("-", "")[4:]
        final_date_text = f"{start_text}-{end_text}"
        retry_on_api_error(target_sheet.update, values=[[final_date_text]], range_name="B1")
        st.write(f"📅 雲端主表 B1 日期同步完成：`{final_date_text}`")

        # 等候 Google Sheets 重新計算透視表，再同步 V 欄 SKU 清單。
        excluded = {"商品SKU", "商品SKU编号", "商品SKU編號", "總和", "总和", "Grand Total"}
        filtered_skus = []
        for _ in range(5):
            # 現有 V 欄透視表前四列是標題／欄位資訊，SKU 從第 5 列開始。
            all_values = retry_on_api_error(source_sheet.col_values, 22)[4:]
            filtered_skus = [
                [str(value).strip()]
                for value in all_values
                if value and str(value).strip() not in excluded
            ]
            if filtered_skus:
                break
            time.sleep(2)

        retry_on_api_error(target_sheet.batch_clear, ["A3:A"])
        if filtered_skus:
            retry_on_api_error(target_sheet.update, values=filtered_skus, range_name="A3")
            st.write("🏷️ 雲端 A 欄品項條碼二次同步成功。")
        else:
            st.warning("⚠️ 03 的 V 欄透視表尚未產生 SKU，採購主表 A 欄未更新。")
        return True
    except Exception as exc:
        st.error(f"❌ 後處理二次同步出錯: {exc}")
        return False


# ================= Streamlit 畫面 =================
PAGE_CSS = """
<style>
.stApp { background-color: #fcfcfc !important; color: #2b303a !important; }
[data-testid="stSidebar"] { background-color: #f3f4f6 !important; border-right: 1px solid #e5e7eb; }
h1, h2, h3, h4, h5, h6, p, label, .stMarkdown { color: #2b303a !important; }
.upload-card {
    background-color: #ffffff; padding: 24px; border-radius: 12px;
    border: 1px solid #e5e7eb; box-shadow: 0 4px 6px -1px rgba(0,0,0,.05);
    margin-bottom: 20px;
}
[data-testid="stFileUploader"] {
    background-color: #fafafa !important; border: 1px dashed #cbd5e1 !important;
    border-radius: 8px !important; padding: 10px !important;
}
[data-testid="stFileUploader"] section, div[data-testid="stFileUploadDropzone"] {
    background-color: transparent !important;
}
button[data-testid*="stBaseButton"] {
    background-color: #e2e8f0 !important; color: #475569 !important;
    border: 1px solid #cbd5e1 !important; transition: all .2s ease;
}
button[data-testid*="stBaseButton"]:hover { background-color: #cbd5e1 !important; color: #334155 !important; }
.custom-main-title {
    color: #1e293b !important; font-size: 1.8rem !important; font-weight: 700 !important;
    margin-bottom: .5rem;
}
.stButton > button {
    background-color: #475569 !important; border-color: #475569 !important;
    color: #ffffff !important; font-weight: 600 !important; font-size: 16px !important;
    border-radius: 8px !important; padding: 12px 0 !important; transition: all .2s ease !important;
    box-shadow: 0 4px 6px -1px rgba(0,0,0,.05) !important;
}
.stButton > button:hover { background-color: #334155 !important; border-color: #334155 !important; color: #fff !important; }
.stButton > button p { color: #fff !important; font-weight: 600 !important; }
.custom-section-title {
    color: #2b5c8f !important; font-size: 18px !important; font-weight: 600 !important;
    margin-bottom: 12px !important; margin-top: 5px !important; display: flex; align-items: center;
}
div[data-testid="stToast"] {
    background-color: #334155 !important; border: 1px solid #1e293b !important;
    border-radius: 8px !important; box-shadow: 0 20px 25px -5px rgba(0,0,0,.1) !important;
}
div[data-testid="stToast"], div[data-testid="stToast"] div,
div[data-testid="stToast"] span, div[data-testid="stToast"] p {
    color: #fff !important; font-weight: 900 !important; font-size: 15px !important;
}
</style>
"""


def save_run_summary(status, task_results=None, elapsed_seconds=None):
    """將執行結果留在 session_state，避免一般 Streamlit rerun 後訊息消失。"""
    st.session_state["last_run_summary"] = {
        "status": status,
        "task_results": dict(task_results or {}),
        "elapsed_seconds": elapsed_seconds,
    }


def render_last_run_summary():
    summary = st.session_state.get("last_run_summary")
    if not summary:
        return

    task_results = summary.get("task_results", {})
    task_text = "｜".join(
        f"{task_name}：{result}"
        for task_name, result in task_results.items()
    )
    elapsed = summary.get("elapsed_seconds")
    elapsed_text = f"｜耗時 {elapsed:,.2f} 秒" if elapsed is not None else ""
    details = f"（{task_text}{elapsed_text}）" if task_text or elapsed_text else ""

    if summary.get("status") == "success":
        st.success(f"✨ 上次流程已完成{details}")
    elif summary.get("status") == "warning":
        st.warning(f"⚠️ 上次流程已結束，但有項目失敗{details}")
    elif summary.get("status") == "error":
        st.error(f"❌ 上次流程未執行完成{details}")
    else:
        st.warning(
            "⚠️ 上次流程沒有留下完成標記，可能在執行中發生頁面重新整理或連線中斷"
            f"{details}。請先核對雲端資料，再決定是否重跑。"
        )


def main():
    st.set_page_config(page_title="BS銷售更新", page_icon="📊", layout="wide")
    st.markdown(PAGE_CSS, unsafe_allow_html=True)

    st.markdown(
        '<div class="custom-main-title">📊 BigSeller 銷售數據更新系統</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <p style='color:#555;margin-bottom:5px;'>請上傳對應的 Excel 報表；系統會先驗證欄位，再同步至雲端。</p>
        <p style='margin-top:0;'><a href='https://docs.google.com/spreadsheets/d/1FLfAbqq1TmQnXFR3rHxGkrXELrlKSMpiXGYXk7hVZm8/edit?gid=1324377276#gid=1324377276' target='_blank' style='color:#2b5c8f;font-weight:bold;text-decoration:underline;'>📊 點此打開 Google Sheets 雲端主表</a></p>
        <p style='margin-top:5px;font-size:12px;color:#94a3b8;'>系統已啟用欄位防呆、安全批次同步與雲端暫時性錯誤重試。</p>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="upload-card">', unsafe_allow_html=True)
        st.markdown('<div class="custom-section-title">📁 01. 庫存清單</div>', unsafe_allow_html=True)
        uploaded_files_01 = st.file_uploader(
            "請上傳「库存清单*.xlsx」檔案（可多選）",
            type=["xlsx"], accept_multiple_files=True, key="u01",
        )
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="custom-section-title">📁 02. 在線產品 更新促銷價</div>', unsafe_allow_html=True)
        uploaded_files_lan = st.file_uploader(
            "① 請上傳「懶餅乾*.xlsx」檔案（可多選）",
            type=["xlsx"], accept_multiple_files=True, key="ulan",
        )
        uploaded_files_onl = st.file_uploader(
            "② 請上傳「其他店鋪」的檔案（可多選）",
            type=["xlsx"], accept_multiple_files=True, key="uonl",
        )
        st.markdown("</div>", unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="upload-card">', unsafe_allow_html=True)
        st.markdown('<div class="custom-section-title">📁 03. 銷量報告</div>', unsafe_allow_html=True)
        uploaded_file_03 = st.file_uploader(
            "請上傳新版「经营分析_店铺SKU_按店铺SKU导出*.xlsx」或舊版「销量报告*.xlsx」（單選）",
            type=["xlsx"], key="u03",
        )
        st.caption("新版請使用：店鋪SKU／按SKU聚合／下單時間／已支付／不看測評訂單／不看贈品。")
        date_col1, date_col2 = st.columns(2)
        with date_col1:
            report_start = st.date_input("03 報表開始日（新版必填）", value=None, key="u03_start")
        with date_col2:
            report_end = st.date_input("03 報表結束日（新版必填）", value=None, key="u03_end")

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="custom-section-title">📁 04. 利潤報告</div>', unsafe_allow_html=True)
        uploaded_file_04 = st.file_uploader(
            "請上傳「商品利润*.xlsx」檔案（單選）",
            type=["xlsx"], key="u04",
        )
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    run_clicked = st.button(
        "🔥 啟動自動化整合全流程", type="primary", use_container_width=True
    )
    if not run_clicked:
        render_last_run_summary()
        return

    if not (uploaded_files_01 or uploaded_files_lan or uploaded_files_onl or uploaded_file_03 or uploaded_file_04):
        st.error("🚨 流程終止：您尚未上傳任何 Excel 檔案。")
        save_run_summary("error", {"流程": "未選擇檔案"}, 0)
        return

    start_time = time.time()
    has_error = False
    task_results = {}
    prepared_uploads = []
    selected_task_keys = []
    save_run_summary("running", task_results)

    with st.spinner("系統正在驗證並同步數據，請稍候..."):
        # 第一階段：先在本機完成所有檔案的讀取與驗證，不修改雲端。
        if uploaded_files_01:
            selected_task_keys.append("01")
            st.toast("⏳ 正在處理 01.庫存清單...")
            dataframes = []
            task_input_error = False
            for file in uploaded_files_01:
                result = safe_read_and_align_uploaded(file, HEADERS_01, "01")
                if result is None:
                    has_error = True
                    task_input_error = True
                else:
                    dataframes.append(result)
            if task_input_error or not dataframes:
                has_error = True
                task_results["01"] = "檢查失敗"
            else:
                try:
                    df01 = pd.concat(dataframes, ignore_index=True)
                    prepared_uploads.append(prepare_cloud_upload(df01, "01"))
                    task_results["01"] = "檢查通過"
                except Exception as exc:
                    st.error(f"❌ 01 報表處理失敗：{exc}")
                    has_error = True
                    task_results["01"] = "檢查失敗"
            save_run_summary("running", task_results)

        if uploaded_files_lan or uploaded_files_onl:
            selected_task_keys.append("02")
            st.toast("⏳ 正在處理 02.在線產品更新...")
            lan_dataframes, other_dataframes = [], []
            task_input_error = False
            for file in uploaded_files_lan:
                result = safe_read_and_align_uploaded(file, HEADERS_02, "02")
                if result is None:
                    has_error = True
                    task_input_error = True
                else:
                    lan_dataframes.append(result)
            for file in uploaded_files_onl:
                result = safe_read_and_align_uploaded(file, HEADERS_02, "02")
                if result is None:
                    has_error = True
                    task_input_error = True
                else:
                    other_dataframes.append(result)

            df_lan = (
                pd.concat(lan_dataframes, ignore_index=True).drop_duplicates(subset=["SKU"])
                if lan_dataframes else pd.DataFrame(columns=HEADERS_02)
            )
            df_other = (
                pd.concat(other_dataframes, ignore_index=True).drop_duplicates(subset=["SKU"])
                if other_dataframes else pd.DataFrame(columns=HEADERS_02)
            )
            df02 = pd.concat([df_lan, df_other], ignore_index=True)
            if task_input_error or df02.empty:
                has_error = True
                task_results["02"] = "檢查失敗"
            else:
                try:
                    df02["priority"] = df02["店铺昵称"].apply(
                        lambda value: 0 if value == "02-懶餅乾家居" else 1
                    )
                    df02 = (
                        df02.sort_values(by=["SKU", "priority"])
                        .drop_duplicates(subset=["SKU"], keep="first")
                        .drop(columns=["priority"])
                    )
                    prepared_uploads.append(prepare_cloud_upload(df02, "02"))
                    task_results["02"] = "檢查通過"
                except Exception as exc:
                    st.error(f"❌ 02 報表處理失敗：{exc}")
                    has_error = True
                    task_results["02"] = "檢查失敗"
            save_run_summary("running", task_results)

        if uploaded_file_03:
            selected_task_keys.append("03")
            st.toast("⏳ 正在處理 03.銷量報告...")
            try:
                df03, title03, metadata03 = prepare_sales_report(
                    uploaded_file_03, report_start, report_end
                )
                show_sales_preflight(metadata03)
                prepared_uploads.append(
                    prepare_cloud_upload(df03, "03", title_list=title03)
                )
                task_results["03"] = "檢查通過"
            except Exception as exc:
                st.error(f"❌ 03 報表處理失敗：{exc}")
                has_error = True
                task_results["03"] = "檢查失敗"
            save_run_summary("running", task_results)

        if uploaded_file_04:
            selected_task_keys.append("04")
            st.toast("⏳ 正在處理 04.利潤報告...")
            try:
                title04 = read_excel(uploaded_file_04, nrows=1, header=None).values.tolist()[0]
                df04 = safe_read_and_align_uploaded(uploaded_file_04, HEADERS_04, "04", header_row=1)
                if df04 is None:
                    has_error = True
                    task_results["04"] = "檢查失敗"
                else:
                    prepared_uploads.append(
                        prepare_cloud_upload(df04, "04", title_list=title04)
                    )
                    task_results["04"] = "檢查通過"
            except Exception as exc:
                st.error(f"❌ 04 報表處理失敗：{exc}")
                has_error = True
                task_results["04"] = "檢查失敗"
            save_run_summary("running", task_results)

        # 第二階段：所有已選檔案都通過後，才共用同一連線批次寫入。
        spreadsheet = None
        worksheet_map = None
        if has_error:
            for task_key in selected_task_keys:
                if task_results.get(task_key) == "檢查通過":
                    task_results[task_key] = "未同步"
            st.error("⛔ 有檔案未通過檢查，本次所有雲端資料均未變更。")
        else:
            st.toast("☁️ 所有檔案檢查通過，正在批次同步雲端...")
            upload_ok, spreadsheet, worksheet_map = batch_upload_to_google_sheets(
                prepared_uploads
            )
            if upload_ok:
                for task_key in selected_task_keys:
                    task_results[task_key] = "成功"
            else:
                has_error = True
                for task_key in selected_task_keys:
                    task_results[task_key] = "同步失敗"
            save_run_summary("running", task_results)

            if upload_ok and "03" in selected_task_keys:
                post_process_ok = post_process_steps(spreadsheet, worksheet_map)
                task_results["後處理"] = "成功" if post_process_ok else "失敗"
                if not post_process_ok:
                    has_error = True
                save_run_summary("running", task_results)

    elapsed_seconds = round(time.time() - start_time, 2)
    save_run_summary(
        "warning" if has_error else "success",
        task_results,
        elapsed_seconds,
    )
    render_last_run_summary()


if __name__ == "__main__":
    main()
