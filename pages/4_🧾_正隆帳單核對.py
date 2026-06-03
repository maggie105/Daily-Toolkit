import streamlit as st
import pandas as pd
import os
import re
import json
import gspread
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="正隆帳單核對", page_icon="🧾", layout="wide")
st.markdown("<h2 style='color: #2b5c8f; font-weight: 700;'>🧾 正隆帳單自動化核對系統</h2>", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["📱 第一階段：整合 LINE 叫貨紀錄", "📄 第二階段：正隆帳單 PDF 核對"])

def get_gspread_client():
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    if len(st.secrets) > 0 and "gcp_service_account" in st.secrets:
        return gspread.authorize(ServiceAccountCredentials.from_json_keyfile_dict(json.loads(st.secrets["gcp_service_account"]) if isinstance(st.secrets["gcp_service_account"], str) else st.secrets["gcp_service_account"], scope))
    if os.path.exists('service_account.json'): return gspread.authorize(ServiceAccountCredentials.from_json_keyfile_name('service_account.json', scope))
    return None

with tab1:
    col_d1, col_d2 = st.columns(2)
    with col_d1: start_date = st.date_input("開始日期", value=pd.Timestamp.now().floor('D') - pd.Timedelta(days=30))
    with col_d2: end_date = st.date_input("結束日期", value=pd.Timestamp.now().floor('D'))
    line_file = st.file_uploader("📁 請上傳 LINE 對話紀錄檔案 (.txt)", type=["txt"])

    if line_file is not None:
        raw_chat_data = ""; file_bytes = line_file.read()
        for enc in ['utf-8', 'utf-16', 'cp950']:
            try: raw_chat_data = file_bytes.decode(enc); break
            except: continue
        if not raw_chat_data: st.error("🚨 無法解析該文字檔編碼。")
        else:
            current_year = start_date.year
            start_dt = datetime(start_date.year, start_date.month, start_date.day)
            end_dt = datetime(end_date.year, end_date.month, end_date.day)
            report_start = raw_chat_data.find("07:38 善敏 01/05(一)排程報告：")
            data_to_process = raw_chat_data[report_start:] if report_start != -1 else raw_chat_data
            order_markers = list(re.finditer(r"您好\s*[，,]\s*與您訂購紙箱", data_to_process))
            all_dates = list(re.finditer(r"(?<![:\d\-])(?:\d{4}/)?(\d{1,2})/(\d{1,2})(?!\d)", data_to_process))
            all_timestamps = list(re.finditer(r"(?m)^\d{1,2}:\d{2}\s", data_to_process))

            results = []
            for o_idx, o_match in enumerate(order_markers):
                order_pos = o_match.start(); nearest_date = None
                for d_match in all_dates:
                    if d_match.start() < order_pos: nearest_date = d_match
                    else: break
                if not nearest_date: continue
                try: msg_dt = datetime(current_year, int(nearest_date.group(1)), int(nearest_date.group(2)))
                except: continue
                if not (start_dt <= msg_dt <= end_dt): continue
                block_end = min([t.start() for t in all_timestamps if t.start() > order_pos] + [order_markers[o_idx+1].start() if o_idx + 1 < len(order_markers) else len(data_to_process)])
                order_block = data_to_process[o_match.start():block_end]
                name = (re.search(r"收件人(?:姓名)?[:：]\s*(.*)", order_block).group(1).strip() if re.search(r"收件人(?:姓名)?[:：]\s*(.*)", order_block) else "未填")
                addr = (re.search(r"(?:收件人地址|收件地址|收貨地址|地址)[:：]\s*(.*)", order_block).group(1).strip() if re.search(r"(?:收件人地址|收件地址|收貨地址|地址)[:：]\s*(.*)", order_block) else "未填")
                all_sizes = list(re.finditer(r"(\d+(?:\.\d+)?\s*\*\s*\d+(?:\.\d+)?\s*\*\s*\d+(?:\.\d+)?)", order_block))
                for s_match in all_sizes:
                    qty_m = re.search(r"(\d+)\s*個", order_block[s_match.end():])
                    if qty_m: results.append({"訂購尺寸": s_match.group(1).replace(" ", ""), "叫貨日期": msg_dt.strftime("%Y/%m/%d"), "數量": int(qty_m.group(1)), "未稅單價": "", "未稅總價": "", "含稅總價": "", "收件人姓名": name, "收件人地址": addr})

            if results:
                df_results = pd.DataFrame(sorted(results, key=lambda x: (x["叫貨日期"], x["訂購尺寸"])))
                st.dataframe(df_results, use_container_width=True)
                if st.button("☁️ 確認無誤，一鍵回填至雲端 Google Sheets", type="primary", use_container_width=True):
                    try:
                        client = get_gspread_client()
                        if client:
                            worksheet = client.open_by_key("1hklqBQ_9Z0HZcgHF-3Kl1jdXqDa13NXOqwHZ1BZnxQg").worksheet("叫貨紀錄")
                            if not worksheet.row_values(1): worksheet.insert_row(["訂購尺寸", "叫貨日期", "數量", "未稅單價", "未稅總價", "含稅總價", "收件人姓名", "收件人地址"], 1)
                            worksheet.append_rows([[d.get(h, "") for h in ["訂購尺寸", "叫貨日期", "數量", "未稅單價", "未稅總價", "含稅總價", "收件人姓名", "收件人地址"]] for d in results], value_input_option='USER_ENTERED')
                            st.success("✨ 雲端同步完成！")
                    except Exception as e: st.error(f"❌ 上傳失敗：{e}")
            else: st.warning("⚠️ 查無訂單數據。")

with tab2:
    pdf_file = st.file_uploader("📄 請上傳正隆帳單混合 PDF 原始檔案 (.pdf)", type=["pdf"])
    if pdf_file is not None and st.button("🚀 啟動輕量引擎進行 49頁全量影像對帳 (完全免費)", type="primary", use_container_width=True):
        with st.spinner("影像提取數字中..."):
            try:
                import pytesseract, gc
                from pdf2image import convert_from_bytes
                def extract_amounts(text):
                    out = []
                    for m in re.finditer(r"(?<!\d)\d{1,3}\.\d{3}(?!\d)|(?<!\d)\d{1,3}(?:,\d{3})+(?!\d)|(?<![\d.,])\d{2,5}(?![\d.])", text):
                        s = m.group(); v = int(s.replace(".", "")) if re.match(r"^\d{1,3}\.\d{3}$", s) else int(s.replace(",", "")) if "," in s else int(s)
                        if 50 <= v < 1e7: out.append((m.start(), v))
                    return out
                def find_triple(amounts):
                    for start in range(len(amounts)):
                        win = amounts[start:start + 6]; m = len(win)
                        for i in range(m):
                            for j in range(m):
                                if i == j: continue
                                for k in range(m):
                                    if k == i or k == j: continue
                                    a, b, c = win[i][1], win[j][1], win[k][1]
                                    if abs(a + b - c) <= 2 and a > b > 0 and b <= a * 0.06 + 1: return (a, b, c)
                    return None

                images = convert_from_bytes(pdf_file.getvalue(), dpi=130)
                stmt_res, inv_res = [], []
                for idx, img in enumerate(images):
                    img.thumbnail((1300, 1300))
                    text = pytesseract.image_to_string(img, lang='chi_tra+eng')
                    if not text.strip(): del img; gc.collect(); continue
                    lines = [l.strip() for l in text.split('\n') if l.strip()]
                    flat = text.replace(" ", "").replace("\n", "")
                    for l_idx, line in enumerate(lines):
                        if re.search(r"([A-Z]{2}\d{8})", line):
                            inv_no = re.search(r"([A-Z]{2}\d{8})", line).group(1)
                            size_m = re.search(r"([\u4e00-\u9fff\w\(\)]+.*?\d+\*\d+\*\d+.*?)(?:\s|$)", line)
                            size_val = size_m.group(1).strip() if size_m else "正隆常規紙箱"
                            triple = find_triple(extract_amounts(line))
                            if not triple and l_idx + 1 < len(lines): triple = find_triple(extract_amounts(line + " " + lines[l_idx+1]))
                            if triple: stmt_res.append({"發票號碼": inv_no, "品名規格": size_val, "數量": "明細", "未稅金額": float(triple[0]), "稅額": float(triple[1]), "總金額": float(triple[2])})
                    if any(k in flat for k in ["電子發票", "發票證明聯", "發票"]):
                        inv_no = re.search(r"([A-Z]{2}\d{8})", text).group(1) if re.search(r"([A-Z]{2}\d{8})", text) else "未知發票"
                        triple = find_triple(extract_amounts(text))
                        if triple: inv_res.append({"發票號碼": inv_no, "品名規格": "電子發票明細", "數量": "1", "單價": float(triple[0]), "營業稅": float(triple[1]), "總計": float(triple[2])})
                    del img; gc.collect()

                df_stmt = pd.DataFrame(stmt_res).drop_duplicates(subset=['發票號碼', '總金額']) if stmt_res else pd.DataFrame()
                df_inv = pd.DataFrame(inv_res).drop_duplicates(subset=['發票號碼', '總計']) if inv_res else pd.DataFrame()
                
                c1, c2 = st.columns(2)
                with c1: st.dataframe(df_stmt, use_container_width=True)
                with c2: st.dataframe(df_inv, use_container_width=True)
                if not df_stmt.empty and not df_inv.empty:
                    df_recon = pd.merge(df_stmt[['發票號碼', '總金額']], df_inv[['發票號碼', '總計']], on='發票號碼', how='outer').fillna(0.0)
                    df_recon.columns = ['發票號碼', '對帳單總金額 (A)', '發票證明聯總計 (B)']
                    df_recon['金額差異 (B-A)'] = df_recon['發票證明聯總計 (B)'] - df_recon['對帳單總金額 (A)']
                    df_recon['核對結果狀態'] = df_recon['金額差異 (B-A)'].apply(lambda d: "✓ 金額完全一致" if abs(d) <= 1 else f"❌ 異常！金額不符 (差額 {d:g})")
                    st.dataframe(df_recon, use_container_width=True)
            except Exception as e: st.error(f"❌ 錯誤: {e}")
