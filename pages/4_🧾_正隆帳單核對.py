import streamlit as st
import pandas as pd
import os
import re
import io
import json
import gspread
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

# ================= 網頁頁面初始設定 =================
st.set_page_config(page_title="正隆帳單核對", page_icon="🧾", layout="wide")

# 🛠️ 終極北歐風裝潢加強版 (超強制 CSS 注入)
st.markdown("""
    <style>
    .stApp { background-color: #fcfcfc !important; color: #2b303a !important; }
    [data-testid="stSidebar"] { background-color: #f3f4f6 !important; border-right: 1px solid #e5e7eb; }
    h1, h2, h3, h4, h5, h6, p, label, .stMarkdown { color: #2b303a !important; }
    .upload-card {
        background-color: #ffffff; padding: 24px; border-radius: 12px; 
        border: 1px solid #e5e7eb; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); margin-bottom: 20px;
    }
    [data-testid="stFileUploader"] {
        background-color: #fafafa !important; border: 1px dashed #cbd5e1 !important; border-radius: 8px !important; padding: 10px !important;
    }
    [data-testid="stFileUploader"] section { background-color: transparent !important; }
    div[data-testid="stFileUploadDropzone"] { background-color: #fafafa !important; }
    
    button[data-testid*="stBaseButton"] {
        background-color: #e2e8f0 !important; color: #475569 !important; border: 1px solid #cbd5e1 !important; transition: all 0.2s ease;
    }
    button[data-testid*="stBaseButton"]:hover { background-color: #cbd5e1 !important; color: #334155 !important; }
    
    .custom-main-title { color: #1e293b !important; font-size: 1.8rem !important; font-weight: 700 !important; margin-bottom: 0.5rem; }
    
    .stButton > button {
        background-color: #475569 !important; border-color: #475569 !important; color: #ffffff !important;
        font-weight: 600 !important; font-size: 16px !important; border-radius: 8px !important; padding: 12px 0 !important;
        transition: all 0.2s ease !important; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05) !important;
    }
    .stButton > button:hover { background-color: #334155 !important; border-color: #334155 !important; }
    .stButton > button p { color: #ffffff !important; font-weight: 600 !important; }
    .custom-section-title { color: #2b5c8f !important; font-size: 18px !important; font-weight: 600 !important; margin-bottom: 12px !important; display: flex; align-items: center; }
    </style>
""", unsafe_allow_html=True)

def get_gspread_client():
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    if len(st.secrets) > 0 and "gcp_service_account" in st.secrets:
        return gspread.authorize(ServiceAccountCredentials.from_json_keyfile_dict(json.loads(st.secrets["gcp_service_account"]) if isinstance(st.secrets["gcp_service_account"], str) else st.secrets["gcp_service_account"], scope))
    if os.path.exists('service_account.json'): 
        return gspread.authorize(ServiceAccountCredentials.from_json_keyfile_name('service_account.json', scope))
    return None

st.markdown('<div class="custom-main-title">🧾 正隆帳單自動化核對系統</div>', unsafe_allow_html=True)
st.markdown("""
    <p style='color: #555; margin-bottom: 5px;'>本系統分為兩個階段：【第一階段】整合 LINE 對話叫貨紀錄，以及【第二階段】核對正隆帳單 PDF。請依序操作。</p>
    <p style='margin-top: 0;'>👉 <a href='https://docs.google.com/spreadsheets/d/1hklqBQ_9Z0HZcgHF-3Kl1jdXqDa13NXOqwHZ1BZnxQg/edit?gid=1929625795#gid=1929625795' target='_blank' style='color: #2b5c8f; font-weight: bold; text-decoration: underline;'>🧾 點此打開 Google Sheets 雲端紙箱叫貨紀錄表</a></p>
""", unsafe_allow_html=True)
st.markdown("---")

tab1, tab2 = st.tabs(["📱 第一階段：整合 LINE 叫貨紀錄", "📄 第二階段：正隆帳單 PDF 核對"])

# ------------------ 📱 TAB 1: LINE 紀錄整合（原樣保留，未更動）------------------
with tab1:
    st.markdown('<div class="upload-card">', unsafe_allow_html=True)
    st.markdown('<div class="custom-section-title">📅 步驟 1：選擇篩選日期區間</div>', unsafe_allow_html=True)
    col_d1, col_d2 = st.columns(2)
    with col_d1: start_date = st.date_input("開始日期", value=pd.Timestamp.now().floor('D') - pd.Timedelta(days=30), key="line_start")
    with col_d2: end_date = st.date_input("結束日期", value=pd.Timestamp.now().floor('D'), key="line_end")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="upload-card">', unsafe_allow_html=True)
    st.markdown('<div class="custom-section-title">📁 步驟 2：上傳 LINE 對話紀錄檔案 (.txt)</div>', unsafe_allow_html=True)
    line_file = st.file_uploader("請上傳 `[LINE]桃園正隆X沐樂.txt` 檔案", type=["txt"], key="uline_chat", label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

    if line_file is not None:
        raw_chat_data = ""
        file_bytes = line_file.read()
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
                    if qty_m: 
                        results.append({"訂購尺寸": s_match.group(1).replace(" ", ""), "叫貨日期": msg_dt.strftime("%Y/%m/%d"), "數量": int(qty_m.group(1)), "未稅單價": "", "未稅總價": "", "含稅總價": "", "收件人姓名": name, "收件人地址": addr})

            if results:
                df_results = pd.DataFrame(sorted(results, key=lambda x: (x["叫貨日期"], x["訂購尺寸"])))
                st.markdown("### 📊 偵測並解析成功之訂單預覽")
                st.dataframe(df_results, use_container_width=True)
                if st.button("☁️ 確認無誤，一鍵回填至雲端 Google Sheets", type="primary", use_container_width=True):
                    try:
                        client = get_gspread_client()
                        if client:
                            worksheet = client.open_by_key("1hklqBQ_9Z0HZcgHF-3Kl1jdXqDa13NXOqwHZ1BZnxQg").worksheet("叫貨紀錄")
                            if not worksheet.row_values(1): worksheet.insert_row(["訂購尺寸", "叫貨日期", "數量", "未稅單價", "未稅總價", "含稅總價", "收件人姓名", "收件人地址"], 1)
                            worksheet.append_rows([[d.get(h, "") for h in ["訂購尺寸", "叫貨日期", "數量", "未稅單價", "未稅總價", "含稅總價", "收件人姓名", "收件人地址"]] for d in results], value_input_option='USER_ENTERED')
                            st.success("✨ 雲端同步完成！")
                        else:
                            st.error("❌ 找不到雲端憑證（service_account.json 或 st.secrets）。")
                    except Exception as e: st.error(f"❌ 上傳失敗：{e}")
            else: st.warning("⚠️ 查無訂單數據。")

# ------------------ 📄 TAB 2: PDF 帳單核對（改用 PaddleOCR 引擎）------------------
with tab2:
    st.markdown('<div class="upload-card">', unsafe_allow_html=True)
    st.markdown('<div class="custom-section-title">📄 步驟 1：請上傳正隆帳單混合 PDF 原始檔案 (.pdf)</div>', unsafe_allow_html=True)
    pdf_file = st.file_uploader("將正隆帳單 PDF 檔案拖放到此處", type=["pdf"], key="updf_billing_main", label_visibility="collapsed")
    col_o1, col_o2 = st.columns(2)
    with col_o1:
        dpi = st.select_slider("OCR 解析度 DPI", options=[150, 200, 250, 300], value=200, key="pdf_dpi",
                               help="PaddleOCR 對解析度不敏感，200 通常已足夠且較快")
    with col_o2:
        debug = st.checkbox("🔧 除錯模式（顯示逐頁判斷）", value=False, key="pdf_debug")
    st.markdown("</div>", unsafe_allow_html=True)
    
    if pdf_file is not None and st.button("🚀 啟動 PaddleOCR 對帳引擎", type="primary", use_container_width=True):
        with st.spinner("PaddleOCR 辨識與對帳進行中…（首次會載入模型，請稍候）"):
            try:
                from recon_engine import process_pdf   # 根目錄的 PaddleOCR 引擎
                df_stmt, df_inv, df_recon, log = process_pdf(
                    pdf_bytes=pdf_file.getvalue(), dpi=dpi, debug=debug)

                # ---- 摘要指標 ----
                vc = df_recon["狀態"].value_counts()
                n_ok = int(vc.get("✓ 一致", 0)); n_total = len(df_recon)
                m1, m2, m3 = st.columns(3)
                m1.metric("✓ 一致",       n_ok)
                m2.metric("⚠ 缺值待人工",  int(vc.get("⚠ 缺值待人工", 0)))
                m3.metric("❌ 差異待人工",  int(vc.get("❌ 差異待人工", 0)))
                if n_total:
                    st.progress(n_ok / n_total, text=f"自動對上 {n_ok}/{n_total}（{n_ok/n_total:.0%}）")

                # ---- 明細兩欄 ----
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("#### 📊 1. 對帳單明細")
                    st.dataframe(df_stmt, use_container_width=True, hide_index=True)
                with c2:
                    st.markdown("#### 📄 2. 電子發票證明聯明細")
                    st.dataframe(df_inv, use_container_width=True, hide_index=True)

                # ---- 交叉對帳（上色）----
                st.markdown("---")
                st.markdown("<h3 style='color: #2b5c8f;'>⚖️ 兩條線自動化交叉對帳結果</h3>", unsafe_allow_html=True)

                def _hl(row):
                    s = row["狀態"]
                    color = ("#E2EFDA" if s.startswith("✓") else
                             "#FFF2CC" if "缺值" in s else "#FCE4D6")
                    return [f"background-color:{color}"] * len(row)
                st.dataframe(df_recon.style.apply(_hl, axis=1), use_container_width=True, hide_index=True)

                need = df_recon[df_recon["狀態"] != "✓ 一致"]
                if len(need) == 0 and n_total > 0:
                    st.balloons()
                    st.success(f"🎉【對帳完美大通關】{n_total} 筆全量交叉撞庫 100% 精準吻合！")
                elif len(need):
                    st.warning(f"以下 {len(need)} 筆無法自動確認，請人工複檢：\n\n"
                               + "、".join(need["發票號碼"].astype(str).tolist()))

                if debug and log:
                    with st.expander("🔧 除錯記錄（逐頁）"):
                        st.code("\n".join(log))

                # ---- 下載 xlsx ----
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as w:
                    df_recon.to_excel(w, sheet_name="對帳結果", index=False)
                    df_inv.to_excel(w, sheet_name="電子發票明細", index=False)
                    df_stmt.to_excel(w, sheet_name="對帳單明細", index=False)
                st.download_button("⬇️ 下載對帳結果 (xlsx)", buf.getvalue(),
                                   file_name="正隆對帳結果.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
            except Exception as e:
                st.error(f"❌ 深度解析時發生非預期錯誤: {e}")
