# -*- coding: utf-8 -*-
# Streamlit「第二階段：正隆帳單核對」分頁接線片段
# 把 recon_engine.py 放在同目錄；本片段示範如何呼叫 process_pdf 並呈現結果。
import io
import streamlit as st
import pandas as pd
from recon_engine import process_pdf, reconcile  # 純函式，與 UI 解耦

def render_tab2():
    st.subheader("正隆紙箱帳單核對（對帳單 ↔ 電子發票證明聯）")
    col1, col2 = st.columns(2)
    dpi   = col1.select_slider("OCR 解析度 DPI", options=[200, 220, 250, 300], value=250,
                               help="密集對帳單建議 250–300；越高越準但越慢")
    debug = col2.checkbox("除錯模式（顯示逐頁轉正/分類/抓到的發票號碼）", value=False)

    up = st.file_uploader("上傳掃描 PDF（單檔含對帳單與電子發票證明聯）", type=["pdf"])
    if not up:
        return

    if st.button("開始核對", type="primary"):
        prog = st.progress(0.0, text="準備中…")
        # process_pdf 一次處理整份；若要逐頁進度，可改用下方「逐頁版本」
        with st.spinner("OCR 與對帳進行中（本地處理，資料不外傳）…"):
            df_stmt, df_inv, df_recon, log = process_pdf(
                pdf_bytes=up.getvalue(), dpi=dpi, debug=debug)
        prog.progress(1.0, text="完成")

        # ---- 摘要 ----
        vc = df_recon["狀態"].value_counts()
        c1, c2, c3 = st.columns(3)
        c1.metric("✓ 一致",       int(vc.get("✓ 一致", 0)))
        c2.metric("⚠ 缺值待人工",  int(vc.get("⚠ 缺值待人工", 0)))
        c3.metric("❌ 差異待人工",  int(vc.get("❌ 差異待人工", 0)))

        # ---- 對帳結果（差異/缺值上色）----
        def hl(row):
            s = row["狀態"]
            color = ("#E2EFDA" if s.startswith("✓") else
                     "#FFF2CC" if "缺值" in s else "#FCE4D6")
            return [f"background-color:{color}"] * len(row)
        st.markdown("**對帳結果**")
        st.dataframe(df_recon.style.apply(hl, axis=1), use_container_width=True, hide_index=True)

        with st.expander("電子發票明細"):
            st.dataframe(df_inv, use_container_width=True, hide_index=True)
        with st.expander("對帳單明細"):
            st.dataframe(df_stmt, use_container_width=True, hide_index=True)
        if debug and log:
            with st.expander("除錯記錄（逐頁）"):
                st.code("\n".join(log))

        # ---- 下載 xlsx ----
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df_recon.to_excel(w, sheet_name="對帳結果", index=False)
            df_inv.to_excel(w, sheet_name="電子發票明細", index=False)
            df_stmt.to_excel(w, sheet_name="對帳單明細", index=False)
        st.download_button("下載對帳結果 (xlsx)", buf.getvalue(),
                           file_name="正隆對帳結果.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---- 逐頁進度版本（若想要更細的進度條，可參考此寫法）----
# from pdf2image import convert_from_bytes
# from recon_engine import _orient_ocr, parse_invoice, parse_statement, reconcile
# def render_tab2_with_perpage_progress(up, dpi=250, debug=False):
#     pages = convert_from_bytes(up.getvalue(), dpi=dpi)
#     prog = st.progress(0.0); inv, stmt = [], []
#     for i, img in enumerate(pages):
#         text, rot = _orient_ocr(img)
#         flat = text.replace(" ", "")
#         if ("電子發票" in flat) or ("證明聯" in flat) or ("銷售額" in flat):
#             inv.append(parse_invoice(i + 1, text))
#         else:
#             stmt += parse_statement(i + 1, text)
#         prog.progress((i + 1) / len(pages), text=f"處理第 {i+1}/{len(pages)} 頁")
#     df_inv, df_stmt = pd.DataFrame(inv), pd.DataFrame(stmt)
#     df_recon = reconcile(df_stmt, df_inv)
#     ...
