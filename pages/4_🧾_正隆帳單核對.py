# -*- coding: utf-8 -*-
import io
import streamlit as st
import pandas as pd
from recon_engine import process_pdf   # 引擎在根目錄，直接 import

st.set_page_config(page_title="正隆帳單核對", page_icon="💲")
st.title("💲 正隆紙箱帳單核對")
st.caption("對帳單 ↔ 電子發票證明聯（本地 OCR，資料不外傳）")

col1, col2 = st.columns(2)
dpi   = col1.select_slider("OCR 解析度 DPI", options=[200, 220, 250, 300], value=250,
                           help="密集對帳單建議 250–300；越高越準但越慢")
debug = col2.checkbox("除錯模式（顯示逐頁轉正/分類/抓到的發票號碼）")

up = st.file_uploader("上傳掃描 PDF", type=["pdf"])

if up and st.button("開始核對", type="primary"):
    with st.spinner("OCR 與對帳進行中…"):
        df_stmt, df_inv, df_recon, log = process_pdf(
            pdf_bytes=up.getvalue(), dpi=dpi, debug=debug)

    vc = df_recon["狀態"].value_counts()
    c1, c2, c3 = st.columns(3)
    c1.metric("✓ 一致",      int(vc.get("✓ 一致", 0)))
    c2.metric("⚠ 缺值待人工", int(vc.get("⚠ 缺值待人工", 0)))
    c3.metric("❌ 差異待人工", int(vc.get("❌ 差異待人工", 0)))

    def hl(row):
        s = row["狀態"]
        c = "#E2EFDA" if s.startswith("✓") else "#FFF2CC" if "缺值" in s else "#FCE4D6"
        return [f"background-color:{c}"] * len(row)

    st.markdown("**對帳結果**")
    st.dataframe(df_recon.style.apply(hl, axis=1), use_container_width=True, hide_index=True)

    with st.expander("電子發票明細"):
        st.dataframe(df_inv, use_container_width=True, hide_index=True)
    with st.expander("對帳單明細"):
        st.dataframe(df_stmt, use_container_width=True, hide_index=True)
    if debug and log:
        with st.expander("除錯記錄（逐頁）"):
            st.code("\n".join(log))

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_recon.to_excel(w, sheet_name="對帳結果", index=False)
        df_inv.to_excel(w, sheet_name="電子發票明細", index=False)
        df_stmt.to_excel(w, sheet_name="對帳單明細", index=False)
    st.download_button("下載對帳結果 (xlsx)", buf.getvalue(),
                       file_name="正隆對帳結果.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
