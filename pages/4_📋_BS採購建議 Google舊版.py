import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="BS採購建議整合", page_icon="📦", layout="centered")

st.title("📋 BS採購建議整合")
st.caption("上傳原始採購建議檔案，自動整合輸出")

KEEP_COLS = [
    "SKU编号", "名称", "仓库", "现有库存", "可用库存",
    "已锁", "在途中", "3天总销量", "7天总销量", "15天总销量", "30天总销量",
]
SUM_COLS = {
    "總現有庫存": "现有库存",
    "已鎖":       "已锁",
    "在途中(總)": "在途中",
    "7天總銷量":  "7天总销量",
    "15天總銷量": "15天总销量",
    "30天總銷量": "30天总销量",
}
NUM_COLS = [
    "现有库存", "可用库存", "已锁", "在途中",
    "3天总销量", "7天总销量", "15天总销量", "30天总销量",
]

st.markdown("## 第一步驟｜採購建議庫存整合")
st.caption("來源：采购建议*.xlsx")

uploaded_step1 = st.file_uploader("請上傳採購建議檔案", type=["xlsx"], key="step1")

if uploaded_step1:
    with st.spinner("處理中..."):
        df_raw = pd.read_excel(uploaded_step1, dtype=str)
        missing = [h for h in KEEP_COLS if h not in df_raw.columns]
        if missing:
            st.error("找不到以下欄位：\n" + "\n".join(f"- {m}" for m in missing))
            st.info(f"檔案中現有欄位：{list(df_raw.columns)}")
            st.stop()
        df = df_raw[KEEP_COLS].copy()
        for c in NUM_COLS:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        for new_col, src_col in SUM_COLS.items():
            df[new_col] = df.groupby("SKU编号")[src_col].transform("sum")
        df = df.drop_duplicates(subset=["SKU编号"], keep="first")
        for c in NUM_COLS + list(SUM_COLS.keys()):
            df[c] = df[c].astype(int)

    st.success(f"完成！共 {len(df)} 筆 SKU")
    c1, c2, c3 = st.columns(3)
    c1.metric("總SKU數", len(df))
    c2.metric("總現有庫存", f"{df['總現有庫存'].sum():,}")
    c3.metric("30天總銷量", f"{df['30天總銷量'].sum():,}")
    st.dataframe(df, use_container_width=True, height=300)
    out1 = io.BytesIO()
    df.to_excel(out1, index=False, engine="openpyxl")
    out1.seek(0)
    st.download_button(
        label="下載第一步驟結果 xlsx",
        data=out1,
        file_name=uploaded_step1.name.replace(".xlsx", "_整合結果.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown("---")

st.markdown("## 第二步驟｜採購數量加總")
st.caption("來源：PurchaseOrder*.xlsx")

uploaded_step2 = st.file_uploader("請上傳採購單檔案", type=["xlsx"], key="step2")

if uploaded_step2:
    with st.spinner("處理中..."):
        df_po = pd.read_excel(uploaded_step2, dtype=str)
        missing2 = [h for h in ["SKU编号", "采购数量"] if h not in df_po.columns]
        if missing2:
            st.error("找不到以下欄位：\n" + "\n".join(f"- {m}" for m in missing2))
            st.info(f"檔案中現有欄位：{list(df_po.columns)}")
            st.stop()
        df2 = df_po[["SKU编号", "采购数量"]].copy()
        df2["采购数量"] = pd.to_numeric(df2["采购数量"], errors="coerce").fillna(0)
        df2 = df2.groupby("SKU编号", as_index=False)["采购数量"].sum()
        df2["采购数量"] = df2["采购数量"].astype(int)

    st.success(f"完成！共 {len(df2)} 筆 SKU")
    c1, c2 = st.columns(2)
    c1.metric("總SKU數", len(df2))
    c2.metric("總採購數量", f"{df2['采购数量'].sum():,}")
    st.dataframe(df2, use_container_width=True, height=300)
    out2 = io.BytesIO()
    df2.to_excel(out2, index=False, engine="openpyxl")
    out2.seek(0)
    st.download_button(
        label="下載第二步驟結果 xlsx",
        data=out2,
        file_name=uploaded_step2.name.replace(".xlsx", "_採購加總.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
