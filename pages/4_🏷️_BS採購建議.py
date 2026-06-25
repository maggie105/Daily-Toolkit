import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="BS採購建議整合", page_icon="📦", layout="centered")

st.title("📦 BS採購建議整合")
st.caption("第一步驟｜上傳原始採購建議檔案，自動整合輸出")

# ── 設定 ──────────────────────────────────────────────────────
KEEP_COLS = [
    "SKU编号",
    "名称",
    "仓库",
    "现有库存",
    "可用库存",
    "已锁",
    "在途中",
    "3天总销量",
    "7天总销量",
    "15天总销量",
    "30天总销量",
]

SUM_COLS = {
    "總現有庫存": "现有库存",
    "已鎖":       "已锁",
    "在途中(總)": "在途中",
    "7天總銷量":  "7天总销量",
    "15天總銷量": "15天总销量",
    "30天總銷量": "30天总销量",
}

NUM_COLS = ["现有库存", "可用库存", "已锁", "在途中",
            "3天总销量", "7天总销量", "15天总销量", "30天总销量"]

# ── 上傳區 ────────────────────────────────────────────────────
uploaded = st.file_uploader("請上傳原始採購建議 xlsx 檔案", type=["xlsx"])

if uploaded:
    with st.spinner("讀取中..."):
        df_raw = pd.read_excel(uploaded, dtype=str)

    # 步驟 a：檢查表頭
    missing = [h for h in KEEP_COLS if h not in df_raw.columns]
    if missing:
        st.error(f"❌ 找不到以下欄位，請確認來源檔案：\n\n" + "\n".join(f"- {m}" for m in missing))
        st.info(f"檔案中現有欄位：{list(df_raw.columns)}")
        st.stop()

    with st.spinner("整合中..."):
        # 步驟 a：只保留需要的欄，依指定順序
        df = df_raw[KEEP_COLS].copy()

        # 步驟 b+c：數字欄轉換，並計算加總欄
        for c in NUM_COLS:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

        for new_col, src_col in SUM_COLS.items():
            df[new_col] = df.groupby("SKU编号")[src_col].transform("sum")

        # 步驟 d：移除重複（依 SKU编号）
        df = df.drop_duplicates(subset=["SKU编号"], keep="first")

        # 數字欄轉為整數顯示
        int_cols = NUM_COLS + list(SUM_COLS.keys())
        for c in int_cols:
            df[c] = df[c].astype(int)

    # ── 結果顯示 ─────────────────────────────────────────────
    st.success(f"✅ 整合完成！共 {len(df)} 筆 SKU")

    col1, col2, col3 = st.columns(3)
    col1.metric("總SKU數", len(df))
    col2.metric("總現有庫存", f"{df['總現有庫存'].sum():,}")
    col3.metric("30天總銷量", f"{df['30天總銷量'].sum():,}")

    st.dataframe(df, use_container_width=True, height=400)

    # ── 下載 ─────────────────────────────────────────────────
    output = io.BytesIO()
    df.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)

    original_name = uploaded.name.replace(".xlsx", "")
    st.download_button(
        label="⬇️ 下載整合結果 xlsx",
        data=output,
        file_name=f"{original_name}_整合結果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
