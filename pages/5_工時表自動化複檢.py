import streamlit as st

# 隱藏原生多頁面選單，消滅那個 app 與工時表自動化複檢的區塊
st.markdown("""
    <style>
        [data-testid="stSidebarNav"] {display: none !important;}
    </style>
""", unsafe_allow_html=True)

st.info("工時表工具已停用，請點選左側或重新整理網頁。")
