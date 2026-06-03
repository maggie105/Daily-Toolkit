import streamlit as st

# ================= 網頁頁面初始設定 =================
st.set_page_config(page_title="Daily Toolkit - 數據處理中心", page_icon="🔧", layout="wide")

# 🔥 終極北歐簡約風裝潢 (超強制 CSS 注入)
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
    /* 儀表板卡片樣式 */
    .dashboard-card {
        background-color: #ffffff;
        padding: 24px; 
        border-radius: 12px; 
        border: 1px solid #e5e7eb;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        margin-bottom: 20px;
        transition: all 0.2s ease;
    }
    .dashboard-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05);
    }
    </style>
""", unsafe_allow_html=True)

# ==================== 主頁面視覺 (極簡儀表板) ====================
st.markdown("<h1 style='color: #2b5c8f; font-weight: 700;'>🔧 Daily Toolkit 數據處理中心</h1>", unsafe_allow_html=True)
st.markdown("<p style='color: #666666; font-size: 1.1rem;'>歡迎回來！請點選左側導覽列，或透過下方快速跳轉入口來啟動您所需的自動化工具。</p>", unsafe_allow_html=True)
st.markdown("---")

# 建立 2x2 的儀表板佈局
col1, col2 = st.columns(2, gap="large")

with col1:
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.markdown("<h3 style='color: #2b5c8f; margin-top:0;'>📊 1. BigSeller 銷售數據更新</h3>", unsafe_allow_html=True)
    st.write("支援本地 01 至 04 資料夾的多檔清洗、字體校正，並全自動將庫存、銷量、利潤數據同步至雲端主表。")
    st.page_link("pages/1_📊_BS銷售更新.py", label="🚀 進入功能頁面", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.markdown("<h3 style='color: #2b5c8f; margin-top:0;'>🏷️ 3. 庫存資料與一維條碼整合</h3>", unsafe_allow_html=True)
    st.write("原 VBA 巨集全自動升級版。一鍵跨表關聯拆櫃、儲位與採購建議，智能取消合併儲格，並自動篩選儲位。")
    st.page_link("pages/3_🏷️_庫存資料一鍵產出.py", label="🚀 進入功能頁面", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.markdown("<h3 style='color: #2b5c8f; margin-top:0;'>📦 2. 貨櫃箱號自動產出</h3>", unsafe_allow_html=True)
    st.write("自動讀取原始拆櫃報表，依據箱數進行智能化列數倍增、自動生成序號，並針對空白異常欄位高亮標記。")
    st.page_link("pages/2_📦_貨櫃箱號自動產出.py", label="🚀 進入功能頁面", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.markdown("<h3 style='color: #2b5c8f; margin-top:0;'>🧾 4. 正隆帳單自動化核對</h3>", unsafe_allow_html=True)
    st.write("雙階段核心審計系統。首階段無縫整合 LINE 對話叫貨紀錄回填雲端，次階段啟動 PDF 帳單全量交叉對帳。")
    st.page_link("pages/4_🧾_正隆帳單核對.py", label="🚀 進入功能頁面", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("---")
st.caption("✨ 目前版本: V4.1 (多頁面架構升級版) | 基於 Streamlit 原生多頁面路由驅動")
