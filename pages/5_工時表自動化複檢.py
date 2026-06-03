import streamlit as st
import pandas as pd
from PIL import Image
import base64
import json
from openai import OpenAI

# 初始化 OpenAI 客戶端（請記得替換成你的真實 sk-... 金鑰）
client = OpenAI(api_key="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")

# 1. 隱藏原生多頁面選單，並利用 CSS 進行版面優化與字體微縮
st.markdown("""
    <style>
        /* 隱藏原生導覽選單 */
        [data-testid="stSidebarNav"] {display: none !important;}
        
        /* 全寬版型優化，靠左對齊減少白邊 */
        .main .block-container {
            max-width: 98% !important;
            padding-left: 1.5rem !important;
            padding-right: 1.5rem !important;
        }
        
        /* 🎨 精細化標題字級 */
        .custom-title {
            font-size: 1.4rem !important;
            font-weight: 700;
            margin-bottom: 0.2rem;
        }
        .custom-subtitle {
            font-size: 1.05rem !important;
            font-weight: 600;
            margin-top: 0.5rem;
            margin-bottom: 0.6rem;
        }
        
        /* 📉 強制將 st.metric 的巨大字體縮小，防止右擠與點點點 (...) */
        [data-testid="stMetricValue"] {
            font-size: 1.3rem !important;
            font-weight: 700 !important;
        }
        [data-testid="stMetricLabel"] {
            font-size: 0.8rem !important;
        }
        [data-testid="stMetricDelta"] {
            font-size: 0.75rem !important;
        }
        
        /* 📝 縮小表格字體，符合緊湊視圖 */
        .stDataFrame div {
            font-size: 0.85rem !important;
        }
    </style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("# 🧰 每日工具包中心")
    st.markdown("---")
    st.markdown("### 🗂️ 功能導覽選單")
    st.page_link("app.py", label="🏠 數據處理中心 (四大工具)", use_container_width=True)
    st.page_link("pages/5_工時表自動化複檢.py", label="📊 工時表自動化複檢", use_container_width=True)
    st.markdown("---")
    st.caption("✨ 目前版本: V4.0 (動態時薪整合版)")

st.markdown('<div class="custom-title">📊 工時表自動化複檢系統</div>', unsafe_allow_html=True)
st.markdown("<p style='color: #666666; font-size: 0.8rem; margin-bottom: 0px;'>上傳手寫工時表照片，系統將自動進行 AI 字體辨識、動態複算工時，並即時審計發放薪資。</p>", unsafe_allow_html=True)
st.markdown("---")

def encode_image_to_base64(uploaded_file):
    bytes_data = uploaded_file.getvalue()
    return base64.b64encode(bytes_data).decode('utf-8')

def analyze_timesheet(base64_image):
    # 升級 Prompt：解除時薪 220 的硬編碼，強控 AI 必須從圖片表頭動態提取手寫時薪
    prompt = """
    你是一個嚴謹的財務審計AI。請仔細辨識手寫工時表照片，執行客觀提取。
    
    【核心辨識與提取規則】
    1. 表格幾何結構：
       - 表格分為左右兩側（左半部為 1-15 日，右半部為 16-31 日）。
       - 請依據每一橫行的「日期數字」精準定位，絕對不可跨行、錯格或漏行。
    2. 欄位精準提取：
       - 上班時間、下班時間：請依據手寫數字原樣提取（例如：09:00, 16:30, 14:00 等）。
       - 備註欄：請完整「原文原樣」提取格子內的手寫中文字或時間範圍（如「中午沒休息」或「13:00-13:30」）。絕對不要自行想像或填寫任何圖上沒有出現的文字。
       - 主管簽名欄：提取該格內填寫的工時數字（可能是整數或帶有.5的分數）。
    3. 表頭與表尾總計提取（⚠️關鍵）：
       - 請仔細尋找並提取表格最外圍手寫的：「時薪」（例如圖中寫 時薪:220，請提取數字 220，若為其他數字請如實提取，絕不可硬編碼）。
       - 提取外圍手寫的：員工姓名、總時數、總薪（或總金額）。
       - 請提取左半部（1-15日）正下方的「手寫加總數字」。
       - 請提取右半部（16-31日）正下方的「手寫加總數字」。
    4. 運算核對邏輯：
       - 每日 AI 理論工時計算：(下班時間 - 上班時間)。若備註欄有寫明確的休息時間範圍（如 13:00-13:30），請扣除對應的時數（0.5小時）；若備註寫「中午沒休息」，則不扣除任何時間。
       - 總薪資理論計算：將所有每日 AI 理論工時加總後，乘以「提取到的手寫時薪」。

    請嚴格以下列 JSON 格式回傳結果，不要包含任何 Markdown 標籤：
    {
      "basic_info": {
        "name": "提取員工姓名",
        "month": "提取月份",
        "hourly_wage": 0, 
        "handwritten_left_total_hours": 0.0,
        "handwritten_right_total_hours": 0.0,
        "handwritten_total_hours": 0.0,
        "handwritten_total_amount": 0
      },
      "daily_records": [
        {
          "date": "5/X",
          "side": "left",
          "start_time": "HH:MM",
          "end_time": "HH:MM",
          "note": "備註原文",
          "manager_hours": 0.0,
          "ai_calculated_hours": 0.0
        }
      ],
      "audit_result": {
        "ai_left_total_hours": 0.0,
        "ai_right_total_hours": 0.0,
        "ai_total_hours": 0.0,
        "ai_total_amount": 0,
        "check_pass": true,
        "summary_notes": "審核結論"
      }
    }
    """
    response = client.chat.completions.create(
        model="gpt-4o",
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}
                ]
            }
        ],
        max_tokens=2000,
    )
    return json.loads(response.choices[0].message.content)

col1, col2 = st.columns([1, 1.2], gap="medium")

with col1:
    st.markdown('<div class="custom-subtitle">📸 1. 上傳來源圖片</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("請選擇或拖曳工時表照片 (JPG/PNG)", type=["jpg", "jpeg", "png"], label_visibility="collapsed")
    
    if uploaded_file is not None:
        image = Image.open(uploaded_file)
        st.image(image, caption="已上傳的工時表照片", use_container_width=True)

with col2:
    st.markdown('<div class="custom-subtitle">👁️ 2. AI 辨識與複算結果</div>', unsafe_allow_html=True)
    
    if uploaded_file is not None:
        if st.button("🚀 開始自動化複檢", type="primary", use_container_width=True):
            with st.spinner("AI 正在辨識字體並動態計算中..."):
                try:
                    base64_img = encode_image_to_base64(uploaded_file)
                    result = analyze_timesheet(base64_img)
                    info = result["basic_info"]
                    audit = result["audit_result"]
                    
                    st.success("🎉 辨識與計算完成！")
                    
                    # 🏠 核心大局核對（欄位調整為 4 欄，加入動態時薪顯示）
                    m_col1, m_col2, m_col3, m_col4 = st.columns(4)
                    m_col1.metric("員工姓名", info["name"])
                    m_col2.metric("手寫時薪", f"${info['hourly_wage']} / 🤹")
                    m_col3.metric("總工時 (手寫)", f"{info['handwritten_total_hours']} 小時", f"AI計算: {audit['ai_total_hours']}")
                    m_col4.metric("總薪資 (手寫)", f"${info['handwritten_total_amount']:,}", f"AI理論: ${audit['ai_total_amount']:,}")
                    
                    st.markdown("<hr style='margin-top:0.5rem; margin-bottom:0.5rem;'>", unsafe_allow_html=True)
                    
                    # 📊 左右半邊工時拆分卡片
                    st.markdown("<p style='font-size:0.9rem; font-weight:600; margin-bottom:5px;'>🧭 左右半邊工時拆分核對</p>", unsafe_allow_html=True)
                    split_col1, split_col2 = st.columns(2)
                    split_col1.metric(
                        "⬅️ 左半部加總 (1-15日)", 
                        f"{info['handwritten_left_total_hours']} 小時", 
                        f"AI算: {audit['ai_left_total_hours']}"
                    )
                    split_col2.metric(
                        "➡️ 右半部加總 (16-31日)", 
                        f"{info['handwritten_right_total_hours']} 小時", 
                        f"AI算: {audit['ai_right_total_hours']}"
                    )
                    
                    st.markdown("<hr style='margin-top:0.5rem; margin-bottom:0.5rem;'>", unsafe_allow_html=True)
                    
                    st.markdown("<p style='font-size:0.9rem; font-weight:600; margin-bottom:5px;'>🔍 審計判定</p>", unsafe_allow_html=True)
                    if audit['check_pass']:
                        st.info(f"✅ {audit['summary_notes']}")
                    else:
                        st.error(f"❌ {audit['summary_notes']}")
                    
                    # 明細轉化
                    df_daily = pd.DataFrame(result["daily_records"])
                    df_daily['side'] = df_daily['side'].map({'left': '左半(1-15)', 'right': '右半(16-31)'})
                    df_daily.columns = ["日期", "表格位置", "上班時間", "下班時間", "備註", "主管簽名時數", "AI理論工時"]
                    
                    st.markdown("<p style='font-size:0.9rem; font-weight:600; margin-top:10px; margin-bottom:5px;'>📅 每日明細比對清單</p>", unsafe_allow_html=True)
                    st.dataframe(df_daily, use_container_width=True)
                    
                except Exception as e:
                    st.error(f"執行過程中發生錯誤: {e}")
    else:
        st.info("請先在左側上傳工時表圖片，系統將自動為您輸出對比明細。")
