import streamlit as st
import pandas as pd
from PIL import Image
import base64
import json
from openai import OpenAI

# 初始化 OpenAI 客戶端
client = OpenAI(api_key="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")

# 1. 隱藏原生多頁面選單，自訂全寬版型
st.markdown("""
    <style>
        [data-testid="stSidebarNav"] {display: none !important;}
        .main .block-container {
            max-width: 95% !important;
            padding-left: 2rem !important;
            padding-right: 2rem !important;
        }
        .custom-title {
            font-size: 1.6rem !important;
            font-weight: 700;
            margin-bottom: 0.2rem;
        }
        .custom-subtitle {
            font-size: 1.1rem !important;
            font-weight: 600;
            margin-top: 0.5rem;
            margin-bottom: 0.8rem;
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
    st.caption("✨ 目前版本: V3.8 (左右拆分精準版)")

st.markdown('<div class="custom-title">📊 工時表自動化複檢系統</div>', unsafe_allow_html=True)
st.markdown("<p style='color: #666666; font-size: 0.85rem; margin-bottom: 0px;'>上傳手寫工時表照片，系統將自動進行 AI 字體辨識、動態複算工時，並即時審計發放薪資。</p>", unsafe_allow_html=True)
st.markdown("---")

def encode_image_to_base64(uploaded_file):
    bytes_data = uploaded_file.getvalue()
    return base64.b64encode(bytes_data).decode('utf-8')

def analyze_timesheet(base64_image):
    # 重新微調的極致強控 Prompt，鎖定左右兩半邊的區域定位與個別加總
    prompt = """
    你是一個極度嚴謹的財務審計專家。眼前這張手寫工時表有很多連筆字，且備註欄的「13:00-13:30」容易與下班時間混淆。
    請嚴格遵循以下「左右分開掃描」規則來解析，不要跳行，也不要讓資料錯格：

    【精準辨識規則】
    1. 這張表格結構分為左右兩側：
       - 左半部（日期 1 到 15日）：最底下的手寫加總數字即為「左側手寫總工時」（例如：37.5）。
       - 右半部（日期 16 到 31日）：最底下的手寫加總數字即為「右側手寫總工時」（例如：30.5）。
    2. 逐行嚴格對齊：
       - 請一行一行掃描，只要有手寫時間的那一天就必須抓取。
       - 看清楚格線，上班、下班、備註、主管簽名時數在同一橫線上。
       - 備註欄通常寫休息時間（如 13:00-13:30），請「原封不動」放進備註，絕對不要跟下班時間搞混！
    3. 工時計算邏輯：
       - AI理論工時 = (下班時間 - 上班時間) - 備註欄的休息時數。
       - 請分別將「左半部所有天數的工時」與「右半部所有天數的工時」各自進行獨立加總。

    請嚴格以下列 JSON 格式回傳結果，不要包含任何 Markdown 標籤，確保能被 json.loads 解析：
    {
      "basic_info": {
        "name": "員工姓名",
        "month": "5月",
        "hourly_wage": 220,
        "handwritten_left_total_hours": 37.5,
        "handwritten_right_total_hours": 30.5,
        "handwritten_total_hours": 68.0,
        "handwritten_total_amount": 14960
      },
      "daily_records": [
        {
          "date": "5/4",
          "side": "left",
          "start_time": "10:30",
          "end_time": "16:30",
          "note": "13:00-13:30",
          "manager_hours": 5.5,
          "ai_calculated_hours": 5.5
        }
      ],
      "audit_result": {
        "ai_left_total_hours": 37.5,
        "ai_right_total_hours": 30.5,
        "ai_total_hours": 68.0,
        "ai_total_amount": 14960,
        "check_pass": true,
        "summary_notes": "審核完成"
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

col1, col2 = st.columns([1, 1.2], gap="large")

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
                    
                    # 🏠 第一層：員工姓名與大局核對
                    m_col1, m_col2, m_col3 = st.columns(3)
                    m_col1.metric("員工姓名", info["name"])
                    m_col2.metric("總時數核對", f"{audit['ai_total_hours']} 小時", f"手寫總計: {info['handwritten_total_hours']}")
                    m_col3.metric("總薪資核對", f"${audit['ai_total_amount']:,}", f"手寫總計: ${info['handwritten_total_amount']:,}")
                    
                    st.markdown("---")
                    
                    # 📊 第二層：新增「左側工時」與「右側工時」的個別加總卡片
                    st.write("### 🧭 左右半邊工時拆分核對")
                    split_col1, split_col2 = st.columns(2)
                    split_col1.metric(
                        "⬅️ 左半部工時加總 (1-15日)", 
                        f"{audit['ai_left_total_hours']} 小時", 
                        f"手寫欄底: {info['handwritten_left_total_hours']}"
                    )
                    split_col2.metric(
                        "➡️ 右半部工時加總 (16-31日)", 
                        f"{audit['ai_right_total_hours']} 小時", 
                        f"手寫欄底: {info['handwritten_right_total_hours']}"
                    )
                    
                    st.markdown("---")
                    
                    st.write("### 🔍 審計判定")
                    if audit['check_pass']:
                        st.info(f"✅ **總數檢查通過**：{audit['summary_notes']}")
                    else:
                        st.error(f"❌ **發現邏輯不符落差**：{audit['summary_notes']}")
                    
                    # 輸出明細
                    df_daily = pd.DataFrame(result["daily_records"])
                    # 將側邊欄英文標記轉為中文顯示，方便閱讀
                    df_daily['side'] = df_daily['side'].map({'left': '左半邊 (1-15)', 'right': '右半邊 (16-31)'})
                    df_daily.columns = ["日期", "表格位置", "上班時間", "下班時間", "備註", "主管簽名時數", "AI理論工時"]
                    
                    st.write("### 📅 每日明細比對清單")
                    st.dataframe(df_daily, use_container_width=True)
                    
                except Exception as e:
                    st.error(f"執行過程中發生錯誤: {e}")
    else:
        st.info("請先在左側上傳工時表圖片，系統將自動為您輸出對比明細。")
