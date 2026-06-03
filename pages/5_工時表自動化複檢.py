import streamlit as st
import pandas as pd
from PIL import Image
import base64
import json
from openai import OpenAI

# 初始化 OpenAI 客戶端（記得把這裡換成你的 sk-... 金鑰喔）
client = OpenAI(api_key="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")

# 1. 隱藏原生多頁面選單，維持高質感側邊欄
st.markdown("""
    <style>
        [data-testid="stSidebarNav"] {display: none !important;}
    </style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("# 🧰 每日工具包中心")
    st.markdown("---")
    st.markdown("### 🗂️ 功能導覽選單")
    st.page_link("app.py", label="🏠 數據處理中心 (四大工具)", use_container_width=True)
    st.page_link("pages/5_工時表自動化複檢.py", label="📊 工時表自動化複檢", use_container_width=True)
    st.markdown("---")
    st.caption("✨ 目前版本: V3.6 (強控防錯版)")

# 2. 主頁面視覺
st.title("📊 工時表自動化複檢系統")
st.markdown("<p style='color: #666666; font-size: 1rem;'>上傳手寫工時表照片，系統將自動進行 AI 字體辨識、動態複算工時，並即時審計發放薪資。</p>", unsafe_allow_html=True)
st.markdown("---")

def encode_image_to_base64(uploaded_file):
    bytes_data = uploaded_file.getvalue()
    return base64.b64encode(bytes_data).decode('utf-8')

def analyze_timesheet(base64_image):
    prompt = """
    你是一個極度嚴謹的財務審計專家。眼前這張手寫工時表有很多連筆字，且備註欄的「12:00-12:30」容易與下班時間混淆。
    請嚴格遵循以下「逐行精準掃描」規則來解析，不要跳行，也不要讓上下行的資料錯格：

    【精準辨識規則】
    1. 定位日期：這張表格分為「左半部 (1-15日)」與「右半部 (16-31日)」。請一行一行對齊，有寫字的那天才能抓取。
    2. 欄位嚴格對齊：
       - 上班時間、下班時間、備註、主管簽名時數這四個欄位在同一橫線上。
       - 例如：表格中許多天的備註欄寫的是休息時間「12:00-12:30」或「12:00~12:30」，請把它乖乖放在備註欄，絕對不要跟下班時間搞混！
    3. 數字微觀辨識：
       - 仔細分辨「16:30」與「12:30」。
       - 仔細分辨日期數字，例如「7」不要看成「6」；「8」不要看成「7」。
       - 主管簽名時數欄位通常是手寫的單個數字（例如 6, 7, 2），請精準抓取。
    4. 工時計算邏輯：
       - 只要備註欄寫了「12:00-12:30」，代表午休 0.5 小時。
       - 計算公式：AI理論工時 = (下班時間 - 上班時間) - 0.5 小時。
       - 比對：檢查「AI理論工時」是否與「主管簽名時數」完全相等。

    請嚴格以下列 JSON 格式回傳結果，不要包含任何 Markdown 標籤，確保能被 json.loads 解析：
    {
      "basic_info": {
        "name": "陳玨伶",
        "month": "5月",
        "hourly_wage": 220,
        "handwritten_total_hours": 73,
        "handwritten_total_amount": 16060
      },
      "daily_records": [
        {
          "date": "5/4",
          "start_time": "10:00",
          "end_time": "16:30",
          "note": "12:00-12:30",
          "manager_hours": 6,
          "ai_calculated_hours": 6.0,
          "is_matched": true
        },
        {
          "date": "5/5",
          "start_time": "9:00",
          "end_time": "16:30",
          "note": "12:00-12:30",
          "manager_hours": 7,
          "ai_calculated_hours": 7.0,
          "is_matched": true
        }
      ],
      "audit_result": {
        "ai_total_hours": 73.0,
        "ai_total_amount": 16060,
        "hours_check_pass": true,
        "amount_check_pass": true,
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

# 左右雙欄版面排版
col1, col2 = st.columns([1, 1.2])

with col1:
    st.subheader("📸 1. 上傳來源圖片")
    uploaded_file = st.file_uploader("請選擇或拖曳工時表照片 (JPG/PNG)", type=["jpg", "jpeg", "png"], label_visibility="collapsed")
    
    if uploaded_file is not None:
        image = Image.open(uploaded_file)
        st.image(image, caption="已上傳的工時表照片", use_container_width=True)

with col2:
    st.subheader("👁️ 2. AI 辨識與複算結果")
    
    if uploaded_file is not None:
        if st.button("🚀 開始自動化複檢", type="primary", use_container_width=True):
            with st.spinner("AI 正在辨識字體並動態計算中..."):
                try:
                    base64_img = encode_image_to_base64(uploaded_file)
                    result = analyze_timesheet(base64_img)
                    info = result["basic_info"]
                    audit = result["audit_result"]
                    
                    st.success("🎉 辨識與計算完成！")
                    
                    m_col1, m_col2, m_col3 = st.columns(3)
                    m_col1.metric("員工姓名", info["name"])
                    m_col2.metric("總時數核對", f"{audit['ai_total_hours']} 小時", f"手寫: {info['handwritten_total_hours']}")
                    m_col3.metric("總薪資核對", f"${audit['ai_total_amount']:,}", f"手寫: ${info['handwritten_total_amount']:,}")
                    
                    st.write("### 🔍 審計判定")
                    if audit['hours_check_pass'] and audit['amount_check_pass']:
                        st.info(f"✅ **總數檢查通過**：{audit['summary_notes']}")
                    else:
                        st.error(f"❌ **發現邏輯不符落差**：{audit['summary_notes']}")
                    
                    df_daily = pd.DataFrame(result["daily_records"])
                    df_daily.columns = ["日期", "上班時間", "下班時間", "備註", "主管簽名時數", "AI理論工時", "工時吻合"]
                    st.write("### 📅 每日明細比對清單")
                    st.dataframe(df_daily, use_container_width=True)
                    
                except Exception as e:
                    st.error(f"執行過程中發生錯誤: {e}")
    else:
        st.info("請先在左側上傳工時表圖片，系統將自動為您輸出對比明細。")
