import streamlit as st
import pandas as pd
from PIL import Image
import base64
import json
from openai import OpenAI

# 初始化 OpenAI 客戶端（請替換為你的 API Key，或使用 st.secrets）
client = OpenAI(api_key="YOUR_OPENAI_API_KEY_HERE")

# 這裡不需再寫 st.set_page_config(layout="wide")，因為主架構會依循 app.py
st.title("📊 工時表自動化複檢系統")
st.subheader("上傳手寫工時表照片，自動辨識、複算工時與薪資")

def encode_image_to_base64(uploaded_file):
    bytes_data = uploaded_file.getvalue()
    return base64.b64encode(bytes_data).decode('utf-8')

def analyze_timesheet(base64_image):
    prompt = """
    你是一個嚴謹的財務審計助手。請仔細辨識這張手寫工時表，並執行以下步驟：
    1. 提取基本資訊：姓名、月份、手寫時薪、手寫總時數、手寫總金額。
    2. 提取每日工時：將有填寫時間的日期列出來，包含日期、上班時間、下班時間、備註、主管簽名時數。
    3. 邏輯計算：
       - 依據上下班時間計算出『AI理論工時』。
       - 注意：若備註有『中午沒休』，則不扣除午休；若無備註且跨越中午，請依常理扣除1小時。
       - 比對『主管簽名時數』與『AI理論工時』是否吻合。
    4. 總計複核：
       - 將每日的『AI理論工時』加總，比對是否等於表格上方的『手寫總時數』。
       - 將 AI 加總的總時數乘以『手寫時薪』，比對是否等於表格上方的『手寫總金額』。

    請嚴格以下列 JSON 格式回傳結果，不要包含任何 Markdown 標籤：
    {
      "basic_info": {"name": "姓名", "month": "月份", "hourly_wage": 220, "handwritten_total_hours": 51, "handwritten_total_amount": 11220},
      "daily_records": [{"date": "5/2", "start_time": "10:00", "end_time": "17:00", "note": "中午沒休", "manager_hours": 7, "ai_calculated_hours": 7.0, "is_matched": true}],
      "audit_result": {"ai_total_hours": 51.0, "ai_total_amount": 11220, "hours_check_pass": true, "amount_check_pass": true, "summary_notes": "審核結果說明"}
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

# 介面排版
col1, col2 = st.columns([1, 1.2])

with col1:
    st.header("📸 1. 上傳來源圖片")
    uploaded_file = st.file_uploader("請選擇或拖曳工時表照片 (JPG/PNG)", type=["jpg", "jpeg", "png"])
    if uploaded_file is not None:
        image = Image.open(uploaded_file)
        st.image(image, caption="已上傳的工時表照片", use_container_width=True)

with col2:
    st.header("🤖 2. AI 辨識與複算結果")
    if uploaded_file is not None:
        if st.button("🚀 開始自動化複檢", type="primary"):
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
        st.info("請先在左側上傳工時表圖片。")
