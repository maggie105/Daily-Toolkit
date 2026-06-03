import streamlit as st
import pandas as pd
from PIL import Image
import base64
import json
from openai import OpenAI

# 初始化 OpenAI 客戶端
client = OpenAI(api_key="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")

# 1. 隱藏原生多頁面自動生成的選單，並維持統一的漂亮側邊欄
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
    st.caption("✨ 目前版本: V3.5 (UI 高質感優化版)")

# --- 主頁面視覺調整 (對齊 BigSeller 系統的字級) ---
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
        "summary_notes": "請在這裡寫下你的審計結論，例如：總金額與時數核對完全正確。"
      }
    }
    """
