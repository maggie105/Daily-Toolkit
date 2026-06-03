import streamlit as st
import pandas as pd
from PIL import Image
import base64
import json
from openai import OpenAI

# 初始化 OpenAI 客戶端
client = OpenAI(api_key="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")

# 1. 隱藏原生多頁面選單，並利用 CSS 進行極致的「版面靠左放寬」與「字體微縮」
st.markdown("""
    <style>
        [data-testid="stSidebarNav"] {display: none !important;}
        .main .block-container {
            max-width: 98% !important;
            padding-left: 1.5rem !important;
            padding-right: 1.5rem !important;
        }
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
        [data-testid="stMetricValue"] {
            font-size: 1.3rem !important;
            font-weight: 700 !important;
        }
        [data-testid="stMetricLabel"] { font-size: 0.8rem; }
        [data-testid="stMetricDelta"] { font-size: 0.75rem; }
        .stDataFrame div { font-size: 0.85rem !important; }
    </style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("# 🧰 每日工具包中心")
    st.markdown("---")
    st.markdown("### 🗂️ 功能導覽選單")
    st.page_link("app.py", label="🏠 數據處理中心 (四大工具)", use_container_width=True)
    st.page_link("pages/5_工時表自動化複檢.py", label="📊 工時表自動化複檢", use_container_width=True)
    st.markdown("---")
    st.caption("✨ 目前版本: V5.0 (鋼鐵幾何防錯版)")

st.markdown('<div class="custom-title">📊 工時表自動化複檢系統</div>', unsafe_allow_html=True)
st.markdown("<p style='color: #666666; font-size: 0.8rem; margin-bottom: 0px;'>上傳手寫工時表照片，系統將自動進行 AI 字體辨識、動態複算工時，並即時審計發放薪資。</p>", unsafe_allow_html=True)
st.markdown("---")

def encode_image_to_base64(uploaded_file):
    bytes_data = uploaded_file.getvalue()
    return base64.b64encode(bytes_data).decode('utf-8')

def analyze_timesheet(base64_image):
    # V5.0 幾何鎖死 Prompt：禁止任何形式的行數位移與日期自定義
    prompt = """
    你是一個沒有感情的精準 OCR 與財務審計機器人。眼前是一張印刷好的工時紀錄表表格。
    表格有兩大欄：左半部為「日期 1 至 15」，右半部為「日期 16 至 31」。
    
    【死命令：嚴格空間幾何對齊】
    1. 每一行都有固定的印刷行號（1 到 31）。你提取的 JSON 中的 "date" 欄位，必須「只能」是純數字字串，如 "1", "2", "3" 到 "31"。
    2. 絕對不可以自作聰明加上 "/9"、"-9" 或任何斜線！印刷行號寫幾，你就提取幾。
    3. 逐行掃描：請按印刷行號 1 到 31 依序對齊檢查。如果某一行的「上班時間」或「下班時間」沒有任何手寫筆跡，代表當天沒上班，請「直接跳過」該日期，絕對不可以把下一行手寫的內容挪上來頂替！
    4. 欄位精準原樣提取：
       - 看清楚橫向格線。上班、下班、備註、主管簽名時數在同一個水平高度。
       - 備註欄如果是寫時間範圍（如 13:00-13:30），請原樣提取。如果是中文（如 中午沒休息），請如實提取。
    
    【表頭總計提取】
    - 仔細提取表頭手寫的：姓名、時薪、總時數、總薪。
    - 仔細提取左半部（1-15日）最底下的手寫加總（數字）、右半部（16-31日）最底下的手寫加總（數字）。

    請嚴格以下列 JSON 格式回傳結果，不要包含任何 Markdown 標籤：
    {
      "basic_info": {
        "name": "姓名",
        "month": "5月",
        "hourly_wage": 0,
        "handwritten_left_total_hours": 0.0,
        "handwritten_right_total_hours": 0.0,
        "handwritten_total_hours": 0.0,
        "handwritten_total_amount": 0
      },
      "daily_records": [
        {
          "date": "4", 
          "side": "left",
          "start_time": "10:00",
          "end_time": "15:00",
          "note": "13:00-13:30",
          "manager_hours": 4.5,
          "ai_calculated_hours": 4.5
        }
      ],
      "audit_result": {
        "ai_left_total_hours": 0.0,
        "ai_right_total_hours": 0.0,
        "ai_total_hours": 0.0,
        "ai_total_amount": 0,
        "check_pass": true,
        "summary_notes": "結論"
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
            with st.spinner("AI 正在使用幾何防錯網進行掃描計算中..."):
                try:
                    base64_img = encode_image_to_base64(uploaded_file)
                    result = analyze_timesheet(base64_img)
                    info = result["basic_info"]
                    audit = result["audit_result"]
                    
                    st.success("🎉 辨識與計算完成！")
                    
                    # 🏠 核心大局核對
                    m_col1, m_col2, m_col3, m_col4 = st.columns(4)
                    m_col1.metric("員工姓名", info["name"])
                    m_col2.metric("手寫時薪", f"${info['hourly_wage']} / 🤹")
                    m_col3.metric("總工時 (手寫)", f"{info['handwritten_total_hours']} 小時", f"AI計算: {audit['ai_total_hours']}")
                    m_col4.metric("總薪資 (手寫)", f"${info['handwritten_total_amount']:,}", f"AI理論: ${audit['ai_total_amount']:,}")
                    
                    st.markdown("<hr style='margin-top:0.5rem; margin-bottom:0.5rem;'>", unsafe_allow_html=True)
                    
                    # 📊 左右半邊工時拆分卡片
                    st.markdown("<p style='font-size:0.9rem; font-weight:600; margin-bottom:5px;'>🧭 左右半邊工時拆分核對</p>", unsafe_allow_html=True)
                    split_col1, split_col2 = st.columns(2)
                    split_col1.metric("左半部加總 (1-15日)", f"{info['handwritten_left_total_hours']} 小時", f"AI算: {audit['ai_left_total_hours']}")
                    split_col2.metric("右半部加總 (16-31日)", f"{info['handwritten_right_total_hours']} 小時", f"AI算: {audit['ai_right_total_hours']}")
                    
                    st.markdown("<hr style='margin-top:0.5rem; margin-bottom:0.5rem;'>", unsafe_allow_html=True)
                    
                    st.markdown("<p style='font-size:0.9rem; font-weight:600; margin-bottom:5px;'>🔍 審計判定</p>", unsafe_allow_html=True)
                    if audit['check_pass']:
                        st.info(f"✅ {audit['summary_notes']}")
                    else:
                        st.error(f"❌ {audit['summary_notes']}")
                    
                    # 明細處理
                    df_daily = pd.DataFrame(result["daily_records"])
                    # 強制格式化日期顯示，避免出現 /9
                    df_daily['date'] = df_daily['date'].apply(lambda x: f"{info['month']}{x}日")
                    df_daily['side'] = df_daily['side'].map({'left': '左半(1-15)', 'right': '右半(16-31)'})
                    df_daily.columns = ["日期", "表格位置", "上班時間", "下班時間", "備註", "主管簽名時數", "AI理論工時"]
                    
                    st.markdown("<p style='font-size:0.9rem; font-weight:600; margin-top:10px; margin-bottom:5px;'>📅 每日明細比對清單</p>", unsafe_allow_html=True)
                    st.dataframe(df_daily, use_container_width=True)
                    
                except Exception as e:
                    st.error(f"執行過程中發生錯誤: {e}")
    else:
        st.info("請先在左側上傳工時表圖片，系統將自動為您輸出對比明細。")
