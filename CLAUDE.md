# Daily-Toolkit 專案規則

## 專案背景
這是一個 Streamlit 多頁面工具網站，部署在 Streamlit Cloud，會自動偵測 GitHub main 分支的更新並重新部署。

結構：
- `app.py`：首頁/導覽
- `pages/1_📊_BS銷售更新.py`：讀取多份 Excel 報表，清洗、比對表頭後同步到 Google Sheets（用 gspread）
- `pages/2_📦_貨櫃箱號自動產出.py`
- `pages/3_🔖_庫存資料一鍵產出.py`
- `pages/4_📄_BS採購建議_Google舊版.py`
- `requirements.txt` / `packages.txt`：套件清單

操作者是 Git / Python 新手，安全邊界要抓嚴一點，寧可多問，不要自作主張。

---

## 開始任何任務前，一定要做
1. `git status` 確認目前是乾淨狀態，沒有未處理的殘留改動
2. `git fetch` + `git pull`，確保是最新版本再開始改
3. 只針對「這次任務有關」的檔案動手，不要順手修改、格式化、或「優化」沒被要求的其他程式碼
4. 不確定需求範圍時，先問清楚，不要自己腦補、自己擴大任務範圍

## 絕對不能碰的東西
- `service_account.json`、`.streamlit/secrets.toml` 或任何看起來像金鑰/密碼/token 的檔案內容，不讀取、不印出、不上傳、不寫進 commit message
- 不要新增、刪除 Google Sheets 上的分頁或欄位結構，除非任務明確要求
- 不要改動 `SPREADSHEET_ID`、`SHEET_CONFIGS` 裡的 `gid` 對應關係，除非任務明確要求

## 修改程式碼時的已知地雷
- `st.markdown("""...""", unsafe_allow_html=True)` 這種多行字串，新增/修改內容時，一定要確認新內容在三個引號 `"""` 之內，並且縮排跟同段落其他行一致，避免造成 SyntaxError
- Google Sheets API（gspread）呼叫有限流風險，避免在短時間內對同一份試算表新增大量無必要的 API 呼叫

## 改完之後，push 之前一定要做
1. 用 `python -m py_compile <改動的檔案路徑>` 或等效方式，先確認語法沒有錯誤
2. 執行 `git diff` 自我檢查這次改動，確認沒有動到不相關的內容
3. commit message 要清楚描述「改了什麼、為什麼改」，用中文寫，禁止空泛的 "update" 這種訊息

## Push 到 GitHub 的規則（重要）
先判斷這次改動屬於「小改動」還是「大改動」：

**小改動（驗證通過後可以直接 push，不用先問）**
- 文字內容、UI 顯示文案、CSS 樣式調整
- 註解、說明文字
- 單一函式內的小幅邏輯修正（例如修一個已知的 bug，改動在 20 行以內）

**大改動（一定要先跟我說明打算怎麼改、為什麼，等我回覆確認後才 push）**
- 任何牽涉 Google Sheets 讀寫邏輯、欄位對應（HEADERS_xx）、資料清洗規則的修改
- 新增或刪除功能、新增依賴套件（requirements.txt 有變動）
- 修改 `SHEET_CONFIGS`、`SPREADSHEET_ID` 等核心設定
- 一次改動涉及超過 2 個檔案，或單一檔案改動超過 50 行
- 任何你自己也不是 100% 確定不會造成副作用的改動

不確定算小改動還是大改動時，一律當作大改動處理，先問再 push。

## 完成後
- 用一句話總結這次做了什麼、改了哪個檔案、屬於小改動還是大改動、是否已經 push
- 不要加不必要的稱讚或客套話，直接講重點
- 如果驗證/測試沒過，明確說「沒過，原因是＿＿＿」，不要淡化或隱藏問題
