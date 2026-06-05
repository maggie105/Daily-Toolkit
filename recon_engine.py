# -*- coding: utf-8 -*-
"""
recon_engine.py — 正隆紙箱帳單 OCR 提取與對帳引擎（沐樂國際香氛）／ PaddleOCR 版

與 UI 解耦的純函式：process_pdf(pdf_bytes=...) 或 process_pdf(pdf_path=...)
                    -> (df_stmt, df_inv, df_recon, log)

為何用 PaddleOCR：
  本檔為純掃描、密集中文表格、且頁面有正/反/橫各種旋轉。PaddleOCR 回傳
  「每個文字框 + 座標」，據此可把同一列欄位重組成一行（還原表格），辨識率
  與穩定度遠勝 Tesseract，且模型輸出固定（你我同一套模型結果一致）。

金額一律以台灣 5% 營業稅推導：未稅×1.05=總計、稅=round(未稅×0.05)、數量×單價=未稅。

相依：
  pip install paddlepaddle==3.1.0 paddleocr pdf2image pillow pandas openpyxl
  系統需 poppler（pdf2image 用）。Windows 把 poppler 的 bin 加進 PATH，或在下方
  POPPLER_PATH 直接指定。
  首次執行 PaddleOCR 會自動下載模型（需連網），之後離線可用。

用法：
  python recon_engine.py 帳單.pdf            # 跑對帳，印逐頁判斷與結果
  python recon_engine.py 帳單.pdf dump        # 另存 paddle_dump.txt（每頁重組後文字）
                                              # 數字若有偏差，把這檔傳回來即可精準重調
"""
import re, sys
import numpy as np
import pandas as pd
from pdf2image import convert_from_bytes, convert_from_path

# 若 poppler 沒進 PATH，可在此直接指定 bin 路徑（否則留 None）
POPPLER_PATH = None   # 例：r"C:\Users\88692\Desktop\poppler-26.02.0\Library\bin"

# ---------- 正規表達式 ----------
ADDR   = re.compile(r"[路段號樓區市縣鄉鎮里鄰巷]")
DIMS   = re.compile(r"\(\s*\d{2,4}\s*[\*xX×]\s*\d{2,4}\s*[\*xX×]\s*\d{2,4}\s*\)")
PRICE5 = re.compile(r"(\d{1,4})\s*[.,．]+\s*(\d{5})")
PRICE2 = re.compile(r"(?<![\d.])(\d{1,4})\s*[.,]\s*(\d{2})(?!\d)")
THOUS  = re.compile(r"\d{1,3}(?:[.,]\d{3})+")
DATE_D = re.compile(r"20\d\d[-/]\d{1,2}[-/]\d{1,2}")
PLANT  = re.compile(r"正隆股份有限公司([\u4e00-\u9fff]{2,8}廠)")
SIZES  = {300,400,450,250,150,200,230,530,120,180,405,420,30,45,65,90,
          350,252,270,190,160,440,600,100,765,35}

def _vat(net): t = round(net*0.05); return t, net+t
def _is_int(x, tol=0.02): return abs(x-round(x)) < tol
def _to_int(tok): return int(re.sub(r"[.,]", "", tok))

def _norm_bc(tail):
    mp = {"I":"1","l":"1","|":"1","O":"0","o":"0","S":"5","B":"8","Z":"2","b":"6","g":"9"}
    t = "".join(mp.get(c, c) for c in tail)
    return "BC"+t if re.fullmatch(r"\d{8}", t) else None

def _find_bc(text):
    out = []
    for m in re.finditer(r"B[C€G]\s?([0-9IlO oSBZbg]{8})", text):
        v = _norm_bc(m.group(1).replace(" ", ""))
        if v: out.append((m.start(), v))
    return out

def _clean(s):
    s = s.replace("MOO","MOQ").replace("MO0","MOQ").replace("FCS","PCS")
    return re.sub(r"\s+", "", s).strip(" )|]")

def _name(text):
    for ln in text.split("\n"):
        if DIMS.search(ln) and not ADDR.search(ln):
            return re.sub(r"^[^\u4e00-\u9fff(]{0,5}", "", _clean(ln))[:40]
    flat = text.replace("\n", ""); m = DIMS.search(flat)
    return _clean(m.group(0)) if m else ""

def _amounts(text):
    out = []
    for m in re.finditer(r"\d{1,3}(?:[.,]\d{3})+|\b\d{3,7}\b", text):
        v = int(re.sub(r"[.,]", "", m.group()))
        if 50 <= v < 1e7: out.append(v)
    return out

def _find_triple(a):
    n = len(a)
    for i in range(n):
        for j in range(n):
            if i == j: continue
            for k in range(n):
                if k in (i, j): continue
                x, y, z = a[i], a[j], a[k]
                if abs(x+y-z) <= 2 and x > y > 0 and y <= x*0.06+2:
                    return (x, y, z)
    return None

# ================= PaddleOCR =================
_OCR = None
def _get_ocr():
    """單例：初始化較慢，整批共用一個。"""
    global _OCR
    if _OCR is None:
        from paddleocr import PaddleOCR
        _OCR = PaddleOCR(
            lang="chinese_cht",
            # mobile 模型：體積小、速度快、記憶體低，雲端免費方案才跑得動；
            # 對印刷體準確度通常與 server 相當。
            text_detection_model_name="PP-OCRv5_mobile_det",
            text_recognition_model_name="PP-OCRv5_mobile_rec",
            use_doc_orientation_classify=True,   # 自動轉正整頁（取代 Tesseract 的 OSD）
            use_doc_unwarping=False,
            use_textline_orientation=True,        # 修正翻轉的文字行
        )
    return _OCR

def _pil_to_bgr(img):
    arr = np.array(img.convert("RGB"))
    return arr[:, :, ::-1]    # RGB -> BGR

def _rows_from_result(res):
    """把 PaddleOCR 的文字框依座標重組成「每列一行」的文字（還原表格列序）。"""
    texts = res.get("rec_texts") or res.get("rec_text") or []
    polys = res.get("rec_polys")
    if polys is None: polys = res.get("dt_polys")
    if polys is None or len(polys) != len(texts):
        return "\n".join(texts)                 # 退路：直接逐行
    items = []
    for t, p in zip(texts, polys):
        p = np.array(p).reshape(-1, 2)
        yc = float(p[:, 1].mean()); xc = float(p[:, 0].mean())
        h  = float(p[:, 1].max() - p[:, 1].min())
        items.append((yc, xc, h, t))
    if not items: return ""
    items.sort(key=lambda z: z[0])
    tol = max(8.0, np.median([it[2] for it in items]) * 0.6)
    rows, cur, cy = [], [], items[0][0]
    for yc, xc, h, t in items:
        if abs(yc - cy) <= tol:
            cur.append((xc, t))
        else:
            rows.append(cur); cur = [(xc, t)]; cy = yc
        cy = (cy + yc) / 2
    if cur: rows.append(cur)
    lines = [" ".join(t for _, t in sorted(r, key=lambda z: z[0])) for r in rows]
    return "\n".join(lines)

def _ocr_page(img):
    """回傳重組後的整頁文字。先靠內建轉正；若抓不到錨點再四角度擇優。"""
    ocr = _get_ocr()
    def run(im):
        res = ocr.predict(input=_pil_to_bgr(im))
        return _rows_from_result(res[0]) if res else ""
    txt = run(img)
    if ("電子發票" in txt) or ("證明聯" in txt) or _find_bc(txt):
        return txt
    best = (txt, _score(txt))
    for ang in (90, 180, 270):
        t = run(img.rotate(ang, expand=True))
        s = _score(t)
        if s > best[1]: best = (t, s)
    return best[0]

def _score(text):
    return (len(_find_bc(text))*100 + len(DATE_D.findall(text))*10
            + (50 if ("電子發票" in text or "證明聯" in text) else 0))

# ================= 解析 =================
def parse_invoice(pg, text):
    fb = _find_bc(text); bc = fb[0][1] if fb else None
    d = DATE_D.search(text); date = d.group(0).replace("-", "/") if d else ""
    plant = PLANT.search(text); plant = plant.group(1) if plant else ""
    price = None; conf = "高"
    pm = PRICE5.search(text)
    if pm: price = round(int(pm.group(1)) + int(pm.group(2))/100000, 4)
    nums = _amounts(text); net = None
    if pm:
        after = text[pm.end():]
        for nm in THOUS.finditer(after):
            if re.match(r"\s*/?\s*(PCS|FCS|P|F)", after[nm.end():nm.end()+5]):
                continue
            net = _to_int(nm.group()); break
    if net is None:
        tri = _find_triple(nums); net = tri[0] if tri else None
        if net is not None: conf = "中"
    if net and price and not _is_int(net/price):
        net2 = round(net/1.05)
        if _is_int(net2/price): net = net2
    if net is None:
        comma = [_to_int(m.group()) for m in THOUS.finditer(text)
                 if not re.match(r"\s*/?\s*(PCS|FCS|P|F)", text[m.end():m.end()+5])]
        comma = [n for n in comma if 1000 <= n <= 200000]
        if len(set(comma)) >= 2 and any(abs(min(comma)*1.05-z) <= 2 for z in comma):
            net = min(comma); conf = "中"
        elif comma:
            t = max(comma)
            net = t if (price and _is_int(t/price)) else round(t/1.05)
            conf = "低"
    qty = round(net/price) if (net and price) else None
    tax, total = _vat(net) if net else (None, None)
    return {"頁": pg, "發票號碼": bc, "發票日": date, "賣方廠": plant, "品名規格": _name(text),
            "數量": qty, "單價": price, "金額(未稅)": net, "營業稅": tax, "總計": total, "信心": conf}

def parse_statement(pg, text):
    out = []; lines = text.split("\n")
    plant = PLANT.search(text); plant = plant.group(1) if plant else ""
    for i, line in enumerate(lines):
        fb = _find_bc(line)
        if not fb: continue
        mpos, bc = fb[0]; combo = " ".join(lines[i:i+3]); conf = "高"
        d = DATE_D.search(line); date = (d.group(0).replace("-", "/") if d else "")
        price = None; pstart = len(combo)
        for pm in PRICE2.finditer(combo):
            v = int(pm.group(1)) + int(pm.group(2))/100
            if v > 0: price = v; pstart = pm.start(); break
        seg = combo[mpos+10:pstart]
        comma = re.findall(r"\d{1,3}(?:,\d{3})+", seg)
        qty = _to_int(comma[-1]) if comma else None
        if qty is None:
            c = [int(x) for x in re.findall(r"\b\d{2,5}\b", seg)
                 if 50 <= int(x) < 100000 and int(x) not in SIZES]
            qty = c[-1] if c else None
        net = round(qty*price) if (qty and price) else None
        if net is None:
            tri = _find_triple(_amounts(combo)); net = tri[0] if tri else None
            conf = "中" if net else "低"
        tax, total = _vat(net) if net else (None, None)
        out.append({"頁": pg, "發票號碼": bc, "發票日": date, "廠商": plant, "品名規格": _name(combo),
                    "數量": qty, "單價": price, "未稅金額": net, "稅額": tax, "總金額": total, "信心": conf})
    return out

# ================= 主流程 =================
def _render(pdf_bytes, pdf_path, dpi):
    kw = {"dpi": dpi}
    if POPPLER_PATH: kw["poppler_path"] = POPPLER_PATH
    return (convert_from_bytes(pdf_bytes, **kw) if pdf_bytes is not None
            else convert_from_path(pdf_path, **kw))

def process_pdf(pdf_bytes=None, pdf_path=None, dpi=200, debug=False, dump_path=None):
    """回傳 (df_stmt, df_inv, df_recon, log)。PaddleOCR 用 200dpi 通常已足夠。"""
    log = []; pages = _render(pdf_bytes, pdf_path, dpi)
    dumpf = open(dump_path, "w", encoding="utf-8") if dump_path else None
    inv, stmt = [], []
    for i, img in enumerate(pages):
        pg = i + 1
        text = _ocr_page(img)
        if dumpf:
            dumpf.write(f"\n===== p{pg} =====\n{text}\n")
        flat = text.replace(" ", "")
        is_inv = ("電子發票" in flat) or ("證明聯" in flat) or ("銷售額" in flat)
        if is_inv:
            r = parse_invoice(pg, text); inv.append(r)
            if debug: log.append(f"p{pg} 發票 {r['發票號碼']} 總計={r['總計']} 信心={r['信心']}")
        else:
            rs = parse_statement(pg, text); stmt += rs
            if debug: log.append(f"p{pg} 對帳單 列數={len(rs)} BC={[x['發票號碼'] for x in rs]}")
    if dumpf: dumpf.close()
    df_inv = pd.DataFrame(inv); df_stmt = pd.DataFrame(stmt)
    # 清掉沒有發票號碼的雜訊列、並去重
    if len(df_inv):  df_inv  = df_inv[df_inv["發票號碼"].notna()].drop_duplicates("發票號碼").reset_index(drop=True)
    if len(df_stmt): df_stmt = df_stmt[df_stmt["發票號碼"].notna()].drop_duplicates("發票號碼").reset_index(drop=True)
    df_recon = reconcile(df_stmt, df_inv)
    return df_stmt, df_inv, df_recon, log

def reconcile(df_stmt, df_inv):
    a = (df_stmt[["發票號碼", "總金額"]].rename(columns={"總金額": "對帳單總計(A)"})
         if len(df_stmt) else pd.DataFrame(columns=["發票號碼", "對帳單總計(A)"]))
    b = (df_inv[["發票號碼", "總計"]].rename(columns={"總計": "發票總計(B)"})
         if len(df_inv) else pd.DataFrame(columns=["發票號碼", "發票總計(B)"]))
    rec = pd.merge(a, b, on="發票號碼", how="outer")
    rec = rec[rec["發票號碼"].notna()]
    def status(r):
        x, y = r["對帳單總計(A)"], r["發票總計(B)"]
        if pd.isna(x) or pd.isna(y): return "⚠ 缺值待人工"
        return "✓ 一致" if abs(y-x) <= 2 else "❌ 差異待人工"
    rec["狀態"] = rec.apply(status, axis=1)
    rec["差異"] = (rec["發票總計(B)"].fillna(0) - rec["對帳單總計(A)"].fillna(0))
    return rec.sort_values("發票號碼").reset_index(drop=True)

if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else "bill.pdf"
    dump = "paddle_dump.txt" if (len(sys.argv) > 2 and sys.argv[2] == "dump") else None
    s, v, r, log = process_pdf(pdf_path=path, dpi=200, debug=True, dump_path=dump)
    print("\n".join(log))
    print(f"\n發票 {len(v)} 張、對帳單 {len(s)} 列")
    print(r["狀態"].value_counts().to_string())
    print(r.to_string(index=False))
    if dump: print(f"\n已輸出 {dump}（每頁重組後文字），數字若有偏差請把此檔傳回。")
