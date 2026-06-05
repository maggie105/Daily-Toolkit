# -*- coding: utf-8 -*-
"""
recon_engine.py — 正隆紙箱帳單 OCR 提取與對帳引擎（沐樂國際香氛）

與 UI 解耦的純函式：process_pdf(pdf_bytes) -> (df_stmt, df_inv, df_recon, log)

實作要點（皆由真實掃描檔 49 頁實測歸納）：
  1. 純掃描、無文字層、且「格式 25 電子發票證明聯」不含 QR Code → 一律走 OCR。
  2. 轉正：OSD 對乾淨發票準，但對密集對帳單信心極低會判成鏡像；故採
     「OSD 快路徑 + 抓不到錨點時四角度擇優（BC 數×100 + 日期數×10 + 電子發票×50）」。
  3. 金額一律以台灣 5% 營業稅推導：未稅×1.05=總計、稅=round(未稅×0.05)、
     數量×單價=未稅。比硬讀被重量欄汙染的金額穩。
  4. 發票淨額：單價(5 位小數，容忍逗號/空白當小數點)右側第一個帶千分位金額；
     失敗則 find_triple（net+稅≈total 且 稅≈5%），再失敗則由總計÷1.05 回推。
  5. BC 號正規化（I→1、O→0、S→5…）以容忍 OCR 誤讀。
  6. 對帳依 BC 分組比對總計；任一側缺值→「缺值待人工」，差異>2→「差異」。

相依：pip install pytesseract pdf2image pillow pandas openpyxl
系統：poppler（pdftoppm）、tesseract 及繁中語言包 chi_tra
  本機 Windows：安裝 Tesseract-OCR 並把 chi_tra.traineddata 放入 tessdata。
  Streamlit Cloud：packages.txt 放 poppler-utils / tesseract-ocr / tesseract-ocr-chi-tra
"""
import re, io
import pandas as pd
import pytesseract
from pytesseract import Output
from pdf2image import convert_from_bytes, convert_from_path

LANG = "chi_tra+eng"

# ---------- 正規表達式 ----------
ADDR   = re.compile(r"[路段號樓區市縣鄉鎮里鄰巷]")
DIMS   = re.compile(r"\(\s*\d{2,4}\s*[\*xX×]\s*\d{2,4}\s*[\*xX×]\s*\d{2,4}\s*\)")
PRICE5 = re.compile(r"(\d{1,4})\s*[.,．]+\s*(\d{5})")           # 發票單價 X.XXXXX（容忍逗號/空白）
PRICE2 = re.compile(r"(?<![\d.])(\d{1,4})\s*[.,]\s*(\d{2})(?!\d)")  # 對帳單單價 X.XX
THOUS  = re.compile(r"\d{1,3}(?:[.,]\d{3})+")                   # 1,941 或 7.043(OCR 誤讀)
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

# ---------- 轉正 + OCR ----------
def _score(text):
    return (len(_find_bc(text))*100 + len(DATE_D.findall(text))*10
            + (50 if ("電子發票" in text or "證明聯" in text) else 0))

def _orient_ocr(img, dpi_note=""):
    """OSD 快路徑；抓不到錨點才四角度擇優。回傳 (text, rot)。"""
    try:
        rot = pytesseract.image_to_osd(img, output_type=Output.DICT).get("rotate", 0)
    except Exception:
        rot = 0
    up = img if rot == 0 else img.rotate(rot, expand=True)
    txt = pytesseract.image_to_string(up, lang=LANG)
    # 乾淨發票或已抓到 BC → 直接採用
    if ("電子發票" in txt) or ("證明聯" in txt) or _find_bc(txt):
        return txt, rot
    # 密集對帳單 OSD 易判錯 → 四角度擇優
    best = (txt, _score(txt), rot)
    for ang in (90, 180, 270):
        t = img.rotate(ang, expand=True)
        tx = pytesseract.image_to_string(t, lang=LANG)
        s = _score(tx)
        if s > best[1]: best = (tx, s, ang)
    return best[0], best[2]

# ---------- 解析 ----------
def parse_invoice(pg, text):
    fb = _find_bc(text); bc = fb[0][1] if fb else None
    d = DATE_D.search(text); date = d.group(0).replace("-", "/") if d else ""
    plant = PLANT.search(text); plant = plant.group(1) if plant else ""
    price = None; conf = "高"
    pm = PRICE5.search(text)
    if pm: price = round(int(pm.group(1)) + int(pm.group(2))/100000, 4)
    nums = _amounts(text); net = None
    if pm:                                  # 主路徑：單價右側第一個千分位金額
        after = text[pm.end():]
        for nm in THOUS.finditer(after):
            if re.match(r"\s*/?\s*(PCS|FCS|P|F)", after[nm.end():nm.end()+5]):
                continue
            net = _to_int(nm.group()); break
    if net is None:
        tri = _find_triple(nums); net = tri[0] if tri else None
        if net is not None: conf = "中"
    if net and price and not _is_int(net/price):   # 候選多半是總計→回推
        net2 = round(net/1.05)
        if _is_int(net2/price): net = net2
    if net is None:                          # 退路：帶千分位金額
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

# ---------- 主流程 ----------
def process_pdf(pdf_bytes=None, pdf_path=None, dpi=250, debug=False):
    """回傳 (df_stmt, df_inv, df_recon, log)。pdf_bytes 或 pdf_path 擇一。"""
    log = []
    if pdf_bytes is not None:
        pages = convert_from_bytes(pdf_bytes, dpi=dpi)
    else:
        pages = convert_from_path(pdf_path, dpi=dpi)
    inv, stmt = [], []
    for i, img in enumerate(pages):
        pg = i + 1
        text, rot = _orient_ocr(img)
        flat = text.replace(" ", "")
        is_inv = ("電子發票" in flat) or ("證明聯" in flat) or ("銷售額" in flat)
        if is_inv:
            r = parse_invoice(pg, text); inv.append(r)
            if debug: log.append(f"p{pg} 發票 rot={rot} {r['發票號碼']} 總計={r['總計']} 信心={r['信心']}")
        else:
            rs = parse_statement(pg, text); stmt += rs
            if debug: log.append(f"p{pg} 對帳單 rot={rot} 列數={len(rs)} BC={[x['發票號碼'] for x in rs]}")
    df_inv = pd.DataFrame(inv)
    df_stmt = pd.DataFrame(stmt)
    df_recon = reconcile(df_stmt, df_inv)
    return df_stmt, df_inv, df_recon, log

def reconcile(df_stmt, df_inv):
    a = (df_stmt[["發票號碼", "總金額"]].rename(columns={"總金額": "對帳單總計(A)"})
         if len(df_stmt) else pd.DataFrame(columns=["發票號碼", "對帳單總計(A)"]))
    b = (df_inv[["發票號碼", "總計"]].rename(columns={"總計": "發票總計(B)"})
         if len(df_inv) else pd.DataFrame(columns=["發票號碼", "發票總計(B)"]))
    rec = pd.merge(a, b, on="發票號碼", how="outer")
    def status(r):
        x, y = r["對帳單總計(A)"], r["發票總計(B)"]
        if pd.isna(x) or pd.isna(y): return "⚠ 缺值待人工"
        return "✓ 一致" if abs(y-x) <= 2 else "❌ 差異待人工"
    rec["狀態"] = rec.apply(status, axis=1)
    rec["差異"] = (rec["發票總計(B)"].fillna(0) - rec["對帳單總計(A)"].fillna(0))
    return rec.sort_values("發票號碼").reset_index(drop=True)

if __name__ == "__main__":
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else "bill.pdf"
    s, v, r, log = process_pdf(pdf_path=path, dpi=250, debug=True)
    print("\n".join(log))
    print(f"\n發票 {len(v)} 張、對帳單 {len(s)} 列")
    print(r["狀態"].value_counts().to_string())
    print(r.to_string(index=False))
