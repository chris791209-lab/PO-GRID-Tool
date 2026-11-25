"""D53 PO GRID engine — PCN-anchored, one sheet per factory.

Data flow:  PCN (scope) -> SPS PO PDFs (quantities) -> WRK workspace (photos) -> GRID workbook.

Scope gates applied to every PO line:
  1. DPCI is in an uploaded PCN (minus anything removed in 進階篩選)
  2. Purpose code is not BAS, and the Ship Window starts on/after the season start date
  3. Latest SPS print of each PO wins; a PO with an "Import PO Change ... CANCEL" notice is excluded
Lines that fail a gate are kept in the 排除清單 sheet with the reason — nothing is dropped silently.

The generated workbook carries its own state in a hidden sheet (_D53_STATE): settings, PCN rows,
every parsed PO document and the change log.  Uploading that workbook again with new PO PDFs
(or another PCN) rebuilds every sheet from that state, keeping the yellow manual cells
(S/C, AC per PO; AE 備註 per DPCI).
"""
import collections
import datetime
import io
import json
import re
import zipfile

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL

STATE_SHEET = "_D53_STATE"
PHOTO_SHEET = "_D53_PHOTOS"
ENGINE_VERSION = "d53-1.0"
AE_LIST = ["Michelle", "Theresa", "Vanessa", "Chris", "Elaine"]


# =====================================================================  PDF parsing
def _pdf_text_pages(data):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        return [pg.extract_text() or "" for pg in pdf.pages]


# SPS repeats a column header block at every page break; it can land in the middle of a line item.
_PAGE_HDR = re.compile(
    r"https://commerce\.spscommerce\.com\S*\s+\d+/\d+\n\d{4}/\d\d/\d\d \d\d:\d\d Fulfillment\n"
    r"(?:(?!Order|Import)[^\n]*\n){0,9}?PRICE\n")
_DOC_SPLIT = re.compile(r"(?m)^(?=\d{4}/\d\d/\d\d \d\d:\d\d Fulfillment\n(?:Import PO Change|Order))")
_LINE = re.compile(r"^(\d+)\s+Buyers Catalog.*?Unit Price:\s*([\d,]+)\s*Each\s*([\d,]*\.\d\d)")


def parse_po_text(fname, full):
    """Return (docs, problems).  A doc is an ORDER or a CHANGE (Import PO Change) document."""
    docs, problems = [], []
    for d in _DOC_SPLIT.split(full):
        if not d.strip():
            continue
        is_chg = "Import PO Change" in d[:120]
        m = re.search(r"Order\s*#\s*:\s*(\d{11})-(\d{4})", d)
        if not m:
            continue
        pm = re.search(r"^(\d{4}/\d\d/\d\d \d\d:\d\d)", d)
        rec = dict(file=fname, po=m.group(1), dc=m.group(2), doc="CHANGE" if is_chg else "ORDER",
                   printed=pm.group(1) if pm else "")
        dm = re.search(r"(\d{7})(Lucky Star[^\n]*?)\s(\d{3})\s*$", d, re.M)
        rec["vendor"] = dm.group(1) if dm else None
        rec["dept"] = dm.group(3) if dm else None
        pd_ = (re.search(r"TARGET IMPORT\s*(?:\n)?\s*(\d\d/\d\d/\d{4})", d)
               or re.search(r"PO Date\s*:\s*Requested Delivery Date\s*:\s*\n(\d\d/\d\d/\d{4})", d))
        rec["po_date"] = pd_.group(1) if pd_ else None
        if is_chg:
            cm = re.search(r"Blanket Order \(Quantity Firm\)\s*-\s*(\w+)", d)
            rec["change"] = cm.group(1) if cm else None
            rec["lines"] = []
            docs.append(rec)
            continue
        sw = re.search(r"(\d\d/\d\d/\d{4})-(\d\d/\d\d/\d{4})", d)
        rec["sw"] = sw.group(0) if sw else None
        fac = re.search(r"Location ID : \d+ Location ID :\s*\nTARGET DC \d+ (.+)", d)
        rec["factory"] = fac.group(1).strip() if fac else None
        sp = re.search(r"Origin \(Shipping Point\)\s*(\S+)", d)
        rec["ship_pt"] = sp.group(1) if sp else None
        s = re.search(r"transaction set sender:\s*([A-Z]{3})", d)
        rec["purpose"] = s.group(1) if s else None
        tq = re.search(r"Total Qty\s*:\s*([\d,\.]+)", d)
        rec["total_qty"] = float(tq.group(1).replace(",", "")) if tq else None
        d2 = _PAGE_HDR.sub("", d)
        L = d2.split("\n")
        lines = []
        for i, ln in enumerate(L):
            mm = _LINE.search(ln)
            if not mm:
                continue
            nxt = L[i + 1] if i + 1 < len(L) else ""
            dp = re.search(r"^(?:Number:\s*)?(\d{9})\b", nxt)
            cost = re.search(r"(\d*\.\d+)\s*$", nxt)
            blk = "\n".join(L[i:i + 8])
            ret = None
            mres = re.search(r"Resale:(.*)", blk)
            if mres:
                for tok in re.findall(r"(?<![\d.])(\d{1,3}\.\d{2})(?![\d])", blk[mres.start() + 7:][:200]):
                    ret = float(tok)
                    break
            st = re.search(r"(?m)^(?:Number:\s*)?\d{7,10}\s+(.+?)\s+(?:#\s*of|SET|Weight)", blk)
            lines.append(dict(line=int(mm.group(1)), dpci=dp.group(1) if dp else None,
                              qty=int(mm.group(2).replace(",", "")),
                              cost=float(cost.group(1)) if cost else None, retail=ret,
                              style=st.group(1) if st else None))
        rec["lines"] = lines
        lq = sum(x["qty"] for x in lines)
        if rec["total_qty"] is not None and abs(lq - rec["total_qty"]) > 0.5:
            problems.append(f"{fname}：PO {rec['po']} 行數量合計 {lq:,} ≠ Total Qty {rec['total_qty']:,.0f}（版型可能改了，請回報）")
        if any(x["dpci"] is None for x in lines):
            problems.append(f"{fname}：PO {rec['po']} 有品項抓不到 DPCI")
        docs.append(rec)
    named = re.search(r"(100\d{8})", fname)
    if named and docs and named.group(1) not in {x["po"] for x in docs}:
        problems.append(f"檔名與內容不符：{fname} 檔名是 {named.group(1)}，內容是 "
                        f"{'、'.join(sorted({x['po'] for x in docs}))} → {named.group(1)} 本身可能不在這批檔案")
    if not docs:
        problems.append(f"{fname}：找不到任何 Order #（可能是掃描檔或不是 SPS PO）")
    return docs, problems


def parse_po_pdf(fname, data):
    return parse_po_text(fname, "\n".join(_pdf_text_pages(data)))


def iter_uploaded_pdfs(files):
    """files: list of (name, bytes).  Expands .zip uploads."""
    for name, data in files:
        if name.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for info in z.infolist():
                    if info.is_dir() or not info.filename.lower().endswith(".pdf"):
                        continue
                    n = info.filename
                    for enc in ("cp950", "utf-8", "cp936"):
                        try:
                            n = info.filename.encode("cp437").decode(enc)
                            break
                        except (UnicodeEncodeError, UnicodeDecodeError):
                            continue
                    yield n.split("/")[-1], z.read(info)
        elif name.lower().endswith(".pdf"):
            yield name, data


# =====================================================================  PCN / WRK
def program_name(fname):
    n = re.sub(r"\.(xlsx|xlsm|xls)$", "", fname, flags=re.I)
    n = re.sub(r"^\d{4}\s*C\d\s+D\d+\s+", "", n.strip())
    return re.sub(r"\s+", " ", n).strip() or fname


def load_pcn(fname, data):
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    ws = wb["Data"] if "Data" in wb.sheetnames else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    hdr_i = next((i for i, r in enumerate(rows[:10]) if r and "DPCI" in [str(c).strip() if c else "" for c in r]), None)
    if hdr_i is None:
        raise ValueError(f"{fname}：找不到 DPCI 欄位，請確認是 PCN 主檔")
    h = [str(c).strip() if c else "" for c in rows[hdr_i]]
    prog = program_name(fname)
    out, nodpci = {}, collections.Counter()

    def g(r, *ks):
        for k in ks:
            if k in r and r[k] not in (None, ""):
                return r[k]
        return None

    def num(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return None

    for raw in rows[hdr_i + 1:]:
        r = dict(zip(h, raw))
        if not any(v not in (None, "") for v in raw):
            continue
        status = g(r, "Item Plan Status (Beginning)") or ""
        dpci = g(r, "DPCI")
        if not dpci:
            nodpci[status or "無狀態"] += 1
            continue
        dpci = str(dpci).strip()
        cq, iq = num(g(r, "Case Unit Quantity")), num(g(r, "Inner Pack Unit Quantity"))
        out[dpci] = dict(
            dpci=dpci, program=prog, desc=str(g(r, "Product Description") or ""),
            style=str(g(r, "Manufacturer Style #", "Manufacturer Style # *") or ""),
            commit=round(num(g(r, "Ent Ttl Rcpt U")) or 0),
            cost=num(g(r, "FCA Factory City Unit Cost", "FOB Unit Cost")),
            retail=num(g(r, "Suggested Unit Retail")),
            pack=f"{int(cq)}/{int(iq)}" if cq and iq else "",
            factory=str(g(r, "Factory Name *", "Factory Name") or ""), status=status)
    return dict(program=prog, file=fname, rows=out, nodpci=dict(nodpci), nrows=len(rows) - hdr_i - 1)


def extract_wrk_photos(fname, data, thumb=240):
    """Map It workspace export: images anchored on the Products sheet, DPCI at the start of column G."""
    from lxml import etree
    from PIL import Image
    NS = {"xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
          "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    z = zipfile.ZipFile(io.BytesIO(data))
    names = set(z.namelist())

    def rels(p):
        rp = p.rsplit("/", 1)[0] + "/_rels/" + p.rsplit("/", 1)[1] + ".rels"
        if rp not in names:
            return {}
        t = etree.fromstring(z.read(rp))
        base = p.rsplit("/", 1)[0]
        out = {}
        for e in t:
            tgt = e.get("Target")
            parts = (base + "/" + tgt).split("/") if not tgt.startswith("/") else tgt.lstrip("/").split("/")
            st = []
            for x in parts:
                if x == "..":
                    st.pop()
                elif x not in ("", "."):
                    st.append(x)
            out[e.get("Id")] = "/".join(st)
        return out

    ss = []
    if "xl/sharedStrings.xml" in names:
        t = etree.fromstring(z.read("xl/sharedStrings.xml"))
        ss = ["".join(si.itertext()) for si in t.findall("m:si", NS)]
    wbx = etree.fromstring(z.read("xl/workbook.xml"))
    wr = rels("xl/workbook.xml")
    photos = {}
    for sh in wbx.find("m:sheets", NS):
        sp = wr[sh.get("{%s}id" % NS["r"])]
        dr = [v for v in rels(sp).values() if "drawing" in v]
        if not dr:
            continue
        sheet = etree.fromstring(z.read(sp))
        gcol = {}
        for c in sheet.iter("{%s}c" % NS["m"]):
            ref = c.get("r")
            if not ref.startswith("G") or not ref[1:].isdigit():
                continue
            v, is_ = c.find("m:v", NS), c.find("m:is", NS)
            val = ss[int(v.text)] if c.get("t") == "s" and v is not None else (
                "".join(is_.itertext()) if is_ is not None else (v.text if v is not None else ""))
            gcol[int(ref[1:])] = val or ""
        drr = rels(dr[0])
        for a in etree.fromstring(z.read(dr[0])):
            fr = a.find("xdr:from", NS)
            blip = a.find(".//a:blip", NS)
            if fr is None or blip is None:
                continue
            m = re.match(r"(\d{3})-(\d{2})-(\d{4})", gcol.get(int(fr.find("xdr:row", NS).text) + 1, ""))
            tgt = drr.get(blip.get("{%s}embed" % NS["r"]))
            if not m or not tgt or tgt not in names:
                continue
            key = "-".join(m.groups())
            if key in photos:
                continue
            try:
                im = Image.open(io.BytesIO(z.read(tgt))).convert("RGBA")
                bg = Image.new("RGB", im.size, "white")
                bg.paste(im, mask=im.split()[3])
                bg.thumbnail((thumb, thumb))
                b = io.BytesIO()
                bg.save(b, "JPEG", quality=85)
                photos[key] = b.getvalue()
            except Exception:
                continue
    return photos


# =====================================================================  scope
def fmt_dpci(s):
    return f"{s[:3]}-{s[3:5]}-{s[5:]}" if s and len(s) == 9 else (s or "")


def _swstart(sw):
    m = re.match(r"(\d\d)/(\d\d)/(\d{4})", sw or "")
    return datetime.date(int(m[3]), int(m[1]), int(m[2])) if m else None


def effective_pcn(state):
    excl_f = set(state["settings"].get("excluded_factories", []))
    excl_d = set(state["settings"].get("excluded_dpci", []))
    out = {}
    for p in state["pcn"].values():
        for d, r in p["rows"].items():
            if d in excl_d or r["factory"] in excl_f:
                continue
            out[d] = r
    return out


def compute(state):
    """Apply the three gates to the stored PO documents."""
    pcn = effective_pcn(state)
    season = datetime.date.fromisoformat(state["settings"]["season_start"])
    docs = state["docs"]
    cancel = {}
    for r in docs:
        if r["doc"] == "CHANGE" and (r.get("change") or "").upper().startswith("CANCEL"):
            cancel.setdefault(r["po"], r)
    orders = {}
    for r in sorted([r for r in docs if r["doc"] == "ORDER"], key=lambda r: r["printed"] or ""):
        orders[r["po"]] = r
    pcn_fac = sorted({v["factory"] for v in pcn.values() if v["factory"]})

    def fullfac(s):
        s0 = (s or "").lower()[:18]
        for f in pcn_fac:
            if s0 and f.lower().startswith(s0):
                return f
        return s or "（工廠不明）"

    inscope, excl = [], []
    for r in orders.values():
        ss = _swstart(r.get("sw"))
        for l in r["lines"]:
            dp = fmt_dpci(l["dpci"])
            if r["po"] in cancel:
                reason = "已取消（Import PO Change CANCEL）"
            elif dp not in pcn:
                reason = "DPCI 不在本次的 PCN" + ("" if r.get("dept") in (None, "053") else f"（Dept {r['dept']}）")
            elif r.get("purpose") == "BAS":
                reason = "Purpose BAS（Inline 補貨，非本季）"
            elif ss and ss < season:
                reason = f"Ship Window 早於本季起始日 {season}"
            else:
                reason = None
            rec = dict(po=r["po"], dc=r["dc"], po_date=r.get("po_date"), printed=r.get("printed"),
                       purpose=r.get("purpose"), sw=r.get("sw"), ship_pt=r.get("ship_pt"),
                       factory=fullfac(r.get("factory")), vendor=r.get("vendor"), dept=r.get("dept"),
                       line=l["line"], dpci=dp, qty=l["qty"], cost=l["cost"], retail=l["retail"],
                       file=r["file"], reason=reason)
            (excl if reason else inscope).append(rec)

    def sig(r):
        return (r.get("factory"), r["dc"], tuple(sorted((l["dpci"], l["qty"]) for l in r["lines"])))
    act = collections.defaultdict(list)
    for r in orders.values():
        if r["po"] not in cancel:
            act[sig(r)].append(r["po"])
    reissue = {p: act.get(sig(orders[p]), []) for p in cancel if p in orders}
    return dict(pcn=pcn, orders=orders, cancel=cancel, inscope=inscope, excl=excl, reissue=reissue)


# =====================================================================  workbook
F = lambda **k: Font(name="Arial", size=k.pop("size", 10), **k)  # noqa: E731
HDR = PatternFill("solid", fgColor="1F4E78"); SUB = PatternFill("solid", fgColor="D9E1F2")
YEL = PatternFill("solid", fgColor="FFF2CC"); RED = PatternFill("solid", fgColor="F8CBAD")
GRN = PatternFill("solid", fgColor="E2EFDA"); BLU = PatternFill("solid", fgColor="BDD7EE")
_th = Side(style="thin", color="BFBFBF"); BOX = Border(top=_th, bottom=_th, left=_th, right=_th)
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LW = Alignment(vertical="center", wrap_text=True)
FIXH = ["STYLE", "DPCI", "品名", "照片", "Program", "Case/\nInner", "PCN 成本", "PCN 零售", "Commit\n(Ent Ttl Rcpt U)"]
FIX = len(FIXH); P0 = FIX + 1; HR = 11
RESERVED = {"使用說明", "Validation Summary", "Program 總覽", "Cost Discrepancy", "排除清單", "PO Ledger", "更新紀錄", STATE_SHEET, PHOTO_SHEET}


def _short(f):
    w = re.sub(r"[^\w& ]", " ", f).split()
    return (" ".join(w[:2]) or "Factory")[:31]


def render(state, photos=None, manual=None):
    """Build the workbook; returns (bytes, stats)."""
    photos = photos or {}
    manual = manual or {"po": {}, "note": {}}
    res = compute(state)
    pcn, inscope, excl = res["pcn"], res["inscope"], res["excl"]
    S = state["settings"]
    ae = S.get("ae") or "（未指定 AE）"
    today = datetime.date.today().isoformat()
    season = S["season_start"]
    costdiff = [x for x in inscope if pcn[x["dpci"]]["cost"] is not None and x["cost"] is not None
                and abs(x["cost"] - pcn[x["dpci"]]["cost"]) > 0.005]
    bad = {(x["dpci"], x["po"]) for x in costdiff}
    retdiff = {x["dpci"]: (x["retail"], pcn[x["dpci"]]["retail"]) for x in inscope
               if x["retail"] is not None and pcn[x["dpci"]]["retail"] is not None
               and abs(x["retail"] - pcn[x["dpci"]]["retail"]) > 0.005}
    byfac = collections.defaultdict(list)
    for x in inscope:
        byfac[x["factory"]].append(x)
    facs = sorted(byfac, key=lambda f: -sum(x["qty"] for x in byfac[f]))
    names = {}
    for f in facs:
        n, i = _short(f), 2
        while n in names.values() or n in RESERVED:
            n = f"{_short(f)[:28]}-{i}"; i += 1
        names[f] = n

    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "使用說明"
    for t, b in [(f"D53 PO GRID｜{ae}", 2), (f"產生 {today}｜資料流：PCN（範圍）→ PO（數量）→ WRK（照片）", 0), ("", 0),
                 ("範圍怎麼決定", 1),
                 ("1. 品項：只看上傳的 PCN（每份 PCN＝一個 Program）。PCN 跟別人共用時，可在 App 的「進階篩選」排除不屬於自己的工廠／品項。", 0),
                 (f"2. 季別：Purpose 不是 BAS，且 Ship Window 開始日 ≥ 本季起始日（本檔：{season}）。", 0),
                 ("3. 版本：同一張 PO 取 SPS 列印時間最新的一份；收到 Import PO Change CANCEL 即整張排除。", 0),
                 ("不符合的 PO 行全部列在「排除清單」並寫明原因。", 0), ("", 0), ("怎麼用", 1),
                 ("• 工廠分頁：第 3–10 列是 PO 表頭（S/C、AC 黃底手填）；品項依 Program 分區塊；最右側是合計、% vs Commit、核對、AE 備註（黃底）。", 0),
                 ("• 紅底數量格＝該 PO 單價 ≠ PCN 成本（PO 為準，只標記）。英文明細在「Cost Discrepancy」。", 0),
                 ("• 之後有新 PO：到 App 選「更新現有 GRID」，上傳這份檔案＋新的 PO PDF 即可；要加 Program 就一起上傳新的 PCN。", 0),
                 ("• 只有黃底欄位會在更新時保留；請不要在其他地方手寫資料，也不要改分頁名稱或刪除隱藏頁。", 0)]:
        ws.append([t]); c = ws.cell(ws.max_row, 1)
        c.font = F(bold=True, size=14) if b == 2 else (F(bold=True) if b == 1 else F())
    ws.column_dimensions["A"].width = 130

    overview = collections.defaultdict(lambda: dict(items=set(), pos=set(), units=0, cost=0.0, diff=0))
    sheetinfo = {}
    for f in facs:
        xs = byfac[f]; ws = wb.create_sheet(names[f])
        pos = {}
        for x in xs:
            pos.setdefault(x["po"], x)
        pos = sorted(pos.values(), key=lambda x: (_swstart(x["sw"]) or datetime.date.max, x["po"]))
        PN = P0 + len(pos) - 1; T = PN + 1
        items = sorted({x["dpci"] for x in xs}, key=lambda d: (pcn[d]["program"], pcn[d]["style"]))
        q = collections.defaultdict(int)
        for x in xs:
            q[(x["dpci"], x["po"])] += x["qty"]
        ws["A1"] = f"D53 PO GRID｜{f}"; ws["A1"].font = F(bold=True, size=14)
        ws["A2"] = (f"{ae}｜產生 {today}｜有效 PO {len(pos)} 張｜品項 {len(items)}｜Program："
                    f"{'、'.join(sorted({pcn[d]['program'] for d in items}))}｜紅底＝PO 單價 ≠ PCN 成本")
        ws["A2"].font = F(color="595959")
        labels = ["PO#-DC", "PO Date", "Purpose", "Ship Point", "Ship Window", "Vendor #", "S/C（手填）", "AC（手填）"]
        for i, lab in enumerate(labels):
            rr = 3 + i; c = ws.cell(rr, FIX, lab); c.font = F(bold=True); c.fill = SUB; c.alignment = CEN; c.border = BOX
            for j, p in enumerate(pos):
                mp = manual["po"].get(p["po"], {})
                v = [f"{p['po']}-{p['dc']}", p["po_date"], p["purpose"], p["ship_pt"],
                     re.sub(r"/20(\d\d)", "", p["sw"] or ""), p["vendor"], mp.get("sc"), mp.get("ac")][i]
                cc = ws.cell(rr, P0 + j, v); cc.alignment = CEN; cc.border = BOX; cc.font = F(size=9, bold=(i == 0))
                if i >= 6:
                    cc.fill = YEL
        for i, h in enumerate(FIXH, 1):
            ws.cell(HR, i, h)
        for j, p in enumerate(pos):
            ws.cell(HR, P0 + j, p["dc"])
        for k, h in enumerate(["PO 合計", "% vs Commit", "未下單量", "核對", "AE 備註（自由欄）"]):
            ws.cell(HR, T + k, h)
        for c in ws[HR]:
            if c.value is not None:
                c.font = F(bold=True, color="FFFFFF"); c.fill = HDR; c.alignment = CEN; c.border = BOX
        ws.row_dimensions[HR].height = 30
        r = HR; prevp = None; R0 = HR + 1
        for d in items:
            p = pcn[d]
            if p["program"] != prevp:
                r += 1; prevp = p["program"]
                ws.cell(r, 1, f"▌Program：{prevp}　｜品項 {sum(1 for z in items if pcn[z]['program'] == prevp)}")
                for cc in range(1, T + 5):
                    ws.cell(r, cc).fill = BLU; ws.cell(r, cc).border = BOX
                ws.cell(r, 1).font = F(bold=True); ws.row_dimensions[r].height = 20
            r += 1
            for i, v in enumerate([p["style"], d, p["desc"], None, p["program"], p["pack"], p["cost"], p["retail"], p["commit"]], 1):
                c = ws.cell(r, i, v); c.border = BOX; c.alignment = LW if i in (1, 3, 5) else CEN; c.font = F(size=9)
            ws.cell(r, 7).number_format = ws.cell(r, 8).number_format = "0.00"; ws.cell(r, 9).number_format = "#,##0"
            for j, po in enumerate(pos):
                c = ws.cell(r, P0 + j, q.get((d, po["po"])))
                c.border = BOX; c.number_format = "#,##0"; c.font = F(size=9); c.alignment = CEN
                if (d, po["po"]) in bad:
                    c.fill = RED
            ws.cell(r, T, f"=SUM({CL(P0)}{r}:{CL(PN)}{r})").number_format = "#,##0"
            ws.cell(r, T + 1, f'=IF({CL(FIX)}{r}>0,{CL(T)}{r}/{CL(FIX)}{r},"")').number_format = "0.0%"
            ws.cell(r, T + 2, f'=IF({CL(FIX)}{r}>0,{CL(FIX)}{r}-{CL(T)}{r},"")').number_format = "#,##0"
            fl = []
            cd = sorted({z["cost"] for z in costdiff if z["dpci"] == d and z["factory"] == f})
            if cd:
                fl.append(f"成本 PO {'/'.join(map(str, cd))} vs PCN {p['cost']}")
            if p["cost"] is None:
                fl.append("PCN 無成本")
            if d in retdiff:
                fl.append(f"零售 PO {retdiff[d][0]} vs PCN {retdiff[d][1]:g}")
            if d not in photos:
                fl.append("缺照片")
            c = ws.cell(r, T + 3, "OK" if not fl else "；".join(fl)); c.alignment = LW
            c.font = F(size=9, color="006100" if not fl else "9C0006"); c.fill = GRN if not fl else RED
            ws.cell(r, T + 4, manual["note"].get(d))
            for k in range(T, T + 5):
                ws.cell(r, k).border = BOX
            for k in range(T, T + 3):
                ws.cell(r, k).alignment = CEN; ws.cell(r, k).font = F(size=9)
            ws.cell(r, T + 4).fill = YEL; ws.cell(r, T + 4).alignment = LW
            ws.row_dimensions[r].height = 62
            if d in photos:
                im = XLImage(io.BytesIO(photos[d]))
                s = 76 / max(im.width, im.height); im.width, im.height = int(im.width * s), int(im.height * s)
                im.anchor = TwoCellAnchor(editAs="twoCell",
                                          _from=AnchorMarker(col=3, colOff=19050, row=r - 1, rowOff=19050),
                                          to=AnchorMarker(col=3, colOff=19050 + im.width * 9525, row=r - 1, rowOff=19050 + im.height * 9525))
                ws.add_image(im)
            o = overview[(p["program"], f)]
            o["items"].add(d)
            o["units"] += sum(v for (dd, _), v in q.items() if dd == d)
            o["cost"] += sum(v * (p["cost"] or 0) for (dd, _), v in q.items() if dd == d)
            o["pos"] |= {po["po"] for po in pos if q.get((d, po["po"]))}
            o["diff"] += sum(1 for z in costdiff if z["dpci"] == d and z["factory"] == f)
        RN = r; t1, t2 = RN + 1, RN + 2
        ws.cell(t1, 3, "Total Units"); ws.cell(t2, 3, "Total 成本 $（PCN 成本 × 數量）")
        for j in range(len(pos) + 1):
            col = CL(P0 + j) if j < len(pos) else CL(T)
            ws[f"{col}{t1}"] = f"=SUM({col}{R0}:{col}{RN})"; ws[f"{col}{t1}"].number_format = "#,##0"
            ws[f"{col}{t2}"] = f"=SUMPRODUCT({col}{R0}:{col}{RN},$G${R0}:$G${RN})"; ws[f"{col}{t2}"].number_format = "#,##0.00"
        for rr in (t1, t2):
            for cc in range(1, T + 5):
                c = ws.cell(rr, cc); c.font = F(bold=True, size=9); c.fill = SUB; c.border = BOX
        for i, w in enumerate([13, 12, 30, 12, 24, 8, 8, 8, 14], 1):
            ws.column_dimensions[CL(i)].width = w
        for j in range(len(pos)):
            ws.column_dimensions[CL(P0 + j)].width = 13
        for k, w in enumerate([11, 10, 11, 30, 22]):
            ws.column_dimensions[CL(T + k)].width = w
        for i in range(3, 11):
            ws.row_dimensions[i].height = 24
        ws.freeze_panes = ws.cell(R0, P0)
        ws.auto_filter.ref = f"A{HR}:{CL(T + 4)}{RN}"
        sheetinfo[f] = (names[f], len(items), len(pos), sum(x["qty"] for x in xs))

    # ---- Program 總覽
    ov = wb.create_sheet("Program 總覽")
    ov.append([f"Program 總覽｜{ae}｜{today}"]); ov["A1"].font = F(bold=True, size=14); ov.append([])
    ov.append(["Program（PCN）", "工廠", "工廠分頁", "品項", "有效 PO", "數量", "Commit", "% vs Commit", "成本 $（PCN）", "成本差異行"])
    for c in ov[3]:
        c.font = F(bold=True, color="FFFFFF"); c.fill = HDR; c.alignment = CEN; c.border = BOX
    for pg in sorted({k[0] for k in overview}):
        rs = ov.max_row + 1
        for (p, f), o in sorted(overview.items(), key=lambda kv: -kv[1]["units"]):
            if p != pg:
                continue
            ov.append([p, f, names[f], len(o["items"]), len(o["pos"]), o["units"],
                       sum(pcn[d]["commit"] for d in o["items"]), None, round(o["cost"], 2), o["diff"]])
            rr = ov.max_row
            ov.cell(rr, 8, f'=IF(G{rr}>0,F{rr}/G{rr},"")')
            ov.cell(rr, 3).hyperlink = f"#'{names[f]}'!A1"; ov.cell(rr, 3).font = F(color="0563C1", underline="single")
            if o["diff"]:
                ov.cell(rr, 10).fill = RED
        re_ = ov.max_row
        ov.append([f"{pg} 小計", None, None, f"=SUM(D{rs}:D{re_})", None, f"=SUM(F{rs}:F{re_})",
                   f"=SUM(G{rs}:G{re_})", None, f"=SUM(I{rs}:I{re_})", f"=SUM(J{rs}:J{re_})"])
        rr = ov.max_row; ov.cell(rr, 8, f'=IF(G{rr}>0,F{rr}/G{rr},"")')
        for c in ov[rr]:
            c.font = F(bold=True); c.fill = SUB
    for row in ov.iter_rows(min_row=4):
        for c in row:
            c.border = BOX
        row[5].number_format = row[6].number_format = "#,##0"; row[7].number_format = "0.0%"; row[8].number_format = "#,##0.00"
    for i, w in enumerate([30, 40, 22, 7, 8, 12, 12, 11, 14, 10], 1):
        ov.column_dimensions[CL(i)].width = w
    ov.freeze_panes = "A4"

    # ---- Cost Discrepancy (English, client-facing)
    cd = wb.create_sheet("Cost Discrepancy")
    cd.append(["Lucky Star – D53 PO Unit Cost vs PCN Cost"]); cd["A1"].font = F(bold=True, size=13)
    cd.append([f"Prepared {today}. PO unit cost differs from the PCN FCA cost. Please review and revise the PO unit cost where applicable."])
    cd.append([])
    cd.append(["PO #", "DC", "PO Date", "Ship Window", "Factory", "DPCI", "Vendor Style #", "Item Description",
               "PO Qty", "PO Unit Cost", "PCN Unit Cost", "Diff / Unit", "Extended Diff (USD)"])
    for c in cd[4]:
        c.font = F(bold=True, color="FFFFFF"); c.fill = HDR; c.alignment = CEN; c.border = BOX
    for x in sorted(costdiff, key=lambda x: (x["po"], x["dpci"])):
        p = pcn[x["dpci"]]; rr = cd.max_row + 1
        cd.append([x["po"], x["dc"], x["po_date"], x["sw"], x["factory"], x["dpci"], p["style"], p["desc"],
                   x["qty"], x["cost"], p["cost"], f"=J{rr}-K{rr}", f"=I{rr}*L{rr}"])
        for c in cd[rr]:
            c.border = BOX; c.font = F(size=9)
        cd[f"I{rr}"].number_format = "#,##0"; cd[f"M{rr}"].number_format = "#,##0.00"
        for k in "JKL":
            cd[f"{k}{rr}"].number_format = "0.000"
    if costdiff:
        rl = cd.max_row
        cd.append(["Total"] + [None] * 7 + [f"=SUM(I5:I{rl})"] + [None] * 3 + [f"=SUM(M5:M{rl})"])
        for c in cd[cd.max_row]:
            c.font = F(bold=True); c.fill = SUB; c.border = BOX
        cd[f"I{cd.max_row}"].number_format = "#,##0"; cd[f"M{cd.max_row}"].number_format = "#,##0.00"
    for i, w in enumerate([13, 6, 11, 22, 30, 12, 14, 38, 10, 11, 11, 10, 14], 1):
        cd.column_dimensions[CL(i)].width = w
    cd.freeze_panes = "A5"

    # ---- 排除清單
    ex = wb.create_sheet("排除清單")
    ex.append(["原因", "PO#", "DC", "Dept", "Purpose", "PO Date", "Ship Window", "工廠", "Line", "DPCI", "數量", "來源檔案"])
    for c in ex[1]:
        c.font = F(bold=True, color="FFFFFF"); c.fill = HDR; c.alignment = CEN
    for x in sorted(excl, key=lambda x: (x["reason"], x["po"], x["line"])):
        ex.append([x["reason"], x["po"], x["dc"], x["dept"], x["purpose"], x["po_date"], x["sw"], x["factory"],
                   x["line"], x["dpci"], x["qty"], x["file"]])
    for i, w in enumerate([34, 13, 6, 6, 8, 11, 22, 30, 5, 12, 9, 40], 1):
        ex.column_dimensions[CL(i)].width = w
    ex.freeze_panes = "A2"; ex.auto_filter.ref = ex.dimensions

    # ---- PO Ledger
    lg = wb.create_sheet("PO Ledger")
    lg.append(["PO#", "DC", "PO Date", "SPS 列印時間", "Purpose", "Ship Window", "Ship Point", "工廠", "Vendor #", "Line",
               "DPCI", "STYLE", "品名", "Program", "數量", "PO 單價", "PCN 成本", "PO 零售", "PCN 零售", "來源檔案"])
    for c in lg[1]:
        c.font = F(bold=True, color="FFFFFF"); c.fill = HDR; c.alignment = CEN
    for x in sorted(inscope, key=lambda x: (x["factory"], x["po"], x["line"])):
        p = pcn[x["dpci"]]
        lg.append([x["po"], x["dc"], x["po_date"], x["printed"], x["purpose"], x["sw"], x["ship_pt"], x["factory"],
                   x["vendor"], x["line"], x["dpci"], p["style"], p["desc"], p["program"], x["qty"], x["cost"],
                   p["cost"], x["retail"], p["retail"], x["file"]])
    for i, w in enumerate([13, 6, 11, 16, 8, 22, 9, 30, 10, 5, 12, 14, 34, 26, 9, 8, 8, 8, 8, 34], 1):
        lg.column_dimensions[CL(i)].width = w
    lg.freeze_panes = "A2"; lg.auto_filter.ref = lg.dimensions

    # ---- 更新紀錄
    cl = wb.create_sheet("更新紀錄")
    cl.append(["時間", "動作", "PO#", "DPCI", "工廠", "原數量", "新數量", "說明"])
    for c in cl[1]:
        c.font = F(bold=True, color="FFFFFF"); c.fill = HDR; c.alignment = CEN
    for e in reversed(state.get("changelog", [])):
        cl.append([e.get("time"), e.get("action"), e.get("po"), e.get("dpci"), e.get("factory"),
                   e.get("old"), e.get("new"), e.get("note")])
    for i, w in enumerate([17, 14, 13, 12, 30, 10, 10, 60], 1):
        cl.column_dimensions[CL(i)].width = w
    cl.freeze_panes = "A2"

    # ---- Validation Summary
    vs = wb.create_sheet("Validation Summary")

    def Hh(t):
        vs.append([t]); vs.cell(vs.max_row, 1).font = F(bold=True, size=12, color="1F4E78")

    def row(v, fill=None, bold=False):
        vs.append(v)
        for c in vs[vs.max_row]:
            if c.value is not None:
                c.font = F(bold=bold); c.alignment = LW
                if fill:
                    c.fill = fill
    vs.append([f"D53 PO GRID｜{ae}｜Validation Summary｜{today}"]); vs["A1"].font = F(bold=True, size=14); vs.append([])
    Hh("1. 範圍")
    row(["項目", "數量", "說明"], SUB, True)
    for p in state["pcn"].values():
        row([f"PCN：{p['program']}", f"{len(p['rows'])} 個 DPCI", f"{p['file']}（{p['nrows']} 列）"])
        if p.get("nodpci"):
            row(["　└ 沒有 DPCI 的列", sum(p["nodpci"].values()),
                 "無法和 PO 比對：" + "、".join(f"{k} {v}" for k, v in p["nodpci"].items())], YEL)
    nx = len(S.get("excluded_factories", [])) + len(S.get("excluded_dpci", []))
    row(["進階篩選", "全選" if not nx else f"排除 {len(S.get('excluded_factories', []))} 家工廠、{len(S.get('excluded_dpci', []))} 個 DPCI", ""])
    row(["本季起始日", season, "Ship Window 開始日早於此日的 PO 行排除"])
    row(["PO 文件（取最新版）", f"{len(res['orders'])} 張", f"另有 {len(res['cancel'])} 張收到 CANCEL"])
    row(["納入本 GRID", f"{len({x['po'] for x in inscope})} 張 PO",
         f"{len(inscope)} 行、{sum(x['qty'] for x in inscope):,} pcs、{len({x['dpci'] for x in inscope})} 個 DPCI、{len(facs)} 家工廠"], GRN, True)
    rc = collections.Counter(x["reason"].split("（Dept")[0] for x in excl)
    for k, v in rc.most_common():
        row([f"排除：{k}", f"{v} 行", f"{len({x['po'] for x in excl if x['reason'].startswith(k)})} 張 PO，明細見「排除清單」"])
    probs = state.get("problems", [])
    if probs:
        vs.append([]); Hh("⚠ 解析提醒")
        for p in probs:
            row([p], RED)
    vs.append([]); Hh("2. 工廠分頁")
    row(["工廠", "分頁", "品項", "有效 PO", "數量"], SUB, True)
    for f in facs:
        row([f, *sheetinfo[f]])
    vs.append([]); Hh("3. 成本差異：PO 單價 ≠ PCN 成本（PO 為準，只標記）")
    hi = [x for x in costdiff if x["cost"] > pcn[x["dpci"]]["cost"]]
    lo = [x for x in costdiff if x["cost"] < pcn[x["dpci"]]["cost"]]
    row(["PO 高於 PCN", f"{len(hi)} 行", f"{len({x['dpci'] for x in hi})} 個 DPCI、{len({x['po'] for x in hi})} 張 PO；合計 USD "
         f"{sum((x['cost'] - pcn[x['dpci']]['cost']) * x['qty'] for x in hi):,.2f}"], RED if hi else None)
    row(["PO 低於 PCN（對我們不利）", f"{len(lo)} 行",
         "；".join(sorted({f"{x['dpci']} PO {x['cost']} vs PCN {pcn[x['dpci']]['cost']}" for x in lo})) or "無"], RED if lo else None)
    if retdiff:
        row(["零售價差異", f"{len(retdiff)} 個 DPCI", "；".join(f"{k} PO {a} vs PCN {b:g}" for k, (a, b) in retdiff.items())])
    vs.append([]); Hh("4. 已取消 PO → 重開對照")
    row(["已取消 PO", "工廠", "原 Ship Window", "重開為"], SUB, True)
    for p, new in sorted(res["reissue"].items(), key=lambda kv: res["orders"][kv[0]].get("factory") or ""):
        o = res["orders"][p]
        if not any(fmt_dpci(l["dpci"]) in pcn for l in o["lines"]):
            continue
        row([f"{p}-{o['dc']}", o.get("factory"), o.get("sw"), ", ".join(new) or "找不到相同內容的新單"], None if new else RED)
    vs.append([]); Hh("5. PCN 有、但還沒收到本季 PO 的品項（Commit > 0，排除 No Longer Under Consideration）")
    row(["DPCI", "STYLE", "品名", "Program", "Commit", "Plan Status", "PCN 工廠"], SUB, True)
    got = {x["dpci"] for x in inscope}
    for d, p in sorted(pcn.items(), key=lambda kv: (kv[1]["program"], kv[0])):
        if d in got or p["commit"] <= 0 or p["status"] == "No Longer Under Consideration":
            continue
        row([d, p["style"], p["desc"], p["program"], p["commit"], p["status"], p["factory"]])
    vs.append([]); Hh("6. 資料缺口")
    noph = sorted({x["dpci"] for x in inscope if x["dpci"] not in photos})
    row(["缺照片", f"{len(noph)} 個 DPCI", "上傳對應的 Map It Workspace 匯出檔（WRK）後重新更新即可"], YEL if noph else None)
    for i, w in enumerate([34, 26, 40, 22, 12, 16, 30], 1):
        vs.column_dimensions[CL(i)].width = w

    # ---- photo store (hidden): every WRK photo for a PCN DPCI, so later updates keep photos
    #      for items that have no PO yet
    ps = wb.create_sheet(PHOTO_SHEET)
    ps.append(["DPCI", "photo"])
    keep = sorted(d for d in photos if d in pcn)
    for i, d in enumerate(keep, start=2):
        ps.cell(i, 1, d)
        im = XLImage(io.BytesIO(photos[d])); s_ = 40 / max(im.width, im.height)
        im.width, im.height = int(im.width * s_), int(im.height * s_)
        im.anchor = TwoCellAnchor(editAs="twoCell", _from=AnchorMarker(col=1, row=i - 1),
                                  to=AnchorMarker(col=1, colOff=im.width * 9525, row=i - 1, rowOff=im.height * 9525))
        ps.add_image(im)
    ps.sheet_state = "hidden"

    # ---- state (hidden)
    st = wb.create_sheet(STATE_SHEET)
    st.append(["key", "part", "json"])
    blob = dict(engine=ENGINE_VERSION, settings=state["settings"], pcn=state["pcn"], docs=state["docs"],
                changelog=state.get("changelog", []), problems=state.get("problems", []))
    for key, val in blob.items():
        s = json.dumps(val, ensure_ascii=False, separators=(",", ":"))
        for i in range(0, max(len(s), 1), 30000):
            st.append([key, i // 30000, s[i:i + 30000]])
    st.sheet_state = "hidden"

    order = ["使用說明", "Validation Summary", "Program 總覽"] + [names[f] for f in facs] + \
            ["Cost Discrepancy", "排除清單", "PO Ledger", "更新紀錄", PHOTO_SHEET, STATE_SHEET]
    wb._sheets = [wb[n] for n in order]
    wb.active = 1
    out = io.BytesIO(); wb.save(out)
    stats = dict(pos=len({x["po"] for x in inscope}), lines=len(inscope), qty=sum(x["qty"] for x in inscope),
                 dpci=len({x["dpci"] for x in inscope}), factories=len(facs), excluded=dict(rc),
                 costdiff=len(costdiff), cost_hi=len(hi), cost_lo=len(lo), missing_photos=len(noph),
                 cancelled=len(res["cancel"]), sheets=[names[f] for f in facs])
    return out.getvalue(), stats


# =====================================================================  read back an existing GRID
def read_grid(data):
    """Return (state, photos, manual) from a workbook produced by render()."""
    wb = openpyxl.load_workbook(io.BytesIO(data))
    if STATE_SHEET not in wb.sheetnames:
        raise ValueError("這份檔案不是 D53 App 產生的 GRID（找不到隱藏頁 _D53_STATE），請改用「新建 GRID」。")
    parts = collections.defaultdict(list)
    for k, i, s in wb[STATE_SHEET].iter_rows(min_row=2, values_only=True):
        if k:
            parts[k].append((int(i), s or ""))
    blob = {k: json.loads("".join(s for _, s in sorted(v))) for k, v in parts.items()}
    state = dict(settings=blob["settings"], pcn=blob["pcn"], docs=blob["docs"],
                 changelog=blob.get("changelog", []), problems=[])
    photos, manual = {}, {"po": {}, "note": {}}
    if PHOTO_SHEET in wb.sheetnames:
        ws = wb[PHOTO_SHEET]
        for im in ws._images:
            try:
                d = ws.cell(im.anchor._from.row + 1, 1).value
                if d:
                    photos[d] = im._data()
            except Exception:
                continue
    for ws in wb.worksheets:
        if ws.title in RESERVED or ws.cell(HR, 1).value != "STYLE":
            continue
        T = next((c for c in range(P0, ws.max_column + 1) if ws.cell(HR, c).value == "PO 合計"), None)
        if not T:
            continue
        for c in range(P0, T):
            po = str(ws.cell(3, c).value or "").split("-")[0]
            sc, ac = ws.cell(9, c).value, ws.cell(10, c).value
            if po and (sc not in (None, "") or ac not in (None, "")):
                manual["po"][po] = {"sc": sc, "ac": ac}
        rowdpci = {}
        for r in range(HR + 1, ws.max_row + 1):
            d = ws.cell(r, 2).value
            if isinstance(d, str) and re.fullmatch(r"\d{3}-\d{2}-\d{4}", d):
                rowdpci[r] = d
                note = ws.cell(r, T + 4).value
                if note not in (None, ""):
                    manual["note"][d] = note
        for im in ws._images:
            try:
                r = im.anchor._from.row + 1
                if r in rowdpci and rowdpci[r] not in photos:
                    photos[rowdpci[r]] = im._data()
            except Exception:
                continue
    return state, photos, manual


# =====================================================================  merge + change log
def _qty_map(state):
    res = compute(state)
    m = {}
    for x in res["inscope"]:
        m[(x["po"], x["dpci"])] = m.get((x["po"], x["dpci"]), 0) + x["qty"]
    fac = {x["po"]: x["factory"] for x in res["inscope"]}
    return m, fac, set(res["cancel"])


def merge_update(state, new_docs, new_pcns=None, new_problems=None, settings_patch=None):
    """Add new PO documents / PCN files to a stored state and append the change log."""
    before, fac_b, canc_b = _qty_map(state)
    old_pcn_dpci = {d for p in state["pcn"].values() for d in p["rows"]}
    seen = {(d["po"], d["doc"], d.get("printed"), d["file"]) for d in state["docs"]}
    added_docs = [d for d in new_docs if (d["po"], d["doc"], d.get("printed"), d["file"]) not in seen]
    state["docs"] = state["docs"] + added_docs
    for p in (new_pcns or []):
        state["pcn"][p["program"]] = p
    if settings_patch:
        state["settings"].update(settings_patch)
    state["problems"] = list(new_problems or [])
    after, fac_a, canc_a = _qty_map(state)
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    log = []
    new_pos = {po for po, _ in after} - {po for po, _ in before}
    for po in sorted(new_pos):
        n = sum(v for (p, _), v in after.items() if p == po)
        log.append(dict(time=now, action="新 PO", po=po, factory=fac_a.get(po), new=n,
                        note=f"{sum(1 for (p, _) in after if p == po)} 個 DPCI"))
    for k in sorted(set(before) | set(after)):
        po, d = k
        if po in new_pos:
            continue
        b, a = before.get(k), after.get(k)
        if b == a:
            continue
        if a is None:
            act = "已取消" if po in canc_a else "移出範圍"
        elif b is None:
            act = "新增品項"
        else:
            act = "改數量"
        log.append(dict(time=now, action=act, po=po, dpci=d, factory=fac_a.get(po) or fac_b.get(po), old=b, new=a))
    for p in (new_pcns or []):
        added = sorted(set(p["rows"]) - old_pcn_dpci)
        if added:
            log.append(dict(time=now, action="PCN 更新", note=f"{p['program']}：新增 {len(added)} 個 DPCI（已納入；不屬於你請用進階篩選排除）："
                            + "、".join(added[:30]) + ("…" if len(added) > 30 else "")))
    if not log:
        log.append(dict(time=now, action="更新", note=f"上傳 {len(new_docs)} 份 PO 文件，GRID 內容沒有變化"))
    state["changelog"] = state.get("changelog", []) + log
    return state, log, added_docs


def new_state(ae, season_start, pcns, docs, problems, excluded_factories=(), excluded_dpci=()):
    return dict(settings=dict(ae=ae, season_start=season_start.isoformat() if hasattr(season_start, "isoformat") else season_start,
                              excluded_factories=list(excluded_factories), excluded_dpci=list(excluded_dpci),
                              created=datetime.datetime.now().strftime("%Y-%m-%d %H:%M")),
                pcn={p["program"]: p for p in pcns}, docs=docs,
                changelog=[dict(time=datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), action="新建",
                                note=f"PCN：{'、'.join(p['program'] for p in pcns)}；PO 文件 {len(docs)} 份")],
                problems=list(problems))
