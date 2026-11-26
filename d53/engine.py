"""D53 PO GRID engine — PCN-anchored, one sheet per factory.

Data flow:  PCN (scope) -> SPS PO PDFs (quantities) -> WRK workspace (photos) -> GRID workbook.

Scope gates applied to every PO line:
  1. DPCI is in an uploaded PCN, narrowed by the 範圍 setting (all / include list / exclude list)
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
ENGINE_VERSION = "d53-1.1"
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
        starts = [i for i, ln in enumerate(L) if _LINE.search(ln)]
        for k, i in enumerate(starts):
            ln = L[i]
            mm = _LINE.search(ln)
            full_blk = "\n".join(L[i:starts[k + 1] if k + 1 < len(starts) else len(L)])
            inn = re.search(r"#\s*of\s*Inners\s*:\s*(\d+)", full_blk)
            upi = re.search(r"#\s*of\s*Units\s*per\s*Inner\s*:\s*(\d+)", full_blk)
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
                              style=st.group(1) if st else None,
                              inners=int(inn.group(1)) if inn else None,
                              upi=int(upi.group(1)) if upi else None))
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


def filter_mode(settings):
    m = settings.get("filter_mode")
    if m:
        return m
    return "exclude" if (settings.get("excluded_factories") or settings.get("excluded_dpci")) else "all"


def match_tokens(pcns, text):
    """Match pasted DPCI / Vendor Style # tokens against PCN rows.
    Returns (matched DPCI list, {token: [dpci...]}, unmatched tokens)."""
    rows = {d: r for p in pcns for d, r in p["rows"].items()}
    by_style = collections.defaultdict(list)
    for d, r in rows.items():
        if r["style"]:
            by_style[r["style"].strip().upper()].append(d)
    hits, unmatched, matched = {}, [], []
    for tok in re.split(r"[\s,;，；、]+", text or ""):
        tok = tok.strip().strip("'\"")
        if not tok:
            continue
        digits = re.sub(r"\D", "", tok)
        cand = []
        if len(digits) == 9 and re.fullmatch(r"[\d\-\s]+", tok):
            d = fmt_dpci(digits)
            cand = [d] if d in rows else []
        if not cand:
            cand = by_style.get(tok.upper(), [])
        if cand:
            hits[tok] = cand
            matched += [d for d in cand if d not in matched]
        else:
            unmatched.append(tok)
    return matched, hits, unmatched


def effective_pcn(state):
    S = state["settings"]
    mode = filter_mode(S)
    inc_f, inc_d = set(S.get("include_factories", [])), set(S.get("include_dpci", []))
    excl_f, excl_d = set(S.get("excluded_factories", [])), set(S.get("excluded_dpci", []))
    out = {}
    for p in state["pcn"].values():
        for d, r in p["rows"].items():
            if mode == "include" and not (d in inc_d or r["factory"] in inc_f):
                continue
            if mode == "exclude" and (d in excl_d or r["factory"] in excl_f):
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
                       inners=l.get("inners"), upi=l.get("upi"), file=r["file"], reason=reason)
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
ORG = PatternFill("solid", fgColor="FFC000")
_md = Side(style="medium", color="404040")
# Program tab (Vanessa layout): PO header block in rows 1-8, column header in row 9, items from row 10.
LABELS = ["PO No.", "Factory", "S/C（手填）", "PO Date", "Assigned by Sender\n(Purpose)", "AC（手填）",
          "Shipping Point", "Shipping Window"]
FIXH = ["Product Description", "Factory", "Image", "DPCI", "STYLE", "Case/\nInner", "Retail$", "Unit Cost$"]
TAILH = ["Q'TY", "Forecasted Units\n(Commit)", "+-20%\n(Forecast×80%)", "差異 %\n(Forecast−Q'TY)/Forecast",
         "核對", "AE 備註（自由欄）"]
FIX = len(FIXH); P0 = FIX + 1; HR = len(LABELS) + 1
C_FAC, C_IMG, C_DPCI, C_COST = 2, 3, 4, 8
RESERVED = {"使用說明", "Validation Summary", "Validation Detail", "Program 總覽", "Cost Discrepancy", "排除清單", "PO Ledger", "更新紀錄", STATE_SHEET, PHOTO_SHEET}


def _short(f):
    w = re.sub(r"[^\w& ]", " ", f).split()
    return (" ".join(w[:2]) or "Factory")[:31]


CHECKS = [
    ("Unit Cost$", "PO 單價 = PCN 成本（FCA／FOB）", "行"),
    ("Retail$", "PO Resale = PCN Suggested Unit Retail", "行"),
    ("Case Pack", "PO # of Inners × Units per Inner = PCN Case；Units per Inner = PCN Inner", "行"),
    ("數量｜整箱", "PO 數量是 Case Pack 的整數倍", "行"),
    ("數量｜vs Commit", "全部 PO 合計 ≤ Commit × 110%（±10% 規則）", "DPCI"),
    ("Factory", "PO 上的工廠 = PCN Factory Name", "行"),
    ("PO 解析驗算", "每張 PO 的行數量合計 = Total Qty", "PO"),
    ("Image", "WRK 有這個 DPCI 的照片", "DPCI"),
]


def _norm_fac(f):
    return re.sub(r"[^a-z]", "", (f or "").lower())


def _pack(p):
    m = re.match(r"(\d+)/(\d+)", p.get("pack") or "")
    return (int(m.group(1)), int(m.group(2))) if m else (None, None)


def validate(state, res, photos):
    """Every check on every in-scope line.  Returns (summary rows, detail rows, flags by DPCI)."""
    pcn, inscope = res["pcn"], res["inscope"]
    det = []
    tested = collections.Counter()
    for x in inscope:
        p = pcn[x["dpci"]]
        base = dict(po=x["po"], dpci=x["dpci"], style=p["style"], factory=x["factory"])
        if p["cost"] is not None and x["cost"] is not None:
            tested["Unit Cost$"] += 1
            if abs(x["cost"] - p["cost"]) > 0.005:
                det.append(dict(base, check="Unit Cost$", po_v=x["cost"], pcn_v=p["cost"],
                                note=f"PO {'高' if x['cost'] > p['cost'] else '低'}於 PCN {abs(x['cost'] - p['cost']):.3f}／pc，"
                                     f"影響 USD {(x['cost'] - p['cost']) * x['qty']:,.2f}"))
        if p["retail"] is not None and x["retail"] is not None:
            tested["Retail$"] += 1
            if abs(x["retail"] - p["retail"]) > 0.005:
                det.append(dict(base, check="Retail$", po_v=x["retail"], pcn_v=p["retail"], note="PO 為準，只標記"))
        case, inner = _pack(p)
        if case and x.get("inners") and x.get("upi"):
            tested["Case Pack"] += 1
            po_case = x["inners"] * x["upi"]
            if po_case != case or x["upi"] != inner:
                det.append(dict(base, check="Case Pack", po_v=f"{po_case}/{x['upi']}", pcn_v=p["pack"],
                                note=f"PO：{x['inners']} inner × {x['upi']} pcs"))
        if case:
            tested["數量｜整箱"] += 1
            if x["qty"] % case:
                det.append(dict(base, check="數量｜整箱", po_v=x["qty"], pcn_v=case,
                                note=f"{x['qty']:,} ÷ {case} = {x['qty'] / case:,.2f} 箱"))
        if p["factory"]:
            tested["Factory"] += 1
            a, b = _norm_fac(x["factory"]), _norm_fac(p["factory"])
            if not (a[:12] and (a.startswith(b[:12]) or b.startswith(a[:12]))):
                det.append(dict(base, check="Factory", po_v=x["factory"], pcn_v=p["factory"], note="PO 工廠與 PCN 不同"))
    tot = collections.Counter()
    for x in inscope:
        tot[x["dpci"]] += x["qty"]
    for d, q in tot.items():
        cm = pcn[d]["commit"]
        if cm > 0:
            tested["數量｜vs Commit"] += 1
            if q > cm * 1.1:
                det.append(dict(po="（全部 PO）", dpci=d, style=pcn[d]["style"], factory=pcn[d]["factory"],
                                check="數量｜vs Commit", po_v=q, pcn_v=cm, note=f"PO 合計為 Commit 的 {q / cm:.1%}，超過 110%"))
    pos = {x["po"] for x in inscope}
    tested["PO 解析驗算"] = len(pos)
    for pr in state.get("problems", []):
        m = re.search(r"PO (\d{11})", pr)
        if m and m.group(1) in pos and "Total Qty" in pr:
            det.append(dict(po=m.group(1), dpci="", style="", factory="", check="PO 解析驗算", po_v="", pcn_v="", note=pr))
    tested["Image"] = len(tot)
    for d in sorted(tot):
        if d not in photos:
            det.append(dict(po="", dpci=d, style=pcn[d]["style"], factory=pcn[d]["factory"], check="Image",
                            po_v="", pcn_v="", note="WRK 沒有這個 DPCI 的照片"))
    nbad = collections.Counter(e["check"] for e in det)
    summary = [(c, rule, unit, tested[c], nbad[c]) for c, rule, unit in CHECKS]
    flags = collections.defaultdict(list)
    for e in det:
        if e["dpci"]:
            flags[(e["dpci"], e["factory"])].append(e)
            flags[(e["dpci"], None)].append(e)
    return summary, det, flags


def render(state, photos=None, manual=None):
    """Build the workbook; returns (bytes, stats)."""
    photos = photos or {}
    manual = manual or {"po": {}, "note": {}}
    res = compute(state)
    pcn, inscope, excl = res["pcn"], res["inscope"], res["excl"]
    vsum, vdet, vflags = validate(state, res, photos)
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
    byprog = collections.defaultdict(list)
    for x in inscope:
        byprog[pcn[x["dpci"]]["program"]].append(x)
    progs = sorted(byprog, key=lambda g: -sum(x["qty"] for x in byprog[g]))
    fqty = collections.Counter()
    for x in inscope:
        fqty[x["factory"]] += x["qty"]
    facs = sorted(fqty, key=lambda f: -fqty[f])
    names = {}
    for g in progs:
        base = re.sub(r"[\[\]:*?/\\']", " ", g).strip()[:31] or "Program"
        n, i = base, 2
        while n in names.values() or n in RESERVED:
            n = f"{base[:28]}-{i}"; i += 1
        names[g] = n

    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "使用說明"
    for t, b in [(f"D53 PO GRID｜{ae}", 2), (f"產生 {today}｜資料流：PCN（範圍）→ PO（數量）→ WRK（照片）", 0), ("", 0),
                 ("範圍怎麼決定", 1),
                 ("1. 品項：只看上傳的 PCN（每份 PCN＝一個 Program）。PCN 跟別人共用時，可在 App 的「品項範圍」貼上自己負責的 DPCI／Style#，只納入這些。", 0),
                 (f"2. 季別：Purpose 不是 BAS，且 Ship Window 開始日 ≥ 本季起始日（本檔：{season}）。", 0),
                 ("3. 版本：同一張 PO 取 SPS 列印時間最新的一份；收到 Import PO Change CANCEL 即整張排除。", 0),
                 ("不符合的 PO 行全部列在「排除清單」並寫明原因。", 0), ("", 0), ("怎麼用", 1),
                 ("• 每個 Program 一個分頁。品項依 Factory 分區塊（粗線隔開），PO 欄也依同樣的工廠順序排列，所以每家工廠的數量落在自己的區塊。", 0),
                 ("• 第 1–8 列是 PO 表頭（S/C、AC 黃底手填）；最右側是 Q'TY、Forecasted Units（Commit）、+-20%、差異 %、核對、AE 備註（黃底）。", 0),
                 ("• 紅底數量格＝該 PO 單價 ≠ Unit Cost$（PO 為準，只標記）。所有核對項目（價錢、Retail、Case Pack、數量、Factory）的結果在 Validation Summary 第 0 節，逐筆明細在「Validation Detail」；英文的成本差異明細在「Cost Discrepancy」。", 0),
                 ("• 之後有新 PO：到 App 選「更新現有 GRID」，上傳這份檔案＋新的 PO PDF 即可；要加 Program 就一起上傳新的 PCN。", 0),
                 ("• 只有黃底欄位會在更新時保留；請不要在其他地方手寫資料，也不要改分頁名稱或刪除隱藏頁。", 0)]:
        ws.append([t]); c = ws.cell(ws.max_row, 1)
        c.font = F(bold=True, size=14) if b == 2 else (F(bold=True) if b == 1 else F())
    ws.column_dimensions["A"].width = 130

    overview = collections.defaultdict(lambda: dict(items=set(), pos=set(), units=0, cost=0.0, diff=0, row=None))
    sheetinfo = {}
    for g in progs:
        xs = byprog[g]; ws = wb.create_sheet(names[g])
        gq = collections.Counter()
        for x in xs:
            gq[x["factory"]] += x["qty"]
        gfacs = sorted(gq, key=lambda f: -gq[f]); forder = {f: i for i, f in enumerate(gfacs)}
        pos = {}
        for x in xs:
            pos.setdefault(x["po"], x)
        pos = sorted(pos.values(), key=lambda x: (forder[x["factory"]], _swstart(x["sw"]) or datetime.date.max, x["po"]))
        PN = P0 + len(pos) - 1; T = PN + 1; LAST = T + len(TAILH) - 1
        items = sorted({(x["factory"], x["dpci"]) for x in xs},
                       key=lambda k: (forder[k[0]], pcn[k[1]]["style"] or "", k[1]))
        q = collections.defaultdict(int)
        for x in xs:
            q[(x["dpci"], x["factory"], x["po"])] += x["qty"]
        # first PO column of each factory -> medium left border (block-diagonal look)
        fstart = {P0 + j for j, p in enumerate(pos) if j == 0 or p["factory"] != pos[j - 1]["factory"]}

        def edge(c, top=False):
            l = _md if c.column in fstart or c.column == T else _th
            c.border = Border(left=l, right=_th, bottom=_th, top=_md if top else _th)

        for i, lab in enumerate(LABELS):
            rr = 1 + i
            c = ws.cell(rr, 1, lab); c.font = F(bold=True); c.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
            c.fill = ORG if i == 0 else SUB
            ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=FIX)
            for j, p in enumerate(pos):
                mp = manual["po"].get(p["po"], {})
                v = [f"{p['po']}-{p['dc']}", _short(p["factory"]), mp.get("sc"), p["po_date"], p["purpose"], mp.get("ac"),
                     p["ship_pt"], re.sub(r"/20(\d\d)", "", p["sw"] or "")][i]
                cc = ws.cell(rr, P0 + j, v); cc.alignment = CEN; cc.font = F(size=9, bold=(i in (0, 1))); edge(cc)
                if i == 0:
                    cc.fill = ORG
                elif i in (2, 5):
                    cc.fill = YEL
            ws.row_dimensions[rr].height = 22
        for i, h in enumerate(FIXH, 1):
            ws.cell(HR, i, h)
        for j, p in enumerate(pos):  # DC code above each PO column, as in the AE templates
            ws.cell(HR, P0 + j, str(p["dc"] or ""))
        for k, h in enumerate(TAILH):
            ws.cell(HR, T + k, h)
        for cc in range(1, LAST + 1):
            c = ws.cell(HR, cc); c.fill = HDR; edge(c)
            if c.value is not None:
                c.font = F(bold=True, color="FFFFFF"); c.alignment = CEN
        ws.row_dimensions[HR].height = 40
        r = HR; R0 = HR + 1; prevf = None
        for f, d in items:
            p = pcn[d]; r += 1
            newblock = prevf is not None and f != prevf
            if f != prevf:
                overview[(g, f)]["row"] = r
            prevf = f
            for i, v in enumerate([p["desc"], f, None, d, p["style"], p["pack"], p["retail"], p["cost"]], 1):
                c = ws.cell(r, i, v); c.alignment = LW if i in (1, 2) else CEN; c.font = F(size=9)
            ws.cell(r, 7).number_format = ws.cell(r, 8).number_format = "0.00"
            ws.cell(r, C_COST).fill = GRN
            for j, po in enumerate(pos):
                c = ws.cell(r, P0 + j, q.get((d, f, po["po"])))
                c.number_format = "#,##0"; c.font = F(size=9); c.alignment = CEN
                if (d, po["po"]) in bad and po["factory"] == f:
                    c.fill = RED
            ws.cell(r, T, f"=SUM({CL(P0)}{r}:{CL(PN)}{r})").number_format = "#,##0"
            ws.cell(r, T + 1, p["commit"]).number_format = "#,##0"
            ws.cell(r, T + 2, f'=IF({CL(T + 1)}{r}>0,{CL(T + 1)}{r}*0.8,"")').number_format = "#,##0"
            ws.cell(r, T + 3, f'=IF({CL(T + 1)}{r}>0,({CL(T + 1)}{r}-{CL(T)}{r})/{CL(T + 1)}{r},"")').number_format = "0.0%"
            fl = []
            for e in vflags.get((d, f), []) + [e for e in vflags.get((d, None), []) if e["check"] in ("數量｜vs Commit", "Image")]:
                ck = e["check"]
                if ck == "Unit Cost$":
                    lab = f"Cost PO {e['po_v']} vs {e['pcn_v']}"
                elif ck == "Retail$":
                    lab = f"Retail PO {e['po_v']} vs {e['pcn_v']}"
                elif ck == "Case Pack":
                    lab = f"Case Pack PO {e['po_v']} vs {e['pcn_v']}"
                else:
                    lab = {"數量｜整箱": "數量非整箱", "數量｜vs Commit": "超過 Commit 110%",
                           "Factory": "Factory 與 PCN 不同", "Image": "缺照片"}.get(ck, ck)
                if lab not in fl:
                    fl.append(lab)
            if p["cost"] is None:
                fl.append("PCN 無成本")
            c = ws.cell(r, T + 4, "OK" if not fl else "；".join(fl)); c.alignment = LW
            c.font = F(size=9, color="006100" if not fl else "9C0006"); c.fill = GRN if not fl else RED
            ws.cell(r, T + 5, manual["note"].get(d))
            for k in range(T, T + 4):
                ws.cell(r, k).alignment = CEN; ws.cell(r, k).font = F(size=9)
            ws.cell(r, T + 5).fill = YEL; ws.cell(r, T + 5).alignment = LW
            for cc in range(1, LAST + 1):
                edge(ws.cell(r, cc), top=newblock)
            ws.row_dimensions[r].height = 62
            if d in photos:
                im = XLImage(io.BytesIO(photos[d]))
                s = 76 / max(im.width, im.height); im.width, im.height = int(im.width * s), int(im.height * s)
                im.anchor = TwoCellAnchor(editAs="twoCell",
                                          _from=AnchorMarker(col=C_IMG - 1, colOff=19050, row=r - 1, rowOff=19050),
                                          to=AnchorMarker(col=C_IMG - 1, colOff=19050 + im.width * 9525, row=r - 1, rowOff=19050 + im.height * 9525))
                ws.add_image(im)
            o = overview[(g, f)]
            o["items"].add(d)
            o["units"] += sum(v for (dd, ff, _), v in q.items() if dd == d and ff == f)
            o["cost"] += sum(v * (p["cost"] or 0) for (dd, ff, _), v in q.items() if dd == d and ff == f)
            o["pos"] |= {po["po"] for po in pos if q.get((d, f, po["po"]))}
            o["diff"] += sum(1 for z in costdiff if z["dpci"] == d and z["factory"] == f)
        RN = r; t1, t2 = RN + 1, RN + 2
        for rr, lab in ((t1, "Total Units"), (t2, "Total Cost $（Unit Cost$ × Q'TY）")):
            ws.cell(rr, 1, lab).alignment = Alignment(horizontal="right", vertical="center")
            ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=FIX)
        cst = f"${CL(C_COST)}${R0}:${CL(C_COST)}${RN}"
        for col in [P0 + j for j in range(len(pos))] + [T, T + 1]:
            L = CL(col)
            ws.cell(t1, col, f"=SUM({L}{R0}:{L}{RN})").number_format = "#,##0"
            ws.cell(t2, col, f"=SUMPRODUCT({L}{R0}:{L}{RN},{cst})").number_format = "#,##0.00"
        for rr in (t1, t2):
            for cc in range(1, LAST + 1):
                c = ws.cell(rr, cc); c.font = F(bold=True, size=9); c.fill = SUB; edge(c, top=(rr == t1))
        for i, w in enumerate([34, 18, 13, 13, 13, 8, 8, 9], 1):
            ws.column_dimensions[CL(i)].width = w
        for j in range(len(pos)):
            ws.column_dimensions[CL(P0 + j)].width = 14
        for k, w in enumerate([11, 12, 11, 12, 30, 22]):
            ws.column_dimensions[CL(T + k)].width = w
        ws.freeze_panes = ws.cell(R0, P0)
        ws.auto_filter.ref = f"A{HR}:{CL(LAST)}{RN}"
        sheetinfo[g] = (names[g], len(gfacs), len({d for _, d in items}), len(pos), sum(x["qty"] for x in xs))

    # ---- Program 總覽
    ov = wb.create_sheet("Program 總覽")
    ov.append([f"Program 總覽｜{ae}｜{today}"]); ov["A1"].font = F(bold=True, size=14); ov.append([])
    ov.append(["Program（PCN）", "Factory", "Program 分頁", "品項", "有效 PO", "數量", "Commit", "% vs Commit", "Cost $（Unit Cost$ × 數量）", "成本差異行"])
    for c in ov[3]:
        c.font = F(bold=True, color="FFFFFF"); c.fill = HDR; c.alignment = CEN; c.border = BOX
    for pg in sorted({k[0] for k in overview}):
        rs = ov.max_row + 1
        for (p, f), o in sorted(overview.items(), key=lambda kv: -kv[1]["units"]):
            if p != pg:
                continue
            ov.append([p, f, names[p], len(o["items"]), len(o["pos"]), o["units"],
                       sum(pcn[d]["commit"] for d in o["items"]), None, round(o["cost"], 2), o["diff"]])
            rr = ov.max_row
            ov.cell(rr, 8, f'=IF(G{rr}>0,F{rr}/G{rr},"")')
            ov.cell(rr, 3).hyperlink = f"#'{names[p]}'!A{o['row'] or 1}"; ov.cell(rr, 3).font = F(color="0563C1", underline="single")
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

    # ---- Validation Detail
    vd = wb.create_sheet("Validation Detail")
    vd.append(["檢查項目", "PO#", "DPCI", "STYLE", "Factory", "PO 值", "PCN 值", "說明"])
    for c in vd[1]:
        c.font = F(bold=True, color="FFFFFF"); c.fill = HDR; c.alignment = CEN
    order_c = {c[0]: i for i, c in enumerate(CHECKS)}
    for e in sorted(vdet, key=lambda e: (order_c.get(e["check"], 99), e["po"], e["dpci"])):
        vd.append([e["check"], e["po"], e["dpci"], e["style"], e["factory"], e["po_v"], e["pcn_v"], e["note"]])
    for i, w in enumerate([16, 13, 12, 14, 32, 14, 14, 60], 1):
        vd.column_dimensions[CL(i)].width = w
    vd.freeze_panes = "A2"; vd.auto_filter.ref = vd.dimensions

    # ---- 排除清單
    ex = wb.create_sheet("排除清單")
    ex.append(["原因", "PO#", "DC", "Dept", "Purpose", "PO Date", "Ship Window", "Factory", "Line", "DPCI", "數量", "來源檔案"])
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
    lg.append(["PO#", "DC", "PO Date", "SPS 列印時間", "Purpose", "Ship Window", "Ship Point", "Factory", "Vendor #", "Line",
               "DPCI", "STYLE", "Product Description", "Program", "數量", "PO 單價", "Unit Cost$", "PO Retail", "Retail$",
               "PO Inners", "PO Units/Inner", "來源檔案"])
    for c in lg[1]:
        c.font = F(bold=True, color="FFFFFF"); c.fill = HDR; c.alignment = CEN
    for x in sorted(inscope, key=lambda x: (x["factory"], x["po"], x["line"])):
        p = pcn[x["dpci"]]
        lg.append([x["po"], x["dc"], x["po_date"], x["printed"], x["purpose"], x["sw"], x["ship_pt"], x["factory"],
                   x["vendor"], x["line"], x["dpci"], p["style"], p["desc"], p["program"], x["qty"], x["cost"],
                   p["cost"], x["retail"], p["retail"], x.get("inners"), x.get("upi"), x["file"]])
    for i, w in enumerate([13, 6, 11, 16, 8, 22, 9, 30, 10, 5, 12, 14, 34, 26, 9, 8, 8, 8, 8, 8, 9, 34], 1):
        lg.column_dimensions[CL(i)].width = w
    lg.freeze_panes = "A2"; lg.auto_filter.ref = lg.dimensions

    # ---- 更新紀錄
    cl = wb.create_sheet("更新紀錄")
    cl.append(["時間", "動作", "PO#", "DPCI", "Factory", "原數量", "新數量", "說明"])
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
    Hh("0. 核對總表 Validation Checklist")
    row(["檢查項目", "規則", "單位", "已檢查", "差異", "結果"], SUB, True)
    for c, rule, unit, n, bad_n in vsum:
        res_txt = ("—（PO 沒有 Inner 資料：這份 GRID 的 PO 是舊版解析的，更新時把 PO PDF 重新上傳一次即可補上）"
                   if n == 0 and c == "Case Pack" and inscope else "—（沒有資料可比對）") if n == 0 else ("✔ 全部相符" if bad_n == 0 else f"✖ {bad_n} 筆，見「Validation Detail」")
        row([c, rule, unit, n, bad_n, res_txt], None if n == 0 else (GRN if bad_n == 0 else RED))
    vs.append([])
    Hh("1. 範圍")
    row(["項目", "數量", "說明"], SUB, True)
    for p in state["pcn"].values():
        row([f"PCN：{p['program']}", f"{len(p['rows'])} 個 DPCI", f"{p['file']}（{p['nrows']} 列）"])
        if p.get("nodpci"):
            row(["　└ 沒有 DPCI 的列", sum(p["nodpci"].values()),
                 "無法和 PO 比對：" + "、".join(f"{k} {v}" for k, v in p["nodpci"].items())], YEL)
    fm = filter_mode(S)
    if fm == "include":
        ftxt = f"只納入：{len(S.get('include_dpci', []))} 個指定 DPCI" + (f"＋{len(S.get('include_factories', []))} 家工廠的全部品項" if S.get("include_factories") else "")
    elif fm == "exclude":
        ftxt = f"排除 {len(S.get('excluded_factories', []))} 家工廠、{len(S.get('excluded_dpci', []))} 個 DPCI"
    else:
        ftxt = "整份 PCN 全部納入"
    row(["品項範圍", ftxt, f"實際納入 {len(pcn)} 個 DPCI"])
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
    vs.append([]); Hh("2. Program 分頁")
    row(["Program", "分頁", "工廠", "品項", "有效 PO", "數量"], SUB, True)
    for g in progs:
        row([g, *sheetinfo[g]])
    vs.append([]); Hh("3. 成本差異：PO 單價 ≠ Unit Cost$（PO 為準，只標記）")
    hi = [x for x in costdiff if x["cost"] > pcn[x["dpci"]]["cost"]]
    lo = [x for x in costdiff if x["cost"] < pcn[x["dpci"]]["cost"]]
    row(["PO 高於 PCN", f"{len(hi)} 行", f"{len({x['dpci'] for x in hi})} 個 DPCI、{len({x['po'] for x in hi})} 張 PO；合計 USD "
         f"{sum((x['cost'] - pcn[x['dpci']]['cost']) * x['qty'] for x in hi):,.2f}"], RED if hi else None)
    row(["PO 低於 PCN（對我們不利）", f"{len(lo)} 行",
         "；".join(sorted({f"{x['dpci']} PO {x['cost']} vs PCN {pcn[x['dpci']]['cost']}" for x in lo})) or "無"], RED if lo else None)
    vs.append([]); Hh("4. 已取消 PO → 重開對照")
    row(["已取消 PO", "Factory", "原 Ship Window", "重開為"], SUB, True)
    for p, new in sorted(res["reissue"].items(), key=lambda kv: res["orders"][kv[0]].get("factory") or ""):
        o = res["orders"][p]
        if not any(fmt_dpci(l["dpci"]) in pcn for l in o["lines"]):
            continue
        row([f"{p}-{o['dc']}", o.get("factory"), o.get("sw"), ", ".join(new) or "找不到相同內容的新單"], None if new else RED)
    vs.append([]); Hh("5. PCN 有、但還沒收到本季 PO 的品項（Commit > 0，排除 No Longer Under Consideration）")
    row(["DPCI", "STYLE", "Product Description", "Program", "Commit", "Plan Status", "Factory"], SUB, True)
    got = {x["dpci"] for x in inscope}
    for d, p in sorted(pcn.items(), key=lambda kv: (kv[1]["program"], kv[0])):
        if d in got or p["commit"] <= 0 or p["status"] == "No Longer Under Consideration":
            continue
        row([d, p["style"], p["desc"], p["program"], p["commit"], p["status"], p["factory"]])
    noph = sorted({x["dpci"] for x in inscope if x["dpci"] not in photos})
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

    order = ["使用說明", "Validation Summary", "Program 總覽"] + [names[g] for g in progs] + \
            ["Validation Detail", "Cost Discrepancy", "排除清單", "PO Ledger", "更新紀錄", PHOTO_SHEET, STATE_SHEET]
    wb._sheets = [wb[n] for n in order]
    wb.active = 1
    out = io.BytesIO(); wb.save(out)
    stats = dict(checks=[(c, n, b) for c, _, _, n, b in vsum], pos=len({x["po"] for x in inscope}), lines=len(inscope), qty=sum(x["qty"] for x in inscope),
                 dpci=len({x["dpci"] for x in inscope}), factories=len(facs), programs=len(progs), factory_names=list(facs), excluded=dict(rc),
                 costdiff=len(costdiff), cost_hi=len(hi), cost_lo=len(lo), missing_photos=len(noph),
                 cancelled=len(res["cancel"]), sheets=[names[g] for g in progs])
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
        if ws.title in RESERVED:
            continue
        lay = _grid_layout(ws)
        if not lay:
            continue
        hr, dcol, T, ncol, porow, scrow, acrow = lay
        for c in range(1, T):
            m = re.match(r"(\d{11})-", str(ws.cell(porow, c).value or ""))
            if not m:
                continue
            sc = ws.cell(scrow, c).value if scrow else None
            ac = ws.cell(acrow, c).value if acrow else None
            if sc not in (None, "") or ac not in (None, ""):
                manual["po"][m.group(1)] = {"sc": sc, "ac": ac}
        rowdpci = {}
        for r in range(hr + 1, ws.max_row + 1):
            d = ws.cell(r, dcol).value
            if isinstance(d, str) and re.fullmatch(r"\d{3}-\d{2}-\d{4}", d):
                rowdpci[r] = d
                note = ws.cell(r, ncol).value if ncol else None
                if note not in (None, "") and d not in manual["note"]:
                    manual["note"][d] = note
        for im in ws._images:
            try:
                r = im.anchor._from.row + 1
                if r in rowdpci and rowdpci[r] not in photos:
                    photos[rowdpci[r]] = im._data()
            except Exception:
                continue
    return state, photos, manual


def _grid_layout(ws):
    """Locate the parts of a GRID tab.  Works for the Program-tab layout (header row 9) and the
    older factory-sheet layout (header row 11), so GRIDs made before the change still update."""
    for hr in range(1, 16):
        vals = {ws.cell(hr, c).value: c for c in range(1, min(ws.max_column, 200) + 1)}
        if "DPCI" not in vals:
            continue
        T = vals.get("Q'TY") or vals.get("PO 合計")
        if not T:
            continue
        ncol = next((c for v, c in vals.items() if isinstance(v, str) and v.startswith("AE 備註")), None)
        po = next(((r, c) for r in range(1, hr) for c in range(1, T)
                   if re.match(r"\d{11}-", str(ws.cell(r, c).value or ""))), None)
        if not po:
            continue
        scrow = acrow = None
        for r in range(1, hr):  # row labels sit left of the first PO column
            for c in range(1, po[1]):
                v = ws.cell(r, c).value
                if isinstance(v, str) and v.startswith("S/C"):
                    scrow = r
                elif isinstance(v, str) and v.startswith("AC"):
                    acrow = r
        return hr, vals["DPCI"], T, ncol, po[0], scrow, acrow
    return None


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
    key = lambda d: (d["po"], d["doc"], d.get("printed"), d["file"])  # noqa: E731
    fresh = {key(d): d for d in new_docs}
    # a PDF uploaded again replaces its stored copy, so fields added to the parser later
    # (e.g. # of Inners / Units per Inner) get filled in for GRIDs built by an older version
    state["docs"] = [fresh.get(key(d), d) for d in state["docs"]]
    seen = {key(d) for d in state["docs"]}
    added_docs = [d for d in new_docs if key(d) not in seen]
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
            how = ("未納入；若是你的請加進「只納入」清單" if filter_mode(state["settings"]) == "include"
                   else "已納入；不屬於你請在範圍設定排除")
            log.append(dict(time=now, action="PCN 更新", note=f"{p['program']}：新增 {len(added)} 個 DPCI（{how}）："
                            + "、".join(added[:30]) + ("…" if len(added) > 30 else "")))
    if not log:
        log.append(dict(time=now, action="更新", note=f"上傳 {len(new_docs)} 份 PO 文件，GRID 內容沒有變化"))
    state["changelog"] = state.get("changelog", []) + log
    return state, log, added_docs


def new_state(ae, season_start, pcns, docs, problems, excluded_factories=(), excluded_dpci=(),
              filter_mode=None, include_factories=(), include_dpci=()):
    if filter_mode is None:  # infer from whichever list was given
        filter_mode = ("include" if (include_factories or include_dpci)
                       else "exclude" if (excluded_factories or excluded_dpci) else "all")
    return dict(settings=dict(ae=ae, season_start=season_start.isoformat() if hasattr(season_start, "isoformat") else season_start,
                              excluded_factories=list(excluded_factories), excluded_dpci=list(excluded_dpci),
                              filter_mode=filter_mode, include_factories=list(include_factories), include_dpci=list(include_dpci),
                              created=datetime.datetime.now().strftime("%Y-%m-%d %H:%M")),
                pcn={p["program"]: p for p in pcns}, docs=docs,
                changelog=[dict(time=datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), action="新建",
                                note=f"PCN：{'、'.join(p['program'] for p in pcns)}；PO 文件 {len(docs)} 份")],
                problems=list(problems))
