#!/usr/bin/env python3
"""
parse_po_pdfs.py — extract line items from Target / SPS Commerce PO PDFs.

Accepts a .zip of PDFs, a folder, or individual PDFs. Handles CJK filenames
inside Windows-made zips. Emits:

  text.json   — {pdf_path: extracted_text}, a cache; extraction is the slow
                step (~1-2 s/PDF on one core), so a re-run reuses it
  items.json  — {"pos": [...headers...], "items": [...line items...]}
  grid.json   — {"auto": {"<DPCI>": {"<PO#>": qty}}, "superseded": [...]},
                the shape po-grid-update's update_grid.py consumes. Built from
                the LIVE version of each PO only: where the same PO number
                appears more than once, the latest-dated copy wins and the
                others contribute nothing. Summing every copy would double the
                quantities.

Every PO is checked against three self-verification gates. Any failure is
reported in `problems`; an empty `problems` list is the signal that the
regexes still match this season's PDF template.

  gate 1  sum of line totals            == "Purchase Order Total:"
  gate 2  qty x unit price              == line total
  gate 3  sum of line quantities        == "Total Qty:"

USAGE
  python parse_po_pdfs.py --input POs.zip --outdir ./work
  python parse_po_pdfs.py --input ./pdf_folder --outdir ./work --no-cache

Requires: pdfplumber
"""
import argparse, glob, json, os, re, sys, tempfile, warnings, zipfile

warnings.filterwarnings('ignore')

try:
    import pdfplumber
except ImportError:
    sys.exit('pdfplumber not installed. Run: pip install pdfplumber --break-system-packages')


def NUM(s):
    return float(str(s).replace(',', '')) if s not in (None, '') else None


# ---------------------------------------------------------------- unpacking
def unpack(inp, workdir):
    """Return a list of PDF paths from a zip / folder / single file.

    Zip entries keep their folder structure — the folder name is often the
    only record of why a PO batch exists (e.g. a factory transfer).
    """
    if os.path.isdir(inp):
        return sorted(glob.glob(os.path.join(inp, '**', '*.pdf'), recursive=True))
    if inp.lower().endswith('.zip'):
        d = os.path.join(workdir, 'pos')
        os.makedirs(d, exist_ok=True)
        with zipfile.ZipFile(inp) as z:
            for info in z.infolist():
                if info.is_dir():
                    continue
                raw = info.filename
                name = raw
                for enc in ('cp950', 'utf-8', 'cp936'):
                    try:
                        name = raw.encode('cp437').decode(enc)
                        break
                    except (UnicodeEncodeError, UnicodeDecodeError):
                        continue
                if not name.lower().endswith('.pdf'):
                    continue
                target = os.path.join(d, name)
                os.makedirs(os.path.dirname(target), exist_ok=True)
                with open(target, 'wb') as fh:
                    fh.write(z.read(info))
        return sorted(glob.glob(os.path.join(d, '**', '*.pdf'), recursive=True))
    return [inp]


def extract_text(pdfs, cache_path, use_cache=True):
    cache = {}
    if use_cache and cache_path and os.path.exists(cache_path):
        try:
            cache = json.load(open(cache_path, encoding='utf-8'))
        except Exception:
            cache = {}
    out, fresh = {}, 0
    for i, p in enumerate(pdfs):
        if p in cache and not cache[p].startswith('__ERROR__'):
            out[p] = cache[p]
            continue
        try:
            with pdfplumber.open(p) as pdf:
                out[p] = '\n'.join((pg.extract_text() or '') for pg in pdf.pages)
        except Exception as e:                                   # noqa: BLE001
            out[p] = '__ERROR__ ' + str(e)
        fresh += 1
        if fresh % 25 == 0:
            print('  extracted %d/%d' % (i + 1, len(pdfs)), file=sys.stderr, flush=True)
    if cache_path:
        json.dump(out, open(cache_path, 'w', encoding='utf-8'), ensure_ascii=False)
    return out


# ------------------------------------------------------------------ parsing
def g(pat, text, grp=1, flags=0):
    m = re.search(pat, text, flags)
    return m.group(grp).strip() if m else ''


# "1 240434424 Vendors Style 199268803079Vendor color description: ..."
SHIP_WIN = r'Shipping Window:(.{0,400}?)(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})'

LINE_START = re.compile(r'^(\d{1,3}) (\d{9}) Vendors Style (\d{11,13})', re.M)

# "1 240437779 21PT11 191908561226 0.465 5088 Each"
# The UPC is sometimes glued to the next field with no space:
# "1 240111245 26HTT05V 199268798313Product: 2.60 808 Each"
# NOTE (v3.2): the space before "Each" can vanish in pdfplumber's extraction
# when the quantity is wide (5-6 digits) — "356400Each" with no space — while
# a narrower quantity on the row above/below keeps it ("3960 Each"). This is
# a per-row rendering quirk, not a per-PO one: the same PO can have both.
# \s*Each (space optional) below is the fix; it silently dropped exactly the
# component with the largest quantity on each assortment line (the retail
# "Spritz" grass-fill row) while still catching the smaller shipper-count row,
# so the miss was invisible unless someone checked component sums by hand.
PREPACK_ROW = re.compile(
    r'^(\d{1,3}) (\d{9}) (\S+) (\d{11,13})[^\d\n]*?(\d+\.\d+) ([\d,]+)\s*Each\s*$', re.M)

# --- 27C1-era template variant ---------------------------------------------
# Here the line-number row never carries the DPCI; it carries only the UPC,
# type, qty and total, and the DPCI wraps to the row below:
#   "1 Buyers Catalog Number: Vendors Style 199592789032Product: REG Unit
#    Price:256464 Each 202,606.56\n240027702 Number: Product: Spritz 0.79\n
#    Buyers Item Number: 27VG004 16ctPrtyFvrs Resale: 3.00\n..."
# Assortment parents drop the "Vendors Style" label entirely:
#   "1 Buyers Catalog Number: 822826394169Product: AST Unit Price: 1858
#    Each 147,599.52\n240020140 Product: 79.44\nBuyers Item Number:
#    ASSORTMENT LS-2: Resale: 216.00\n..."
LINE_START_V2 = re.compile(
    r'^(\d{1,2}) Buyers Catalog Number:\s*(?:Vendors Style\s*)?'
    r'(\d{11,13})Product:\s*(\w+)\s*Unit\s*Price:\s*([\d,]+)\s*Each\s*([\d,]+\.\d{2})',
    re.M)
DPCI_V2 = re.compile(r'^(\d{9})\s*(?:Number:\s*)?Product:\s*(.*?)\s*([\d,]+\.\d{2,4})\s*$', re.M)
STYLE_V2 = re.compile(r'Buyers Item Number:\s*(.*?)\s*Resale:\s*([\d.]+)', re.S)

# "1 240024442 199592788790 1.10 89184 Each" — no vendor-style token at all
# in this template variant, unlike PREPACK_ROW above.
PREPACK_ROW_V2 = re.compile(
    r'^(\d{1,3}) (\d{9}) (\d{11,13}) ([\d.]+) ([\d,]+)\s*Each\s*$', re.M)

# --- 27C2-era template variant ---------------------------------------------
# A hybrid of the two above: the line-number row wraps "Buyers Catalog
# Number:" like V2, so the DPCI drops to the row below — but the unit price
# stays ON the line-number row like the original, ahead of the quantity:
#   "1 Buyers Catalog Number: Vendors Style 822826351834Product: REG Unit
#    Price: 1.09 590 Each 643.10\n240048085 Number: Product: Spritz Dec
#    Resale: 5.00\nBuyers Item Number: EB2702 Basket Green Wholesale"
# V2 cannot match it: V2 expects the quantity immediately after "Unit Price:".
# The distinguishing token is that extra decimal before the integer quantity.
LINE_START_V3 = re.compile(
    r'^(\d{1,3}) Buyers Catalog Number:\s*(?:Vendors Style\s*)?'
    r'(\d{11,13})Product:\s*(\w+)\s*Unit\s*Price:\s*([\d.]+)\s+([\d,]+)\s*Each\s*'
    r'([\d,]+\.\d{2})', re.M)
# DPCI on the following row; the vendor style follows "Buyers Item Number:".
DPCI_V3 = re.compile(r'^(\d{9})\b', re.M)
STYLE_V3 = re.compile(r'Buyers Item Number:\s*(\S+)')

# "1 240046252 822826351360 Basket Yellow 1.09 14161 Each" — no vendor-style
# token ahead of the UPC (unlike PREPACK_ROW) and a free-text description
# *between* the UPC and the unit price (unlike PREPACK_ROW_V2, which requires
# them adjacent). Anchor on the trailing " Each" and let the description float.
PREPACK_ROW_V3 = re.compile(
    r'^(\d{1,3}) (\d{9}) (\d{11,13})\s*(.*?)\s*(\d+\.\d+) ([\d,]+)\s*Each\s*$', re.M)


def _resale(blk):
    """Resale is the first money-shaped token after 'Resale:'.

    When the PO's columns wrap, the 10-digit Buyers Item Number lands between
    the label and its value ("Resale: 1011139857 # of Inners: 10 5.00"), so a
    plain [\\d.]+ grab returns the item number and every retail check fails.
    Require a decimal and skip bare integers.
    """
    m = re.search(r'Resale:', blk)
    if not m:
        return ''
    tail = blk[m.end():m.end() + 200]
    mm = re.search(r'(\d{1,4}\.\d{2})\b', tail)
    return mm.group(1) if mm else ''


def parse_text(path, txt):
    """Parse one PO's extracted text into (header, [items])."""
    # read the retrieval stamp BEFORE stripping the page chrome that carries it
    stamp = g(r'(\d{4}/\d{1,2}/\d{1,2} \d{1,2}:\d{2}) Fulfillment', txt)
    txt = re.sub(r'^\d{4}/\d{1,2}/\d{1,2} \d{1,2}:\d{2} Fulfillment\s*$', '', txt, flags=re.M)
    txt = re.sub(r'^https://\S+\s+\d+/\d+\s*$', '', txt, flags=re.M)

    # "Order #: 0240-1016166-3891"  ->  dept 0240, PO 1016166, DC 3891
    m = re.search(r'Order #:\s*(\d{4})-(\d{5,8})-(\w{4})', txt)
    if m:
        dept, po, dc = m.group(1), m.group(2), m.group(3)
    else:
        # 27C1-era template variant: no dept prefix baked into Order #, just
        # "Order #: 10001939567-0581" (PO-DC). Dept isn't printed as a single
        # field anywhere reliable in this variant, and nothing downstream
        # consumes it, so fall back to the filename's "240-02" prefix.
        m2 = re.search(r'Order #:\s*(\d{6,11})-(\w{3,5})', txt)
        po, dc = (m2.group(1), m2.group(2)) if m2 else ('', '')
        dept = g(r'^(\d{3}-\d{2})', os.path.basename(path))

    fname = os.path.basename(path)
    head = {
        'po': po, 'dept': dept, 'dc': dc,
        'file': fname, 'path': path,
        'folder': os.path.basename(os.path.dirname(path)),
        # page-1 retrieval stamp — the reliable version marker
        'retrieved': stamp,
        'fname_date': g(r'_(\d{4})\.pdf$', fname),
        # Filenames vary: "..._1044568-T2_1224.pdf", "..._1044568_T2_0506.pdf",
        # "..._4429652-TS_0417_Promotion date 修改.pdf". Match the label
        # wherever it sits rather than anchoring to the end of the name.
        'dclabel': g(r'[-_]([A-Z]\d|TS)(?=[-_.])', fname) or
                   # Some programs (e.g. Mini Seasonal) never put the tranche
                   # label in the filename at all — it's TS/T1/T2/T3/H1/H3 for
                   # full seasonal programs, but BAS/SEA/SET ("Basic" /
                   # "Seasonal" / "Set Order") for this one, printed as
                   # "Assigned by transaction set sender: BAS Basic:" on
                   # page 1. Don't assume the label vocabulary — read
                   # whatever short code sits there.
                   g(r'Assigned by transaction set sender:\s*([A-Z]{2,5})\s+\w+', txt),
        'is_revised': bool(re.search(r'revis', fname, re.I)),
        'po_date': g(r'Order PO Date:.*?\n(\d{2}/\d{2}/\d{4})', txt, 1, re.S),
        'cancel': g(r'Cancel Date:.*?\n(\d{2}/\d{2}/\d{4})', txt, 1, re.S),
        # The date range sits near the "Shipping Window:" label, but on either
        # side of the word "Standard" depending on the PDF variant — anchoring
        # on "Standard" missed 58 of 382 POs in 26C5. Anchor on the label and
        # take the first range after it.
        'ship_start': g(SHIP_WIN, txt, 2, re.S),
        'ship_end': g(SHIP_WIN, txt, 3, re.S),
        'po_type': g(r'PO Type:\s*([^\n]+?)\s+Original Delivery', txt),
        'coo': g(r'Country of Origin:\s*\n?([A-Z]{2})', txt),
        'port': g(r'Origin \(Shipping Point\)\s*([A-Z]{5})', txt),
        'po_total': g(r'Purchase Order Total:\s*([\d,]+\.\d{2})', txt),
        'total_qty': g(r'Total Qty:\s*([\d,]+)', txt),
        'promo_start': g(r'Promotion Start:\s*(\d{2}/\d{2}/\d{4})', txt),
    }

    starts = [(mm.start(), mm) for mm in LINE_START.finditer(txt)]
    items = []
    # A single 27C2 PO can mix both wrapped shapes: most lines carry the unit
    # price on the line-number row (V3), but a line whose price wraps to the
    # row below is V2-shaped. The two patterns are mutually exclusive on any
    # given line — V2 requires the quantity immediately after "Unit Price:",
    # V3 requires a decimal there — so union the matches rather than picking
    # one variant per PO. Choosing either/or silently dropped line 1 of
    # PO 10002036378 (15,720 pieces of 240-04-1101).
    v3_hits = [(mm.start(), 'v3', mm) for mm in LINE_START_V3.finditer(txt)]
    v2_hits = [(mm.start(), 'v2', mm) for mm in LINE_START_V2.finditer(txt)]
    if not starts and (v3_hits or v2_hits):
        starts_v3 = sorted(v3_hits + v2_hits, key=lambda x: x[0])
        for idx, (pos_, variant, mm) in enumerate(starts_v3):
            end = starts_v3[idx + 1][0] if idx + 1 < len(starts_v3) else len(txt)
            blk = txt[pos_:end]
            if variant == 'v3':
                lineno, upc, ptype, unit, qty_s, total = mm.groups()
            else:
                lineno, upc, ptype, qty_s, total = mm.groups()
                dm2 = DPCI_V2.search(blk)
                unit = dm2.group(3) if dm2 else ''
            qty = int(qty_s.replace(',', ''))

            dm = DPCI_V3.search(blk)
            if not dm:
                continue
            sku = dm.group(1)

            sm = STYLE_V3.search(blk)
            style_tok = sm.group(1).strip() if sm else ''
            style = style_tok if re.match(r'^\d{0,2}[A-Za-z]', style_tok) else ''

            resale = _resale(blk)

            items.append({
                'kind': 'ast' if ptype.upper() == 'AST' else 'line',
                'line': int(lineno), 'sku': sku, 'upc': upc,
                'style': style, 'type': ptype.upper(), 'qty': qty,
                'unit': unit, 'total': total, 'resale': resale,
            })

            seen_comp = set()
            for pm in PREPACK_ROW_V3.finditer(blk):
                _, csku, cupc, _desc, cunit, cqty = pm.groups()
                seen_comp.add(csku)
                items.append({
                    'kind': 'prepack', 'parent_sku': sku, 'parent_line': int(lineno),
                    'sku': csku, 'upc': cupc, 'style': '',
                    'unit': cunit, 'qty': int(cqty.replace(',', '')),
                })
            for pm in PREPACK_ROW.finditer(blk):
                _, csku, cstyle, cupc, cunit, cqty = pm.groups()
                if csku in seen_comp:
                    continue
                seen_comp.add(csku)
                items.append({
                    'kind': 'prepack', 'parent_sku': sku, 'parent_line': int(lineno),
                    'sku': csku, 'upc': cupc, 'style': cstyle,
                    'unit': cunit, 'qty': int(cqty.replace(',', '')),
                })
            for pm in PREPACK_ROW_V2.finditer(blk):
                _, csku, cupc, cunit, cqty = pm.groups()
                if csku in seen_comp:
                    continue
                seen_comp.add(csku)
                items.append({
                    'kind': 'prepack', 'parent_sku': sku, 'parent_line': int(lineno),
                    'sku': csku, 'upc': cupc, 'style': '',
                    'unit': cunit, 'qty': int(cqty.replace(',', '')),
                })
    elif starts:
        for idx, (pos_, mm) in enumerate(starts):
            end = starts[idx + 1][0] if idx + 1 < len(starts) else len(txt)
            blk = txt[pos_:end]
            lineno, sku, upc = mm.groups()

            qm = re.search(r'Unit Price:\s*([\d,]+)\s*Each\s*([\d,]+\.\d{2})', blk)
            if not qm:
                continue
            qty, total = int(qm.group(1).replace(',', '')), qm.group(2)
            ptype = 'AST' if re.search(r'Product:\s*AST\b|ASSORTMENT', blk) else 'REG'

            unit = g(r'^Number:[^\n]*?([\d]+\.[\d]+)\s*$', blk, 1, re.M)
            if not unit:
                unit = g(r'^[^\n]*?([\d]+\.[\d]{2,4})\s*$', blk, 1, re.M)
            resale = g(r'Resale:\s*([\d.]+)', blk)
            if not resale:                       # assortment: value wraps to next line
                resale = g(r'Resale:\s*\n[^\n]*?([\d]+\.[\d]{2})\s*$', blk, 1, re.M)

            items.append({
                'kind': 'ast' if ptype == 'AST' else 'line',
                'line': int(lineno), 'sku': sku, 'upc': upc,
                'style': g(r'^(\d{2}[A-Z]{2,4}\d{0,3})\b', blk, 1, re.M),
                'type': ptype, 'qty': qty, 'unit': unit,
                'total': total, 'resale': resale,
            })

            for pm in PREPACK_ROW.finditer(blk):
                _, csku, cstyle, cupc, cunit, cqty = pm.groups()
                items.append({
                    'kind': 'prepack', 'parent_sku': sku, 'parent_line': int(lineno),
                    'sku': csku, 'upc': cupc, 'style': cstyle,
                    'unit': cunit, 'qty': int(cqty.replace(',', '')),
                })
    else:
        starts_v2 = [(mm.start(), mm) for mm in LINE_START_V2.finditer(txt)]
        for idx, (pos_, mm) in enumerate(starts_v2):
            end = starts_v2[idx + 1][0] if idx + 1 < len(starts_v2) else len(txt)
            blk = txt[pos_:end]
            lineno, upc, ptype, qty_s, total = mm.groups()
            qty = int(qty_s.replace(',', ''))

            dm = DPCI_V2.search(blk)
            sku = dm.group(1) if dm else ''
            unit = dm.group(3) if dm else ''

            sm = STYLE_V2.search(blk)
            style_raw = sm.group(1).strip() if sm else ''
            resale = sm.group(2) if sm else ''
            # First whitespace-free token looks like a real vendor style code
            # ("27VG004"); assortment parents instead start with the word
            # "ASSORTMENT" — leave style blank there since the GRID's STYLE
            # column is filled from the master, not this field.
            style_tok = style_raw.split(' ', 1)[0] if style_raw else ''
            style = style_tok if re.match(r'^\d{2}[A-Za-z]', style_tok) else ''

            if not sku:
                continue

            items.append({
                'kind': 'ast' if ptype.upper() == 'AST' else 'line',
                'line': int(lineno), 'sku': sku, 'upc': upc,
                'style': style, 'type': ptype.upper(), 'qty': qty,
                'unit': unit, 'total': total, 'resale': resale,
            })

            for pm in PREPACK_ROW_V2.finditer(blk):
                _, csku, cupc, cunit, cqty = pm.groups()
                items.append({
                    'kind': 'prepack', 'parent_sku': sku, 'parent_line': int(lineno),
                    'sku': csku, 'upc': cupc, 'style': '',
                    'unit': cunit, 'qty': int(cqty.replace(',', '')),
                })

    for it in items:
        it.update({'po': po, 'dc': dc, 'file': fname})
    return head, items


def _version_key(h):
    """Date decides which copy of a PO number is live. Never `PO Type:` —
    SPS reprints a revised PO still labelled Original."""
    m = re.match(r'(\d{4})/(\d{1,2})/(\d{1,2})\s+(\d{1,2}):(\d{2})', h.get('retrieved') or '')
    if m:
        return tuple(int(x) for x in m.groups())
    fd = h.get('fname_date') or ''
    if len(fd) == 4:
        mo, d = int(fd[:2]), int(fd[2:])
        return (0, mo, d, 0, 0)
    return (0, 0, 0, 0, 0)


def build_grid_json(heads, items):
    groups = {}
    for h in heads:
        groups.setdefault((h['po'], h['dc']), []).append(h)
    live, superseded = set(), []
    for key, hs in groups.items():
        hs.sort(key=_version_key, reverse=True)
        live.add(hs[0]['file'])
        for lose in hs[1:]:
            superseded.append({'po': key[0], 'dc': key[1],
                               'dropped': lose['file'], 'kept': hs[0]['file'],
                               'qty_changed': lose.get('total_qty') != hs[0].get('total_qty'),
                               'from_qty': lose.get('total_qty'),
                               'to_qty': hs[0].get('total_qty')})
    grid = {}
    for i in items:
        if i['file'] not in live or i['kind'] not in ('line', 'prepack'):
            continue
        d = '%s-%s-%s' % (i['sku'][:3], i['sku'][3:5], i['sku'][5:])
        grid.setdefault(d, {})
        grid[d][i['po']] = grid[d].get(i['po'], 0) + i['qty']

    # Same-number revisions are handled above. The dangerous duplicates carry
    # DIFFERENT PO numbers — a factory transfer re-issued under a new number, or
    # a same-day twin — and no date rule can see them. Fingerprint each live PO
    # by its (sku, qty) set and flag collisions so the caller resolves them
    # before any of this is written to a GRID.
    per_file = {}
    for i in items:
        if i['file'] in live and i['kind'] in ('line', 'prepack'):
            per_file.setdefault(i['file'], []).append((i['sku'], i['qty']))
    fp, byfile = {}, {h['file']: h for h in heads}
    for f, rowset in per_file.items():
        fp.setdefault(tuple(sorted(rowset)), []).append(f)
    dupes = []
    for key, files in fp.items():
        pos_ = {byfile[f]['po'] for f in files}
        if len(pos_) > 1:
            dupes.append({'skus': len(key), 'units': sum(q for _, q in key),
                          'copies': [{'po': byfile[f]['po'], 'dc': byfile[f]['dc'],
                                      'folder': byfile[f].get('folder'),
                                      'po_total': byfile[f].get('po_total'),
                                      'file': f} for f in sorted(files)]})
    return {'auto': grid, 'superseded': superseded, 'duplicate_po_groups': dupes}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', required=True)
    ap.add_argument('--outdir', default='./work')
    ap.add_argument('--no-cache', action='store_true')
    a = ap.parse_args()
    os.makedirs(a.outdir, exist_ok=True)

    pdfs = unpack(a.input, a.outdir)
    if not pdfs:
        sys.exit('No PDFs found in %s' % a.input)

    texts = extract_text(pdfs, os.path.join(a.outdir, 'text.json'),
                         use_cache=not a.no_cache)

    heads, all_items, problems = [], [], []
    for path in sorted(texts):
        txt = texts[path]
        if txt.startswith('__ERROR__'):
            problems.append({'file': os.path.basename(path), 'kind': 'extract',
                             'issue': txt[:150]})
            continue
        try:
            h, items = parse_text(path, txt)
        except Exception as e:                                   # noqa: BLE001
            problems.append({'file': os.path.basename(path), 'kind': 'exception',
                             'issue': repr(e)})
            continue
        if not h['po']:
            problems.append({'file': os.path.basename(path), 'kind': 'no_po',
                             'issue': 'PO number not found — PDF template may have changed'})
            continue

        # --- combined-print guard (v3.3) ------------------------------------
        # An SPS "print all" PDF holds many orders in one file. Fed here
        # directly, parse_text() reads the FIRST order's header and sums EVERY
        # order's line items against it, so the gates fire as total_mismatch +
        # qty_mismatch and the real cause is invisible. Count distinct
        # "Order #:" values and say what is actually wrong.
        _orders = set(re.findall(r'Order #:\s*(\d{6,11})', txt))
        if len(_orders) > 1:
            problems.append({
                'file': os.path.basename(path), 'po': h['po'],
                'kind': 'combined_print',
                'issue': ('%d distinct orders in one PDF — this is a combined '
                          'SPS print. Split it first: python scripts/'
                          'split_combined.py --input "%s" --outdir <dir> '
                          '--prefix <program DPCI prefix>'
                          % (len(_orders), os.path.basename(path)))})
            continue
        # --------------------------------------------------------------------

        lines = [i for i in items if i['kind'] in ('line', 'ast')]
        lt = sum(NUM(i['total']) for i in lines if i.get('total'))
        pt = NUM(h['po_total'])
        if pt is not None and abs(lt - pt) > 0.05:
            problems.append({'file': h['file'], 'po': h['po'], 'kind': 'total_mismatch',
                             'issue': 'line totals %.2f != PO total %.2f' % (lt, pt)})
        for i in lines:
            if i.get('unit') and i.get('total'):
                calc = NUM(i['unit']) * i['qty']
                if abs(calc - NUM(i['total'])) > 0.05:
                    problems.append({'file': h['file'], 'po': h['po'], 'kind': 'unit_math',
                                     'issue': 'sku %s: %s x %d = %.2f != %s'
                                              % (i['sku'], i['unit'], i['qty'], calc, i['total'])})
        tq = NUM(h['total_qty'])
        sq = sum(i['qty'] for i in lines)
        if tq is not None and abs(sq - tq) > 0.5:
            problems.append({'file': h['file'], 'po': h['po'], 'kind': 'qty_mismatch',
                             'issue': 'line qty %d != Total Qty %d' % (sq, tq)})

        h['n_lines'] = len(lines)
        h['n_prepack'] = sum(1 for i in items if i['kind'] == 'prepack')
        heads.append(h)
        all_items.extend(items)

    json.dump({'pos': heads, 'items': all_items},
              open(os.path.join(a.outdir, 'items.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=1)

    gj = build_grid_json(heads, all_items)
    json.dump(gj, open(os.path.join(a.outdir, 'grid.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=1)

    by_kind = {}
    for pr in problems:
        by_kind[pr['kind']] = by_kind.get(pr['kind'], 0) + 1

    print(json.dumps({
        'pdfs': len(pdfs), 'pos_parsed': len(heads),
        'distinct_po_numbers': len({h['po'] for h in heads}),
        'regular_lines': sum(1 for i in all_items if i['kind'] == 'line'),
        'assortment_lines': sum(1 for i in all_items if i['kind'] == 'ast'),
        'prepack_rows': sum(1 for i in all_items if i['kind'] == 'prepack'),
        'distinct_sku': len({i['sku'] for i in all_items}),
        'pos_superseded': len(gj['superseded']),
        'superseded_with_qty_change': sum(1 for x in gj['superseded'] if x['qty_changed']),
        'duplicate_po_groups': len(gj['duplicate_po_groups']),
        'problems_by_kind': by_kind,
        'problems_sample': problems[:20],
        'outdir': a.outdir,
    }, ensure_ascii=False, indent=2))
    return 1 if problems else 0


if __name__ == '__main__':
    sys.exit(main())
