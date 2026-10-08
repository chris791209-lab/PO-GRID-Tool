#!/usr/bin/env python3
"""
update_grid.py — Write PO quantities into a fixed-format PO GRID workbook
using XML surgery, preserving in-cell images (richData), styles and layout.

WHY XML SURGERY: PO GRID files store product photos as *in-cell images*
(xl/richData/* + vm="" attributes). openpyxl does NOT support richData and
silently destroys those images on any read/write round-trip. Editing the
sheet XML directly is the only safe method.

USAGE
  python update_grid.py --grid GRID.xlsx --data data.json --out OUT.xlsx
  python update_grid.py --grid GRID.xlsx --data data.json --out OUT.xlsx --dry-run

data.json shape:
{
  "<sheet name or 'auto'>": {          # 'auto' = locate DPCI on any sheet
     "240-02-4442": {"10001939610": 89184, "10001939999": 1200},
     "240-02-5838": {"10001939610": 44592}
  }
}
Quantities REPLACE the existing cell value for that (DPCI, PO) pair.

Exit code 0 = success, 1 = error. A JSON report is printed to stdout.
"""
import argparse, json, os, re, shutil, sys, tempfile, zipfile
from collections import OrderedDict

COL_RE = re.compile(r'^([A-Z]+)(\d+)$')

# DC code -> destination, kept in step with po-grid-create/scripts/build_grid.py.
# An unknown code is written through as the raw number, never guessed.
DEST = {
    '0581': 'LAS',
    '3891': 'SAVANNAH',
    '3758': 'CHARLESTON',
    '3887': 'HOUSTON',
    '3890': 'PNW',
}

META_BY_PO = {}


def col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def num_to_col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


class Sheet:
    """Minimal read/modify wrapper over one worksheet XML."""

    def __init__(self, path, shared):
        self.path = path
        self.xml = open(path, encoding='utf-8').read()
        self.shared = shared

    # ---------- reading ----------
    def cell_raw(self, ref):
        # IMPORTANT: try the self-closing form FIRST and never let [^>]* cross a
        # '>'. A naive combined pattern lets a self-closing cell fall through to
        # the '>.*?</c>' branch, which then swallows every following cell up to
        # the next </c> — silently deleting cells and even whole rows.
        m = re.search(r'<c r="%s"[^>]*/>' % ref, self.xml)
        if m:
            return m.group(0)
        m = re.search(r'<c r="%s"[^>]*>.*?</c>' % ref, self.xml, re.S)
        return m.group(0) if m else None

    def is_text(self, ref):
        """True if the cell currently holds text (shared/inline string)."""
        raw = self.cell_raw(ref)
        return bool(raw) and ('t="s"' in raw or 't="inlineStr"' in raw)

    def cell_text(self, ref):
        raw = self.cell_raw(ref)
        if not raw:
            return None
        if 't="s"' in raw:
            m = re.search(r'<v>(\d+)</v>', raw)
            return self.shared[int(m.group(1))] if m else None
        if 't="inlineStr"' in raw:
            m = re.search(r'<t[^>]*>(.*?)</t>', raw, re.S)
            return m.group(1) if m else None
        m = re.search(r'<v>([^<]*)</v>', raw)
        return m.group(1) if m else None

    def row_cells(self, row):
        m = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % row, self.xml, re.S)
        if not m:
            return []
        return re.findall(r'<c r="([A-Z]+%d)"' % row, m.group(1))

    # ---------- writing ----------
    def set_number(self, ref, value):
        """Replace cell value with a number, keeping its style (s=) attribute."""
        raw = self.cell_raw(ref)
        if raw is None:
            return False
        sm = re.search(r'\ss="(\d+)"', raw)
        style = ' s="%s"' % sm.group(1) if sm else ''
        new = '<c r="%s"%s><v>%s</v></c>' % (ref, style, value)
        self.xml = self.xml.replace(raw, new, 1)
        return True

    def save(self):
        open(self.path, 'w', encoding='utf-8').write(self.xml)


def load_shared(base):
    p = os.path.join(base, 'xl', 'sharedStrings.xml')
    if not os.path.exists(p):
        return []
    x = open(p, encoding='utf-8').read()
    out = []
    for si in re.findall(r'<si>(.*?)</si>', x, re.S):
        out.append(''.join(re.findall(r'<t[^>]*>(.*?)</t>', si, re.S)))
    return out


def sheet_name_map(base):
    """Return OrderedDict {display name: worksheets/sheetN.xml path}."""
    wb = open(os.path.join(base, 'xl', 'workbook.xml'), encoding='utf-8').read()
    rels = open(os.path.join(base, 'xl', '_rels', 'workbook.xml.rels'), encoding='utf-8').read()
    # attribute order differs by writer: Excel/LibreOffice put Id first,
    # openpyxl puts Target first — read each <Relationship> tag on its own
    rid2t = {}
    for tag in re.findall(r'<Relationship\b[^>]*>', rels):
        i_ = re.search(r'\bId="([^"]+)"', tag)
        t_ = re.search(r'\bTarget="([^"]+)"', tag)
        if i_ and t_:
            rid2t[i_.group(1)] = t_.group(1)
    out = OrderedDict()
    for name, rid in re.findall(r'<sheet\b[^>]*?\bname="([^"]+)"[^>]*?\br:id="([^"]+)"', wb):
        t = rid2t.get(rid, '')
        t = t.split('/')[-1]
        out[name] = os.path.join(base, 'xl', 'worksheets', t)
    return out


# ---------------------------------------------------------------------------
# Column insertion (added 2026-08-18) — see the module notes below.
# ---------------------------------------------------------------------------
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"')
PO_COL_WIDTH = 17.375          # frozen layout
FIRST_PO_COL = 14              # column N



def _mmdd(date_str):
    """'11/24/2026' -> '11/24'. Returns '' for anything unparseable."""
    m = re.match(r'(\d{1,2})/(\d{1,2})/\d{4}', str(date_str or ''))
    return '%02d/%02d' % (int(m.group(1)), int(m.group(2))) if m else ''


def ship_window(meta):
    a, b = _mmdd(meta.get('ship_start')), _mmdd(meta.get('ship_end'))
    return '%s-%s' % (a, b) if a and b else ''


def sort_key(meta):
    """POs sort by ship-start date; the PO number breaks ties."""
    m = re.match(r'(\d{1,2})/(\d{1,2})/(\d{4})', str(meta.get('ship_start') or ''))
    if m:
        return ('%s%02d%02d' % (m.group(3), int(m.group(1)), int(m.group(2))),
                meta.get('po', ''))
    return ('9999', meta.get('po', ''))


# ---------------------------------------------------------------- XML helpers

def shift_cells_right(xml, at_col):
    """Move every cell whose column >= at_col one column to the right."""
    def repl(mo):
        col, row = mo.group(1), mo.group(2)
        n = col_to_num(col)
        if n >= at_col:
            col = num_to_col(n + 1)
        return '<c r="%s%s"' % (col, row)

    xml = CELL_RE.sub(repl, xml)
    # spans are a hint Excel recomputes; stale ones truncate a row's display
    xml = re.sub(r'\s+spans="[^"]*"', '', xml)
    # widen the declared dimension so Excel does not clip the new column
    m = re.search(r'<dimension ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/>', xml)
    if m:
        end = num_to_col(col_to_num(m.group(3)) + 1)
        xml = xml.replace(m.group(0),
                          '<dimension ref="%s%s:%s%s"/>'
                          % (m.group(1), m.group(2), end, m.group(4)))
    return xml


def shift_cols_element(xml, at_col, width=PO_COL_WIDTH):
    """Shift <col> width entries to make room at at_col.

    The create-side build writes ONE <col> entry spanning the whole PO block,
    so the new column should widen that range rather than add an entry of its
    own — an overlapping entry is legal but leaves the file different from a
    freshly built one for no reason.
    """
    m = re.search(r'<cols>(.*?)</cols>', xml, re.S)
    if not m:
        return xml
    entries = re.findall(r'<col [^/]*/>', m.group(1))
    out, covered = [], False
    for e in entries:
        mn = int(re.search(r'min="(\d+)"', e).group(1))
        mx = int(re.search(r'max="(\d+)"', e).group(1))
        if mn <= at_col <= mx:
            e = re.sub(r'max="\d+"', 'max="%d"' % (mx + 1), e)
            covered = True
        elif mn > at_col:
            e = re.sub(r'min="\d+"', 'min="%d"' % (mn + 1), e)
            e = re.sub(r'max="\d+"', 'max="%d"' % (mx + 1), e)
        out.append(e)
    if not covered:
        out.append('<col collapsed="false" customWidth="true" hidden="false" '
                   'outlineLevel="0" max="%d" min="%d" style="0" width="%s"/>'
                   % (at_col, at_col, width))
    out.sort(key=lambda e: int(re.search(r'min="(\d+)"', e).group(1)))
    return xml.replace(m.group(0), '<cols>' + ''.join(out) + '</cols>')


def formulaize_totals(xml, total_col_num, first_po, last_po):
    """Restore the PO TOTAL column to =SUM() over the PO block.

    update_grid recomputes each touched row's total and writes it as a plain
    number; left that way the GRID stops self-updating the moment anyone edits
    a quantity by hand. The computed figure is kept as the cached value so the
    file still reads correctly before Excel recalculates.
    """
    tc = num_to_col(total_col_num)
    n = 0

    def repl(mo):
        nonlocal n
        n += 1
        row, style, val = mo.group(1), mo.group(2), mo.group(3)
        return ('<c r="%s%s" s="%s" t="n"><f aca="false">SUM(%s%s:%s%s)</f>'
                '<v>%s</v></c>'
                % (tc, row, style, num_to_col(first_po), row,
                   num_to_col(last_po), row, val))

    # set_number writes <c r="T7" s="26"><v>55308</v></c> — no t="n" — so the
    # type attribute is optional here.
    xml = re.sub(r'<c r="%s(\d+)" s="(\d+)"(?: t="n")?><v>([^<]*)</v></c>' % tc,
                 repl, xml)
    return xml, n


def _row_block(xml, row):
    """Return (whole_match, inner_xml, is_self_closing) for one <row>."""
    m = re.search(r'<row r="%d"[^>]*?/>' % row, xml)
    if m:
        return m.group(0), '', True
    m = re.search(r'<row r="%d"[^>]*?>(.*?)</row>' % row, xml, re.S)
    if m:
        return m.group(0), m.group(1), False
    return None, None, None


def put_cell(xml, row, col_num, style, value=None, kind='inline'):
    """Insert or replace one cell, keeping the row's cells in column order.

    kind: 'inline' writes an inline string (no sharedStrings edit needed),
          'n' writes a number, 'blank' writes a styled empty cell.
    """
    ref = '%s%d' % (num_to_col(col_num), row)
    if kind == 'n':
        cell = '<c r="%s" s="%s" t="n"><v>%s</v></c>' % (ref, style, value)
    elif kind == 'blank':
        cell = '<c r="%s" s="%s"/>' % (ref, style)
    else:
        text = (str(value).replace('&', '&amp;').replace('<', '&lt;')
                .replace('>', '&gt;'))
        cell = ('<c r="%s" s="%s" t="inlineStr"><is><t xml:space="preserve">'
                '%s</t></is></c>' % (ref, style, text))

    whole, inner, self_closing = _row_block(xml, row)
    if whole is None:
        return xml                                    # row absent; nothing to do
    if self_closing:
        opening = whole[:-2] + '>'
        return xml.replace(whole, opening + cell + '</row>')

    cells = re.findall(r'<c r="[A-Z]+\d+".*?(?:/>|</c>)', inner, re.S)
    kept = [c for c in cells
            if col_to_num(CELL_RE.match(c).group(1)) != col_num]
    kept.append(cell)
    kept.sort(key=lambda c: col_to_num(CELL_RE.match(c).group(1)))
    opening = whole[:whole.index('>') + 1]
    return xml.replace(whole, opening + ''.join(kept) + '</row>')


def cell_style(xml, ref):
    m = re.search(r'<c r="%s"[^>]*\bs="(\d+)"' % ref, xml)
    return m.group(1) if m else '0'


# PO-type fills — must match build_grid.py's TYPE_FILL (and verify_layout.py)
TYPE_FILL = {'TS': 'FFF2CC', 'T1': 'E2EFDA', 'T2': 'DDEBF7', 'T3': 'E4DFEC', 'H1': 'D9E1F2',
             'H3': 'EDEDED', 'BAS': 'FCE4D6', 'SEA': 'D9D2E9', 'SET': 'D0E0E3',
             'NWT': 'FFF2CC', 'SPO': 'E2EFDA'}


def clone_style_with_fill(base, xf_id, hex_rgb):
    """Append a copy of cellXfs[xf_id] whose fill is a solid hex_rgb; return the new xf index.
    Used when a PO-type label is new to the workbook, so the inserted column gets the
    correct colour instead of the neighbour's."""
    p = os.path.join(base, 'xl', 'styles.xml')
    x = open(p, encoding='utf-8').read()
    m = re.search(r'(<cellXfs\b[^>]*>)(.*?)(</cellXfs>)', x, re.S)
    xfs = re.findall(r'<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>', m.group(2), re.S)
    src = xfs[int(xf_id)] if xf_id is not None and int(xf_id) < len(xfs) else '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    fill = ('<fill><patternFill patternType="solid"><fgColor rgb="FF%s"/>'
            '<bgColor indexed="64"/></patternFill></fill>' % hex_rgb)
    x, fill_id = _append_children(x, 'fills', 'fill', [fill])
    new_xf = re.sub(r'fillId="\d+"', 'fillId="%d"' % fill_id, src, count=1)
    if 'fillId=' not in new_xf:
        new_xf = new_xf.replace('<xf ', '<xf fillId="%d" ' % fill_id, 1)
    if 'applyFill' not in new_xf:
        new_xf = new_xf.replace('<xf ', '<xf applyFill="1" ', 1)
    x, new_id = _append_children(x, 'cellXfs', 'xf', [new_xf])
    open(p, 'w', encoding='utf-8').write(x)
    return str(new_id)


def find_type_style(sheets_xml, label):
    """Style ids already used for this PO-type label, as (row1, row3).

    The frozen layout fills BOTH row 1 and row 3 with the type colour, and the
    two rows use different style ids. Copying row 3 from a neighbouring column
    gives it the neighbour's colour, which verify_layout catches. So find a
    column that already carries this label and take both of its styles.

    Returns (None, None) if the label has never appeared in this workbook —
    there is no style with the right fill and the caller must say so rather
    than pick a wrong colour silently.
    """
    for xml, strings in sheets_xml:
        for mo in re.finditer(r'<c r="([A-Z]+)1" s="(\d+)" t="s"><v>(\d+)</v></c>',
                              xml):
            idx = int(mo.group(3))
            if idx >= len(strings) or strings[idx] != label:
                continue
            col = mo.group(1)
            m3 = re.search(r'<c r="%s3"[^>]*\bs="(\d+)"' % col, xml)
            return mo.group(2), (m3.group(1) if m3 else None)
    return None, None


def rewrite_totals(xml, total_col_num, first_po, last_po):
    """Re-emit every SUM in the PO TOTAL column over the widened range."""
    tc = num_to_col(total_col_num)
    n = 0

    def repl(mo):
        nonlocal n
        n += 1
        row = mo.group(1)
        return ('<c r="%s%s" s="%s" t="n"><f aca="false">SUM(%s%s:%s%s)</f>'
                % (tc, row, mo.group(2), num_to_col(first_po), row,
                   num_to_col(last_po), row))

    # t="n" is optional here: formulaize_totals adds it to the row it just
    # touched, but every other row's pre-existing SUM formula never had it.
    # Requiring it caused rewrite_totals to silently skip all untouched
    # rows on column insert, leaving their ranges stale (undercounting).
    xml = re.sub(r'<c r="%s(\d+)" s="(\d+)"(?: t="n")?><f[^>]*>SUM\([^)]*\)</f>' % tc,
                 repl, xml)

    # A row whose TOTAL cell was a *shared-formula child* — a self-closing
    # <f t="shared" si="N"/> with no visible formula text — never matches
    # the regex above, since there is no "SUM(...)" text to find. If the
    # group's master cell (the one row that actually carried the formula
    # text and a ref="..." range) gets caught and rewritten by the pass
    # above, its t="shared"/ref attributes are dropped and it becomes a
    # plain formula — orphaning every child still pointing at that si
    # group. Excel then reports "removed records: shared formula" on open.
    # Found 2026-08-31 on a GRID whose total column had a shared group
    # spanning 6 rows where only 5 had ever been touched by an update run;
    # the 6th (never touched) stayed an orphaned reference indefinitely.
    # Fix: any leftover <f t="shared" si="N"/> in this column, regardless
    # of whether its group's master survived, gets its own standalone SUM
    # too — matching what every sibling row in the same column already is.
    def repl_orphan(mo):
        nonlocal n
        n += 1
        row = mo.group(1)
        return ('<c r="%s%s" s="%s" t="n"><f aca="false">SUM(%s%s:%s%s)</f>'
                % (tc, row, mo.group(2), num_to_col(first_po), row,
                   num_to_col(last_po), row))

    xml = re.sub(
        r'<c r="%s(\d+)" s="(\d+)"(?: t="n")?><f t="shared" si="\d+"/>' % tc,
        repl_orphan, xml)
    return xml, n



VALIDATION_SHEET = 'Validation Summary'


# ---------------------------------------------------------------------------
# Optional: rewrite the Validation Summary tab from a caller-supplied summary
# (--summary-json). Used by the TG Team PO platform so the updated GRID shows
# the same Chinese summary as its validation report, followed by this run's
# call-outs. Pure XML: new styles are appended to styles.xml, the summary
# sheet's sheetData is replaced, nothing else in the workbook is touched.
# ---------------------------------------------------------------------------
def _xml_text(v):
    return (str(v).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            .replace('"', '&quot;'))


def _append_children(x, tag, child, new_items):
    """Append child elements to <tag> in styles.xml; return (xml, first new index)."""
    m = re.search(r'<%s\b[^>]*/>' % tag, x)
    if m:                                   # self-closing, empty
        x = x[:m.start()] + '<%s count="0"></%s>' % (tag, tag) + x[m.end():]
    m = re.search(r'(<%s\b[^>]*>)(.*?)(</%s>)' % (tag, tag), x, re.S)
    if not m:
        return x, None
    start = len(re.findall(r'<%s\b' % child, m.group(2)))
    inner = m.group(2) + ''.join(new_items)
    open_tag = re.sub(r'count="\d+"', 'count="%d"' % (start + len(new_items)), m.group(1))
    if 'count=' not in open_tag:
        open_tag = open_tag[:-1] + ' count="%d">' % (start + len(new_items))
    return x[:m.start()] + open_tag + inner + m.group(3) + x[m.end():], start


def add_summary_styles(base):
    p = os.path.join(base, 'xl', 'styles.xml')
    x = open(p, encoding='utf-8').read()
    # number formats
    money_code = '&quot;$&quot;#,##0.00##'
    pct_code = '0.0&quot;%&quot;'
    x = re.sub(r'<numFmts\b([^>]*?)\s*/>', r'<numFmts\1></numFmts>', x)
    if '<numFmts' not in x:
        x = re.sub(r'(<styleSheet\b[^>]*>)', r'\1<numFmts count="0"></numFmts>', x, count=1)
    ids = [int(i) for i in re.findall(r'numFmtId="(\d+)"', re.search(r'<numFmts\b.*?</numFmts>', x, re.S).group(0))]
    money_id = max(ids + [190]) + 1
    pct_id = money_id + 1
    x, _ = _append_children(x, 'numFmts', 'numFmt', [
        '<numFmt numFmtId="%d" formatCode="%s"/>' % (money_id, money_code),
        '<numFmt numFmtId="%d" formatCode="%s"/>' % (pct_id, pct_code)])
    fonts = ['<font><sz val="10"/><name val="Arial"/></font>',
             '<font><b/><sz val="10"/><name val="Arial"/></font>',
             '<font><b/><sz val="14"/><name val="Arial"/></font>',
             '<font><b/><i/><sz val="12"/><name val="Arial"/></font>',
             '<font><sz val="9"/><color rgb="FF808080"/><name val="Arial"/></font>',
             '<font><b/><sz val="11"/><name val="Arial"/></font>',
             '<font><b/><sz val="10"/><color rgb="FFC00000"/><name val="Arial"/></font>',
             '<font><b/><sz val="10"/><color rgb="FF548235"/><name val="Arial"/></font>']
    x, f0 = _append_children(x, 'fonts', 'font', fonts)
    fills = ['<fill><patternFill patternType="solid"><fgColor rgb="FF%s"/><bgColor indexed="64"/></patternFill></fill>' % c
             for c in ('D9D9D9', 'E2EFDA', 'F8CBAD')]
    x, fl0 = _append_children(x, 'fills', 'fill', fills)
    spec = {   # name: (numFmtId, font offset, fill offset or None)
        'plain': (0, 0, None), 'bold': (0, 1, None), 'title': (0, 2, None), 'subtitle': (0, 3, None),
        'run': (0, 4, None), 'section': (0, 5, None), 'hdr': (0, 1, 0), 'ok': (0, 0, 1), 'review': (0, 0, 2),
        'money': (money_id, 0, None), 'qty': (3, 0, None), 'pct': (pct_id, 0, None),
        'red': (0, 6, None), 'green': (0, 7, None)}
    xfs, names = [], []
    for name, (nf, fo, fi) in spec.items():
        xfs.append('<xf numFmtId="%d" fontId="%d" fillId="%d" borderId="0" xfId="0" applyFont="1"%s%s/>'
                   % (nf, f0 + fo, 0 if fi is None else fl0 + fi,
                      ' applyFill="1"' if fi is not None else '',
                      ' applyNumberFormat="1"' if nf else ''))
        names.append(name)
    x, x0 = _append_children(x, 'cellXfs', 'xf', xfs)
    open(p, 'w', encoding='utf-8').write(x)
    return {n: x0 + i for i, n in enumerate(names)}


def write_summary_xml(sheet_path, data, styles, inserted, attention):
    money = set(data.get('money_cols', []))
    qty = set(data.get('qty_cols', []))
    rows = []

    def put(r, cells):
        out = []
        for col, v, st in cells:
            ref = '%s%d' % (num_to_col(col), r)
            if v is None or v == '':
                continue
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                out.append('<c r="%s" s="%d"><v>%s</v></c>' % (ref, styles[st], repr(v) if isinstance(v, float) else v))
            else:
                out.append('<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                           % (ref, styles[st], _xml_text(v)))
        rows.append('<row r="%d">%s</row>' % (r, ''.join(out)))

    put(1, [(1, data.get('title', ''), 'title')])
    put(2, [(1, 'PO Validation Summary', 'subtitle')])
    if data.get('run'):
        put(3, [(1, data['run'], 'run')])
    r = 5
    for k, v in data.get('headline', []):
        st = 'plain'
        if k == '結論':
            st = 'red' if '需確認' in str(v) else 'green'
        put(r, [(1, k, 'bold'), (2, v, st)])
        r += 1
    r += 1
    put(r, [(1, 'Check 檢核項目', 'hdr'), (2, 'Count', 'hdr'), (3, 'Status', 'hdr')])
    r += 1
    for label, n, status in data.get('checks', []):
        st = 'review' if status == 'REVIEW' else ('ok' if status == 'OK' else 'plain')
        put(r, [(1, label, st), (2, n, st), (3, status, st)])
        r += 1

    details = list(data.get('details', []))
    if inserted:
        details.append({'title': '本次更新：新增的 PO 欄',
                        'columns': ['PO', '工作表', '欄', '類型', 'Ship Window', '目的地'],
                        'rows': [[i.get('po'), i.get('sheet'), i.get('column'), i.get('label'),
                                  i.get('window'), i.get('dest')] for i in inserted]})
    if attention:
        details.append({'title': '本次更新：需要人工處理',
                        'columns': ['DPCI / PO', '工作表', '問題', '建議'],
                        'rows': [[a_.get('dpci') or a_.get('po') or '', a_.get('sheet', ''),
                                  a_.get('issue', '') + (' (POs: %s)' % ', '.join(a_['pos'][:6]) if a_.get('pos') else ''),
                                  a_.get('next_step', '')] for a_ in attention]})
    for d in details:
        r += 1
        put(r, [(1, d['title'], 'section')])
        r += 1
        cols = d['columns']
        put(r, [(j, h, 'hdr') for j, h in enumerate(cols, 1)])
        r += 1
        for row in d['rows']:
            cells = []
            for j, v in enumerate(row, 1):
                st = 'plain'
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    st = ('money' if cols[j - 1] in money else 'qty' if cols[j - 1] in qty
                          else 'pct' if cols[j - 1] == '差異 %' else 'plain')
                cells.append((j, v, st))
            put(r, cells)
            r += 1

    xml = open(sheet_path, encoding='utf-8').read()
    body = '<sheetData>%s</sheetData>' % ''.join(rows)
    if re.search(r'<sheetData\s*/>', xml):
        xml = re.sub(r'<sheetData\s*/>', body, xml, count=1)
    else:
        xml = re.sub(r'<sheetData>.*?</sheetData>', lambda m: body, xml, count=1, flags=re.S)
    cols_xml = ('<cols><col min="1" max="1" width="54" customWidth="1"/>'
                + ''.join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i, i, w)
                          for i, w in zip(range(2, 9), [26, 22, 22, 22, 40, 14, 14])) + '</cols>')
    if '<cols>' in xml:
        xml = re.sub(r'<cols>.*?</cols>', lambda m: cols_xml, xml, count=1, flags=re.S)
    else:
        xml = xml.replace('<sheetData>', cols_xml + '<sheetData>', 1)
    for tag in ('mergeCells', 'conditionalFormatting', 'dataValidations'):
        xml = re.sub(r'<%s\b.*?</%s>' % (tag, tag), '', xml, flags=re.S)
    xml = re.sub(r'<dimension ref="[^"]*"/>', '<dimension ref="A1:H%d"/>' % max(r, 1), xml)
    open(sheet_path, 'w', encoding='utf-8').write(xml)
    return r


def add_summary_sheet_first(base, name=None):
    """Create an empty first-tab worksheet (for GRIDs built by hand without one)."""
    name = name or VALIDATION_SHEET
    ws_dir = os.path.join(base, 'xl', 'worksheets')
    n = 1
    while os.path.exists(os.path.join(ws_dir, 'sheet%d.xml' % n)):
        n += 1
    fname = 'sheet%d.xml' % n
    open(os.path.join(ws_dir, fname), 'w', encoding='utf-8').write(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/><sheetData/>'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')
    rels_p = os.path.join(base, 'xl', '_rels', 'workbook.xml.rels')
    rels = open(rels_p, encoding='utf-8').read()
    used = set(re.findall(r'\bId="([^"]+)"', rels))
    k = 1
    while 'rIdVS%d' % k in used:
        k += 1
    rid = 'rIdVS%d' % k
    rels = rels.replace('</Relationships>',
                        '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                        'relationships/worksheet" Target="worksheets/%s"/></Relationships>' % (rid, fname))
    open(rels_p, 'w', encoding='utf-8').write(rels)
    ct_p = os.path.join(base, '[Content_Types].xml')
    ct = open(ct_p, encoding='utf-8').read()
    ct = ct.replace('</Types>', '<Override PartName="/xl/worksheets/%s" ContentType="application/'
                    'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>' % fname)
    open(ct_p, 'w', encoding='utf-8').write(ct)
    wb_p = os.path.join(base, 'xl', 'workbook.xml')
    wb = open(wb_p, encoding='utf-8').read()
    ids = [int(i) for i in re.findall(r'sheetId="(\d+)"', wb)]
    ns = ('' if re.search(r'<workbook\b[^>]*xmlns:r=', wb)
          else ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
    wb = re.sub(r'(<sheets>)', lambda m: m.group(1) + '<sheet%s name="%s" sheetId="%d" r:id="%s"/>'
                % (ns, name, max(ids + [0]) + 1, rid), wb, count=1)
    # every sheet moved one place right: keep sheet-scoped names pointing at the same sheet
    wb = re.sub(r'localSheetId="(\d+)"', lambda m: 'localSheetId="%d"' % (int(m.group(1)) + 1), wb)
    wb = re.sub(r'activeTab="\d+"', 'activeTab="0"', wb)
    wb = re.sub(r'firstSheet="\d+"', 'firstSheet="0"', wb)
    open(wb_p, 'w', encoding='utf-8').write(wb)
    return os.path.join(ws_dir, fname)


def append_callouts(sh, inserted, attention):
    """Write an update-run block onto the Validation Summary sheet.

    Standing instruction from the account team: a PO-vs-GRID problem has to be
    visible in the workbook itself, not only in a JSON report that can get
    separated from the file. Column-insert results go here too, so it is
    obvious which columns were not in the original GRID.
    """
    rows = [int(m) for m in re.findall(r'<row r="(\d+)"', sh.xml)]
    r = (max(rows) if rows else 0) + 2
    body = []

    def line(vals, row):
        cells = []
        for i, v in enumerate(vals, start=1):
            if v is None or v == '':
                continue
            text = (str(v).replace('&', '&amp;').replace('<', '&lt;')
                    .replace('>', '&gt;'))
            cells.append('<c r="%s%d" s="0" t="inlineStr"><is>'
                         '<t xml:space="preserve">%s</t></is></c>'
                         % (num_to_col(i), row, text))
        return ('<row r="%d" customFormat="false" ht="15" hidden="false" '
                'customHeight="false" outlineLevel="0" collapsed="false">%s</row>'
                % (row, ''.join(cells)))

    body.append(line(['Update run — columns added and items needing a decision'], r))
    r += 1
    if inserted:
        body.append(line(['PO columns inserted', 'Sheet', 'Column', 'Type',
                          'Ship window', 'Dest'], r))
        r += 1
        for i in inserted:
            body.append(line([i['po'], i['sheet'], i['column'], i['label'],
                              i['window'], i['dest']], r))
            r += 1
        r += 1
    if attention:
        body.append(line(['Needs attention', 'Detail', 'Next step'], r))
        r += 1
        for a_ in attention:
            what = a_.get('dpci') or a_.get('po') or ''
            detail = a_.get('issue', '')
            if a_.get('pos'):
                detail += ' (POs: %s)' % ', '.join(a_['pos'][:6])
            if a_.get('sheet'):
                detail += ' [sheet: %s]' % a_['sheet']
            body.append(line([what, detail, a_.get('next_step', '')], r))
            r += 1
    if not inserted and not attention:
        body.append(line(['Update run', 'no columns inserted, nothing outstanding'], r))

    m = re.search(r'</sheetData>', sh.xml)
    if not m:
        return 0
    sh.xml = sh.xml[:m.start()] + ''.join(body) + sh.xml[m.start():]
    return len(body)


def insert_po_column(sh, po, meta, qtys, type_styles, dest_map):
    """Insert one PO column on one sheet and write its header + quantities.

    qtys maps row number -> (kind, value): ('n', pieces) for an item row,
    ('inline', '240-04-0531\n(2023)') for an assortment header row.
    Returns a dict describing what was written.
    """
    po_row, po_cols, total_col = find_layout(sh)
    if not po_row or not po_cols:
        return {'error': 'no PO header row'}

    ordered = sorted(po_cols.items(), key=lambda kv: col_to_num(kv[1]))
    after = [c for p, c in ordered if sort_key(meta) > sort_key(META_BY_PO.get(p, {'po': p}))]
    at = (col_to_num(after[-1]) + 1) if after else col_to_num(ordered[0][1])

    # photos live in column C; the PO block starts at N. Never shift a photo.
    if at <= 3:
        return {'error': 'refusing to insert left of the photo column'}

    sh.xml = shift_cells_right(sh.xml, at)
    sh.xml = shift_cols_element(sh.xml, at)

    neighbour = num_to_col(at + 1 if not after else at - 1)
    label = meta.get('dclabel') or ''
    dest = dest_map.get(meta.get('dc', ''), meta.get('dc', ''))

    st1, st3 = type_styles
    sh.xml = put_cell(sh.xml, 1, at,
                      st1 or cell_style(sh.xml, '%s1' % neighbour), label)
    sh.xml = put_cell(sh.xml, 2, at, cell_style(sh.xml, '%s2' % neighbour),
                      kind='blank')
    sh.xml = put_cell(sh.xml, 3, at,
                      st3 or cell_style(sh.xml, '%s3' % neighbour), po)
    sh.xml = put_cell(sh.xml, 4, at, cell_style(sh.xml, '%s4' % neighbour),
                      ship_window(meta))
    sh.xml = put_cell(sh.xml, 5, at, cell_style(sh.xml, '%s5' % neighbour), dest)

    written = 0
    for row, (kind, value) in sorted(qtys.items()):
        style = cell_style(sh.xml, '%s%d' % (neighbour, row))
        sh.xml = put_cell(sh.xml, row, at, style, value, kind)
        written += 1

    # every data row keeps a styled cell in the new column, even when empty
    for row in find_dpci_rows(sh):
        pass

    po_row2, po_cols2, total_col2 = find_layout(sh)
    first = min(col_to_num(c) for c in po_cols2.values())
    last = max(col_to_num(c) for c in po_cols2.values())
    totals = 0
    if total_col2:
        sh.xml, totals = rewrite_totals(sh.xml, col_to_num(total_col2), first, last)

    return {'po': po, 'column': num_to_col(at), 'label': label,
            'window': ship_window(meta), 'dest': dest,
            'cells': written, 'totals_reformulated': totals,
            'type_style_found': bool(st1 and st3)}


def find_layout(sh, max_scan_row=6, max_col=60):
    """Locate: PO-number header row, DPCI column, TOTAL column.

    The GRID layout is frozen: column M carries the row labels TG PO# / S/W /
    DES PORT, and the PO numbers sit in that same row from column N rightwards.
    Anchoring on the "TG PO#" label is exact, so it is tried first; it also
    removes any chance of reading a UPC or a commit quantity as a PO number.

    The heuristic fallback stays for GRIDs that predate the label. Its digit
    range is 6-12: Target PO numbers were 7 digits on 26C5 but are 11 on the
    27C1/27C2 template ("10002036206"), and the old 6-9 bound matched none of
    them — every sheet was silently skipped with "PO column not found".
    """
    po_row, po_cols, total_col = None, {}, None

    # 1. exact: the row whose label cell reads "TG PO#"
    for r in range(1, max_scan_row + 1):
        label_col = None
        for ref in sh.row_cells(r):
            v = sh.cell_text(ref)
            if v and str(v).strip().upper().replace(' ', '') == 'TGPO#':
                label_col = col_to_num(COL_RE.match(ref).group(1))
                break
        if label_col is None:
            continue
        found = {}
        for ref in sh.row_cells(r):
            col = COL_RE.match(ref).group(1)
            if col_to_num(col) <= label_col:
                continue
            v = sh.cell_text(ref)
            if v and re.fullmatch(r'\d{6,12}', str(v).strip()):
                found[str(v).strip()] = col
        if found:
            po_row, po_cols = r, found
            break

    # 2. fallback for GRIDs with no TG PO# label
    if not po_cols:
        for r in range(1, max_scan_row + 1):
            found = {}
            for ref in sh.row_cells(r):
                col = COL_RE.match(ref).group(1)
                v = sh.cell_text(ref)
                if v and re.fullmatch(r'\d{6,12}', str(v).strip()):
                    found[str(v).strip()] = col
            if len(found) >= 2 and len(found) > len(po_cols):
                po_cols, po_row = found, r

    # TOTAL column: header cell containing TOTAL in the rows above data
    for r in range(1, max_scan_row + 1):
        for ref in sh.row_cells(r):
            v = sh.cell_text(ref)
            if v and 'TOTAL' in str(v).upper():
                total_col = COL_RE.match(ref).group(1)
    return po_row, po_cols, total_col


def find_dpci_rows(sh, max_row=400):
    """Map normalized DPCI -> row number (column A)."""
    out = {}
    for m in re.finditer(r'<row r="(\d+)"', sh.xml):
        r = int(m.group(1))
        if r > max_row:
            break
        v = sh.cell_text('A%d' % r)
        if v and re.fullmatch(r'\d{3}-\d{2}-\d{4}', str(v).strip()):
            out[re.sub(r'\D', '', v)] = r
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--grid', required=True)
    ap.add_argument('--data', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--dry-run', action='store_true')
    ap.add_argument('--po-meta',
                    help='items.json from parse_po_pdfs.py. Supplying it lets '
                         'the script INSERT a column for a PO that is not in '
                         'the GRID yet, instead of skipping it.')
    ap.add_argument('--no-insert', action='store_true',
                    help='never insert a column; only fill existing ones')
    ap.add_argument('--summary-json',
                    help='rewrite the Validation Summary tab from this summary JSON '
                         '(this run\'s call-outs are appended to it)')
    ap.add_argument('--force-text', action='store_true',
                    help='allow overwriting cells that currently hold text labels')
    a = ap.parse_args()

    data = json.load(open(a.data, encoding='utf-8'))
    tmp = tempfile.mkdtemp()
    base = os.path.join(tmp, 'x')
    with zipfile.ZipFile(a.grid) as z:
        names = z.namelist()
        z.extractall(base)
    media_before = len([n for n in names if '/media/' in n])

    shared = load_shared(base)
    sheets = sheet_name_map(base)

    report = {'written': [], 'skipped': [], 'totals_updated': [],
              'inserted': [], 'needs_attention': [],
              'media_before': media_before}

    # PO header metadata: ship window, DC and PO-type label for any PO we may
    # have to create a column for.
    global META_BY_PO
    META_BY_PO = {}
    box_text = {}
    if a.po_meta:
        meta_doc = json.load(open(a.po_meta, encoding='utf-8'))
        for h in meta_doc.get('pos', []):
            META_BY_PO[re.sub(r'\\D', '', h.get('po', ''))] = h
        for it in meta_doc.get('items', []):
            if it.get('kind') == 'assortment':
                p = re.sub(r'\\D', '', str(it.get('po', '')))
                d = re.sub(r'\\D', '', str(it.get('sku', '')))
                if p and d:
                    box_text.setdefault(d, {})[p] = it.get('qty')

    # normalize requested data: dpci -> {po: qty}
    flat = {}
    for sname, items in data.items():
        # grid.json also carries diagnostics ("superseded",
        # "duplicate_po_groups") alongside the quantity blocks; only the
        # dict-valued keys are sheet payloads.
        if not isinstance(items, dict):
            continue
        for dpci, pos in items.items():
            if not isinstance(pos, dict):
                continue
            flat.setdefault(re.sub(r'\D', '', dpci), {}).update(
                {re.sub(r'\D', '', p): q for p, q in pos.items()})

    # every sheet's XML, for locating an existing style that already carries a
    # given PO-type fill
    all_sheet_xml = [(open(sp, encoding='utf-8').read(), shared)
                     for sp in sheets.values()]
    grid_dpcis = set()
    new_type_styles = {}

    for sname, spath in sheets.items():
        sh = Sheet(spath, shared)
        po_row, po_cols, total_col = find_layout(sh)
        if not po_row or not po_cols:
            # Never skip a sheet silently — a layout the finder cannot read is
            # exactly how an update run reports success with holes in it.
            report['skipped'].append(
                {'sheet': sname,
                 'reason': 'no PO header row found on this sheet — sheet not updated'})
            continue
        rows = find_dpci_rows(sh)
        grid_dpcis.update(rows)
        touched_rows = set()

        # ---- insert columns for POs this sheet needs but does not have -----
        if META_BY_PO and not a.no_insert:
            needed = {}
            for dpci, pos in flat.items():
                if dpci not in rows:
                    continue
                for po, qty in pos.items():
                    if po not in po_cols:
                        needed.setdefault(po, {})[rows[dpci]] = ('n', qty)
            for dpci, pos in box_text.items():
                if dpci not in rows:
                    continue
                for po, boxes in pos.items():
                    if po not in po_cols:
                        pretty = '%s-%s-%s' % (dpci[:3], dpci[3:5], dpci[5:])
                        needed.setdefault(po, {})[rows[dpci]] = (
                            'inline', '%s\n(%s)' % (pretty, boxes))

            for po in sorted(needed, key=lambda p: sort_key(META_BY_PO.get(p, {'po': p}))):
                meta = META_BY_PO.get(po)
                if not meta:
                    report['needs_attention'].append(
                        {'sheet': sname, 'po': po,
                         'issue': 'PO has no column and no header metadata',
                         'next_step': 'include this PO in --po-meta, or rebuild '
                                      'with po-grid-create'})
                    continue
                if not meta.get('dclabel') or not ship_window(meta):
                    report['needs_attention'].append(
                        {'sheet': sname, 'po': po,
                         'issue': 'PO type label or shipping window missing',
                         'next_step': 'the build gate forbids a blank header — '
                                      'check the PO PDF, then rerun'})
                    continue
                tstyle = new_type_styles.get(meta['dclabel']) or find_type_style(all_sheet_xml, meta['dclabel'])
                made_style = False
                if not (tstyle[0] and tstyle[1]) and meta['dclabel'] in TYPE_FILL and po_cols:
                    # label new to this workbook: clone a neighbouring PO header style with the right fill
                    ncol = sorted(po_cols.values(), key=col_to_num)[-1]
                    s1 = re.search(r'<c r="%s1"[^>]*\bs="(\d+)"' % ncol, sh.xml)
                    s3 = re.search(r'<c r="%s3"[^>]*\bs="(\d+)"' % ncol, sh.xml)
                    fill = TYPE_FILL[meta['dclabel']]
                    tstyle = (clone_style_with_fill(base, s1.group(1) if s1 else None, fill),
                              clone_style_with_fill(base, s3.group(1) if s3 else None, fill))
                    new_type_styles[meta['dclabel']] = tstyle
                    made_style = True
                res = insert_po_column(sh, po, meta, needed[po], tstyle, DEST)
                if made_style:
                    res['type_style_found'] = True
                    res['note'] = 'PO 類型 %s 原檔沒有，已依標準配色新增' % meta['dclabel']
                if res.get('error'):
                    report['needs_attention'].append(
                        {'sheet': sname, 'po': po, 'issue': res['error'],
                         'next_step': 'rebuild with po-grid-create'})
                    continue
                res['sheet'] = sname
                if not res['type_style_found']:
                    res['note'] = ('PO type "%s" is new to this workbook — the '
                                   'column copies its neighbour\'s fill; recolour '
                                   'by hand or rebuild' % meta['dclabel'])
                    report['needs_attention'].append(
                        {'sheet': sname, 'po': po,
                         'issue': 'new PO-type label %s has no fill in this workbook'
                                  % meta['dclabel'],
                         'next_step': 'recolour the column, or rebuild with '
                                      'po-grid-create to get the correct fill'})
                report['inserted'].append(res)
            po_row, po_cols, total_col = find_layout(sh)

        for dpci, pos in flat.items():
            if dpci not in rows:
                # Reported once, workbook-wide, after the sheet loop — a DPCI
                # normally lives on exactly one factory sheet, so flagging it
                # here would fire on all the others too.
                continue
            r = rows[dpci]
            for po, qty in pos.items():
                col = po_cols.get(po)
                if not col:
                    report['skipped'].append(
                        {'sheet': sname, 'dpci': dpci, 'po': po,
                         'reason': 'PO column not found on this sheet'})
                    continue
                ref = '%s%d' % (col, r)
                # Assortment master rows carry a TEXT label in the PO column
                # (e.g. "240-02-0140-(1,858)"). Never silently replace text with
                # a number — that destroys the sheet's intended presentation.
                if sh.is_text(ref) and not a.force_text:
                    report['skipped'].append(
                        {'sheet': sname, 'dpci': dpci, 'po': po, 'cell': ref,
                         'current': sh.cell_text(ref),
                         'reason': 'cell holds a text label (assortment master); '
                                   'use --force-text to overwrite'})
                    continue
                if sh.set_number(ref, qty):
                    report['written'].append(
                        {'sheet': sname, 'dpci': dpci, 'po': po,
                         'cell': ref, 'qty': qty})
                    touched_rows.add(r)
                else:
                    report['skipped'].append(
                        {'sheet': sname, 'dpci': dpci, 'po': po,
                         'reason': 'cell %s absent in XML' % ref})

        # recompute row TOTAL for touched rows
        if total_col and touched_rows:
            for r in sorted(touched_rows):
                s = 0
                for po, col in po_cols.items():
                    v = sh.cell_text('%s%d' % (col, r))
                    try:
                        s += float(v)
                    except (TypeError, ValueError):
                        pass
                s = int(s) if s == int(s) else s
                tref = '%s%d' % (total_col, r)
                if sh.is_text(tref) and not a.force_text:
                    continue
                if sh.set_number(tref, s):
                    report['totals_updated'].append(
                        {'sheet': sname, 'cell': '%s%d' % (total_col, r), 'value': s})

        # Restore =SUM in the PO TOTAL column over the current PO block.
        if total_col and po_cols:
            first = min(col_to_num(c) for c in po_cols.values())
            last = max(col_to_num(c) for c in po_cols.values())
            sh.xml, nf = formulaize_totals(sh.xml, col_to_num(total_col),
                                           first, last)
            if nf:
                report.setdefault('totals_reformulated', []).append(
                    {'sheet': sname, 'cells': nf})

        if not a.dry_run:
            sh.save()

    # ---- items this GRID has no place for, judged across the whole workbook,
    # not per sheet: a DPCI normally lives on exactly one factory sheet, so a
    # per-sheet check would fire on all the others.
    for dpci, pos in sorted(flat.items()):
        if dpci in grid_dpcis:
            continue
        pretty = '%s-%s-%s' % (dpci[:3], dpci[3:5], dpci[5:])
        report['needs_attention'].append(
            {'dpci': pretty, 'pos': sorted(pos),
             'issue': 'DPCI is on a PO but has no row anywhere in this GRID',
             'next_step': 'new item — update cannot add a row (it has no master '
                          'row data and no photo). Rebuild with po-grid-create, '
                          'or add the row by hand and rerun'})

    # ---- call-outs onto the Validation Summary tab ------------------------
    vs_path = sheets.get(VALIDATION_SHEET)
    if a.summary_json and not a.dry_run:
        if not vs_path:
            vs_path = add_summary_sheet_first(base)
            report['validation_sheet_added'] = True
        styles = add_summary_styles(base)
        report['validation_rows_added'] = write_summary_xml(
            vs_path, json.load(open(a.summary_json, encoding='utf-8')), styles,
            report['inserted'], report['needs_attention'])
    elif vs_path and not a.dry_run:
        vs = Sheet(vs_path, shared)
        n = append_callouts(vs, report['inserted'], report['needs_attention'])
        if n:
            vs.save()
            report['validation_rows_added'] = n
    elif not vs_path:
        report.setdefault('notes', []).append(
            'no "%s" sheet in this GRID — call-outs are in this report only'
            % VALIDATION_SHEET)

    if a.dry_run:
        report['dry_run'] = True
        shutil.rmtree(tmp)
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0

    # calcChain.xml goes stale on any column insert (it still lists the
    # pre-shift cell refs), which makes Excel open with a "removed records:
    # formula" repair prompt. Excel rebuilds the chain on its own — drop it
    # and its two registry entries, and force a full recalc so every
    # formula (not just the ones we touched) gets a fresh value.
    calc_chain = os.path.join(base, 'xl', 'calcChain.xml')
    if os.path.exists(calc_chain):
        os.remove(calc_chain)
        ct_path = os.path.join(base, '[Content_Types].xml')
        ct = open(ct_path, encoding='utf-8').read()
        ct2 = re.sub(r'<Override[^>]*calcChain\.xml[^>]*/>', '', ct)
        if ct2 != ct:
            open(ct_path, 'w', encoding='utf-8').write(ct2)
        rels_path = os.path.join(base, 'xl', '_rels', 'workbook.xml.rels')
        rels = open(rels_path, encoding='utf-8').read()
        rels2 = re.sub(r'<Relationship[^>]*calcChain\.xml[^>]*/>', '', rels)
        if rels2 != rels:
            open(rels_path, 'w', encoding='utf-8').write(rels2)
        report['calc_chain_removed'] = True

    wb_path = os.path.join(base, 'xl', 'workbook.xml')
    wb_xml = open(wb_path, encoding='utf-8').read()
    if 'fullCalcOnLoad' not in wb_xml:
        wb_xml2 = re.sub(r'(<calcPr\b[^>]*)/>', r'\1 fullCalcOnLoad="1"/>', wb_xml, count=1)
        if wb_xml2 != wb_xml:
            open(wb_path, 'w', encoding='utf-8').write(wb_xml2)

    # repack, preserving every original part
    if os.path.exists(a.out):
        os.remove(a.out)
    with zipfile.ZipFile(a.out, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(base):
            for f in files:
                full = os.path.join(root, f)
                zf.write(full, os.path.relpath(full, base))

    with zipfile.ZipFile(a.out) as z:
        after = z.namelist()

    report['media_after'] = len([n for n in after if '/media/' in n])
    report['richdata_ok'] = any('richData' in n for n in after) if \
        any('richData' in n for n in names) else True
    report['image_integrity'] = (report['media_after'] == media_before
                                 and report['richdata_ok'])
    shutil.rmtree(tmp)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report['image_integrity'] else 1


if __name__ == '__main__':
    sys.exit(main())
