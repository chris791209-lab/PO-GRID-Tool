import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime

# ==========================================
# 0. 密碼保護機制
# ==========================================
def check_password():
    """回傳 True 代表使用者輸入了正確的密碼"""
    def password_entered():
        if st.session_state["password"] == st.secrets["app_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input(
            "🔒 請輸入 AE 部門共用密碼以啟用工具：",
            type="password",
            on_change=password_entered,
            key="password"
        )
        return False
    elif not st.session_state["password_correct"]:
        st.text_input(
            "🔒 請輸入 AE 部門共用密碼以啟用工具：",
            type="password",
            on_change=password_entered,
            key="password"
        )
        st.error("❌ 密碼錯誤，請重新輸入。")
        return False
    else:
        return True

if not check_password():
    st.stop()

# ==========================================
# 1. 共通資料清理函數
# ==========================================
def clean_dpci(series):
    """清理 DPCI 字串，強制移除所有空白、斜線與隱藏字元"""
    if series is None:
        return series
    cleaned = series.astype(str)
    cleaned = cleaned.str.replace(r'\s+', '', regex=True)
    cleaned = cleaned.str.replace(r'[/\\]', '-', regex=True)
    cleaned = cleaned.str.replace(r'\.0$', '', regex=True)
    return cleaned

def clean_upc(series):
    """清理 UPC/Barcode 字串，避免因 Excel 浮點數轉換產生 .0 導致比對失敗"""
    if series is None:
        return series
    cleaned = series.astype(str)
    cleaned = cleaned.str.replace(r'\.0$', '', regex=True)
    cleaned = cleaned.str.replace(r'\s+', '', regex=True)
    cleaned = cleaned.replace('nan', np.nan)
    return cleaned

def fuzzy_col(df_cols, *keywords, require_all=False):
    """
    G10 模糊欄位搜尋：正規化空白與大小寫後，尋找包含所有/任一關鍵字的欄位。
    require_all=True → 所有 keywords 都需命中；False → 任一即可。
    """
    normalized = {c: re.sub(r'\s+', '', c).lower() for c in df_cols}
    for col, norm in normalized.items():
        hits = [kw.lower().replace(' ', '') in norm for kw in keywords]
        if (all(hits) if require_all else any(hits)):
            return col
    return None

# ==========================================
# 2. G3 重複 PO 偵測
# ==========================================
def detect_duplicate_pos(po_df, po_number_col='PO NUMBER'):
    """
    偵測兩種重複情形：
    (a) 同 PO# 有多個版本 → 保留最新（依列序），回傳警告訊息
    (b) 不同 PO# 但內容完全相同 → 回傳警告訊息
    回傳 (cleaned_df, warnings_list)
    """
    warnings = []
    df = po_df.copy()

    # (a) 同 PO# 多版本：以 PO# + DPCI 的組合為 key，若同 PO# 出現完全相同 DPCI 組合多次
    dpci_col = 'Final_DPCI' if 'Final_DPCI' in df.columns else 'Original_DPCI'
    if dpci_col in df.columns:
        po_signatures = df.groupby(po_number_col)[dpci_col].apply(lambda x: frozenset(x)).reset_index()
        po_signatures.columns = [po_number_col, 'dpci_set']

        # 找出相同 PO# 但 DPCI set 完全重複（代表同 PO# 被上傳多次）
        dup_pos = df[df.duplicated(subset=[po_number_col, dpci_col], keep=False)][po_number_col].unique()
        if len(dup_pos) > 0:
            warnings.append(
                f"⚠️ **版本重複偵測**：以下 PO# 有完全重複的品項列，已自動保留最後出現的版本（請確認是否為最新修訂單）：\n"
                + ", ".join(str(p) for p in dup_pos)
            )
            df = df.drop_duplicates(subset=[po_number_col, dpci_col], keep='last')

        # (b) 不同 PO# 但 DPCI 組合完全相同（可能是誤傳兩張不同號碼的相同內容 PO）
        po_signatures2 = df.groupby(po_number_col)[dpci_col].apply(lambda x: frozenset(x)).reset_index()
        po_signatures2.columns = [po_number_col, 'dpci_set']
        sig_groups = po_signatures2.groupby('dpci_set')[po_number_col].apply(list).reset_index()
        dupe_sig_groups = sig_groups[sig_groups[po_number_col].map(len) > 1]
        if len(dupe_sig_groups) > 0:
            for _, row in dupe_sig_groups.iterrows():
                po_list = row[po_number_col]
                warnings.append(
                    f"⚠️ **內容相同 PO 偵測**：以下 PO# 含有完全相同的品項組合，請確認是否為重複上傳：\n"
                    + " / ".join(str(p) for p in po_list)
                )

    return df, warnings

# ==========================================
# 3. G5 PO 自我驗證（訂單內部一致性）
# ==========================================
def po_self_verify(po_df, mode='standard'):
    """
    G5：驗證 PO 內部一致性
    - 每行 qty × unit cost ≈ line total（若欄位存在）
    - 所有行小計加總 ≈ PO Grand Total（若欄位存在）
    回傳 warnings list
    """
    warnings = []
    df = po_df.copy()

    if mode == 'standard':
        # 標準版：TOTAL ITEM QTY × ITEM UNIT COST ≈ EXTENDED COST（若有）
        ext_cost_col = fuzzy_col(df.columns, 'extended', 'cost')
        if ext_cost_col and 'ITEM UNIT COST' in df.columns and 'TOTAL ITEM QTY' in df.columns:
            df['_calc_ext'] = pd.to_numeric(df['TOTAL ITEM QTY'], errors='coerce') * pd.to_numeric(df['ITEM UNIT COST'], errors='coerce')
            df['_stated_ext'] = pd.to_numeric(df[ext_cost_col], errors='coerce')
            mismatch = df[np.abs(df['_calc_ext'] - df['_stated_ext']) > 0.05].dropna(subset=['_calc_ext', '_stated_ext'])
            if len(mismatch) > 0:
                dpcis = mismatch.get('Original_DPCI', mismatch.get('Final_DPCI', pd.Series())).tolist()
                warnings.append(
                    f"⚠️ **PO 內部驗算警告**：{len(mismatch)} 筆品項的「數量 × 單價」與「Extended Cost」不符（差異 > $0.05）：\n"
                    + ", ".join(str(d) for d in dpcis[:10])
                    + ("..." if len(dpcis) > 10 else "")
                )

        # PO Grand Total 驗算
        total_col = fuzzy_col(df.columns, 'grand', 'total') or fuzzy_col(df.columns, 'po', 'total')
        if total_col and ext_cost_col:
            df['_stated_ext2'] = pd.to_numeric(df[ext_cost_col], errors='coerce')
            calc_total = df['_stated_ext2'].sum()
            stated_total = pd.to_numeric(df[total_col], errors='coerce').dropna()
            if len(stated_total) > 0:
                stated_val = stated_total.iloc[0]
                if abs(calc_total - stated_val) > 1.0:
                    warnings.append(
                        f"⚠️ **PO 總金額驗算警告**：所有品項 Extended Cost 合計 ${calc_total:,.2f}，"
                        f"但 PO 表頭 Grand Total 為 ${stated_val:,.2f}，差異 ${abs(calc_total - stated_val):,.2f}。"
                    )

    elif mode == 'modern':
        # 現代版：ORIGINAL QUANTITY × COST $ ≈ 該列金額
        orig_qty_col = fuzzy_col(df.columns, 'original', 'quantity')
        cost_col = fuzzy_col(df.columns, 'cost', '$') or fuzzy_col(df.columns, 'revcost', '$')
        line_total_col = fuzzy_col(df.columns, 'extended') or fuzzy_col(df.columns, 'line', 'total')
        if orig_qty_col and cost_col and line_total_col:
            qty = pd.to_numeric(df[orig_qty_col], errors='coerce')
            cost = pd.to_numeric(df[cost_col], errors='coerce')
            stated = pd.to_numeric(df[line_total_col], errors='coerce')
            calc = qty * (cost / qty.replace(0, np.nan))  # cost is already line total in modern
            # Modern PO cost col is already the line total (cost$ = qty * unit cost)
            # Verify: cost$ / qty ≈ unit cost is internally consistent — skip if no separate unit cost col
            pass

    return warnings

# ==========================================
# 4. 資料處理函數
# ==========================================
def process_standard_po(df):
    """處理【標準版】訂單原始資料"""
    df.columns = df.columns.astype(str).str.replace(r'[﻿\n\r]', '', regex=True).str.strip()

    po_col = next((c for c in df.columns if 'PO NUMBER' in c.upper()), None)

    if not po_col:
        if next((c for c in df.columns if 'PO' in c.upper() and '#' in c.upper()), None):
            st.error("❌ 檔案讀取錯誤：您上傳的似乎是【現代版 PO】。請切換到「📈 現代版 (Modern PO)」分頁！")
        else:
            st.error("❌ 檔案讀取錯誤：找不到 'PO NUMBER' 欄位！請確認您上傳的是正確的【標準版 PO】檔案。")
        st.stop()

    df[po_col] = df[po_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df = df[df[po_col].str.match(r'^\d+$', na=False)].copy()

    df['PO NUMBER'] = df[po_col]
    df['ASSORTMENT ITEM?'] = df['ASSORTMENT ITEM?'].fillna('N').astype(str).str.strip().str.upper()

    cols_to_clean = ['DEPARTMENT', 'CLASS', 'ITEM', 'COMPONENT DEPARTMENT', 'COMPONENT CLASS', 'COMPONENT ITEM']
    for col in cols_to_clean:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

    df['Original_DPCI'] = clean_dpci(df['DEPARTMENT'].str.zfill(3) + "-" + df['CLASS'].str.zfill(2) + "-" + df['ITEM'].str.zfill(4))
    df['Final_DPCI'] = np.where(
        df['ASSORTMENT ITEM?'] == 'Y',
        clean_dpci(df['COMPONENT DEPARTMENT'].str.zfill(3) + "-" + df['COMPONENT CLASS'].str.zfill(2) + "-" + df['COMPONENT ITEM'].str.zfill(4)),
        df['Original_DPCI']
    )

    qty_series = np.where(df['ASSORTMENT ITEM?'] == 'Y', df['COMPONENT ITEM TOTAL QTY'], df['TOTAL ITEM QTY'])
    df['Final_QTY'] = pd.Series(qty_series).astype(str).str.replace(',', '', regex=False).astype(float)

    for col in ['ITEM UNIT COST', 'ITEM UNIT RETAIL', 'VCP QUANTITY', 'COMPONENT ASSORT QTY']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    df['PO UPC'] = clean_upc(df.get('ITEM BAR CODE', pd.Series(np.nan, index=df.index)))

    # G9: 標記 Shipper Display DPCIs（ASSORTMENT ITEM? == 'Y' 且 Final_DPCI == Original_DPCI 的視為 box 本身）
    # Shipper display 通常在 description 含 "SHIPPER" 或 "DISPLAY" 字樣
    if 'ITEM DESCRIPTION' in df.columns:
        shipper_mask = df['ITEM DESCRIPTION'].astype(str).str.upper().str.contains(r'\bSHIPPER\b|\bDISPLAY\b', regex=True, na=False)
        df['Is_Shipper_Display'] = shipper_mask
    else:
        df['Is_Shipper_Display'] = False

    return df

def process_modern_po(df):
    """處理【現代版】訂單原始資料"""
    df.columns = df.columns.astype(str).str.replace(r'[﻿\n\r]', '', regex=True).str.strip()

    po_col = next((c for c in df.columns if 'PO' in c.upper() and '#' in c.upper()), None)

    cost_col = next((c for c in df.columns if c.upper() == 'COST $'), None) or \
               next((c for c in df.columns if 'REV COST' in c.upper() and '$' in c.upper()), None) or \
               next((c for c in df.columns if 'COST' in c.upper() and '$' in c.upper()), None)

    retail_col = next((c for c in df.columns if c.upper() == 'RETAIL $'), None) or \
                 next((c for c in df.columns if 'REV RETAIL' in c.upper() and '$' in c.upper()), None) or \
                 next((c for c in df.columns if 'RETAIL' in c.upper() and '$' in c.upper()), None)

    if not po_col or not cost_col:
        if next((c for c in df.columns if 'PO NUMBER' in c.upper()), None):
            st.error("❌ 檔案讀取錯誤：您上傳的似乎是【標準版 PO】。請切換到「📊 標準版 (Standard PO)」分頁！")
        else:
            st.error("❌ 檔案讀取錯誤：找不到 'PO #' 或 'COST $' 相關欄位！請確認您上傳的是正確的【現代版 PO】檔案。")
        st.stop()

    df[po_col] = df[po_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df = df[df[po_col].str.match(r'^\d+$', na=False)].copy()

    orig_qty_col = next((c for c in df.columns if 'ORIGINAL QUANTITY' in c.upper()), None)
    rev_qty_col = next((c for c in df.columns if 'REVISED QUANTITY' in c.upper()), None)

    for col in [orig_qty_col, rev_qty_col, cost_col, retail_col, 'VCP QUANTITY']:
        if col and col in df.columns:
            df[col] = df[col].astype(str).str.replace(',', '', regex=False)
            df[col] = pd.to_numeric(df[col], errors='coerce')

    df['PO NUMBER'] = df[po_col].astype(str)
    df['Original_DPCI'] = clean_dpci(df['DPCI'])
    df['Final_DPCI'] = df['Original_DPCI']

    if orig_qty_col and cost_col:
        df['ITEM UNIT COST'] = df[cost_col] / df[orig_qty_col]
    else:
        df['ITEM UNIT COST'] = np.nan

    if orig_qty_col and retail_col:
        df['ITEM UNIT RETAIL'] = df[retail_col] / df[orig_qty_col]
    else:
        df['ITEM UNIT RETAIL'] = np.nan

    df['Final_QTY'] = df[rev_qty_col] if rev_qty_col else (df[orig_qty_col] if orig_qty_col else np.nan)
    df['REVISED QUANTITY'] = df['Final_QTY']
    df['ASSORTMENT ITEM?'] = 'N'
    df['COMPONENT ASSORT QTY'] = np.nan
    df['PO UPC'] = clean_upc(df.get('UPC', pd.Series(np.nan, index=df.index)))

    # G9: Shipper Display 標記
    desc_col = fuzzy_col(df.columns, 'description') or fuzzy_col(df.columns, 'item', 'desc')
    if desc_col:
        shipper_mask = df[desc_col].astype(str).str.upper().str.contains(r'\bSHIPPER\b|\bDISPLAY\b', regex=True, na=False)
        df['Is_Shipper_Display'] = shipper_mask
    else:
        df['Is_Shipper_Display'] = False

    return df

def process_products(files):
    df_list = []
    for f in files:
        df = pd.read_csv(f) if f.name.endswith('.csv') else pd.read_excel(f)
        df_list.append(df)
    if not df_list:
        return pd.DataFrame()

    master_product_df = pd.concat(df_list, ignore_index=True)
    if 'DPCI' in master_product_df.columns:
        master_product_df['DPCI'] = clean_dpci(master_product_df['DPCI'])

    if 'Barcode' in master_product_df.columns:
        master_product_df['Target UPC'] = clean_upc(master_product_df['Barcode'])
    else:
        master_product_df['Target UPC'] = clean_upc(master_product_df.get('UPC', pd.Series(np.nan)))

    numeric_cols = ['FCA Factory City Unit Cost', 'FOB Unit Cost', 'Suggested Unit Retail', 'Case Unit Quantity', 'Ent Ttl Rcpt U']
    for col in numeric_cols:
        if col in master_product_df.columns:
            master_product_df[col] = pd.to_numeric(master_product_df[col], errors='coerce')

    if 'FCA Factory City Unit Cost' in master_product_df.columns and 'FOB Unit Cost' in master_product_df.columns:
        master_product_df['Final_Product_Cost'] = master_product_df['FCA Factory City Unit Cost'].fillna(master_product_df['FOB Unit Cost'])
    elif 'FCA Factory City Unit Cost' in master_product_df.columns:
        master_product_df['Final_Product_Cost'] = master_product_df['FCA Factory City Unit Cost']
    elif 'FOB Unit Cost' in master_product_df.columns:
        master_product_df['Final_Product_Cost'] = master_product_df['FOB Unit Cost']
    else:
        master_product_df['Final_Product_Cost'] = np.nan

    return master_product_df

def process_assortments(files):
    """G10 改善版：模糊欄位名稱匹配"""
    df_list = []
    for f in files:
        if f.name.endswith('.csv'):
            raw_sheets = [pd.read_csv(f, header=None)]
        else:
            xf = pd.ExcelFile(f)
            raw_sheets = [xf.parse(s, header=None) for s in xf.sheet_names]

        for raw_df in raw_sheets:
            # 找 header row：含 'assortment' 且含 'dpci' 的列
            header_idx = -1
            for i, row in raw_df.iterrows():
                row_str = row.astype(str).str.replace(r'\s+', '', regex=True).str.lower()
                if row_str.str.contains('assortment', na=False).any() and row_str.str.contains('dpci', na=False).any():
                    header_idx = i
                    break
            if header_idx == -1:
                continue

            df = raw_df.iloc[header_idx + 1:].reset_index(drop=True)
            df.columns = raw_df.iloc[header_idx].astype(str).str.replace(r'[\n\r]', ' ', regex=True).str.strip()

            # G10 模糊欄位搜尋
            master_col = fuzzy_col(df.columns, 'assortment', 'dpci', require_all=True) or \
                         fuzzy_col(df.columns, 'assortmentdpci')
            sub_col = fuzzy_col(df.columns, 'component', 'dpci', require_all=True) or \
                      fuzzy_col(df.columns, 'itemdpci')
            cost_col = fuzzy_col(df.columns, 'asst', 'cost') or \
                       fuzzy_col(df.columns, 'fa', 'box', 'cost', require_all=False) or \
                       fuzzy_col(df.columns, 'boxcost')
            units_col = fuzzy_col(df.columns, 'units', 'assortment', require_all=True) or \
                        fuzzy_col(df.columns, 'unitsinassortment')

            if not all([master_col, sub_col, cost_col, units_col]):
                continue

            temp_df = df[[master_col, sub_col, cost_col, units_col]].copy()
            temp_df.columns = ['Assortment_DPCI', 'Component_DPCI', 'Asst_Box_Cost', 'Units_in_Assortment']
            temp_df['Assortment_DPCI'] = temp_df['Assortment_DPCI'].replace(r'^\s*$', np.nan, regex=True).ffill()
            temp_df['Asst_Box_Cost'] = temp_df['Asst_Box_Cost'].replace(r'^\s*$', np.nan, regex=True).ffill()
            temp_df = temp_df.dropna(subset=['Assortment_DPCI', 'Component_DPCI'])
            temp_df['Assortment_DPCI'] = clean_dpci(temp_df['Assortment_DPCI'])
            temp_df['Component_DPCI'] = clean_dpci(temp_df['Component_DPCI'])
            temp_df = temp_df[~temp_df['Assortment_DPCI'].str.lower().str.contains('iafillsout|nan|none', na=False)]
            temp_df['Asst_Box_Cost'] = pd.to_numeric(temp_df['Asst_Box_Cost'], errors='coerce')
            temp_df['Units_in_Assortment'] = pd.to_numeric(temp_df['Units_in_Assortment'], errors='coerce')
            df_list.append(temp_df)

    if df_list:
        return pd.concat(df_list, ignore_index=True).drop_duplicates(subset=['Assortment_DPCI', 'Component_DPCI'])
    return pd.DataFrame(columns=['Assortment_DPCI', 'Component_DPCI', 'Asst_Box_Cost', 'Units_in_Assortment'])

def process_dispatch(files):
    """G7: 讀取 Dispatch 對照表，輸出 DPCI → Factory / AC / AE 對應"""
    if not files:
        return pd.DataFrame()
    df_list = []
    for f in files:
        df = pd.read_csv(f) if f.name.endswith('.csv') else pd.read_excel(f)
        df_list.append(df)
    if not df_list:
        return pd.DataFrame()
    dispatch_df = pd.concat(df_list, ignore_index=True)
    dispatch_df.columns = dispatch_df.columns.astype(str).str.strip()

    # 模糊找 DPCI, Factory, AC, AE 欄
    dpci_col = fuzzy_col(dispatch_df.columns, 'dpci')
    factory_col = fuzzy_col(dispatch_df.columns, 'factory')
    ac_col = fuzzy_col(dispatch_df.columns, 'ac') or fuzzy_col(dispatch_df.columns, 'account')
    ae_col = fuzzy_col(dispatch_df.columns, 'ae') or fuzzy_col(dispatch_df.columns, 'program')

    keep = {}
    if dpci_col:
        keep['DPCI'] = dispatch_df[dpci_col].pipe(clean_dpci)
    if factory_col:
        keep['Dispatch_Factory'] = dispatch_df[factory_col]
    if ac_col:
        keep['Dispatch_AC'] = dispatch_df[ac_col]
    if ae_col:
        keep['Dispatch_AE'] = dispatch_df[ae_col]

    if 'DPCI' not in keep:
        return pd.DataFrame()

    result = pd.DataFrame(keep)
    result = result.drop_duplicates(subset=['DPCI'])
    return result

# ==========================================
# 5. G1: PDF 解析函數
# ==========================================
def parse_sps_pdf(pdf_file):
    """
    G1: 解析 SPS Commerce PO PDF，回傳 (po_df, parse_warnings)
    支援標準版格式（PO NUMBER / ITEM NUMBER / COST）
    """
    try:
        import pdfplumber
    except ImportError:
        return None, ["❌ 缺少 pdfplumber 套件，請在 requirements.txt 加入 pdfplumber 並重新部署。"]

    warnings = []
    all_rows = []
    current_po = None
    current_dpci_parts = {}

    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                tables = page.extract_tables()
                text = page.extract_text() or ''

                # 嘗試從文字中提取 PO NUMBER
                po_match = re.search(r'PO\s*(?:NUMBER|#)[:\s]+(\d{8,12})', text, re.IGNORECASE)
                if po_match:
                    current_po = po_match.group(1)

                for table in tables:
                    if not table:
                        continue
                    # 找 header row
                    header_row_idx = None
                    for i, row in enumerate(table):
                        row_str = ' '.join(str(c) for c in row if c).upper()
                        if ('DPCI' in row_str or 'ITEM' in row_str) and ('COST' in row_str or 'QTY' in row_str or 'QUANTITY' in row_str):
                            header_row_idx = i
                            break

                    if header_row_idx is None:
                        continue

                    headers = [str(c).strip() if c else '' for c in table[header_row_idx]]

                    # 找各欄位 index
                    def find_col_idx(keywords, require_all=False):
                        for idx, h in enumerate(headers):
                            h_norm = re.sub(r'\s+', '', h).lower()
                            hits = [kw.lower().replace(' ', '') in h_norm for kw in keywords]
                            if (all(hits) if require_all else any(hits)):
                                return idx
                        return None

                    dpci_idx = find_col_idx(['dpci'])
                    qty_idx = find_col_idx(['qty', 'quantity'])
                    cost_idx = find_col_idx(['cost'])
                    retail_idx = find_col_idx(['retail'])
                    desc_idx = find_col_idx(['description', 'desc'])
                    upc_idx = find_col_idx(['upc', 'barcode', 'bar code'])
                    asst_idx = find_col_idx(['assortment'])
                    vcp_idx = find_col_idx(['vcp', 'case'])

                    if dpci_idx is None or qty_idx is None:
                        continue

                    for row in table[header_row_idx + 1:]:
                        if not row or all(c is None or str(c).strip() == '' for c in row):
                            continue
                        def get(idx):
                            if idx is None or idx >= len(row):
                                return None
                            return str(row[idx]).strip() if row[idx] is not None else None

                        dpci_raw = get(dpci_idx)
                        qty_raw = get(qty_idx)
                        if not dpci_raw or not re.search(r'\d{3}-\d{2}-\d{4}', dpci_raw or ''):
                            continue

                        all_rows.append({
                            'PO NUMBER': current_po,
                            'Original_DPCI': clean_dpci(pd.Series([dpci_raw])).iloc[0],
                            'Final_DPCI': clean_dpci(pd.Series([dpci_raw])).iloc[0],
                            'Final_QTY': pd.to_numeric(str(qty_raw).replace(',', ''), errors='coerce') if qty_raw else np.nan,
                            'ITEM UNIT COST': pd.to_numeric(str(get(cost_idx) or '').replace(',', '').replace('$', ''), errors='coerce'),
                            'ITEM UNIT RETAIL': pd.to_numeric(str(get(retail_idx) or '').replace(',', '').replace('$', ''), errors='coerce'),
                            'ITEM DESCRIPTION': get(desc_idx),
                            'PO UPC': clean_upc(pd.Series([get(upc_idx) or np.nan])).iloc[0],
                            'ASSORTMENT ITEM?': 'Y' if (get(asst_idx) or '').upper() in ['Y', 'YES'] else 'N',
                            'VCP QUANTITY': pd.to_numeric(str(get(vcp_idx) or ''), errors='coerce'),
                            'COMPONENT ASSORT QTY': np.nan,
                            'Is_Shipper_Display': False,
                        })

    except Exception as e:
        return None, [f"❌ PDF 解析錯誤：{e}"]

    if not all_rows:
        return None, ["⚠️ 在 PDF 中找不到可解析的 PO 表格，請確認格式為 SPS Commerce 標準版 PO。"]

    po_df = pd.DataFrame(all_rows)
    po_df = po_df[po_df['PO NUMBER'].notna()].copy()

    if po_df['PO NUMBER'].isna().any():
        warnings.append("⚠️ 部分列無法對應 PO NUMBER，已略過。")

    return po_df, warnings

# ==========================================
# 6. 共通驗證引擎（G2 / G4 / G6 整合）
# ==========================================
def run_validation(po_df, prod_df, asst_df, mode='standard'):
    """
    共通驗證引擎，整合 G2 / G4 / G6 改善：
    - G4: Cost tolerance = 0.005
    - G2: QTY tolerance = case_pack rounding OR ±10% of plan
    - G6: Assortment cost reverse-check
    - G9: Shipper Display 排除於 QTY 加總
    """
    merged_df = pd.merge(
        po_df,
        prod_df[[c for c in ['DPCI', 'Final_Product_Cost', 'Suggested Unit Retail',
                              'Case Unit Quantity', 'Ent Ttl Rcpt U', 'Target UPC']
                 if c in prod_df.columns]].drop_duplicates(subset=['DPCI']),
        left_on='Final_DPCI', right_on='DPCI', how='left'
    )

    # 混裝箱處理
    if asst_df is not None and len(asst_df) > 0:
        if mode == 'standard':
            merged_df = pd.merge(
                merged_df, asst_df,
                left_on=['Original_DPCI', 'Final_DPCI'],
                right_on=['Assortment_DPCI', 'Component_DPCI'],
                how='left'
            )
            merged_df['Target_Cost'] = np.where(
                merged_df['ASSORTMENT ITEM?'] == 'Y',
                merged_df['Asst_Box_Cost'],
                merged_df['Final_Product_Cost']
            )
        else:  # modern
            condensed_asst = asst_df.groupby('Assortment_DPCI', as_index=False).agg({'Asst_Box_Cost': 'first'})
            merged_df = pd.merge(merged_df, condensed_asst, left_on='Original_DPCI', right_on='Assortment_DPCI', how='left')
            merged_df['ASSORTMENT ITEM?'] = np.where(merged_df['Asst_Box_Cost'].notna(), 'Y', 'N')
            merged_df['Target_Cost'] = np.where(merged_df['ASSORTMENT ITEM?'] == 'Y', merged_df['Asst_Box_Cost'], merged_df['Final_Product_Cost'])
    else:
        merged_df['Target_Cost'] = merged_df['Final_Product_Cost']

    # G4: 成本比對（tolerance = 0.005）
    merged_df['Cost Match'] = np.isclose(
        merged_df['ITEM UNIT COST'].fillna(-1),
        merged_df['Target_Cost'].fillna(-1),
        atol=0.005  # G4: 從 0.01 → 0.005
    )
    merged_df['Cost Match'] = np.where(merged_df['Target_Cost'].isna(), False, merged_df['Cost Match'])

    # 零售價比對
    asst_mask = merged_df.get('ASSORTMENT ITEM?', pd.Series('N', index=merged_df.index)) == 'Y'
    if 'Suggested Unit Retail' in merged_df.columns:
        merged_df['Retail Match'] = np.where(
            asst_mask,
            True,  # 混裝箱 master row 零售標 n/a
            np.isclose(
                merged_df['ITEM UNIT RETAIL'].fillna(0),
                merged_df['Suggested Unit Retail'].fillna(0),
                atol=0.01
            )
        )
    else:
        merged_df['Retail Match'] = True

    # 裝箱數比對
    if mode == 'standard':
        merged_df['Target Case / Assort QTY'] = np.where(
            asst_mask,
            merged_df.get('Units_in_Assortment', pd.Series(np.nan, index=merged_df.index)),
            merged_df.get('Case Unit Quantity', pd.Series(np.nan, index=merged_df.index))
        )
        merged_df['PO VCP / Assort QTY'] = np.where(
            asst_mask,
            merged_df.get('COMPONENT ASSORT QTY', pd.Series(np.nan, index=merged_df.index)),
            merged_df.get('VCP QUANTITY', pd.Series(np.nan, index=merged_df.index))
        )
    else:
        merged_df['Target Case / Assort QTY'] = merged_df.get('Case Unit Quantity', pd.Series(np.nan, index=merged_df.index))
        merged_df['PO VCP / Assort QTY'] = merged_df.get('VCP QUANTITY', pd.Series(np.nan, index=merged_df.index))
        merged_df['Target Case / Assort QTY'] = np.where(
            asst_mask, merged_df['PO VCP / Assort QTY'], merged_df['Target Case / Assort QTY']
        )

    merged_df['Case QTY Match'] = np.isclose(
        merged_df['PO VCP / Assort QTY'].fillna(-1),
        merged_df['Target Case / Assort QTY'].fillna(-1),
        atol=0.01
    )
    merged_df['Case QTY Match'] = np.where(merged_df['Target Case / Assort QTY'].isna(), False, merged_df['Case QTY Match'])
    if mode == 'modern':
        merged_df['Case QTY Match'] = np.where(asst_mask, True, merged_df['Case QTY Match'])

    # G2: 總數量比對（case_pack rounding + 10% 容忍）
    merged_df['Target Commit QTY'] = merged_df.get('Ent Ttl Rcpt U', pd.Series(np.nan, index=merged_df.index))
    case_pack = merged_df.get('Case Unit Quantity', pd.Series(1, index=merged_df.index)).fillna(1)

    # G9: Shipper Display 排除於 QTY 加總
    shipper_mask = merged_df.get('Is_Shipper_Display', pd.Series(False, index=merged_df.index)).astype(bool)
    merged_df['Final_QTY_for_count'] = np.where(shipper_mask, 0, merged_df['Final_QTY'])
    merged_df['PO Total QTY'] = merged_df.groupby('Final_DPCI')['Final_QTY_for_count'].transform('sum')

    qty_diff = (merged_df['PO Total QTY'] - merged_df['Target Commit QTY']).abs()
    merged_df['Total QTY Match'] = np.where(
        merged_df['Target Commit QTY'].isna(),
        False,
        (qty_diff < case_pack) | (qty_diff <= 0.10 * merged_df['Target Commit QTY'])
    )
    merged_df['QTY Diff'] = merged_df['PO Total QTY'] - merged_df['Target Commit QTY']
    merged_df['QTY Diff %'] = np.where(
        merged_df['Target Commit QTY'].replace(0, np.nan).isna(),
        np.nan,
        (merged_df['QTY Diff'] / merged_df['Target Commit QTY'] * 100).round(1)
    )

    # UPC 比對（只在兩側都有值時才比對）
    upc_both_exist = merged_df['PO UPC'].notna() & merged_df['Target UPC'].notna()
    merged_df['UPC Match'] = np.where(
        upc_both_exist,
        merged_df['PO UPC'] == merged_df['Target UPC'],
        True  # 缺一側 → N/A，不算失敗
    )
    merged_df['UPC Status'] = np.where(
        ~upc_both_exist, '⚪ 無資料',
        np.where(merged_df['UPC Match'], '✅ 相符', '❌ 不符')
    )

    # G6: Assortment 成本反算驗證
    if asst_df is not None and len(asst_df) > 0 and 'Units_in_Assortment' in merged_df.columns:
        # 計算每個 assortment box 的理論成本 = Σ(component FCA × units_in_asst)
        asst_cost_check = merged_df[asst_mask & merged_df['Final_Product_Cost'].notna() & merged_df['Units_in_Assortment'].notna()].copy()
        if len(asst_cost_check) > 0:
            asst_cost_check['_component_contribution'] = asst_cost_check['Final_Product_Cost'] * asst_cost_check['Units_in_Assortment']
            calc_box_cost = asst_cost_check.groupby('Original_DPCI')['_component_contribution'].sum().reset_index()
            calc_box_cost.columns = ['Original_DPCI', 'Calc_Box_Cost']
            merged_df = pd.merge(merged_df, calc_box_cost, on='Original_DPCI', how='left')
            merged_df['Asst_Cost_Check'] = np.where(
                asst_mask & merged_df['Asst_Box_Cost'].notna() & merged_df['Calc_Box_Cost'].notna(),
                np.isclose(merged_df['Asst_Box_Cost'], merged_df['Calc_Box_Cost'], atol=0.005),
                np.nan  # 非混裝或資料不足 → N/A
            )
            # 轉為可顯示狀態
            merged_df['Asst_Cost_Status'] = np.where(
                merged_df['Asst_Cost_Check'].isna(), '⚪ N/A',
                np.where(merged_df['Asst_Cost_Check'] == 1.0, '✅ 反算相符', '❌ 反算不符')
            )
        else:
            merged_df['Calc_Box_Cost'] = np.nan
            merged_df['Asst_Cost_Status'] = '⚪ N/A'
    else:
        merged_df['Calc_Box_Cost'] = np.nan
        merged_df['Asst_Cost_Status'] = '⚪ N/A'

    # 總覽判定（All Match）
    merged_df['All Match (Pass)'] = (
        merged_df['Cost Match'] &
        merged_df['Retail Match'] &
        merged_df['Case QTY Match'] &
        merged_df['Total QTY Match'] &
        merged_df['UPC Match']
    )

    return merged_df

# ==========================================
# 7. 結果顯示與下載
# ==========================================
def style_result(df):
    """顏色標示：True=綠, False=紅, 中性欄=白"""
    bool_cols = [c for c in ['Cost Match', 'Retail Match', 'Case QTY Match', 'Total QTY Match', 'UPC Match', 'All Match (Pass)'] if c in df.columns]

    def color_row(row):
        styles = [''] * len(row)
        for col in bool_cols:
            if col in row.index:
                idx = row.index.get_loc(col)
                if row[col] is True:
                    styles[idx] = 'background-color: #d4edda; color: #155724'
                elif row[col] is False:
                    styles[idx] = 'background-color: #f8d7da; color: #721c24'
        # All Match 整列著色
        if 'All Match (Pass)' in row.index:
            if row['All Match (Pass)'] is False:
                return ['background-color: #fff3cd' if s == '' else s for s in styles]
        return styles

    return df.style.apply(color_row, axis=1)

def show_results(merged_df, source_label, run_meta=None):
    """
    統一結果顯示 + 彩色 + 摘要統計 + G8 Excel 下載（含執行摘要 sheet）
    run_meta: dict with keys: input_files, timestamp
    """
    display_cols = [
        'PO NUMBER', 'ASSORTMENT ITEM?', 'Is_Shipper_Display', 'Original_DPCI', 'Final_DPCI',
        'ITEM DESCRIPTION', 'Final_QTY', 'Final_QTY_for_count',
        'Cost Match', 'ITEM UNIT COST', 'Target_Cost',
        'Retail Match', 'ITEM UNIT RETAIL', 'Suggested Unit Retail',
        'Case QTY Match', 'PO VCP / Assort QTY', 'Target Case / Assort QTY',
        'Total QTY Match', 'PO Total QTY', 'Target Commit QTY', 'QTY Diff', 'QTY Diff %',
        'UPC Status', 'PO UPC', 'Target UPC',
        'Asst_Cost_Status', 'Asst_Box_Cost', 'Calc_Box_Cost',
        'Dispatch_Factory', 'Dispatch_AC', 'Dispatch_AE',
        'All Match (Pass)'
    ]
    result_df = merged_df[[c for c in display_cols if c in merged_df.columns]].copy()
    errors_df = result_df[result_df['All Match (Pass)'] == False]

    # 摘要統計
    total = len(result_df)
    pass_count = (result_df['All Match (Pass)'] == True).sum()
    fail_count = (result_df['All Match (Pass)'] == False).sum()

    col1, col2, col3 = st.columns(3)
    col1.metric("📋 總筆數", total)
    col2.metric("✅ 全部相符", pass_count, delta=None)
    col3.metric("❌ 發現異常", fail_count, delta=None, delta_color="inverse")

    if fail_count > 0:
        # 各欄位異常數
        check_cols = ['Cost Match', 'Retail Match', 'Case QTY Match', 'Total QTY Match']
        col_errors = {c: int((result_df[c] == False).sum()) for c in check_cols if c in result_df.columns}
        upc_mismatch = int((result_df.get('UPC Status', pd.Series()) == '❌ 不符').sum())
        if upc_mismatch > 0:
            col_errors['UPC 不符'] = upc_mismatch
        asst_mismatch = int((result_df.get('Asst_Cost_Status', pd.Series()) == '❌ 反算不符').sum())
        if asst_mismatch > 0:
            col_errors['混裝成本反算不符'] = asst_mismatch

        st.markdown("**各欄位異常數**")
        err_summary = pd.DataFrame({'欄位': list(col_errors.keys()), '異常筆數': list(col_errors.values())})
        st.dataframe(err_summary, hide_index=True)

        st.markdown("**❌ 異常明細**")
        st.dataframe(style_result(errors_df), use_container_width=True)
    else:
        st.balloons()
        st.success("🎉 所有資料皆一致！")

    st.markdown("**📄 完整核對結果**")
    st.dataframe(style_result(result_df), use_container_width=True)

    # G8: 下載 Excel（含執行摘要 sheet）
    def strip_emoji(s):
        return re.sub(r'[^\x00-\x7F一-鿿Ā-ɏ]+', '', str(s))

    def make_excel_bytes(result_df, run_meta):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Sheet 1: 核對結果（去 emoji）
            clean_df = result_df.copy()
            for col in clean_df.columns:
                if clean_df[col].dtype == object:
                    clean_df[col] = clean_df[col].astype(str).apply(strip_emoji)
            clean_df.to_excel(writer, sheet_name='核對結果', index=False)

            # G8 Sheet 2: 執行摘要
            summary_data = {
                '項目': ['執行時間', '資料來源', '總筆數', '全部相符', '發現異常', '相符率 (%)'],
                '內容': [
                    run_meta.get('timestamp', datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
                    run_meta.get('input_files', source_label),
                    total,
                    pass_count,
                    fail_count,
                    f"{pass_count/total*100:.1f}%" if total > 0 else 'N/A'
                ]
            }
            pd.DataFrame(summary_data).to_excel(writer, sheet_name='執行摘要', index=False)

        return output.getvalue()

    run_meta = run_meta or {'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 'input_files': source_label}
    excel_bytes = make_excel_bytes(result_df, run_meta)
    safe_label = re.sub(r'[^\w\-]', '_', source_label)
    st.download_button(
        "📥 下載核對報告 (Excel，含執行摘要)",
        data=excel_bytes,
        file_name=f'PO_Validation_{safe_label}_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

# ==========================================
# 8. Streamlit 網頁介面
# ==========================================
st.set_page_config(page_title="訂單自動核對系統", layout="wide")
st.title("📦 跨專案訂單自動核對系統")

# ---- Sidebar ----
st.sidebar.header("📂 步驟 1：上傳共通資料庫")
product_files = st.sidebar.file_uploader("上傳 產品資料表 (可多選)", type=['csv', 'xlsx'], accept_multiple_files=True)
asst_files = st.sidebar.file_uploader("上傳 混裝箱表單 (可多選/選填)", type=['csv', 'xlsx'], accept_multiple_files=True)

# G7: Dispatch table（選填）
st.sidebar.markdown("---")
st.sidebar.header("🗂 步驟 2（選填）：上傳 Dispatch 對照表")
dispatch_files = st.sidebar.file_uploader("上傳 Dispatch 表（Factory / AC / AE 對照）", type=['csv', 'xlsx'], accept_multiple_files=True)
dispatch_df_global = process_dispatch(dispatch_files) if dispatch_files else pd.DataFrame()

# ---- Tabs ----
tab1, tab2, tab3 = st.tabs(["📊 標準版 (Standard PO) 核對", "📈 現代版 (Modern PO) 核對", "📄 PDF 上傳解析"])

# ==========================================
# Tab 1: 標準版 CSV
# ==========================================
with tab1:
    st.subheader("上傳標準版 PO 並執行核對")
    po_file_std = st.file_uploader("📥 上傳 Purchase Order Item Details (CSV)", type=['csv'], key="std_po")

    if st.button("🚀 開始核對標準版", type="primary", key="btn_std"):
        if not product_files or not po_file_std:
            st.warning("⚠️ 請確保已在側邊欄上傳「產品資料表」，並在上方上傳「標準版 PO」！")
        else:
            with st.spinner("標準版資料清洗與比對中..."):
                raw_po_df = process_standard_po(pd.read_csv(po_file_std))

                # G3: 重複 PO 偵測
                clean_po_df, dup_warnings = detect_duplicate_pos(raw_po_df)
                for w in dup_warnings:
                    st.warning(w)

                # G5: PO 自我驗證
                self_verify_warnings = po_self_verify(clean_po_df, mode='standard')
                for w in self_verify_warnings:
                    st.warning(w)

                prod_df = process_products(product_files)
                asst_df = process_assortments(asst_files) if asst_files else None

                merged_df = run_validation(clean_po_df, prod_df, asst_df, mode='standard')

                # G7: 加入 Dispatch 欄
                if len(dispatch_df_global) > 0:
                    merged_df = pd.merge(merged_df, dispatch_df_global, left_on='Final_DPCI', right_on='DPCI', how='left', suffixes=('', '_dispatch'))

                run_meta = {
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'input_files': f"PO: {po_file_std.name} | Products: {', '.join(f.name for f in product_files)}"
                }
                show_results(merged_df, 'Standard', run_meta=run_meta)

# ==========================================
# Tab 2: 現代版 CSV
# ==========================================
with tab2:
    st.subheader("上傳現代版 PO 並執行核對")
    po_file_mod = st.file_uploader("📥 上傳 Modern PO Visibility (CSV)", type=['csv'], key="mod_po")

    if st.button("🚀 開始核對現代版", type="primary", key="btn_mod"):
        if not product_files or not po_file_mod:
            st.warning("⚠️ 請確保已在側邊欄上傳「產品資料表」，並在上方上傳「現代版 PO」！")
        else:
            with st.spinner("現代版資料清洗與比對中..."):
                raw_po_df = process_modern_po(pd.read_csv(po_file_mod))

                # G3
                clean_po_df, dup_warnings = detect_duplicate_pos(raw_po_df)
                for w in dup_warnings:
                    st.warning(w)

                # G5
                self_verify_warnings = po_self_verify(clean_po_df, mode='modern')
                for w in self_verify_warnings:
                    st.warning(w)

                prod_df = process_products(product_files)
                asst_df = process_assortments(asst_files) if asst_files else None

                merged_df = run_validation(clean_po_df, prod_df, asst_df, mode='modern')

                if len(dispatch_df_global) > 0:
                    merged_df = pd.merge(merged_df, dispatch_df_global, left_on='Final_DPCI', right_on='DPCI', how='left', suffixes=('', '_dispatch'))

                run_meta = {
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'input_files': f"PO: {po_file_mod.name} | Products: {', '.join(f.name for f in product_files)}"
                }
                show_results(merged_df, 'Modern', run_meta=run_meta)

# ==========================================
# Tab 3: G1 PDF 上傳解析
# ==========================================
with tab3:
    st.subheader("📄 直接上傳 SPS Commerce PO PDF")
    st.info("免匯出 CSV！直接上傳 SPS Commerce 標準版 PO 的 PDF 檔案，系統將自動解析並執行核對。")

    pdf_files = st.file_uploader("📥 上傳 SPS Commerce PO PDF（可多選）", type=['pdf'], accept_multiple_files=True, key="pdf_po")

    if st.button("🚀 解析 PDF 並執行核對", type="primary", key="btn_pdf"):
        if not product_files or not pdf_files:
            st.warning("⚠️ 請確保已在側邊欄上傳「產品資料表」，並在上方上傳 PDF！")
        else:
            all_parsed_dfs = []
            all_parse_warnings = []

            with st.spinner("PDF 解析中..."):
                for pdf_file in pdf_files:
                    po_df_parsed, parse_warnings = parse_sps_pdf(pdf_file)
                    all_parse_warnings.extend(parse_warnings)
                    if po_df_parsed is not None and len(po_df_parsed) > 0:
                        all_parsed_dfs.append(po_df_parsed)

            for w in all_parse_warnings:
                st.warning(w)

            if not all_parsed_dfs:
                st.error("❌ 所有 PDF 均無法解析出訂單資料，請確認格式後再試。")
            else:
                combined_po_df = pd.concat(all_parsed_dfs, ignore_index=True)
                st.success(f"✅ 成功從 {len(pdf_files)} 份 PDF 解析出 {len(combined_po_df)} 筆訂單列。")

                # G3
                clean_po_df, dup_warnings = detect_duplicate_pos(combined_po_df)
                for w in dup_warnings:
                    st.warning(w)

                prod_df = process_products(product_files)
                asst_df = process_assortments(asst_files) if asst_files else None

                merged_df = run_validation(clean_po_df, prod_df, asst_df, mode='standard')

                if len(dispatch_df_global) > 0:
                    merged_df = pd.merge(merged_df, dispatch_df_global, left_on='Final_DPCI', right_on='DPCI', how='left', suffixes=('', '_dispatch'))

                run_meta = {
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'input_files': f"PDFs: {', '.join(f.name for f in pdf_files)} | Products: {', '.join(f.name for f in product_files)}"
                }
                show_results(merged_df, 'PDF', run_meta=run_meta)
