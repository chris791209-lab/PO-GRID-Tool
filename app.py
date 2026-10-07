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
# 2. G3 重複 PO 偵測（含 N7: superseded_with_qty_change）
# ==========================================
def detect_duplicate_pos(po_df, po_number_col='PO NUMBER'):
    """
    偵測兩種重複情形：
    (a) 同 PO# 有多個版本 → 保留最新（依列序），回傳警告訊息
        N7: 若同 PO# 的數量跨版本有變動，標示 superseded_with_qty_change
    (b) 不同 PO# 但內容完全相同 → 回傳警告訊息
    回傳 (cleaned_df, warnings_list)
    """
    warnings = []
    df = po_df.copy()

    dpci_col = 'Final_DPCI' if 'Final_DPCI' in df.columns else 'Original_DPCI'
    qty_col = 'Final_QTY'

    if dpci_col in df.columns:
        # (a) 同 PO# 多版本：偵測重複
        dup_pos = df[df.duplicated(subset=[po_number_col, dpci_col], keep=False)][po_number_col].unique()
        if len(dup_pos) > 0:
            # N7: 在去重之前，偵測哪些 PO# 的數量跨版本有變動
            if qty_col in df.columns:
                qty_changed_pos = []
                for po_num in dup_pos:
                    po_rows = df[df[po_number_col] == po_num]
                    # 按 DPCI group，看是否有不同數量
                    qty_by_dpci = po_rows.groupby(dpci_col)[qty_col].nunique()
                    if (qty_by_dpci > 1).any():
                        qty_changed_pos.append(str(po_num))
                if qty_changed_pos:
                    warnings.append(
                        f"⚠️ **[N7] 版本更新且數量變動**：以下 PO# 有多個版本且至少一個品項的訂購數量已更改，"
                        f"請確認使用最新版本的數量：\n"
                        + ", ".join(qty_changed_pos)
                    )

            warnings.append(
                f"⚠️ **版本重複偵測**：以下 PO# 有完全重複的品項列，已自動保留最後出現的版本（請確認是否為最新修訂單）：\n"
                + ", ".join(str(p) for p in dup_pos)
            )
            df = df.drop_duplicates(subset=[po_number_col, dpci_col], keep='last')

        # (b) 不同 PO# 但 DPCI 組合完全相同
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
    """
    G7 + N6: 讀取工廠&人員隸屬清單，輸出 DPCI → Factory / AC / AE 對應。
    N6: 同時保留 Factory_Name_from_Dispatch 供 PCN Factory Name 比對。
    """
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

    dpci_col = fuzzy_col(dispatch_df.columns, 'dpci')
    factory_col = fuzzy_col(dispatch_df.columns, 'factory')
    ac_col = fuzzy_col(dispatch_df.columns, 'ac') or fuzzy_col(dispatch_df.columns, 'account')
    ae_col = fuzzy_col(dispatch_df.columns, 'ae') or fuzzy_col(dispatch_df.columns, 'program')

    keep = {}
    if dpci_col:
        keep['DPCI'] = dispatch_df[dpci_col].pipe(clean_dpci)
    if factory_col:
        keep['Dispatch_Factory'] = dispatch_df[factory_col].astype(str).str.strip()
    if ac_col:
        keep['Dispatch_AC'] = dispatch_df[ac_col]
    if ae_col:
        keep['Dispatch_AE'] = dispatch_df[ae_col]

    # N6: also keep program column separately if it exists (AE follows Program rule)
    program_col = fuzzy_col(dispatch_df.columns, 'program')
    if program_col and program_col != ae_col:
        keep['Dispatch_Program'] = dispatch_df[program_col]

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

    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                tables = page.extract_tables()
                text = page.extract_text() or ''

                po_match = re.search(r'PO\s*(?:NUMBER|#)[:\s]+(\d{8,12})', text, re.IGNORECASE)
                if po_match:
                    current_po = po_match.group(1)

                for table in tables:
                    if not table:
                        continue
                    header_row_idx = None
                    for i, row in enumerate(table):
                        row_str = ' '.join(str(c) for c in row if c).upper()
                        if ('DPCI' in row_str or 'ITEM' in row_str) and ('COST' in row_str or 'QTY' in row_str or 'QUANTITY' in row_str):
                            header_row_idx = i
                            break

                    if header_row_idx is None:
                        continue

                    headers = [str(c).strip() if c else '' for c in table[header_row_idx]]

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
# 6. 共通驗證引擎（G2 / G4 / G6 / N2 / N3 / N4 / N6 整合）
# ==========================================
def run_validation(po_df, prod_df, asst_df, mode='standard', dispatch_df=None):
    """
    共通驗證引擎：
    - G4: Cost tolerance = 0.005
    - N2: Retail tolerance = 0.005（修正自 0.01）
    - G2: QTY tolerance = case_pack rounding OR ±10% of plan
    - G6: Assortment cost reverse-check
    - N3: Assortment QTY expansion check（boxes × units_per_box vs PO prepack qty）
    - G9: Shipper Display 排除於 QTY 加總
    - N4: AC/AE 規則驗證（AC follows factory, AE follows Program）
    - N6: Factory Name from PCN → Dispatch 工廠比對
    """
    # N6: 取得 PCN 的 Factory Name 欄位（若存在）
    factory_col_in_prod = None
    if prod_df is not None and 'Factory Name' in prod_df.columns:
        factory_col_in_prod = 'Factory Name'

    # ---- 合併 PCN ----
    prod_cols = [c for c in ['DPCI', 'Final_Product_Cost', 'Suggested Unit Retail',
                              'Case Unit Quantity', 'Ent Ttl Rcpt U', 'Target UPC', 'Factory Name']
                 if c in prod_df.columns]
    merged_df = pd.merge(
        po_df,
        prod_df[prod_cols].drop_duplicates(subset=['DPCI']),
        left_on='Final_DPCI', right_on='DPCI', how='left'
    )

    # ---- 混裝箱處理 ----
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
        else:
            condensed_asst = asst_df.groupby('Assortment_DPCI', as_index=False).agg({'Asst_Box_Cost': 'first'})
            merged_df = pd.merge(merged_df, condensed_asst, left_on='Original_DPCI', right_on='Assortment_DPCI', how='left')
            merged_df['ASSORTMENT ITEM?'] = np.where(merged_df['Asst_Box_Cost'].notna(), 'Y', 'N')
            merged_df['Target_Cost'] = np.where(merged_df['ASSORTMENT ITEM?'] == 'Y', merged_df['Asst_Box_Cost'], merged_df['Final_Product_Cost'])
    else:
        merged_df['Target_Cost'] = merged_df['Final_Product_Cost']
        if 'Units_in_Assortment' not in merged_df.columns:
            merged_df['Units_in_Assortment'] = np.nan
        if 'Asst_Box_Cost' not in merged_df.columns:
            merged_df['Asst_Box_Cost'] = np.nan

    asst_mask = merged_df.get('ASSORTMENT ITEM?', pd.Series('N', index=merged_df.index)) == 'Y'

    # ---- G4: 成本比對（tolerance = 0.005）----
    merged_df['Cost Match'] = np.isclose(
        merged_df['ITEM UNIT COST'].fillna(-1),
        merged_df['Target_Cost'].fillna(-1),
        atol=0.005
    )
    merged_df['Cost Match'] = np.where(merged_df['Target_Cost'].isna(), False, merged_df['Cost Match'])

    # ---- N2: 零售價比對（tolerance = 0.005）----
    if 'Suggested Unit Retail' in merged_df.columns:
        merged_df['Retail Match'] = np.where(
            asst_mask,
            True,  # 混裝箱 master row 零售標 n/a
            np.isclose(
                merged_df['ITEM UNIT RETAIL'].fillna(0),
                merged_df['Suggested Unit Retail'].fillna(0),
                atol=0.005  # N2: 修正自 0.01 → 0.005
            )
        )
    else:
        merged_df['Retail Match'] = True

    # ---- 裝箱數比對 ----
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

    # ---- G2: 總數量比對（case_pack rounding + 10% 容忍）----
    merged_df['Target Commit QTY'] = merged_df.get('Ent Ttl Rcpt U', pd.Series(np.nan, index=merged_df.index))
    case_pack = merged_df.get('Case Unit Quantity', pd.Series(1, index=merged_df.index)).fillna(1)

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

    # ---- UPC 比對 ----
    upc_both_exist = merged_df['PO UPC'].notna() & merged_df['Target UPC'].notna()
    merged_df['UPC Match'] = np.where(
        upc_both_exist,
        merged_df['PO UPC'] == merged_df['Target UPC'],
        True
    )
    merged_df['UPC Status'] = np.where(
        ~upc_both_exist, '⚪ 無資料',
        np.where(merged_df['UPC Match'], '✅ 相符', '❌ 不符')
    )

    # ---- G6: Assortment 成本反算驗證 ----
    if asst_df is not None and len(asst_df) > 0 and 'Units_in_Assortment' in merged_df.columns:
        asst_cost_check = merged_df[asst_mask & merged_df['Final_Product_Cost'].notna() & merged_df['Units_in_Assortment'].notna()].copy()
        if len(asst_cost_check) > 0:
            asst_cost_check['_component_contribution'] = asst_cost_check['Final_Product_Cost'] * asst_cost_check['Units_in_Assortment']
            calc_box_cost = asst_cost_check.groupby('Original_DPCI')['_component_contribution'].sum().reset_index()
            calc_box_cost.columns = ['Original_DPCI', 'Calc_Box_Cost']
            merged_df = pd.merge(merged_df, calc_box_cost, on='Original_DPCI', how='left')
            merged_df['Asst_Cost_Check'] = np.where(
                asst_mask & merged_df['Asst_Box_Cost'].notna() & merged_df['Calc_Box_Cost'].notna(),
                np.isclose(merged_df['Asst_Box_Cost'], merged_df['Calc_Box_Cost'], atol=0.005),
                np.nan
            )
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

    # ---- N3: Assortment QTY 展開驗證 ----
    # 對每個 assortment box DPCI：boxes in PO × units_per_box = expected component qty
    # 比較 PO 中記錄的 component qty（Final_QTY for asst rows）
    merged_df['Asst_QTY_Check'] = '⚪ N/A'
    if asst_df is not None and len(asst_df) > 0 and 'Units_in_Assortment' in merged_df.columns:
        # 找出 assortment 行（box level）的訂購數量
        # Standard PO: assortment rows have Original_DPCI = box DPCI, ASSORTMENT ITEM? = Y
        asst_rows = merged_df[asst_mask].copy()
        if len(asst_rows) > 0 and 'COMPONENT ASSORT QTY' in asst_rows.columns:
            # boxes = Original_DPCI 的 Total qty (sum of Final_QTY for original asst rows)
            # For each component: expected = boxes × units_in_assortment
            # Actual = COMPONENT ASSORT QTY (the prepack quantity in the PO for that component)
            box_qtys = po_df[po_df.get('ASSORTMENT ITEM?', pd.Series('N')) == 'Y'].groupby(
                po_df.get('Original_DPCI', po_df.index)
            )['Final_QTY'].sum().reset_index() if 'Original_DPCI' in po_df.columns else pd.DataFrame()

            if len(box_qtys) > 0:
                box_qtys.columns = ['Original_DPCI', 'Total_Box_QTY']
                asst_check = asst_rows.merge(box_qtys, on='Original_DPCI', how='left')
                asst_check['Expected_Component_QTY'] = asst_check['Total_Box_QTY'] * asst_check['Units_in_Assortment']
                asst_check['Asst_QTY_OK'] = np.where(
                    asst_check['Expected_Component_QTY'].notna() & asst_check['COMPONENT ASSORT QTY'].notna(),
                    np.isclose(asst_check['Expected_Component_QTY'], asst_check['COMPONENT ASSORT QTY'], atol=0.5),
                    True
                )
                # Map results back to merged_df
                asst_check['Asst_QTY_Check'] = np.where(
                    asst_check['Expected_Component_QTY'].isna() | asst_check['COMPONENT ASSORT QTY'].isna(),
                    '⚪ N/A',
                    np.where(asst_check['Asst_QTY_OK'], '✅ 數量相符', '❌ 數量不符')
                )
                # Store expected component qty
                for idx_val, row_val in asst_check.iterrows():
                    if idx_val in merged_df.index:
                        merged_df.loc[idx_val, 'Asst_QTY_Check'] = row_val['Asst_QTY_Check']
                        merged_df.loc[idx_val, 'Expected_Component_QTY'] = row_val.get('Expected_Component_QTY', np.nan)

    if 'Expected_Component_QTY' not in merged_df.columns:
        merged_df['Expected_Component_QTY'] = np.nan

    # ---- N4: Dispatch AC/AE 規則驗證 ----
    # 規則：AC follows factory（用 Factory Name 從 PCN 找 factory → dispatch AC 應與工廠一致）
    #       AE follows Program（dispatch AE 應與 Program 一致）
    merged_df['AC_Rule_Check'] = '⚪ N/A'
    merged_df['AE_Rule_Check'] = '⚪ N/A'

    if dispatch_df is not None and len(dispatch_df) > 0:
        # N6: 使用 Factory Name 從 PCN 進行工廠比對（若 dispatch_df 有 DPCI 欄可直接 join）
        # Merge dispatch by DPCI
        dispatch_cols = ['DPCI'] + [c for c in ['Dispatch_Factory', 'Dispatch_AC', 'Dispatch_AE', 'Dispatch_Program'] if c in dispatch_df.columns]
        merged_df = pd.merge(
            merged_df,
            dispatch_df[dispatch_cols],
            left_on='Final_DPCI', right_on='DPCI', how='left',
            suffixes=('', '_disp')
        )
        # Remove duplicate DPCI col from dispatch merge
        if 'DPCI_disp' in merged_df.columns:
            merged_df.drop(columns=['DPCI_disp'], inplace=True, errors='ignore')

        # N6: If PCN has Factory Name, check if it matches Dispatch_Factory
        if factory_col_in_prod and 'Factory Name' in merged_df.columns and 'Dispatch_Factory' in merged_df.columns:
            both_have_factory = merged_df['Factory Name'].notna() & merged_df['Dispatch_Factory'].notna()
            # Normalize for comparison (strip whitespace, case-insensitive)
            pcn_factory_norm = merged_df['Factory Name'].astype(str).str.strip().str.lower()
            disp_factory_norm = merged_df['Dispatch_Factory'].astype(str).str.strip().str.lower()
            merged_df['Factory_Match'] = np.where(
                both_have_factory,
                pcn_factory_norm == disp_factory_norm,
                np.nan
            )
            merged_df['Factory_Match_Status'] = np.where(
                ~both_have_factory, '⚪ N/A',
                np.where(merged_df['Factory_Match'] == 1.0, '✅ 工廠一致', '⚠️ 工廠不符')
            )
        else:
            merged_df['Factory_Match_Status'] = '⚪ N/A'

        # N4: AC follows factory — if Factory_Match_Status is a mismatch, AC assignment may be wrong
        # We flag if: Dispatch_Factory in PCN doesn't match Dispatch table factory
        if 'Factory_Match_Status' in merged_df.columns:
            merged_df['AC_Rule_Check'] = np.where(
                merged_df['Factory_Match_Status'] == '⚠️ 工廠不符',
                '⚠️ 工廠不符→AC待確認',
                np.where(merged_df['Dispatch_AC'].notna(), '✅ 已對應 AC', '⚪ N/A')
            )

        # N4: AE follows Program — check AE assignment
        if 'Dispatch_AE' in merged_df.columns:
            merged_df['AE_Rule_Check'] = np.where(
                merged_df['Dispatch_AE'].notna(), '✅ 已對應 AE', '⚪ N/A'
            )

    else:
        merged_df['Factory_Match_Status'] = '⚪ N/A'

    # ---- 總覽判定（All Match）----
    merged_df['All Match (Pass)'] = (
        merged_df['Cost Match'] &
        merged_df['Retail Match'] &
        merged_df['Case QTY Match'] &
        merged_df['Total QTY Match'] &
        merged_df['UPC Match']
    )

    return merged_df

# ==========================================
# 7. 結果顯示與下載（含 N1 GRID + N5 Validation Summary）
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
        if 'All Match (Pass)' in row.index:
            if row['All Match (Pass)'] is False:
                return ['background-color: #fff3cd' if s == '' else s for s in styles]
        return styles

    return df.style.apply(color_row, axis=1)

def build_po_grid(merged_df):
    """
    N1: 建立 PO GRID 矩陣
    - 列：Final_DPCI（排除 Shipper Display）
    - 欄：各 PO NUMBER
    - 值：該 PO 對該 DPCI 的 Final_QTY
    - 最後欄：TOTAL（各列加總）
    回傳 grid_df (DataFrame)
    """
    df = merged_df.copy()

    # 排除 Shipper Display rows
    if 'Is_Shipper_Display' in df.columns:
        df = df[~df['Is_Shipper_Display'].astype(bool)]

    # 只保留有 PO NUMBER 和 DPCI 的行
    df = df[df['Final_DPCI'].notna() & df['PO NUMBER'].notna()]

    if len(df) == 0:
        return pd.DataFrame()

    # 聚合（同一 PO+DPCI 可能有多行，取 sum）
    grid_pivot = df.groupby(['Final_DPCI', 'PO NUMBER'])['Final_QTY'].sum().unstack(fill_value=0)
    grid_pivot = grid_pivot.reset_index()
    grid_pivot.columns.name = None

    # 排序 PO 欄（字串排序）
    po_cols = sorted([c for c in grid_pivot.columns if c != 'Final_DPCI'])
    grid_pivot = grid_pivot[['Final_DPCI'] + po_cols]

    # TOTAL 欄
    grid_pivot['TOTAL'] = grid_pivot[po_cols].sum(axis=1)

    # 加入 Plan QTY (Ent Ttl Rcpt U) 若存在
    if 'Ent Ttl Rcpt U' in merged_df.columns and 'Target Commit QTY' in merged_df.columns:
        plan_lookup = merged_df[['Final_DPCI', 'Target Commit QTY']].drop_duplicates(subset=['Final_DPCI'])
        grid_pivot = grid_pivot.merge(plan_lookup, on='Final_DPCI', how='left')
        grid_pivot.rename(columns={'Target Commit QTY': 'PLAN QTY'}, inplace=True)

        # QTY Diff 欄
        if 'PLAN QTY' in grid_pivot.columns:
            grid_pivot['QTY vs PLAN'] = grid_pivot['TOTAL'] - grid_pivot['PLAN QTY']

    return grid_pivot

def make_excel_bytes(result_df, run_meta, source_label, validation_notes=None, merged_df_full=None):
    """
    N1 + N5 + G8: 建立含多 sheet 的 Excel
    - Sheet 1: 核對結果
    - Sheet 2: PO GRID（N1）
    - Sheet 3: 驗核摘要（N5）
    - Sheet 4: 執行摘要（G8）
    """
    def strip_emoji(s):
        return re.sub(r'[^\x00-\x7F一-鿿Ā-ɏ]+', '', str(s))

    total = len(result_df)
    pass_count = int((result_df['All Match (Pass)'] == True).sum())
    fail_count = int((result_df['All Match (Pass)'] == False).sum())

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:

        # ---- Sheet 1: 核對結果 ----
        clean_df = result_df.copy()
        for col in clean_df.columns:
            if clean_df[col].dtype == object:
                clean_df[col] = clean_df[col].astype(str).apply(strip_emoji)
        clean_df.to_excel(writer, sheet_name='核對結果', index=False)

        # 格式：凍結首列、自動欄寬
        ws = writer.sheets['核對結果']
        ws.freeze_panes = 'A2'
        for col_cells in ws.columns:
            max_len = max((len(str(cell.value)) for cell in col_cells if cell.value), default=8)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 30)

        # ---- Sheet 2: PO GRID（N1）----
        if merged_df_full is not None:
            grid_df = build_po_grid(merged_df_full)
        else:
            grid_df = build_po_grid(result_df)

        if len(grid_df) > 0:
            # 清理 emoji
            grid_clean = grid_df.copy()
            for col in grid_clean.columns:
                if grid_clean[col].dtype == object:
                    grid_clean[col] = grid_clean[col].astype(str).apply(strip_emoji)
            grid_clean.to_excel(writer, sheet_name='PO GRID', index=False)

            ws_grid = writer.sheets['PO GRID']
            ws_grid.freeze_panes = 'B2'

            # 嘗試套用顏色：QTY vs PLAN 欄 — 紅色代表超出 ±10%
            try:
                from openpyxl.styles import PatternFill, Font
                header_row = [cell.value for cell in ws_grid[1]]
                if 'QTY vs PLAN' in header_row:
                    diff_col_idx = header_row.index('QTY vs PLAN') + 1
                    plan_col_idx = header_row.index('PLAN QTY') + 1 if 'PLAN QTY' in header_row else None
                    red_fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
                    green_fill = PatternFill(start_color='CCFFCC', end_color='CCFFCC', fill_type='solid')
                    for row_idx in range(2, ws_grid.max_row + 1):
                        diff_cell = ws_grid.cell(row=row_idx, column=diff_col_idx)
                        plan_cell = ws_grid.cell(row=row_idx, column=plan_col_idx) if plan_col_idx else None
                        try:
                            diff_val = float(diff_cell.value or 0)
                            plan_val = float(plan_cell.value or 0) if plan_cell else 0
                            if plan_val > 0:
                                pct = abs(diff_val) / plan_val
                                if pct > 0.10:
                                    diff_cell.fill = red_fill
                                else:
                                    diff_cell.fill = green_fill
                        except (ValueError, TypeError):
                            pass
            except Exception:
                pass  # 顏色套用失敗不影響主功能

            for col_cells in ws_grid.columns:
                max_len = max((len(str(cell.value)) for cell in col_cells if cell.value), default=8)
                ws_grid.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 20)

        # ---- Sheet 3: 驗核摘要（N5）----
        check_cols = ['Cost Match', 'Retail Match', 'Case QTY Match', 'Total QTY Match']
        summary_rows = []
        for c in check_cols:
            if c in result_df.columns:
                n_fail = int((result_df[c] == False).sum())
                n_pass = int((result_df[c] == True).sum())
                summary_rows.append({'檢核項目': c, '相符筆數': n_pass, '異常筆數': n_fail, '建議動作': '請確認 PO 內容' if n_fail > 0 else '無需處理'})

        # UPC
        upc_fail = int((result_df.get('UPC Status', pd.Series()) == '❌ 不符').sum())
        upc_na = int((result_df.get('UPC Status', pd.Series()) == '⚪ 無資料').sum())
        summary_rows.append({'檢核項目': 'UPC', '相符筆數': total - upc_fail - upc_na, '異常筆數': upc_fail, '建議動作': '請確認條碼' if upc_fail > 0 else '無需處理'})

        # Assortment cost
        asst_fail = int((result_df.get('Asst_Cost_Status', pd.Series()) == '❌ 反算不符').sum())
        summary_rows.append({'檢核項目': '混裝成本反算', '相符筆數': 0, '異常筆數': asst_fail, '建議動作': '請確認混裝箱成本計算' if asst_fail > 0 else '無需處理'})

        # Factory match
        factory_fail = int((result_df.get('Factory_Match_Status', pd.Series()) == '⚠️ 工廠不符').sum())
        if factory_fail > 0:
            summary_rows.append({'檢核項目': 'Factory 工廠比對 (N6)', '相符筆數': 0, '異常筆數': factory_fail, '建議動作': '請確認 PCN 工廠與 Dispatch 清單是否一致'})

        # Skipped DPCIs (no PCN match)
        no_prod_match = int(result_df['Target_Cost'].isna().sum()) if 'Target_Cost' in result_df.columns else 0
        if no_prod_match > 0:
            summary_rows.append({'檢核項目': '無 PCN 對應 (略過)', '相符筆數': 0, '異常筆數': no_prod_match, '建議動作': '請確認 PCN 是否包含所有 DPCI'})

        # Add validation notes
        if validation_notes:
            for note in validation_notes:
                summary_rows.append({'檢核項目': '系統警告', '相符筆數': 0, '異常筆數': 1, '建議動作': strip_emoji(note)[:200]})

        val_summary_df = pd.DataFrame(summary_rows)
        val_summary_df.to_excel(writer, sheet_name='驗核摘要', index=False)
        ws_val = writer.sheets['驗核摘要']
        ws_val.freeze_panes = 'A2'
        for col_cells in ws_val.columns:
            max_len = max((len(str(cell.value)) for cell in col_cells if cell.value), default=10)
            ws_val.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 50)

        # ---- Sheet 4: 執行摘要（G8）----
        summary_data = {
            '項目': ['執行時間', '資料來源', '總筆數', '全部相符', '發現異常', '相符率 (%)',
                    '成本容忍值', '零售容忍值', '數量容忍規則'],
            '內容': [
                run_meta.get('timestamp', datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
                run_meta.get('input_files', source_label),
                total,
                pass_count,
                fail_count,
                f"{pass_count/total*100:.1f}%" if total > 0 else 'N/A',
                '±$0.005 (FCA/FOB)',
                '±$0.005 (Suggested Unit Retail)',
                '|diff| < case_pack OR |diff| ≤ 10% of plan'
            ]
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='執行摘要', index=False)
        ws_exec = writer.sheets['執行摘要']
        for col_cells in ws_exec.columns:
            max_len = max((len(str(cell.value)) for cell in col_cells if cell.value), default=10)
            ws_exec.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 60)

    return output.getvalue()

def show_results(merged_df, source_label, run_meta=None, validation_notes=None):
    """
    統一結果顯示 + 彩色 + 摘要統計 + Excel 下載（含 GRID / 驗核摘要 / 執行摘要）
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
        'Asst_QTY_Check', 'COMPONENT ASSORT QTY', 'Expected_Component_QTY',
        'Factory_Match_Status', 'Factory Name', 'Dispatch_Factory',
        'Dispatch_AC', 'Dispatch_AE', 'AC_Rule_Check', 'AE_Rule_Check',
        'All Match (Pass)'
    ]
    result_df = merged_df[[c for c in display_cols if c in merged_df.columns]].copy()
    errors_df = result_df[result_df['All Match (Pass)'] == False]

    total = len(result_df)
    pass_count = (result_df['All Match (Pass)'] == True).sum()
    fail_count = (result_df['All Match (Pass)'] == False).sum()

    col1, col2, col3 = st.columns(3)
    col1.metric("📋 總筆數", total)
    col2.metric("✅ 全部相符", pass_count)
    col3.metric("❌ 發現異常", fail_count)

    if fail_count > 0:
        check_cols = ['Cost Match', 'Retail Match', 'Case QTY Match', 'Total QTY Match']
        col_errors = {c: int((result_df[c] == False).sum()) for c in check_cols if c in result_df.columns}
        upc_mismatch = int((result_df.get('UPC Status', pd.Series()) == '❌ 不符').sum())
        if upc_mismatch > 0:
            col_errors['UPC 不符'] = upc_mismatch
        asst_mismatch = int((result_df.get('Asst_Cost_Status', pd.Series()) == '❌ 反算不符').sum())
        if asst_mismatch > 0:
            col_errors['混裝成本反算不符'] = asst_mismatch
        factory_mismatch = int((result_df.get('Factory_Match_Status', pd.Series()) == '⚠️ 工廠不符').sum())
        if factory_mismatch > 0:
            col_errors['Factory 工廠不符 (N6)'] = factory_mismatch

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

    # N1: 預覽 PO GRID
    st.markdown("---")
    st.markdown("**📊 PO GRID 訂購矩陣**")
    grid_df = build_po_grid(merged_df)
    if len(grid_df) > 0:
        # 顏色標示 QTY vs PLAN
        if 'QTY vs PLAN' in grid_df.columns and 'PLAN QTY' in grid_df.columns:
            def color_grid_row(row):
                styles = [''] * len(row)
                try:
                    diff = float(row.get('QTY vs PLAN', 0) or 0)
                    plan = float(row.get('PLAN QTY', 0) or 0)
                    if plan > 0 and abs(diff) / plan > 0.10:
                        return ['background-color: #FFCCCC'] * len(row)
                except (ValueError, TypeError):
                    pass
                return styles
            st.dataframe(grid_df.style.apply(color_grid_row, axis=1), use_container_width=True)
        else:
            st.dataframe(grid_df, use_container_width=True)
    else:
        st.info("GRID 矩陣暫無資料（需有 PO NUMBER 與 Final_DPCI）。")

    # 下載 Excel（N1 GRID + N5 摘要 + G8 執行摘要）
    run_meta = run_meta or {'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 'input_files': source_label}
    excel_bytes = make_excel_bytes(result_df, run_meta, source_label, validation_notes=validation_notes, merged_df_full=merged_df)
    safe_label = re.sub(r'[^\w\-]', '_', source_label)
    st.download_button(
        "📥 下載完整核對報告 (Excel：核對結果 + PO GRID + 驗核摘要 + 執行摘要)",
        data=excel_bytes,
        file_name=f'PO_Validation_{safe_label}_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

# ==========================================
# 8. Streamlit 網頁介面
# ==========================================
st.set_page_config(page_title="TG Team PO 驗證管理平台", layout="wide")

# LuckyStar Logo SVG（藍底白色人形金字塔，圓頭+弧形身體）
LOGO_SVG = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 60 60" width="52" height="52">
  <rect width="60" height="60" rx="8" fill="#1e6bb8"/>
  <!-- 底排 3人 -->
  <!-- 左 -->
  <circle cx="11" cy="43" r="5" fill="white"/>
  <path d="M5,54 Q11,49 17,54" fill="white"/>
  <!-- 中 -->
  <circle cx="30" cy="43" r="5" fill="white"/>
  <path d="M24,54 Q30,49 36,54" fill="white"/>
  <!-- 右 -->
  <circle cx="49" cy="43" r="5" fill="white"/>
  <path d="M43,54 Q49,49 55,54" fill="white"/>
  <!-- 中排 2人 -->
  <!-- 左中 -->
  <circle cx="20" cy="28" r="5" fill="white"/>
  <path d="M14,39 Q20,34 26,39" fill="white"/>
  <!-- 右中 -->
  <circle cx="40" cy="28" r="5" fill="white"/>
  <path d="M34,39 Q40,34 46,39" fill="white"/>
  <!-- 頂 1人 -->
  <circle cx="30" cy="13" r="5" fill="white"/>
  <path d="M24,24 Q30,19 36,24" fill="white"/>
</svg>
"""

col_logo, col_title = st.columns([0.07, 0.93])
with col_logo:
    st.markdown(LOGO_SVG, unsafe_allow_html=True)
with col_title:
    st.markdown("<h1 style='margin-top:4px; font-size:2rem;'>TG Team PO 驗證管理平台</h1>", unsafe_allow_html=True)

# ---- Sidebar ----
st.sidebar.header("📂 步驟 1：上傳共通資料庫")
product_files = st.sidebar.file_uploader("產品資料表 PCN（可多選）", type=['csv', 'xlsx'], accept_multiple_files=True)
asst_files = st.sidebar.file_uploader("混裝箱表單（可多選／選填）", type=['csv', 'xlsx'], accept_multiple_files=True)

st.sidebar.markdown("---")
st.sidebar.header("🏭 步驟 2（選填）：工廠 & 人員")
dispatch_files = st.sidebar.file_uploader("工廠&人員隸屬清單（Factory / AE / AC）", type=['csv', 'xlsx'], accept_multiple_files=True)
dispatch_df_global = process_dispatch(dispatch_files) if dispatch_files else pd.DataFrame()

# ==========================================
# 主介面：PDF 上傳解析
# ==========================================
if True:
    st.subheader("📄 上傳 SPS Commerce PO PDF")

    pdf_files = st.file_uploader("", type=['pdf'], accept_multiple_files=True, key="pdf_po", label_visibility="collapsed")

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

                clean_po_df, dup_warnings = detect_duplicate_pos(combined_po_df)
                all_warnings = all_parse_warnings + dup_warnings
                for w in dup_warnings:
                    st.warning(w)

                prod_df = process_products(product_files)
                asst_df = process_assortments(asst_files) if asst_files else None

                dispatch_arg = dispatch_df_global if len(dispatch_df_global) > 0 else None
                merged_df = run_validation(clean_po_df, prod_df, asst_df, mode='standard', dispatch_df=dispatch_arg)

                run_meta = {
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'input_files': f"PDFs: {', '.join(f.name for f in pdf_files)} | Products: {', '.join(f.name for f in product_files)}"
                }
                show_results(merged_df, 'PDF', run_meta=run_meta, validation_notes=all_warnings)

