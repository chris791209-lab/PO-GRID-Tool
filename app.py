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
    normalized = {c: re.sub(r'\s+', '', str(c)).lower() for c in df_cols if c is not None}
    for col, norm in normalized.items():
        hits = [kw.lower().replace(' ', '') in norm for kw in keywords]
        if (all(hits) if require_all else any(hits)):
            return col
    return None

def normalize_factory_name(s: str) -> str:
    """
    N6 工廠名稱正規化：
    - 移除常見法律後綴（Co., Ltd., Corp., Inc., LLC 等）
    - 移除標點、括號內容、多餘空白
    - 統一小寫
    讓 PCN 英文名與 Dispatch 表中英混用名仍可比對到最長公共核心。
    """
    if not isinstance(s, str):
        return ''
    result = s.lower()
    # 移除括號及其內容（含中文括號）
    result = re.sub(r'[\(（][^)）]*[\)）]', '', result)
    # 移除常見法律後綴
    suffixes = [
        r'\bco\.?,?\s*ltd\.?', r'\bcorp\.?', r'\binc\.?', r'\bllc\.?',
        r'\blimited', r'\bcompany', r'\bfactory', r'\bindustrial',
        r'\bmanufacturing', r'\benterprise', r'\bgroup',
        r'有限公司', r'工廠', r'工業', r'集團',
    ]
    for sfx in suffixes:
        result = re.sub(sfx, '', result)
    # 移除標點與多餘空白
    result = re.sub(r'[^\w\s]', ' ', result)
    result = re.sub(r'\s+', ' ', result).strip()
    return result

# ==========================================
# 2. G3 重複 PO 偵測（含 N7: superseded_with_qty_change）
# ==========================================
def detect_duplicate_pos(po_df, po_number_col='PO NUMBER'):
    """
    偵測兩種重複情形：
    (a) 同 PO# 有多個版本 → 依 PO 日期排序後保留最新版本，回傳警告訊息
        N7: 若同 PO# 的數量跨版本有變動，標示 superseded_with_qty_change
    (b) 不同 PO# 但內容完全相同 → 回傳警告訊息
    回傳 (cleaned_df, warnings_list)
    """
    warnings = []
    df = po_df.copy()

    dpci_col = 'Final_DPCI' if 'Final_DPCI' in df.columns else 'Original_DPCI'
    qty_col = 'Final_QTY'

    # ── 嘗試偵測 PO 日期欄位（PO DATE / ISSUE DATE / ORDER DATE / CANCEL DATE）──
    date_col = None
    for candidate in df.columns:
        c_norm = re.sub(r'\s+', '', str(candidate)).lower()
        if any(kw in c_norm for kw in ['podate', 'issuedate', 'orderdate', 'receiveddate', 'poreceiveddate']):
            date_col = candidate
            break
    # 若找到日期欄，嘗試解析；解析失敗則忽略
    if date_col:
        try:
            df['_sort_date'] = pd.to_datetime(df[date_col], errors='coerce')
            if df['_sort_date'].isna().all():
                date_col = None
                df.drop(columns=['_sort_date'], inplace=True)
        except Exception:
            date_col = None
            df.pop('_sort_date', None)

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

            # ── 依日期排序（若有），使最新版本落在最後，再用 keep='last' 保留 ──
            if date_col and '_sort_date' in df.columns:
                df = df.sort_values('_sort_date', ascending=True, na_position='first')
                version_note = "（已依 PO 收到日期自動選取最新版本）"
            else:
                version_note = "（依上傳順序保留最後出現的版本，建議上傳 PO 日期欄位以自動比對）"

            warnings.append(
                f"⚠️ **版本重複偵測**：以下 PO# 有完全重複的品項列，已自動保留最新版本{version_note}：\n"
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

    # 清理排序輔助欄
    if '_sort_date' in df.columns:
        df = df.drop(columns=['_sort_date'])

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

@st.cache_data(show_spinner=False)
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

@st.cache_data(show_spinner=False)
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
            raw_headers = raw_df.iloc[header_idx].astype(str).str.replace(r'[\n\r]', ' ', regex=True).str.strip()
            df.columns = raw_headers

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

@st.cache_data(show_spinner=False)
def process_dispatch(files):
    """
    G7 + N6: 讀取工廠&人員隸屬清單，輸出 DPCI → Factory / AC / AE 對應。
    N6: 同時保留 Factory_Name_from_Dispatch 供 PCN Factory Name 比對。
    files 可以是 file-like 物件的 list，或已是 DataFrame 的 list（圖片辨識後直傳）。
    """
    if not files:
        return pd.DataFrame()
    df_list = []
    for f in files:
        if isinstance(f, pd.DataFrame):
            df_list.append(f)
        else:
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
    G1: 解析 SPS Commerce / Target Import PO PDF，回傳 (po_df, parse_warnings)
    支援三種路線：
      1. pdfplumber extract_tables()：有真實表格結構時使用
      2. 純文字逐行解析（fallback）：SPS Commerce 文字版
      3. OCR fallback：Target Import PO（全向量圖形、無文字層）
         - 使用 pdf2image 光柵化 + pytesseract OCR
         - 支援 "Buyers Catalog Number: XXXXXXXXX" → DPCI DDD-CC-XXXX 轉換
    """
    try:
        import pdfplumber
    except ImportError:
        return None, ["❌ 缺少 pdfplumber 套件，請在 requirements.txt 加入 pdfplumber 並重新部署。"]

    warnings = []
    all_rows = []
    current_po = None

    # PO NUMBER 多種格式：
    # "Purchase Order: 1234567890" / "PO NUMBER: 1234567890" / "Order #: 10002036126-0581"
    PO_PATTERNS = [
        r'Order\s*#?\s*[:#]?\s*(\d{10,13}(?:-\d{3,5})?)',
        r'Purchase\s+Order\s*[:#]?\s*(\d{7,12})',
        r'PO\s*(?:NUMBER|#|No\.?)[:\s]+(\d{7,12})',
        r'\bP\.?O\.?\s*[:#]?\s*(\d{7,12})\b',
    ]
    DPCI_RE = re.compile(r'\b(\d{3}-\d{2}-\d{4})\b')
    # Buyers Catalog Number（Target Import format）: 9 digits → DDD-CC-XXXX
    CATALOG_RE = re.compile(r'Buyers\s+Catalog\s+Number\s*[:\s]+(\d{9})', re.IGNORECASE)
    # Also match plain 9-digit block that could be catalog number
    CATALOG_BARE_RE = re.compile(r'\b(\d{9})\b')

    def extract_po_number(text):
        for pat in PO_PATTERNS:
            m = re.search(pat, text, re.IGNORECASE)
            if m:
                raw = m.group(1)
                # Strip suffix like "-0581" if present — keep base PO number
                return raw.split('-')[0] if '-' in raw and len(raw.split('-')[0]) >= 10 else raw
        return None

    def catalog_to_dpci(catalog_num):
        """Convert 9-digit Buyers Catalog Number to DDD-CC-XXXX DPCI format."""
        s = str(catalog_num).zfill(9)
        return f"{s[:3]}-{s[3:5]}-{s[5:]}"

    def _build_row(po, dpci_raw, qty_raw, cost_raw=None, retail_raw=None,
                   desc_raw=None, upc_raw=None, asst_raw=None, vcp_raw=None):
        def to_num(s):
            if not s:
                return np.nan
            return pd.to_numeric(str(s).replace(',', '').replace('$', ''), errors='coerce')
        return {
            'PO NUMBER': po,
            'Original_DPCI': clean_dpci(pd.Series([dpci_raw])).iloc[0],
            'Final_DPCI': clean_dpci(pd.Series([dpci_raw])).iloc[0],
            'Final_QTY': to_num(qty_raw),
            'ITEM UNIT COST': to_num(cost_raw),
            'ITEM UNIT RETAIL': to_num(retail_raw),
            'ITEM DESCRIPTION': desc_raw,
            'PO UPC': clean_upc(pd.Series([upc_raw if upc_raw else np.nan])).iloc[0],
            'ASSORTMENT ITEM?': 'Y' if str(asst_raw or '').upper() in ['Y', 'YES'] else 'N',
            'VCP QUANTITY': to_num(vcp_raw),
            'COMPONENT ASSORT QTY': np.nan,
            'Is_Shipper_Display': False,
        }

    def _ocr_parse(pdf_file):
        """
        路線 3：OCR 解析 Target Import PO（向量圖形 PDF，無文字層）
        使用 pdf2image 光柵化後以 pytesseract 辨識。
        支援 "Buyers Catalog Number" 跨行格式（目錄編號在標題下一行）。
        """
        ocr_rows = []
        ocr_po = None
        try:
            from pdf2image import convert_from_bytes
            import pytesseract
        except ImportError as e:
            return [], [f"❌ OCR 套件缺失：{e}"]

        try:
            pdf_file.seek(0)
            pdf_bytes = pdf_file.read()
            images = convert_from_bytes(pdf_bytes, dpi=200, fmt='png')
        except Exception as e:
            return [], [f"❌ PDF 光柵化失敗：{e}"]

        ocr_warnings = []
        full_ocr_text = ''

        for img in images:
            ocr_text = pytesseract.image_to_string(img, lang='eng', config='--psm 6 --oem 3')
            full_ocr_text += ocr_text + '\n'

        lines = [l.rstrip() for l in full_ocr_text.splitlines()]

        i = 0
        while i < len(lines):
            line = lines[i]

            # ── Extract PO number ──
            po_found = extract_po_number(line)
            if po_found:
                ocr_po = po_found
                i += 1
                continue

            # ── Detect "Buyers Catalog Number" header line ──
            # Target Import format: "Buyers Catalog Number:" on one line,
            # then the 9-digit catalog number at the start of the NEXT line.
            if re.search(r'Buyers\s+Catalog\s+Number', line, re.IGNORECASE):
                # The catalog number is on the next non-empty line
                catalog_num = None
                # Check same line first (inline format)
                same_line_m = re.search(r'Buyers\s+Catalog\s+Number[:\s]+(\d{9})', line, re.IGNORECASE)
                if same_line_m:
                    catalog_num = same_line_m.group(1)
                else:
                    # Look at next line for a 9-digit number at start
                    for j in range(i + 1, min(i + 3, len(lines))):
                        next_m = re.match(r'\s*(\d{9})\b', lines[j])
                        if next_m:
                            catalog_num = next_m.group(1)
                            break

                if not catalog_num:
                    i += 1
                    continue

                dpci_raw = catalog_to_dpci(catalog_num)

                # Gather context: current line and surrounding ±6 lines
                ctx_start = max(0, i - 2)
                ctx_end = min(len(lines), i + 12)
                context_lines = lines[ctx_start:ctx_end]
                context_text = ' '.join(context_lines)

                # ── UPC (12-13 digit number) ──
                upc_m = re.search(r'\b(\d{12,13})\b', context_text)
                upc_raw = upc_m.group(1) if upc_m else None

                # ── Unit Price and Resale — use labeled patterns first ──
                up_m = re.search(r'Unit\s+Price\s*[:\s]+(\d+\.?\d*)', context_text, re.IGNORECASE)
                res_m = re.search(r'Resale\s*[:\s]+(\d+\.?\d*)', context_text, re.IGNORECASE)
                cost_raw = up_m.group(1) if up_m else None
                retail_raw = res_m.group(1) if res_m else None

                # Fallback: extract floats if labeled patterns didn't match
                if not cost_raw or not retail_raw:
                    float_nums = re.findall(r'\b(\d+\.\d{2})\b', context_text)
                    # Filter: remove UPC-digit runs and numbers > 9999 (totals)
                    price_candidates = []
                    for f in float_nums:
                        try:
                            v = float(f)
                            if 0 < v < 9999 and (upc_raw is None or f not in upc_raw):
                                price_candidates.append(f)
                        except ValueError:
                            pass
                    if not cost_raw and price_candidates:
                        cost_raw = price_candidates[0]
                    if not retail_raw and len(price_candidates) > 1:
                        retail_raw = price_candidates[1]

                # ── QTY: prefer "N Each" pattern, then fallback ──
                qty_m = re.search(r'\b(\d{1,5})\s+Each\b', context_text, re.IGNORECASE)
                if not qty_m:
                    qty_m = re.search(r'QTY\s*[:\s]+(\d+)', context_text, re.IGNORECASE)
                qty_raw = qty_m.group(1) if qty_m else None

                # ── Description: prefer "Product: <NAME>" label (before Resale/Price/Wholesale) ──
                desc_raw = None
                # Target Import has two "Product:" occurrences; the real one ends before Resale/Wholesale
                prod_m = re.search(
                    r'Product\s*:\s*([A-Z][A-Z0-9&\s\'"]{5,}?)(?=\s+(?:Resale|Unit Price|Wholesale|\Z))',
                    context_text, re.IGNORECASE
                )
                if prod_m:
                    desc_raw = prod_m.group(1).strip()
                else:
                    SKIP_KEYWORDS = {'ORDER', 'TARGET', 'VENDOR', 'BUYER', 'CATALOG',
                                     'UNIT PRICE', 'RESALE', 'TOTAL', 'SHIP', 'CANCEL',
                                     'FREIGHT', 'CONTACT', 'CURRENCY', 'INCOTERM',
                                     'TERMS', 'RELEASE', 'CONTRACT', 'PURCHASING'}
                    for cl in context_lines:
                        cl = cl.strip()
                        if len(cl) < 8:
                            continue
                        if not re.search(r'[A-Za-z]{3,}', cl):
                            continue
                        cl_up = cl.upper()
                        if any(kw in cl_up for kw in SKIP_KEYWORDS):
                            continue
                        if re.match(r'^[\d\s\.\-\#\$:]+$', cl):
                            continue
                        desc_raw = cl
                        break

                ocr_rows.append(_build_row(
                    ocr_po, dpci_raw, qty_raw,
                    cost_raw=cost_raw, retail_raw=retail_raw,
                    desc_raw=desc_raw, upc_raw=upc_raw,
                ))
                i += 1
                continue

            # ── Also handle plain DPCI format (DDD-CC-XXXX) in OCR text ──
            dpci_m = DPCI_RE.search(line)
            if dpci_m:
                dpci_raw = dpci_m.group(1)
                nums = re.findall(r'[\$]?([\d,]+\.?\d*)', line)
                nums_clean = []
                for n in nums:
                    n_plain = n.replace(',', '')
                    try:
                        val = float(n_plain)
                        if val > 0 and n_plain not in dpci_raw.replace('-', ''):
                            nums_clean.append(val)
                    except ValueError:
                        pass
                qty_raw = str(int(nums_clean[0])) if nums_clean and nums_clean[0] == int(nums_clean[0]) else (str(nums_clean[0]) if nums_clean else None)
                cost_raw = str(nums_clean[1]) if len(nums_clean) > 1 else None
                retail_raw = str(nums_clean[2]) if len(nums_clean) > 2 else None
                if qty_raw:
                    ocr_rows.append(_build_row(ocr_po, dpci_raw, qty_raw,
                                               cost_raw=cost_raw, retail_raw=retail_raw))
            i += 1

        if not ocr_rows and ocr_po:
            ocr_warnings.append("⚠️ OCR 已辨識 PO 編號但未找到商品明細，請確認 PDF 格式或提高掃描品質。")
        elif not ocr_rows:
            ocr_warnings.append("⚠️ OCR 無法辨識 PO 內容，PDF 可能掃描品質過低。")

        return ocr_rows, ocr_warnings

    try:
        with pdfplumber.open(pdf_file) as pdf:
            full_text_pages = []
            has_text_layer = False

            for page_num, page in enumerate(pdf.pages, 1):
                text = page.extract_text() or ''
                full_text_pages.append(text)
                if len(text.strip()) > 20:
                    has_text_layer = True

                # ── 更新 PO 編號 ──
                po_found = extract_po_number(text)
                if po_found:
                    current_po = po_found

                # ── 路線 1：extract_tables ──
                tables = page.extract_tables()
                for table in tables:
                    if not table:
                        continue
                    header_row_idx = None
                    for i, row in enumerate(table):
                        row_str = ' '.join(str(c) for c in row if c).upper()
                        if ('DPCI' in row_str or 'ITEM' in row_str) and \
                           ('COST' in row_str or 'QTY' in row_str or 'QUANTITY' in row_str):
                            header_row_idx = i
                            break

                    if header_row_idx is None:
                        continue

                    headers = [str(c).strip() if c else '' for c in table[header_row_idx]]

                    def find_col_idx(keywords):
                        for idx, h in enumerate(headers):
                            h_norm = re.sub(r'\s+', '', h).lower()
                            if any(kw.lower().replace(' ', '') in h_norm for kw in keywords):
                                return idx
                        return None

                    dpci_idx = find_col_idx(['dpci'])
                    qty_idx  = find_col_idx(['qty', 'quantity'])
                    cost_idx = find_col_idx(['cost'])
                    retail_idx = find_col_idx(['retail'])
                    desc_idx   = find_col_idx(['description', 'desc'])
                    upc_idx    = find_col_idx(['upc', 'barcode', 'bar code'])
                    asst_idx   = find_col_idx(['assortment'])
                    vcp_idx    = find_col_idx(['vcp', 'case'])

                    if dpci_idx is None or qty_idx is None:
                        continue

                    for row in table[header_row_idx + 1:]:
                        if not row or all(c is None or str(c).strip() == '' for c in row):
                            continue

                        def get(idx, r=row):
                            if idx is None or idx >= len(r):
                                return None
                            return str(r[idx]).strip() if r[idx] is not None else None

                        dpci_raw = get(dpci_idx)
                        qty_raw  = get(qty_idx)
                        if not dpci_raw or not DPCI_RE.search(dpci_raw):
                            continue
                        all_rows.append(_build_row(
                            current_po, dpci_raw, qty_raw,
                            cost_raw=get(cost_idx), retail_raw=get(retail_idx),
                            desc_raw=get(desc_idx), upc_raw=get(upc_idx),
                            asst_raw=get(asst_idx), vcp_raw=get(vcp_idx),
                        ))

            # ── 路線 2：純文字 fallback（SPS Commerce 文字版）──
            if not all_rows and has_text_layer:
                full_text = '\n'.join(full_text_pages)
                current_po = None

                for pat in PO_PATTERNS:
                    m = re.search(pat, full_text, re.IGNORECASE)
                    if m:
                        current_po = m.group(1).split('-')[0]
                        break

                lines = full_text.splitlines()
                for line in lines:
                    po_in_line = extract_po_number(line)
                    if po_in_line:
                        current_po = po_in_line
                        continue

                    dpci_m = DPCI_RE.search(line)
                    if not dpci_m:
                        continue

                    dpci_raw = dpci_m.group(1)
                    nums = re.findall(r'[\$]?([\d,]+\.?\d*)', line)
                    nums_clean = []
                    for n in nums:
                        n_plain = n.replace(',', '')
                        try:
                            val = float(n_plain)
                            if val > 0 and n_plain not in dpci_raw.replace('-', ''):
                                nums_clean.append(val)
                        except ValueError:
                            pass

                    qty_raw = str(int(nums_clean[0])) if len(nums_clean) > 0 and nums_clean[0] == int(nums_clean[0]) else (str(nums_clean[0]) if nums_clean else None)
                    cost_raw = str(nums_clean[1]) if len(nums_clean) > 1 else None
                    retail_raw = str(nums_clean[2]) if len(nums_clean) > 2 else None

                    if qty_raw is None:
                        continue

                    all_rows.append(_build_row(
                        current_po, dpci_raw, qty_raw,
                        cost_raw=cost_raw, retail_raw=retail_raw,
                    ))

    except Exception as e:
        return None, [f"❌ PDF 解析錯誤：{e}"]

    # ── 路線 3：OCR fallback（Target Import PO，全向量圖形，無文字層）──
    if not all_rows:
        pdf_file.seek(0)
        ocr_rows, ocr_warnings = _ocr_parse(pdf_file)
        warnings.extend(ocr_warnings)
        if ocr_rows:
            all_rows = ocr_rows
            warnings.append("__OCR_USED__")

    if not all_rows:
        return None, [
            "⚠️ 在 PDF 中找不到可解析的 PO 表格，已嘗試 OCR 辨識。"
            "請確認 PDF 來自 SPS Commerce 或 Target Import，或將原始檔案傳給 TG Team 檢視。"
        ]

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

    # ---- 混裝箱處理（改善版：支援三種比對策略）----
    if asst_df is not None and len(asst_df) > 0:
        # 建立查找用 set
        asst_box_dpcis = set(asst_df['Assortment_DPCI'].dropna())      # 混裝 box DPCI
        asst_comp_dpcis = set(asst_df['Component_DPCI'].dropna())       # 混裝內容品 DPCI

        if mode == 'standard':
            # ── 策略 1：以 Final_DPCI 比對 Component_DPCI（最常見：PO 以零件 DPCI 下單）──
            merge1 = pd.merge(
                merged_df,
                asst_df.rename(columns={'Assortment_DPCI': '_Asst_Box_DPCI'}),
                left_on='Final_DPCI',
                right_on='Component_DPCI',
                how='left',
                suffixes=('', '_asst')
            )
            # 若同一 Final_DPCI 在混裝表中對應多個 box，取 Asst_Box_Cost 最小（保守）
            # dedup: keep first hit per Final_DPCI (sort by cost ascending, NaN last)
            merge1 = merge1.sort_values('Asst_Box_Cost', ascending=True, na_position='last')
            # Use original columns as dedup key to avoid duplicates from merge
            orig_cols = list(po_df.columns) + [c for c in merged_df.columns if c not in po_df.columns and c != 'Asst_Box_Cost' and c != 'Units_in_Assortment']
            key_cols = [c for c in ['PO NUMBER', 'Final_DPCI', 'Original_DPCI'] if c in merge1.columns]
            merge1 = merge1.drop_duplicates(subset=key_cols, keep='first')
            # 保留 box DPCI 追溯欄（Gap 3 修正：供 Asst_Role 診斷使用）
            merge1['Matched_Box_DPCI'] = merge1.get('_Asst_Box_DPCI', pd.Series(np.nan, index=merge1.index))
            # Clean up extra merge cols
            merge1.drop(columns=[c for c in ['_Asst_Box_DPCI', 'Component_DPCI'] if c in merge1.columns], inplace=True, errors='ignore')
            merged_df = merge1

            # 以混裝表資訊標記是否為混裝品：記錄哪些 Final_DPCI 找到了零件對應
            merged_df['_is_comp_flag'] = merged_df['Asst_Box_Cost'].notna()

            # ── 策略 2：以 Original_DPCI 比對 Assortment_DPCI（PO 中的 box 層級行）──
            # 若 Original_DPCI 本身是 box DPCI，需拉出 box cost
            if 'Original_DPCI' in merged_df.columns:
                box_cost_lookup = asst_df.drop_duplicates(subset=['Assortment_DPCI'])[['Assortment_DPCI', 'Asst_Box_Cost']].rename(
                    columns={'Assortment_DPCI': '_box_dpci_key', 'Asst_Box_Cost': '_box_cost_direct'}
                )
                merged_df = pd.merge(merged_df, box_cost_lookup, left_on='Original_DPCI', right_on='_box_dpci_key', how='left')
                is_comp = merged_df['_is_comp_flag']
                is_box = merged_df['_box_cost_direct'].notna() & ~is_comp
                # 填補 Asst_Box_Cost for box rows
                merged_df.loc[is_box, 'Asst_Box_Cost'] = merged_df.loc[is_box, '_box_cost_direct']
                merged_df.drop(columns=['_box_dpci_key', '_box_cost_direct', '_is_comp_flag'], inplace=True, errors='ignore')
            else:
                is_comp = merged_df['_is_comp_flag']
                merged_df.drop(columns=['_is_comp_flag'], inplace=True, errors='ignore')
                is_box = pd.Series(False, index=merged_df.index)

            # ── 更新 ASSORTMENT ITEM? 欄位 ──
            merged_df['ASSORTMENT ITEM?'] = np.where(
                is_comp | is_box, 'Y',
                merged_df.get('ASSORTMENT ITEM?', pd.Series('N', index=merged_df.index))
            )

            # ── 新增診斷欄：Asst_Role（Gap 3：含 Matched_Box_DPCI 追溯）──
            # 建立 Final_DPCI → 所屬 Box DPCI 的查找表（供「零件無對應」診斷用）
            comp_to_box_map = (
                asst_df[asst_df['Component_DPCI'].notna()]
                .groupby('Component_DPCI')['Assortment_DPCI']
                .first()
                .to_dict()
            )
            def _asst_role(row):
                fdpci = row.get('Final_DPCI', '')
                if is_comp[row.name]:
                    box_dpci = row.get('Matched_Box_DPCI', '')
                    box_str = f' ({box_dpci})' if pd.notna(box_dpci) and box_dpci else ''
                    return f'🔹 混裝零件{box_str}'
                if is_box[row.name]:
                    return '📦 混裝 Box'
                if fdpci in asst_comp_dpcis:
                    box_ref = comp_to_box_map.get(fdpci, '')
                    box_str = f' → 應屬 {box_ref}' if box_ref else ''
                    return f'⚠️ 零件無對應{box_str}'
                if fdpci in asst_box_dpcis:
                    return '⚠️ Box無對應'
                return '—'

            merged_df['Asst_Role'] = merged_df.apply(_asst_role, axis=1)

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
            merged_df['Asst_Role'] = '—'
    else:
        merged_df['Target_Cost'] = merged_df['Final_Product_Cost']
        if 'Units_in_Assortment' not in merged_df.columns:
            merged_df['Units_in_Assortment'] = np.nan
        if 'Asst_Box_Cost' not in merged_df.columns:
            merged_df['Asst_Box_Cost'] = np.nan
        merged_df['Asst_Role'] = '—'

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

    # Case QTY 比對：僅當 VCP 資料確實存在時才比對
    # 若整欄皆為 NaN（SPS 標準 PO 通常無 VCP 欄），略過比對改標 True
    vcp_col_present = merged_df['PO VCP / Assort QTY'].notna().any()
    if vcp_col_present:
        merged_df['Case QTY Match'] = np.where(
            merged_df['PO VCP / Assort QTY'].isna(),
            True,  # 此列無 VCP 資料 → 略過
            np.where(
                merged_df['Target Case / Assort QTY'].isna(),
                False,
                np.isclose(
                    merged_df['PO VCP / Assort QTY'].fillna(-1),
                    merged_df['Target Case / Assort QTY'].fillna(-1),
                    atol=0.01
                )
            )
        )
    else:
        # VCP 欄完全缺失（SPS 標準 PO）→ 不執行此項核對
        merged_df['Case QTY Match'] = True
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
            # Normalize for comparison: strip legal suffixes, punctuation, case
            # (handles English PCN name vs Chinese/mixed Dispatch name)
            pcn_factory_norm = merged_df['Factory Name'].astype(str).apply(normalize_factory_name)
            disp_factory_norm = merged_df['Dispatch_Factory'].astype(str).apply(normalize_factory_name)
            # Primary: exact match after normalization
            exact_match = pcn_factory_norm == disp_factory_norm
            # Fallback: one name is a substring of the other (handles truncated names)
            subset_match = (
                pcn_factory_norm.str.len().gt(3) & disp_factory_norm.str.len().gt(3) &
                (pcn_factory_norm.apply(lambda x: any(x in d or d in x
                    for d in [disp_factory_norm.iloc[i] if i < len(disp_factory_norm) else '' for i in [pcn_factory_norm.tolist().index(x) if x in pcn_factory_norm.tolist() else 0]])))
            )
            # Simpler vectorised substring check
            factory_match_vec = []
            for pcn_n, disp_n in zip(pcn_factory_norm, disp_factory_norm):
                if pcn_n == disp_n:
                    factory_match_vec.append(True)
                elif len(pcn_n) > 3 and len(disp_n) > 3 and (pcn_n in disp_n or disp_n in pcn_n):
                    factory_match_vec.append(True)
                else:
                    factory_match_vec.append(False)
            merged_df['Factory_Match'] = np.where(
                both_have_factory,
                factory_match_vec,
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
        'PO NUMBER', 'ASSORTMENT ITEM?', 'Asst_Role', 'Matched_Box_DPCI', 'Is_Shipper_Display', 'Original_DPCI', 'Final_DPCI',
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

# LuckyStar Logo（幾何藍C形+黃三角，Base64 PNG）
LOGO_B64 = "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"
LOGO_HTML = f'''<img src="data:image/png;base64,{LOGO_B64}" style="width:52px;height:52px;object-fit:contain;" />'''

col_logo, col_title = st.columns([0.07, 0.93])
with col_logo:
    st.markdown(LOGO_HTML, unsafe_allow_html=True)
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

            ocr_pdf_names = []
            with st.spinner("PDF 解析中..."):
                for pdf_file in pdf_files:
                    po_df_parsed, parse_warnings = parse_sps_pdf(pdf_file)
                    for w in parse_warnings:
                        if w == "__OCR_USED__":
                            ocr_pdf_names.append(pdf_file.name)
                        else:
                            all_parse_warnings.append(w)
                    if po_df_parsed is not None and len(po_df_parsed) > 0:
                        all_parsed_dfs.append(po_df_parsed)

            # 彙整 OCR 警告（合併為一條訊息）
            if ocr_pdf_names:
                if len(ocr_pdf_names) == 1:
                    st.info(f"ℹ️ 以下 PDF 為向量圖形格式（Target Import PO），已透過 OCR 辨識解析：{ocr_pdf_names[0]}")
                else:
                    with st.expander(f"ℹ️ {len(ocr_pdf_names)} 份 PDF 透過 OCR 解析（點擊展開檔名）", expanded=False):
                        for n in ocr_pdf_names:
                            st.write(f"• {n}")

            # 顯示其他解析警告（非 OCR、非重複 PO）
            for w in all_parse_warnings:
                st.warning(w)

            if not all_parsed_dfs:
                st.error("❌ 所有 PDF 均無法解析出訂單資料，請確認格式後再試。")
            else:
                combined_po_df = pd.concat(all_parsed_dfs, ignore_index=True)
                st.success(f"✅ 成功從 {len(pdf_files)} 份 PDF 解析出 {len(combined_po_df)} 筆訂單列。")

                clean_po_df, dup_warnings = detect_duplicate_pos(combined_po_df)
                all_warnings = all_parse_warnings + dup_warnings
                # 彙整重複 PO 警告（折疊顯示，避免過多警告訊息）
                if dup_warnings:
                    if len(dup_warnings) == 1:
                        st.warning(dup_warnings[0])
                    else:
                        with st.expander(f"⚠️ {len(dup_warnings)} 則重複 PO 偵測警告（點擊展開）", expanded=False):
                            for w in dup_warnings:
                                st.warning(w)

                prod_df = process_products(product_files)
                asst_df = process_assortments(asst_files) if asst_files else None

                # ── G5 PO 內部一致性自我驗證 ──
                self_verify_warnings = po_self_verify(clean_po_df, mode='standard')
                if self_verify_warnings:
                    with st.expander(f"⚠️ G5 PO 內部驗算：{len(self_verify_warnings)} 則警告（點擊展開）", expanded=True):
                        for w in self_verify_warnings:
                            st.warning(w)
                all_warnings = all_warnings + self_verify_warnings

                dispatch_arg = dispatch_df_global if len(dispatch_df_global) > 0 else None
                merged_df = run_validation(clean_po_df, prod_df, asst_df, mode='standard', dispatch_df=dispatch_arg)

                run_meta = {
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'input_files': f"PDFs: {', '.join(f.name for f in pdf_files)} | Products: {', '.join(f.name for f in product_files)}"
                }
                show_results(merged_df, 'PDF', run_meta=run_meta, validation_notes=all_warnings)

