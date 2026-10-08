import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import os
import json
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
    # 用 Python 的 re 逐筆處理：雲端的 pandas 字串引擎（Arrow/RE2）的 \s 不含不換行空白，
    # 也不接受 \u 跳脫，會讓「240-04-0531\xa0」這類儲存格比對不到
    def _one(v):
        v = re.sub(r'[\s\u00a0\u200b\u3000\ufeff]+', '', str(v))
        v = re.sub(r'[/\\]', '-', v)
        return re.sub(r'\.0$', '', v)
    return series.map(_one)

def clean_upc(series):
    """清理 UPC/Barcode 字串，避免因 Excel 浮點數轉換產生 .0 導致比對失敗"""
    if series is None:
        return series
    def _one(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return np.nan
        v = re.sub(r'[\s\u00a0]+', '', re.sub(r'\.0$', '', str(v)))
        return np.nan if v in ('', 'nan', 'None', '<NA>') else v
    return series.map(_one)

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
            # 成本欄必須同時含 box/asst 與 cost；舊寫法會誤選到「Vendor Asst Style #」
            cost_col = fuzzy_col(df.columns, 'box', 'cost', require_all=True) or \
                       fuzzy_col(df.columns, 'asst', 'cost', require_all=True) or \
                       fuzzy_col(df.columns, 'assortment', 'cost', require_all=True)
            units_col = fuzzy_col(df.columns, 'units', 'assortment', require_all=True) or \
                        fuzzy_col(df.columns, 'unitsinassortment')

            if not all([master_col, sub_col, cost_col, units_col]):
                continue

            temp_df = df[[master_col, sub_col, cost_col, units_col]].copy()
            temp_df.columns = ['Assortment_DPCI', 'Component_DPCI', 'Asst_Box_Cost', 'Units_in_Assortment']
            blank = r'^[\s\u00a0]*$'
            temp_df['Assortment_DPCI'] = temp_df['Assortment_DPCI'].replace(blank, np.nan, regex=True)
            temp_df['Asst_Box_Cost'] = temp_df['Asst_Box_Cost'].replace(blank, np.nan, regex=True)
            # 每個 Box 以「該列有 Box 成本或 Box DPCI」為起點；只在同一個 Box 內往下填，
            # 避免尚未配發 DPCI 的 Box 繼承上一個 Box 的 DPCI
            grp = (temp_df['Asst_Box_Cost'].notna() | temp_df['Assortment_DPCI'].notna()).cumsum()
            temp_df['Assortment_DPCI'] = temp_df.groupby(grp)['Assortment_DPCI'].ffill()
            temp_df['Asst_Box_Cost'] = temp_df.groupby(grp)['Asst_Box_Cost'].ffill()
            temp_df = temp_df.dropna(subset=['Assortment_DPCI', 'Component_DPCI'])
            temp_df['Assortment_DPCI'] = clean_dpci(temp_df['Assortment_DPCI'])
            temp_df['Component_DPCI'] = clean_dpci(temp_df['Component_DPCI'])
            temp_df = temp_df[temp_df['Assortment_DPCI'].str.match(r'^\d{3}-\d{2}-\d{4}$', na=False)
                              & temp_df['Component_DPCI'].str.match(r'^\d{3}-\d{2}-\d{4}$', na=False)]
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
# 5. G1: PDF 解析（與 po-grid Skill 的 parse_po_pdfs.py 同一套規則）
# ==========================================
# 文字層：有 → pdfplumber；無（Microsoft Print To PDF）→ 300dpi 灰階 OCR
# 版型：V1（DPCI 在行首）、V2（27C1：數量緊接 Unit Price）、V3（27C2：單價在數量前）
# 自我驗證三道 gate：Σ行小計 = PO Total、數量×單價 = 行小計、Σ數量 = Total Qty

_PO_LINE_START = re.compile(r'^(\d{1,3}) (\d{9}) Vendors Style (\d{11,13})', re.M)
_PO_PREPACK = re.compile(
    r'^(\d{1,3}) (\d{9}) (\S+) (\d{11,13})[^\d\n]*?(\d+\.\d+) ([\d,]+)\s*Each\s*$', re.M)
_PO_LINE_START_V2 = re.compile(
    r'^(\d{1,2}) Buyers Catalog Number:\s*(?:Vendors Style\s*)?'
    r'(\d{11,13})Product:\s*(\w+)\s*Unit\s*Price:\s*([\d,]+)\s*Each\s*([\d,]+\.\d{2})', re.M)
_PO_DPCI_V2 = re.compile(r'^(\d{9})\s*(?:Number:\s*)?Product:\s*(.*?)\s*([\d,]+\.\d{2,4})\s*$', re.M)
_PO_STYLE_V2 = re.compile(r'Buyers Item Number:\s*(.*?)\s*Resale:\s*([\d.]+)', re.S)
_PO_PREPACK_V2 = re.compile(r'^(\d{1,3}) (\d{9}) (\d{11,13}) ([\d.]+) ([\d,]+)\s*Each\s*$', re.M)
_PO_LINE_START_V3 = re.compile(
    r'^(\d{1,3}) Buyers Catalog Number:\s*(?:Vendors Style\s*)?'
    r'(\d{11,13})Product:\s*(\w+)\s*Unit\s*Price:\s*([\d.]+)\s+([\d,]+)\s*Each\s*'
    r'([\d,]+\.\d{2})', re.M)
_PO_DPCI_V3 = re.compile(r'^(\d{9})\b', re.M)
_PO_STYLE_V3 = re.compile(r'Buyers Item Number:\s*(\S+)')
_PO_PREPACK_V3 = re.compile(
    r'^(\d{1,3}) (\d{9}) (\d{11,13})\s*(.*?)\s*(\d+\.\d+) ([\d,]+)\s*Each\s*$', re.M)
_PO_SHIP_WIN = r'Shipping Window:(.{0,400}?)(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})'


def _po_num(s):
    try:
        return float(str(s).replace(',', '').replace(' ', '')) if s not in (None, '') else None
    except ValueError:
        return None


def _po_g(pat, text, grp=1, flags=0):
    m = re.search(pat, text, flags)
    return m.group(grp).strip() if m else ''


def _sku_to_dpci(sku):
    s = str(sku).zfill(9)
    return f"{s[:3]}-{s[3:5]}-{s[5:]}"


def normalize_ocr_text(t):
    """把 OCR 文字整理成與原生文字層相同的形狀，讓同一組 regex 可用。"""
    out = []
    for ln in t.splitlines():
        ln = ln.replace(' ', ' ').rstrip()
        ln = re.sub(r'[ \t]{2,}', ' ', ln).strip()
        out.append(ln)
    t = '\n'.join(out)
    t = re.sub(r'(\d{11,13})\s+Product:', r'\1Product:', t)
    t = re.sub(r'(Order #:\s*\d{10,11})\s+-\s*(\d{4})', r'\1-\2', t)
    t = re.sub(r'Unit Price:\s*', 'Unit Price: ', t)
    t = re.sub(r'Resale:\s*', 'Resale: ', t)
    t = re.sub(r'(\d+\.\d+)\.(?=\s)', r'\1', t)      # 表格線被讀成小數點
    t = re.sub(r'(\d) ,(\d{3})', r'\1,\2', t)         # "468 ,002.66"
    return t


@st.cache_data(show_spinner=False)
def extract_po_text(pdf_bytes):
    """回傳 (text, used_ocr, error)。以檔案內容快取，重跑不需再 OCR。"""
    try:
        import pdfplumber
    except ImportError:
        return '', False, "缺少 pdfplumber 套件，請在 requirements.txt 加入 pdfplumber。"
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            n_pages = len(pdf.pages)
            text = '\f'.join((pg.extract_text() or '') for pg in pdf.pages)
    except Exception as e:
        return '', False, f"PDF 讀取錯誤：{e}"
    if 'Order #' in text and len(text.strip()) > 200:
        return text, False, None

    # 無文字層 → OCR。300dpi：150dpi 曾把 1.675 讀成 1.875。
    try:
        from pdf2image import convert_from_bytes
        import pytesseract
    except ImportError as e:
        return '', True, f"OCR 套件缺失：{e}"
    pages = []
    try:
        for p in range(1, n_pages + 1):
            img = convert_from_bytes(pdf_bytes, dpi=300, grayscale=True,
                                     first_page=p, last_page=p, fmt='png')[0]
            pages.append(pytesseract.image_to_string(
                img, lang='eng', config='--psm 6 -c preserve_interword_spaces=1'))
            del img
    except Exception as e:
        return '', True, f"OCR 失敗：{e}"
    return '\f'.join(normalize_ocr_text(pg) for pg in pages), True, None


def _po_resale(blk):
    """Resale 後第一個金額形狀的數字（欄位換行時 10 位 Item Number 會插在中間）。"""
    m = re.search(r'Resale:', blk)
    if not m:
        return ''
    mm = re.search(r'(\d{1,4}\.\d{2})\b', blk[m.end():m.end() + 200])
    return mm.group(1) if mm else ''


def parse_po_text(fname, txt):
    """解析單一 PO 文字 → (head, items)。items.kind = line / ast / prepack。"""
    stamp = _po_g(r'(\d{4}/\d{1,2}/\d{1,2} \d{1,2}:\d{2}) Fulfillment', txt)
    txt = re.sub(r'^\d{4}/\d{1,2}/\d{1,2} \d{1,2}:\d{2} Fulfillment\s*$', '', txt, flags=re.M)
    txt = re.sub(r'^https://\S+\s+\d+/\d+\s*$', '', txt, flags=re.M)

    m = re.search(r'Order #:\s*(\d{4})-(\d{5,8})-(\w{4})', txt)
    if m:
        po, dc = m.group(2), m.group(3)
    else:
        m2 = re.search(r'Order #:\s*(\d{6,11})-(\w{3,5})', txt)
        po, dc = (m2.group(1), m2.group(2)) if m2 else ('', '')

    head = {
        'po': po, 'dc': dc, 'file': fname, 'retrieved': stamp,
        'fname_date': _po_g(r'_(\d{4})(?:\D[^_]*)?\.pdf$', fname),
        'is_change': bool(re.search(r'Import PO Change|PO Change Date', txt)),
        'is_cancel': bool(re.search(r'CANCEL ORDER', txt, re.I)),
        'doc_date': _po_g(r'(?:PO Change Date|PO Date):.*?\n\s*(\d{2}/\d{2}/\d{4})', txt, 1, re.S),
        'po_total': _po_g(r'Purchase Order Total:\s*([\d, ]+\.\d{2})', txt),
        'total_qty': _po_g(r'Total Qty:\s*([\d,]+)', txt),
        'n_orders': len(set(re.findall(r'Order #:\s*(\d{6,11})', txt))),
    }

    items = []

    def add_prepacks(blk, sku, lineno, patterns):
        seen = set()
        for kind, pat in patterns:
            for pm in pat.finditer(blk):
                gr = pm.groups()
                if kind == 'v3':
                    _, csku, cupc, _d, cunit, cqty = gr
                elif kind == 'v1':
                    _, csku, _s, cupc, cunit, cqty = gr
                else:
                    _, csku, cupc, cunit, cqty = gr
                if csku in seen:
                    continue
                seen.add(csku)
                items.append({'kind': 'prepack', 'parent_sku': sku, 'parent_line': int(lineno),
                              'sku': csku, 'upc': cupc, 'unit': cunit,
                              'qty': int(cqty.replace(',', ''))})

    starts = [(mm.start(), mm) for mm in _PO_LINE_START.finditer(txt)]
    v3_hits = [(mm.start(), 'v3', mm) for mm in _PO_LINE_START_V3.finditer(txt)]
    v2_hits = [(mm.start(), 'v2', mm) for mm in _PO_LINE_START_V2.finditer(txt)]

    if not starts and (v3_hits or v2_hits):
        # 27C2：同一張 PO 可同時出現 V2 與 V3 形狀 → 取聯集
        hits = sorted(v3_hits + v2_hits, key=lambda x: x[0])
        for idx, (pos_, variant, mm) in enumerate(hits):
            end = hits[idx + 1][0] if idx + 1 < len(hits) else len(txt)
            blk = txt[pos_:end]
            if variant == 'v3':
                lineno, upc, ptype, unit, qty_s, total = mm.groups()
            else:
                lineno, upc, ptype, qty_s, total = mm.groups()
                dm2 = _PO_DPCI_V2.search(blk)
                unit = dm2.group(3) if dm2 else ''
            dm = _PO_DPCI_V3.search(blk)
            if not dm:
                continue
            sku = dm.group(1)
            sm = _PO_STYLE_V3.search(blk)
            style_tok = sm.group(1).strip() if sm else ''
            items.append({'kind': 'ast' if ptype.upper() == 'AST' else 'line',
                          'line': int(lineno), 'sku': sku, 'upc': upc,
                          'style': style_tok if re.match(r'^\d{0,2}[A-Za-z]', style_tok) else '',
                          'qty': int(qty_s.replace(',', '')), 'unit': unit,
                          'total': total, 'resale': _po_resale(blk)})
            add_prepacks(blk, sku, lineno,
                         [('v3', _PO_PREPACK_V3), ('v1', _PO_PREPACK), ('v2', _PO_PREPACK_V2)])
    elif starts:
        for idx, (pos_, mm) in enumerate(starts):
            end = starts[idx + 1][0] if idx + 1 < len(starts) else len(txt)
            blk = txt[pos_:end]
            lineno, sku, upc = mm.groups()
            qm = re.search(r'Unit Price:\s*([\d,]+)\s*Each\s*([\d,]+\.\d{2})', blk)
            if not qm:
                continue
            ptype = 'AST' if re.search(r'Product:\s*AST\b|ASSORTMENT', blk) else 'REG'
            unit = _po_g(r'^Number:[^\n]*?([\d]+\.[\d]+)\s*$', blk, 1, re.M) or \
                   _po_g(r'^[^\n]*?([\d]+\.[\d]{2,4})\s*$', blk, 1, re.M)
            resale = _po_g(r'Resale:\s*([\d.]+)', blk) or \
                     _po_g(r'Resale:\s*\n[^\n]*?([\d]+\.[\d]{2})\s*$', blk, 1, re.M)
            items.append({'kind': 'ast' if ptype == 'AST' else 'line',
                          'line': int(lineno), 'sku': sku, 'upc': upc,
                          'style': _po_g(r'^(\d{2}[A-Z]{2,4}\d{0,3})\b', blk, 1, re.M),
                          'qty': int(qm.group(1).replace(',', '')), 'unit': unit,
                          'total': qm.group(2), 'resale': resale})
            add_prepacks(blk, sku, lineno, [('v1', _PO_PREPACK)])

    for it in items:
        it.update({'po': po, 'dc': dc, 'file': fname})
    return head, items


def _po_gate_check(head, items):
    """三道自我驗證 gate，回傳問題字串 list（空 = 通過）。"""
    probs = []
    lines = [i for i in items if i['kind'] in ('line', 'ast')]
    if not lines:
        return ["找不到任何品項列（PDF 版型可能已變更）"]
    lt = sum(_po_num(i['total']) or 0 for i in lines)
    pt = _po_num(head['po_total'])
    if pt is None:
        probs.append("讀不到 Purchase Order Total，無法驗算總金額")
    elif abs(lt - pt) > 0.05:
        probs.append(f"行小計合計 {lt:,.2f} ≠ PO Total {pt:,.2f}")
    for i in lines:
        u, t = _po_num(i.get('unit')), _po_num(i.get('total'))
        if u is None:
            probs.append(f"{_sku_to_dpci(i['sku'])} 讀不到單價")
        elif abs(u * i['qty'] - t) > 0.05:
            probs.append(f"{_sku_to_dpci(i['sku'])}：{u} × {i['qty']:,} ≠ {t:,.2f}")
    tq = _po_num(head['total_qty'])
    sq = sum(i['qty'] for i in lines)
    if tq is None:
        probs.append("讀不到 Total Qty，無法驗算總數量")
    elif abs(sq - tq) > 0.5:
        probs.append(f"行數量合計 {sq:,} ≠ Total Qty {tq:,.0f}")
    return probs


def _po_version_key(h):
    """同一 PO# 多份文件時決定何者為現行版：PO Change 優先於原單，其次文件日期、擷取時間、檔名日期。"""
    def dkey(s):
        m = re.match(r'(\d{2})/(\d{2})/(\d{4})', s or '')
        return (int(m.group(3)), int(m.group(1)), int(m.group(2))) if m else (0, 0, 0)
    m = re.match(r'(\d{4})/(\d{1,2})/(\d{1,2})\s+(\d{1,2}):(\d{2})', h.get('retrieved') or '')
    stamp = tuple(int(x) for x in m.groups()) if m else (0, 0, 0, 0, 0)
    fd = h.get('fname_date') or ''
    fdk = (int(fd[:2]), int(fd[2:])) if len(fd) == 4 and fd.isdigit() else (0, 0)
    return (dkey(h.get('doc_date')), 1 if h.get('is_change') else 0, stamp, fdk)


def split_combined_print(text):
    """
    SPS「合併列印」：一份 PDF 內含多張訂單。依每頁的 Order # 分組（無 Order # 的頁歸前一張），
    回傳 [(po_dc, 該訂單文字)]；只有一張訂單時回傳單一元素。
    """
    pages = text.split('\f')
    groups, cur = [], None
    for pg in pages:
        m = re.search(r'Order #:\s*(?:\d{4}-)?(\d{5,11})-(\w{3,5})', pg)
        key = (m.group(1), m.group(2)) if m else None
        if key and key != cur:
            groups.append([key, [pg]])
            cur = key
        elif groups:
            groups[-1][1].append(pg)
        else:
            groups.append([None, [pg]])
    return [(k, '\n'.join(pgs)) for k, pgs in groups]


def parse_po_pdfs(pdf_files, progress=None, known_dpcis=None):
    """
    解析多份 PO PDF → (po_df, info)
    info: ocr_files / warnings / gate_problems / superseded / cancelled / duplicates / n_files / n_pos
    po_df 每列 Row_Type = line（一般品項）/ box（混裝 Box）/ component（Box 內零件）
    """
    info = {'ocr_files': [], 'warnings': [], 'gate_problems': [], 'superseded': [],
            'cancelled': [], 'duplicates': [], 'n_files': len(pdf_files), 'n_pos': 0,
            'combined': [], 'skipped_other_program': [], 'live_docs': []}
    parsed = []
    for n, f in enumerate(pdf_files, 1):
        name = getattr(f, 'name', str(f))
        if progress:
            progress(n, len(pdf_files), name)
        try:
            f.seek(0)
        except Exception:
            pass
        text, used_ocr, err = extract_po_text(f.read())
        if used_ocr:
            info['ocr_files'].append(name)
        if err:
            info['warnings'].append(f"❌ {name}：{err}")
            continue
        orders = split_combined_print(text)
        is_combined = len(orders) > 1
        kept = 0
        for key, otext in orders:
            label = f"{name} ▸ {key[0]}" if is_combined and key else name
            head, items = parse_po_text(label, otext)
            if not head['po']:
                if not is_combined:
                    info['warnings'].append(f"⚠️ {name}：找不到 Order #，已略過（PDF 版型可能不同）。")
                continue
            # 合併列印常混入其他部門／Program 的訂單：沒有任何品項在主檔內的訂單直接略過
            if is_combined and known_dpcis is not None:
                skus = {_sku_to_dpci(i['sku']) for i in items}
                if not (skus & known_dpcis):
                    info['skipped_other_program'].append(
                        {'po': head['po'], 'file': name, 'dpcis': sorted(skus)[:3]})
                    continue
            head['_text'] = otext
            parsed.append((head, items))
            kept += 1
        if is_combined:
            info['combined'].append({'file': name, 'orders': len(orders), 'kept': kept})

    # ── 版本判定：同 PO#+DC 取最新；CANCEL ORDER 的 PO Change → 整張作廢 ──
    groups = {}
    for head, items in parsed:
        groups.setdefault((head['po'], head['dc']), []).append((head, items))
    live = []
    for (po, dc), docs in groups.items():
        docs.sort(key=lambda d: _po_version_key(d[0]), reverse=True)
        cancel = next((h for h, _ in docs if h['is_cancel']), None)
        if cancel:
            info['cancelled'].append({'po': po, 'dc': dc, 'file': cancel['file'],
                                      'dropped': [h['file'] for h, _ in docs if h is not cancel]})
            continue
        head, items = docs[0]
        if head['is_change'] and not items:
            info['warnings'].append(f"⚠️ {head['file']}：為 PO Change 文件但讀不到變更後的品項，請人工確認 PO {po}。")
            base = next(((h, it) for h, it in docs[1:] if it), None)
            if not base:
                continue
            head, items = base
        def _sig(its):
            return sorted((i['kind'], i['sku'], i['qty'], str(i.get('unit'))) for i in its)
        for lose, lose_items in docs:
            if lose is not head and _sig(lose_items) != _sig(items):   # 內容相同的重複檔不提示
                info['superseded'].append({'po': po, 'kept': head['file'], 'dropped': lose['file'],
                                           'from_qty': lose.get('total_qty'), 'to_qty': head.get('total_qty')})
        probs = _po_gate_check(head, items)
        if probs:
            info['gate_problems'].append({'po': po, 'file': head['file'], 'issues': probs})
        live.append((head, items))
        info['live_docs'].append({'po': head['po'], 'dc': head['dc'], 'file': head['file'],
                                  'fname_date': head.get('fname_date') or
                                  ''.join((head.get('doc_date') or '')[:5].split('/')),
                                  'text': head.get('_text', '')})

    # ── 不同 PO# 但 (DPCI, 數量) 完全相同 → 疑似重複下單 ──
    fp = {}
    for head, items in live:
        key = tuple(sorted((i['sku'], i['qty']) for i in items if i['kind'] in ('line', 'prepack')))
        if key:
            fp.setdefault(key, set()).add(head['po'])
    info['duplicates'] = [sorted(v) for v in fp.values() if len(v) > 1]

    rows = []
    for head, items in live:
        box_qty = {(i['sku'], i['line']): i['qty'] for i in items if i['kind'] == 'ast'}
        for i in items:
            dpci = _sku_to_dpci(i['sku'])
            base = {'PO NUMBER': head['po'], 'DC': head['dc'], 'PO_FILE': head['file'],
                    'PO UPC': i.get('upc') or np.nan, 'VCP QUANTITY': np.nan,
                    'Is_Shipper_Display': False, 'ITEM DESCRIPTION': i.get('style', '')}
            if i['kind'] == 'prepack':
                parent = _sku_to_dpci(i['parent_sku'])
                bq = box_qty.get((i['parent_sku'], i['parent_line']))
                base.update({'Row_Type': 'component', 'ASSORTMENT ITEM?': 'Y',
                             'Original_DPCI': parent, 'Final_DPCI': dpci,
                             'Final_QTY': float(i['qty']), 'Box_QTY': bq,
                             'ITEM UNIT COST': _po_num(i.get('unit')), 'ITEM UNIT RETAIL': np.nan,
                             'LINE TOTAL': np.nan,
                             'COMPONENT ASSORT QTY': (i['qty'] / bq) if bq else np.nan})
            else:
                is_box = i['kind'] == 'ast'
                base.update({'Row_Type': 'box' if is_box else 'line',
                             'ASSORTMENT ITEM?': 'Y' if is_box else 'N',
                             'Original_DPCI': dpci, 'Final_DPCI': dpci,
                             'Final_QTY': float(i['qty']), 'Box_QTY': float(i['qty']) if is_box else np.nan,
                             'ITEM UNIT COST': _po_num(i.get('unit')),
                             'ITEM UNIT RETAIL': _po_num(i.get('resale')),
                             'LINE TOTAL': _po_num(i.get('total')),
                             'COMPONENT ASSORT QTY': np.nan})
            rows.append(base)
    info['n_pos'] = len(live)
    cols = ['PO NUMBER', 'DC', 'PO_FILE', 'Row_Type', 'ASSORTMENT ITEM?', 'Original_DPCI', 'Final_DPCI',
            'ITEM DESCRIPTION', 'Final_QTY', 'Box_QTY', 'ITEM UNIT COST', 'ITEM UNIT RETAIL', 'LINE TOTAL',
            'PO UPC', 'VCP QUANTITY', 'COMPONENT ASSORT QTY', 'Is_Shipper_Display']
    return pd.DataFrame(rows, columns=cols), info

# ==========================================
# 5b. PO GRID（呼叫與 po-grid Skill 相同的腳本：grid_engine/ 資料夾）
# ==========================================
def _grid_engine_dir():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'grid_engine')


def extract_spk_images(xlsx_bytes, out_dir):
    """
    從 SPK Workspace 活頁簿抽出產品圖，依 DPCI 存成 <DPCI>.png。
    多個工作表都有圖時只取 Products（Claims 是瑕疵照，不是產品照）。
    回傳 (已存張數, 訊息 list)
    """
    import openpyxl
    from PIL import Image as _PILImage
    notes, saved = [], 0
    try:
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
    except Exception as e:
        return 0, [f"SPK 檔讀取失敗：{e}"]
    with_imgs = [ws for ws in wb.worksheets if getattr(ws, '_images', None)]
    if not with_imgs:
        return 0, ["SPK 檔內找不到浮動圖片（若圖片是「儲存格內圖片」格式，請改上傳圖片 zip）。"]
    if len(with_imgs) > 1:
        prod = [ws for ws in with_imgs if ws.title.strip().lower() == 'products']
        if prod:
            notes.append(f"SPK 有 {len(with_imgs)} 個工作表含圖片，只取 Products。")
            with_imgs = prod
    for ws in with_imgs:
        dpci_col = None
        for row in ws.iter_rows(min_row=1, max_row=15):
            for c in row:
                if isinstance(c.value, str) and c.value.strip().upper() == 'DPCI':
                    dpci_col, hdr_row = c.column, c.row
                    break
            if dpci_col:
                break
        if not dpci_col:
            notes.append(f"工作表 {ws.title}：找不到 DPCI 欄，已略過。")
            continue
        missing = 0
        for img in ws._images:
            try:
                r = img.anchor._from.row + 1
                v = ws.cell(r, dpci_col).value
                d = re.sub(r'[\s ]', '', str(v or '')).replace('/', '-')
                if not re.match(r'^\d{3}-\d{2}-\d{4}$', d):
                    missing += 1
                    continue
                im = _PILImage.open(io.BytesIO(img._data()))
                im.save(os.path.join(out_dir, d + '.png'))
                saved += 1
            except Exception:
                missing += 1
        if missing:
            notes.append(f"工作表 {ws.title}：{missing} 張圖片所在列沒有有效 DPCI，已略過。")
    return saved, notes


def collect_grid_images(image_files, out_dir):
    """圖片 zip（可巢狀）、單張圖片、SPK .xlsx → 全部攤平成 out_dir/<DPCI>.png|jpg。回傳 (張數, 訊息)。"""
    import zipfile
    notes, count = [], 0

    def take_zip(data, depth=0):
        nonlocal count
        try:
            z = zipfile.ZipFile(io.BytesIO(data))
        except Exception as e:
            notes.append(f"圖片 zip 讀取失敗：{e}")
            return
        for n in z.namelist():
            low = n.lower()
            base = os.path.basename(n)
            if low.endswith('.zip') and depth < 3:
                take_zip(z.read(n), depth + 1)
            elif low.endswith(('.png', '.jpg', '.jpeg')) and base and not base.startswith('.'):
                stem = re.sub(r'[\s ]', '', os.path.splitext(base)[0])
                m = re.search(r'(\d{3})[-_ ]?(\d{2})[-_ ]?(\d{4})', stem)
                if not m:
                    continue
                with open(os.path.join(out_dir, '%s-%s-%s%s' % (m.group(1), m.group(2), m.group(3),
                                                              os.path.splitext(base)[1].lower())), 'wb') as fh:
                    fh.write(z.read(n))
                count += 1

    for f in image_files or []:
        name = getattr(f, 'name', '')
        try:
            f.seek(0)
        except Exception:
            pass
        data = f.read()
        low = name.lower()
        if low.endswith('.zip'):
            take_zip(data)
        elif low.endswith(('.xlsx', '.xlsm')):
            n, ns = extract_spk_images(data, out_dir)
            count += n
            notes += [f"{name}：{x}" for x in ns]
        elif low.endswith(('.png', '.jpg', '.jpeg')):
            m = re.search(r'(\d{3})[-_ ]?(\d{2})[-_ ]?(\d{4})', name)
            if m:
                with open(os.path.join(out_dir, '%s-%s-%s%s' % (m.group(1), m.group(2), m.group(3),
                                                              os.path.splitext(low)[1])), 'wb') as fh:
                    fh.write(data)
                count += 1
    return count, notes


def _json_tail(text):
    """腳本 stdout 末段的 JSON 物件（解析不到回傳 {}）。"""
    try:
        return json.loads(text[text.index('{'):text.rindex('}') + 1])
    except Exception:
        return {}


def run_grid_engine(live_docs, master_files, asst_files, image_files, existing_grid, title):
    """
    產生 PO GRID。existing_grid 有上傳 → 更新模式（保留原檔圖片與手填內容，只補數量／插入新 PO 欄）；
    沒上傳 → 重建模式（每間工廠一個工作表）。
    回傳 dict：ok / mode / data(bytes) / filename / report / notes / log
    """
    import subprocess, sys, tempfile
    eng = _grid_engine_dir()
    res = {'ok': False, 'mode': 'update' if existing_grid is not None else 'create',
           'data': None, 'filename': None, 'report': {}, 'notes': [], 'log': ''}
    need = ['parse_po_pdfs.py', 'reconcile.py', 'build_grid.py', 'update_grid.py']
    lost = [n for n in need if not os.path.exists(os.path.join(eng, n))]
    if lost:
        res['notes'].append("找不到 grid_engine 資料夾內的腳本：" + ", ".join(lost) + "。請確認已連同 app.py 一起部署。")
        return res

    def run(args):
        p = subprocess.run([sys.executable] + args, capture_output=True, text=True, timeout=900)
        res['log'] += f"\n$ {os.path.basename(args[0])}\n{p.stdout[-3000:]}\n{p.stderr[-1500:]}"
        return p

    with tempfile.TemporaryDirectory() as td:
        pos_dir = os.path.join(td, 'pos')
        os.makedirs(pos_dir)
        cache, used = {}, set()
        for d in live_docs:
            mmdd = d.get('fname_date') or '0101'
            name = f"240_{d['po']}_{mmdd}.pdf"
            if name in used:
                name = f"240_{d['po']}-{d['dc']}_{mmdd}.pdf"
            used.add(name)
            path = os.path.join(pos_dir, name)
            open(path, 'wb').close()               # 佔位檔；文字直接由快取提供，不需再讀 PDF
            cache[path] = d['text']
        work = os.path.join(td, 'work')
        os.makedirs(work)
        with open(os.path.join(work, 'text.json'), 'w', encoding='utf-8') as fh:
            json.dump(cache, fh, ensure_ascii=False)

        p = run([os.path.join(eng, 'parse_po_pdfs.py'), '--input', pos_dir, '--outdir', work])
        parse_rep = _json_tail(p.stdout)
        if not os.path.exists(os.path.join(work, 'items.json')):
            res['notes'].append("PO 解析步驟失敗，無法產生 GRID。")
            return res
        if parse_rep.get('problems_by_kind'):
            res['notes'].append(f"PO 內部驗算有問題：{parse_rep['problems_by_kind']}（GRID 仍會產生，請複核這些 PO）")

        safe_title = re.sub(r'[^\w\-]+', '_', title).strip('_') or 'PO_GRID'

        # ───────── 更新既有 GRID ─────────
        if existing_grid is not None:
            gpath = os.path.join(td, 'grid_in.xlsx')
            existing_grid.seek(0)
            with open(gpath, 'wb') as fh:
                fh.write(existing_grid.read())
            out = os.path.join(td, 'grid_out.xlsx')
            p = run([os.path.join(eng, 'update_grid.py'), '--grid', gpath,
                     '--data', os.path.join(work, 'grid.json'),
                     '--po-meta', os.path.join(work, 'items.json'), '--out', out])
            res['report'] = _json_tail(p.stdout)
            if not os.path.exists(out):
                res['notes'].append("更新既有 GRID 失敗（版面可能與標準 GRID 不同）。可改為不上傳既有 GRID，直接重建。")
                return res
            res['data'] = open(out, 'rb').read()
            base = re.sub(r'\.xlsx$', '', getattr(existing_grid, 'name', safe_title), flags=re.I)
            res['filename'] = f"{base}_updated_{datetime.now().strftime('%m%d')}.xlsx"
            res['ok'] = True
            return res

        # ───────── 重建 ─────────
        args = [os.path.join(eng, 'reconcile.py'), '--items', os.path.join(work, 'items.json'), '--outdir', work]
        xl_masters = [m for m in master_files if getattr(m, 'name', '').lower().endswith(('.xlsx', '.xlsm'))]
        if not xl_masters:
            res['notes'].append("產生 GRID 需要 Excel 格式（.xlsx）的產品資料表。")
            return res
        for k, mf in enumerate(xl_masters):
            mpath = os.path.join(td, f'master_{k}.xlsx')
            mf.seek(0)
            with open(mpath, 'wb') as fh:
                fh.write(mf.read())
            label = title if len(xl_masters) == 1 else re.sub(r'\.xls[xm]$', '', mf.name, flags=re.I)
            args += ['--master', f"{label}={mpath}"]
        import openpyxl
        for k, af in enumerate(asst_files or []):
            apath = os.path.join(td, f'asst_{k}.xlsx')
            af.seek(0)
            with open(apath, 'wb') as fh:
                fh.write(af.read())
            try:
                wb = openpyxl.load_workbook(apath, read_only=True, data_only=True)
                for ws in wb.worksheets:
                    hit = False
                    for row in ws.iter_rows(min_row=1, max_row=40, values_only=True):
                        cells = {str(c).strip() for c in row if c}
                        if cells & {'Component Item DPCI', 'Item DPCI'} and 'Assortment DPCI' in cells:
                            hit = True
                            break
                    if hit:
                        args += ['--assortment', f"{apath}:{ws.title}"]
                wb.close()
            except Exception as e:
                res['notes'].append(f"混裝表 {getattr(af, 'name', '')} 讀取失敗：{e}")
        p = run(args)
        if not os.path.exists(os.path.join(work, 'recon.json')):
            msg = (p.stdout + p.stderr).strip().splitlines()[-1:] or ['']
            res['notes'].append("主檔比對步驟失敗：" + msg[0][:300])
            return res
        recon_rep = _json_tail(p.stdout)

        img_dir = os.path.join(td, 'imgs')
        os.makedirs(img_dir)
        n_img, img_notes = collect_grid_images(image_files, img_dir)
        res['notes'] += img_notes

        out = os.path.join(td, 'grid.xlsx')
        rpt = os.path.join(td, 'build_report.json')
        build = [os.path.join(eng, 'build_grid.py'), '--recon', os.path.join(work, 'recon.json'),
                 '--items', os.path.join(work, 'items.json'), '--title', title,
                 '--out', out, '--report', rpt, '--no-recalc']
        if n_img:
            build += ['--images', img_dir]
        p = run(build)
        if not os.path.exists(out):
            gaps = _json_tail(p.stdout).get('gaps')
            if gaps:
                res['notes'].append("以下 PO 讀不到 PO 類型或 Shipping Window，該欄表頭會留空，請人工補上：" +
                                    ", ".join(f"{g_['po']}（{g_['missing']}）" for g_ in gaps))
                p = run(build + ['--allow-header-gaps'])
        if not os.path.exists(out):
            res['notes'].append("GRID 建立步驟失敗。")
            return res
        rep = json.load(open(rpt, encoding='utf-8')) if os.path.exists(rpt) else {}
        rep['findings_counts'] = recon_rep.get('findings_counts', {})
        rep['images_supplied'] = n_img
        res['report'] = rep
        res['data'] = open(out, 'rb').read()
        res['filename'] = f"{safe_title}_PO_GRID_{datetime.now().strftime('%Y%m%d')}.xlsx"
        res['ok'] = True
        return res


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
                              'Case Unit Quantity', 'Ent Ttl Rcpt U', 'Target UPC', 'Factory Name',
                              'Product Description']
                 if c in prod_df.columns]
    merged_df = pd.merge(
        po_df,
        prod_df[prod_cols].drop_duplicates(subset=['DPCI']),
        left_on='Final_DPCI', right_on='DPCI', how='left'
    )

    # ---- 列類型：line（一般品項）/ box（混裝 Box）/ component（Box 內零件）----
    has_asst = asst_df is not None and len(asst_df) > 0
    if 'Row_Type' not in merged_df.columns:
        # CSV 來源沒有列類型：標準版以 ASSORTMENT ITEM? 判斷零件；現代版以混裝表判斷 Box
        is_y = merged_df.get('ASSORTMENT ITEM?', pd.Series('N', index=merged_df.index)).astype(str) == 'Y'
        merged_df['Row_Type'] = np.where(is_y, 'component', 'line')
        if has_asst:
            box_set = set(asst_df['Assortment_DPCI'].dropna())
            merged_df.loc[(~is_y) & merged_df['Original_DPCI'].isin(box_set), 'Row_Type'] = 'box'
    is_line = merged_df['Row_Type'] == 'line'
    is_box = merged_df['Row_Type'] == 'box'
    is_comp = merged_df['Row_Type'] == 'component'
    merged_df['ASSORTMENT ITEM?'] = np.where(is_line, 'N', 'Y')
    for c in ['COMPONENT ASSORT QTY', 'VCP QUANTITY', 'Box_QTY']:
        if c not in merged_df.columns:
            merged_df[c] = np.nan

    # ---- 混裝表查找：(Box, 零件) → 每箱入數；Box → proposal 成本 ----
    pair_units, box_cost_prop, box_comps = {}, {}, {}
    if has_asst:
        for _, a in asst_df.iterrows():
            pair_units[(a['Assortment_DPCI'], a['Component_DPCI'])] = a['Units_in_Assortment']
            box_comps.setdefault(a['Assortment_DPCI'], set()).add(a['Component_DPCI'])
            if pd.notna(a['Asst_Box_Cost']):
                box_cost_prop.setdefault(a['Assortment_DPCI'], a['Asst_Box_Cost'])
    fca = prod_df.drop_duplicates(subset=['DPCI']).set_index('DPCI')['Final_Product_Cost'].to_dict() \
        if 'Final_Product_Cost' in prod_df.columns else {}

    merged_df['Units_in_Assortment'] = [
        pair_units.get((o, f), np.nan) if t == 'component' else np.nan
        for o, f, t in zip(merged_df['Original_DPCI'], merged_df['Final_DPCI'], merged_df['Row_Type'])]
    merged_df['Matched_Box_DPCI'] = np.where(is_comp, merged_df['Original_DPCI'], np.nan)
    merged_df['Asst_Box_Cost'] = np.where(is_box, merged_df['Original_DPCI'].map(box_cost_prop), np.nan)

    # 反算 Box 成本 = Σ(零件 FCA × 每箱入數)；優先用混裝表的入數，沒有混裝表時用 PO 自身的零件列
    calc_prop = {}
    for b, comps in box_comps.items():
        vals = [fca.get(c, np.nan) * pair_units[(b, c)] for c in comps]
        calc_prop[b] = np.nan if any(pd.isna(v) for v in vals) else float(np.sum(vals))
    comp_rows = merged_df[is_comp].copy()
    comp_rows['_contrib'] = comp_rows['Final_Product_Cost'] * comp_rows['COMPONENT ASSORT QTY']
    calc_po = comp_rows.groupby(['PO NUMBER', 'Original_DPCI'])['_contrib'].agg(
        lambda s: np.nan if s.isna().any() else s.sum()).to_dict()
    merged_df['Calc_Box_Cost'] = [
        (calc_prop.get(o) if pd.notna(calc_prop.get(o, np.nan)) else calc_po.get((p, o), np.nan)) if t == 'box' else np.nan
        for p, o, t in zip(merged_df['PO NUMBER'], merged_df['Original_DPCI'], merged_df['Row_Type'])]

    def _role(row):
        if row['Row_Type'] == 'box':
            if has_asst and row['Original_DPCI'] not in box_comps:
                return '⚠️ Box 不在混裝表'
            return '📦 混裝 Box'
        if row['Row_Type'] == 'component':
            if has_asst and pd.isna(row['Units_in_Assortment']):
                return f"⚠️ 零件不在混裝表的 {row['Original_DPCI']}"
            return f"🔹 混裝零件 ({row['Original_DPCI']})"
        return '—'
    merged_df['Asst_Role'] = merged_df.apply(_role, axis=1)

    # ---- G4: 成本比對（tolerance = 0.005）----
    # 一般品項、零件 → FCA/FOB；Box → 混裝表 Box 成本（無混裝表時用反算成本）
    merged_df['Target_Cost'] = np.where(
        is_box, merged_df['Asst_Box_Cost'].fillna(merged_df['Calc_Box_Cost']), merged_df['Final_Product_Cost'])
    merged_df['Cost Match'] = np.where(
        merged_df['Target_Cost'].isna() | merged_df['ITEM UNIT COST'].isna(), False,
        (merged_df['ITEM UNIT COST'] - merged_df['Target_Cost']).abs() <= 0.005 + 1e-9)

    # ---- N2: 零售價比對（tolerance = 0.005）；Box 與零件無單一零售價 → n/a ----
    if 'Suggested Unit Retail' in merged_df.columns:
        merged_df['Retail Match'] = np.where(
            ~is_line, True,
            np.where(merged_df['ITEM UNIT RETAIL'].isna() | merged_df['Suggested Unit Retail'].isna(), False,
                     (merged_df['ITEM UNIT RETAIL'] - merged_df['Suggested Unit Retail']).abs() <= 0.005 + 1e-9))
    else:
        merged_df['Retail Match'] = True

    # ---- 裝箱數比對：一般品項 VCP vs Case；零件 每箱入數 vs 混裝表（PO 無資料則略過）----
    merged_df['Target Case / Assort QTY'] = np.where(
        is_comp, merged_df['Units_in_Assortment'],
        np.where(is_line, merged_df.get('Case Unit Quantity', pd.Series(np.nan, index=merged_df.index)), np.nan))
    merged_df['PO VCP / Assort QTY'] = np.where(
        is_comp, merged_df['COMPONENT ASSORT QTY'], np.where(is_line, merged_df['VCP QUANTITY'], np.nan))
    po_pack = pd.to_numeric(merged_df['PO VCP / Assort QTY'], errors='coerce')
    tg_pack = pd.to_numeric(merged_df['Target Case / Assort QTY'], errors='coerce')
    merged_df['Case QTY Match'] = np.where(
        po_pack.isna(), True,                                   # PO 無此資料 → 不檢核
        np.where(tg_pack.isna(), ~(is_comp & has_asst) if has_asst else True,
                 (po_pack - tg_pack).abs() <= 0.01))
    merged_df['Case QTY Match'] = merged_df['Case QTY Match'].astype(bool)

    # ---- G2: 總數量 = 單品 PO 數量 + Box 內含數量；Box 列本身不計（避免重複）----
    merged_df['Target Commit QTY'] = np.where(
        is_box, np.nan, merged_df.get('Ent Ttl Rcpt U', pd.Series(np.nan, index=merged_df.index)))
    case_pack = pd.to_numeric(
        merged_df.get('Case Unit Quantity', pd.Series(1, index=merged_df.index)), errors='coerce').fillna(1)
    shipper_mask = merged_df.get('Is_Shipper_Display', pd.Series(False, index=merged_df.index)).fillna(False).astype(bool)
    merged_df['Final_QTY_for_count'] = np.where(shipper_mask | is_box, 0, merged_df['Final_QTY'])
    merged_df['PO Total QTY'] = merged_df.groupby('Final_DPCI')['Final_QTY_for_count'].transform('sum')
    merged_df['PO Total QTY'] = np.where(is_box, merged_df['Final_QTY'], merged_df['PO Total QTY'])
    qty_diff = (merged_df['PO Total QTY'] - merged_df['Target Commit QTY']).abs()
    merged_df['Total QTY Match'] = np.where(
        is_box, True,                                           # Box 無計畫量 → n/a
        np.where(merged_df['Target Commit QTY'].isna(), False,
                 (qty_diff < case_pack) | (qty_diff <= 0.10 * merged_df['Target Commit QTY'])))
    merged_df['Total QTY Match'] = merged_df['Total QTY Match'].astype(bool)
    merged_df['QTY Diff'] = merged_df['PO Total QTY'] - merged_df['Target Commit QTY']
    merged_df['QTY Diff %'] = np.where(
        merged_df['Target Commit QTY'].replace(0, np.nan).isna(), np.nan,
        (merged_df['QTY Diff'] / merged_df['Target Commit QTY'] * 100).round(1))

    # ---- UPC 比對（去除前導 0 後比較；Excel 常把 Barcode 存成數字）----
    def _upc_key(s):
        return s.astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.lstrip('0')
    upc_both_exist = merged_df['PO UPC'].notna() & merged_df['Target UPC'].notna() & ~is_box
    merged_df['UPC Match'] = np.where(
        upc_both_exist, _upc_key(merged_df['PO UPC']) == _upc_key(merged_df['Target UPC']), True)
    merged_df['UPC Status'] = np.where(
        ~upc_both_exist, '⚪ 無資料', np.where(merged_df['UPC Match'], '✅ 相符', '❌ 不符'))

    # ---- N3: 零件數量展開：Box 箱數 × 每箱入數（混裝表）vs PO 零件數量 ----
    merged_df['Expected_Component_QTY'] = np.where(
        is_comp, pd.to_numeric(merged_df['Box_QTY'], errors='coerce') * merged_df['Units_in_Assortment'], np.nan)
    merged_df['Asst_QTY_Check'] = np.where(
        ~is_comp | merged_df['Expected_Component_QTY'].isna(), '⚪ N/A',
        np.where((merged_df['Expected_Component_QTY'] - merged_df['Final_QTY']).abs() <= 0.5,
                 '✅ 數量相符', '❌ 數量不符'))

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

# ==========================================
# 顯示格式：金額加 $、數量加千分位（畫面與 Excel 共用同一份欄位清單）
# ==========================================
MONEY_COLS = {'主檔成本', 'PO 單價', '主檔零售', 'PO 零售',
              'ITEM UNIT COST', 'Target_Cost', 'ITEM UNIT RETAIL', 'Suggested Unit Retail', 'Asst_Box_Cost'}
QTY_COLS = {'PO 數量', 'PO 數量（含 Assortment 內含）', 'PCN Commit Qty', '差異', 'Case Pack',
            '應為（箱數×入數）', 'DPCI 合計',
            'Final_QTY', 'Box_QTY', 'Final_QTY_for_count', 'PO Total QTY', 'QTY Diff',
            'PO VCP / Assort QTY', 'Target Case / Assort QTY', 'COMPONENT ASSORT QTY', 'Expected_Component_QTY'}
XL_MONEY_FMT = '"$"#,##0.00##'
XL_QTY_FMT = '#,##0'


def fmt_money(v):
    """$1.09、$0.219、$1,234.50；空值回傳空字串。"""
    try:
        if v is None or pd.isna(v):
            return ''
        v = float(v)
    except (TypeError, ValueError):
        return str(v)
    txt = f"{abs(v):,.4f}".rstrip('0')
    if len(txt.split('.')[1]) < 2:
        txt = f"{abs(v):,.2f}"
    return ('-$' if v < 0 else '$') + txt


def fmt_qty(v):
    try:
        if v is None or pd.isna(v):
            return ''
        return f"{float(v):,.0f}"
    except (TypeError, ValueError):
        return str(v)


def styled_for_screen(df):
    """依欄名套用 $ 與千分位（只改顯示，不改資料）。"""
    fmt = {c: fmt_money for c in df.columns if c in MONEY_COLS and pd.api.types.is_numeric_dtype(df[c])}
    fmt.update({c: fmt_qty for c in df.columns if c in QTY_COLS and pd.api.types.is_numeric_dtype(df[c])})
    if '差異 %' in df.columns and pd.api.types.is_numeric_dtype(df['差異 %']):
        fmt['差異 %'] = lambda v: '' if pd.isna(v) else f"{v:,.1f}%"
    return df.style.format(fmt, na_rep='')


def apply_excel_number_formats(ws, header_row=1, first_data_row=None, last_row=None):
    """依表頭名稱幫整欄設定 Excel 數字格式（數值仍是數字，可加總）。"""
    first_data_row = first_data_row or header_row + 1
    last_row = last_row or ws.max_row
    for c in ws[header_row]:
        name = str(c.value).strip() if c.value is not None else ''
        f = XL_MONEY_FMT if name in MONEY_COLS else (XL_QTY_FMT if name in QTY_COLS else None)
        if f:
            for r in range(first_data_row, last_row + 1):
                ws.cell(r, c.column).number_format = f


def build_validation_summary(merged_df, ctx=None):
    """
    第一頁總結（對齊 po-grid Skill 的 Validation Summary）：
    回傳 (checks, details, headline)
      checks  = [(檢核項目, 數量, 狀態)]；狀態 'OK' / 'REVIEW' / ''（純資訊）
      details = [(標題, DataFrame)]，只列有內容的項目
      headline = [(標籤, 值)]
    一律以 DPCI 為單位彙總，同一問題出現在多張 PO 只算一項，PO 號列在明細。
    """
    ctx = ctx or {}
    m = merged_df
    rt = m['Row_Type'] if 'Row_Type' in m.columns else pd.Series('line', index=m.index)
    not_box = rt != 'box'
    desc_col = 'Product Description' if 'Product Description' in m.columns else None

    def _pos(s):
        u = sorted(set(s.astype(str)))
        return ', '.join(u)

    def _desc(g):
        return str(g[desc_col].iloc[0])[:60] if desc_col and pd.notna(g[desc_col].iloc[0]) else ''

    details = []

    # 1. PO 品項不在主檔
    unk = m[not_box & m['Final_Product_Cost'].isna() & m.get('Target Commit QTY', pd.Series(np.nan, index=m.index)).isna()] \
        if 'Final_Product_Cost' in m.columns else m.iloc[0:0]
    unk_df = pd.DataFrame([{'DPCI': d, 'PO': _pos(g['PO NUMBER']), 'PO 數量': g['Final_QTY'].sum()}
                           for d, g in unk.groupby('Final_DPCI')])
    # 2. 成本不符
    cm = m[(m['Cost Match'] == False) & ~m.index.isin(unk.index)]
    cost_df = pd.DataFrame([{
        'DPCI': d, '品名': _desc(g), '類型': {'line': '一般', 'box': 'Assortment', 'component': 'Assortment'}.get(g['Row_Type'].iloc[0], ''),
        '主檔成本': g['Target_Cost'].iloc[0], 'PO 單價': ', '.join(fmt_money(v) for v in sorted(set(g['ITEM UNIT COST'].dropna()))) or '讀不到',
        'PO': _pos(g['PO NUMBER'])} for d, g in cm.groupby('Final_DPCI')])
    # 3. 零售不符
    rm = m[(m['Retail Match'] == False) & ~m.index.isin(unk.index)]
    retail_df = pd.DataFrame([{
        'DPCI': d, '品名': _desc(g), '主檔零售': g['Suggested Unit Retail'].iloc[0] if 'Suggested Unit Retail' in g else np.nan,
        'PO 零售': ', '.join(fmt_money(v) for v in sorted(set(g['ITEM UNIT RETAIL'].dropna()))) or '讀不到',
        'PO': _pos(g['PO NUMBER'])} for d, g in rm.groupby('Final_DPCI')])
    # 4. 數量 vs 計畫
    qm = m[not_box & (m['Total QTY Match'] == False) & m['Target Commit QTY'].notna()].drop_duplicates('Final_DPCI')
    qty_df = pd.DataFrame([{
        'DPCI': r['Final_DPCI'], '品名': str(r[desc_col])[:60] if desc_col and pd.notna(r[desc_col]) else '',
        'PO 數量（含 Assortment 內含）': r['PO Total QTY'], 'PCN Commit Qty': r['Target Commit QTY'],
        '差異': r['QTY Diff'], '差異 %': r['QTY Diff %'], 'Case Pack': r.get('Case Unit Quantity', np.nan)}
        for _, r in qm.iterrows()])
    # 5. 混裝不符（零件數量 ≠ 箱數 × 每箱入數，或 Box／零件不在混裝表）
    am = m[(m.get('Asst_QTY_Check', '') == '❌ 數量不符') | m['Asst_Role'].astype(str).str.startswith('⚠️')]
    asst_df_ = pd.DataFrame([{
        'PO': r['PO NUMBER'], 'Box DPCI': r['Original_DPCI'], '零件 DPCI': r['Final_DPCI'] if r['Row_Type'] == 'component' else '',
        'PO 數量': r['Final_QTY'], '應為（箱數×入數）': r.get('Expected_Component_QTY', np.nan),
        '說明': re.sub(r'^[^\w]+', '', str(r['Asst_Role']))} for _, r in am.iterrows()])
    # 7. UPC
    um = m[m.get('UPC Status', '') == '❌ 不符']
    upc_df = pd.DataFrame([{'DPCI': d, 'PO UPC': g['PO UPC'].iloc[0], '主檔 Barcode': g['Target UPC'].iloc[0],
                            'PO': _pos(g['PO NUMBER'])} for d, g in um.groupby('Final_DPCI')])

    no_df = pd.DataFrame(ctx.get('not_ordered', []))
    dup_df = pd.DataFrame([{'PO（品項與數量完全相同）': ' / '.join(g)} for g in ctx.get('duplicates', [])])
    sup_df = pd.DataFrame([{'PO': x['po'], '採用': x['kept'], '捨棄': x['dropped'],
                            'Total Qty（舊→新）': f"{x.get('from_qty') or '?'} → {x.get('to_qty') or '?'}"}
                           for x in ctx.get('superseded', [])])
    can_df = pd.DataFrame([{'PO': x['po'], '取消通知檔': x['file'], '已排除的原單': ', '.join(x['dropped']) or '—'}
                           for x in ctx.get('cancelled', [])])
    gate_df = pd.DataFrame([{'PO': x['po'], '檔案': x['file'], '問題': '；'.join(x['issues'])}
                            for x in ctx.get('gate_problems', [])])
    skip_df = pd.DataFrame([{'PO': x['po'], '來源檔': x['file'], '品項（前 3 個）': ', '.join(x['dpcis'])}
                            for x in ctx.get('skipped_other_program', [])])
    warn_df = pd.DataFrame([{'訊息': re.sub(r'[*]', '', w)} for w in ctx.get('warnings', [])])

    def st_(n):
        return 'OK' if n == 0 else 'REVIEW'
    checks = [
        ('PO 品項不在主檔 (Unknown DPCI)', len(unk_df), st_(len(unk_df))),
        ('成本不符：PO 單價 vs 主檔 FCA/FOB (Cost mismatch)', len(cost_df), st_(len(cost_df))),
        ('零售不符：PO Resale vs 主檔 (Retail mismatch)', len(retail_df), st_(len(retail_df))),
        ('PO數量 vs PCN Commit：超出整箱進位與 ±10% (Qty vs plan)', len(qty_df), st_(len(qty_df))),
        ('混裝不符：Box 與零件數量／混裝表 (Assortment mismatch)', len(asst_df_), st_(len(asst_df_))),
        ('UPC 不符', len(upc_df), st_(len(upc_df))),
        ('尚未下單(未收到PO的Item) (Not yet ordered)', len(no_df), st_(len(no_df))),
        ('疑似重複 PO：不同 PO# 內容相同 (Duplicate PO groups)', len(dup_df), st_(len(dup_df))),
        ('PO 內部驗算未通過 (Self-check failed)', len(gate_df), st_(len(gate_df))),
        ('無法解析的檔案／其他警告', len(warn_df), st_(len(warn_df))),
        ('已取消的 PO（已排除）', len(can_df), ''),
        ('多版本 PO（已採用最新版）', len(sup_df), ''),
        ('其他 Program 的訂單（已略過）', len(skip_df), ''),
    ]
    for title, df in [
        ('PO 品項不在主檔 — 明細', unk_df), ('成本不符 — 明細', cost_df), ('零售不符 — 明細', retail_df),
        ('PO數量 vs PCN Commit — 明細', qty_df), ('混裝不符 — 明細', asst_df_),
        ('UPC 不符 — 明細', upc_df), ('尚未下單(未收到PO的Item) — 明細', no_df), ('疑似重複 PO — 明細', dup_df),
        ('PO 內部驗算未通過 — 明細', gate_df), ('無法解析的檔案／其他警告 — 明細', warn_df),
        ('已取消的 PO — 明細', can_df), ('多版本 PO — 明細', sup_df), ('其他 Program 的訂單 — 明細', skip_df)]:
        if len(df) > 0:
            details.append((title, df))

    ordered = float(m.get('Final_QTY_for_count', pd.Series(dtype=float)).sum())
    plan = ctx.get('plan_total')
    n_review = sum(1 for _, n, s_ in checks if s_ == 'REVIEW')
    headline = [
        ('有效 PO 張數', int(m['PO NUMBER'].nunique())),
        ('品項列數', len(m)),
        ('DPCI 數（不含 Box）', int(m.loc[not_box, 'Final_DPCI'].nunique())),
        ('已下單總數量', f"{ordered:,.0f}"),
    ]
    if plan:
        headline.append(('PCN Commit Qty / 達成率', f"{plan:,.0f} / {ordered / plan * 100:.1f}%"))
    headline.append(('結論', '全部檢核通過' if n_review == 0 else f'{n_review} 項需確認（見下方明細）'))
    return checks, details, headline


def write_summary_sheet(wb, title, checks, details, headline, run_meta):
    """把總結寫成活頁簿的第一個工作表。"""
    from openpyxl.styles import Font, PatternFill, Alignment
    ws = wb.create_sheet('Validation Summary', 0)
    green = PatternFill('solid', start_color='E2EFDA')
    orange = PatternFill('solid', start_color='F8CBAD')
    grey = PatternFill('solid', start_color='D9D9D9')
    ws['A1'] = title
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = 'PO Validation Summary'
    ws['A2'].font = Font(italic=True, bold=True, size=12)
    ws['A3'] = f"執行時間 {run_meta.get('timestamp', '')}　｜　{run_meta.get('input_files', '')}"
    ws['A3'].font = Font(color='808080', size=9)
    r = 5
    for label, val in headline:
        ws.cell(r, 1, label).font = Font(bold=True)
        c = ws.cell(r, 2, val)
        c.alignment = Alignment(horizontal='left')
        if label == '結論':
            c.font = Font(bold=True, color='C00000' if '需確認' in str(val) else '548235')
        r += 1
    r += 1
    for j, h in enumerate(['Check 檢核項目', 'Count', 'Status'], 1):
        c = ws.cell(r, j, h)
        c.font = Font(bold=True)
        c.fill = grey
    r += 1
    for label, n, status in checks:
        fill = orange if status == 'REVIEW' else (green if status == 'OK' else None)
        for j, v in enumerate([label, n, status], 1):
            c = ws.cell(r, j, v)
            if fill:
                c.fill = fill
        r += 1
    for dtitle, df in details:
        r += 1
        ws.cell(r, 1, dtitle).font = Font(bold=True, size=11)
        r += 1
        for j, h in enumerate(df.columns, 1):
            c = ws.cell(r, j, h)
            c.font = Font(bold=True)
            c.fill = grey
        r += 1
        cols = list(df.columns)
        for row in df.itertuples(index=False):
            for j, v in enumerate(row, 1):
                if isinstance(v, (np.floating, float)):
                    v = None if pd.isna(v) else (int(v) if float(v).is_integer() else float(v))
                elif isinstance(v, np.integer):
                    v = int(v)
                c = ws.cell(r, j, v)
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    if cols[j - 1] in MONEY_COLS:
                        c.number_format = XL_MONEY_FMT
                    elif cols[j - 1] in QTY_COLS:
                        c.number_format = XL_QTY_FMT
                    elif cols[j - 1] == '差異 %':
                        c.number_format = '0.0"%"'
            r += 1
    ws.column_dimensions['A'].width = 54
    for col, w in zip('BCDEFGH', [26, 22, 22, 22, 40, 14, 14]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A5'
    wb.active = 0


def make_excel_bytes(result_df, run_meta, source_label, validation_notes=None, merged_df_full=None, summary_ctx=None):
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
        apply_excel_number_formats(ws)
        for col_cells in ws.columns:
            max_len = max((len(str(cell.value)) for cell in col_cells if cell.value), default=8)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 30)

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

        # ---- 第一頁：Validation Summary（對齊 Skill 版）----
        if merged_df_full is not None and 'Cost Match' in merged_df_full.columns:
            ctx = summary_ctx or {}
            checks, details, headline = build_validation_summary(merged_df_full, ctx)
            write_summary_sheet(writer.book, ctx.get('title') or 'PO Validation', checks, details, headline, run_meta)

    return output.getvalue()

def show_results(merged_df, source_label, run_meta=None, validation_notes=None, summary_ctx=None):
    """
    統一結果顯示 + 彩色 + 摘要統計 + Excel 下載（含 GRID / 驗核摘要 / 執行摘要）
    """
    display_cols = [
        'PO NUMBER', 'DC', 'ASSORTMENT ITEM?', 'Asst_Role', 'Matched_Box_DPCI', 'Is_Shipper_Display', 'Original_DPCI', 'Final_DPCI',
        'ITEM DESCRIPTION', 'Final_QTY', 'Box_QTY', 'Final_QTY_for_count',
        'Cost Match', 'ITEM UNIT COST', 'Target_Cost',
        'Retail Match', 'ITEM UNIT RETAIL', 'Suggested Unit Retail',
        'Case QTY Match', 'PO VCP / Assort QTY', 'Target Case / Assort QTY',
        'Total QTY Match', 'PO Total QTY', 'Target Commit QTY', 'QTY Diff', 'QTY Diff %',
        'UPC Status', 'PO UPC', 'Target UPC',
        'Asst_Box_Cost',
        'Asst_QTY_Check', 'COMPONENT ASSORT QTY', 'Expected_Component_QTY',
        'Factory_Match_Status', 'Factory Name', 'Dispatch_Factory',
        'Dispatch_AC', 'Dispatch_AE', 'AC_Rule_Check', 'AE_Rule_Check',
        'All Match (Pass)'
    ]
    result_df = merged_df[[c for c in display_cols if c in merged_df.columns]].copy()
    result_df = result_df.rename(columns={'Target Commit QTY': 'PCN Commit Qty'})

    # ── 總結（與 Excel 第一頁相同）──
    checks, details, headline = build_validation_summary(merged_df, summary_ctx)
    st.markdown("### 📌 Validation Summary")
    st.markdown("　｜　".join(f"**{k}**：{v}" for k, v in headline))
    chk_df = pd.DataFrame(checks, columns=['檢核項目', 'Count', 'Status'])
    def _chk_color(row):
        c = '#F8CBAD' if row['Status'] == 'REVIEW' else ('#E2EFDA' if row['Status'] == 'OK' else '')
        return [f'background-color: {c}; color: #000' if c else ''] * len(row)
    st.dataframe(chk_df.style.apply(_chk_color, axis=1), hide_index=True, use_container_width=True)
    for dtitle, ddf in details:
        with st.expander(f"{dtitle}（{len(ddf)}）", expanded=len(ddf) <= 10 and '已略過' not in dtitle and '其他 Program' not in dtitle):
            st.dataframe(styled_for_screen(ddf), hide_index=True, use_container_width=True)

    # ── 逐筆異常（只列有問題的 PO 品項列；完整逐筆結果在 Excel「核對結果」）──
    errors_df = result_df[result_df['All Match (Pass)'] == False]
    if len(errors_df) == 0:
        st.success("🎉 所有品項列皆一致！")
    else:
        screen_cols = {
            'PO NUMBER': 'PO', 'DC': 'DC', 'Asst_Role': '類型', 'Final_DPCI': 'DPCI',
            'Final_QTY': 'PO 數量', 'ITEM UNIT COST': 'PO 單價', 'Target_Cost': '主檔成本',
            'ITEM UNIT RETAIL': 'PO 零售', 'Suggested Unit Retail': '主檔零售',
            'PO Total QTY': 'DPCI 合計', 'PCN Commit Qty': 'PCN Commit Qty', 'QTY Diff %': '差異 %',
            'UPC Status': 'UPC', 'Cost Match': '成本', 'Retail Match': '零售', 'Total QTY Match': '數量'}
        view = errors_df[[c for c in screen_cols if c in errors_df.columns]].rename(columns=screen_cols)
        if '類型' in view.columns:
            view['類型'] = view['類型'].astype(str).str.replace(r'^[^\w(]+', '', regex=True)
        bool_cols = [c for c in ['成本', '零售', '數量'] if c in view.columns]
        def _bad(v):
            return 'background-color: #F8CBAD; color: #000' if v is False or v == False else ''
        styler = styled_for_screen(view)
        styler = styler.map(_bad, subset=bool_cols) if hasattr(styler, 'map') else styler.applymap(_bad, subset=bool_cols)
        with st.expander(f"❌ 逐筆異常明細（{len(errors_df)} 筆 PO 品項列）", expanded=False):
            st.dataframe(styler, hide_index=True, use_container_width=True)
            st.caption("完整逐筆核對結果（含相符的列與所有欄位）請見下載的 Excel「核對結果」工作表。")

    # 報表只產生一次，存起來供下載（按下載不會重算、結果不會消失）
    res = st.session_state.get('results')
    if res is not None and res.get('excel') is None:
        run_meta = run_meta or {'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 'input_files': source_label}
        res['excel'] = make_excel_bytes(result_df, run_meta, source_label, validation_notes=validation_notes,
                                        merged_df_full=merged_df, summary_ctx=summary_ctx)
        res['excel_name'] = f'PO_Validation_{re.sub(r"[^\w\-]", "_", source_label)}_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx'

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

st.sidebar.markdown("---")
st.sidebar.header("📊 步驟 3（選填）：PO GRID")
st.sidebar.caption("核對完成後會一併產生 PO GRID（每間工廠一個工作表）。以下三項都可以不填。")
grid_title_input = st.sidebar.text_input(
    "GRID 標題", value="", placeholder="例：D240 27C2 EASTER",
    help="顯示在每個工作表左上角，也會用在下載的檔名。空白時用產品資料表的檔名。")
st.sidebar.markdown("**產品圖片**")
st.sidebar.caption("放進 GRID 的 PICTURE 欄，二選一或混用：\n\n"
                   "• **圖片 zip**：檔名要含 DPCI（例：240-04-8085.png），zip 裡再包 zip 也可以\n\n"
                   "• **SPK Workspace 匯出的 .xlsx**：自動抽出 Products 工作表的縮圖，依 DPCI 對應\n\n"
                   "不上傳 → PICTURE 欄留空。")
grid_image_files = st.sidebar.file_uploader(
    "產品圖片", type=['zip', 'xlsx', 'xlsm', 'png', 'jpg', 'jpeg'],
    accept_multiple_files=True, key="grid_imgs", label_visibility="collapsed")
st.sidebar.markdown("**上一版 PO GRID**")
st.sidebar.caption("• **不上傳** → 用本次上傳的全部 PO 產生一份全新的 GRID（第一次做、或想重整版面時用）\n\n"
                   "• **上傳** → 在這份 GRID 上補數量、為新 PO 加欄，原檔的圖片與手填內容（AGE、工廠料號等）都保留。"
                   "若有新品項需要加列，會另外附一份全新版本。")
grid_existing_file = st.sidebar.file_uploader(
    "上一版 PO GRID", type=['xlsx'], key="grid_existing", label_visibility="collapsed")

# ==========================================
# 主介面：PDF 上傳解析
# ==========================================
if True:
    st.subheader("📄 上傳 SPS Commerce PO PDF")

    pdf_files = st.file_uploader("", type=['pdf'], accept_multiple_files=True, key="pdf_po", label_visibility="collapsed")

    if st.button("🚀 解析 PDF 並執行核對", type="primary", key="btn_pdf"):
        st.session_state.pop('grid_out', None)
        st.session_state.pop('results', None)
        if not product_files or not pdf_files:
            st.warning("⚠️ 請確保已在側邊欄上傳「產品資料表」，並在上方上傳 PDF！")
        else:
            prog = st.progress(0.0, text="PDF 解析中...")
            def _tick(i, n, name):
                prog.progress(i / n, text=f"PDF 解析中 {i}/{n}：{name}（無文字層的 PDF 需 OCR，每份約 10–20 秒）")
            prod_df = process_products(product_files)
            known = set(prod_df['DPCI'].dropna().astype(str)) if 'DPCI' in prod_df.columns else set()
            asst_df = process_assortments(asst_files) if asst_files else None
            if asst_df is not None and len(asst_df) > 0:
                known |= set(asst_df['Assortment_DPCI']) | set(asst_df['Component_DPCI'])
            po_df, pinfo = parse_po_pdfs(pdf_files, progress=_tick, known_dpcis=known or None)
            prog.empty()

            for cb in pinfo['combined']:
                st.info(f"📑 {cb['file']} 為合併列印，已自動拆成 {cb['orders']} 張訂單，其中 {cb['kept']} 張屬於本主檔。")
            if pinfo['skipped_other_program']:
                with st.expander(f"ℹ️ {len(pinfo['skipped_other_program'])} 張訂單的品項不在主檔內（其他 Program），已略過（點擊展開）", expanded=False):
                    for sk in pinfo['skipped_other_program']:
                        st.write(f"• PO {sk['po']}（{sk['file']}）：{', '.join(sk['dpcis'])}")

            all_warnings = list(pinfo['warnings'])

            if pinfo['ocr_files']:
                with st.expander(f"ℹ️ {len(pinfo['ocr_files'])} 份 PDF 無文字層，已透過 OCR 解析（點擊展開檔名）", expanded=False):
                    for n in pinfo['ocr_files']:
                        st.write(f"• {n}")
            for w in pinfo['warnings']:
                st.warning(w)

            if pinfo['cancelled']:
                msg = [f"PO {c['po']}（取消通知：{c['file']}；已排除：{', '.join(c['dropped']) or '—'}）" for c in pinfo['cancelled']]
                st.warning("🚫 **以下 PO 已被客人取消（CANCEL ORDER），不計入核對**：\n\n" + "\n\n".join("• " + m for m in msg))
                all_warnings.append("已取消 PO（不計入）：" + ", ".join(c['po'] for c in pinfo['cancelled']))
            if pinfo['superseded']:
                with st.expander(f"⚠️ {len(pinfo['superseded'])} 張 PO 有多個版本，已採用最新版（點擊展開）", expanded=False):
                    for s in pinfo['superseded']:
                        st.write(f"• PO {s['po']}：採用 {s['kept']}，捨棄 {s['dropped']}（Total Qty {s['from_qty'] or '?'} → {s['to_qty'] or '?'}）")
                all_warnings.append("多版本 PO（已取最新）：" + ", ".join(s['po'] for s in pinfo['superseded']))
            if pinfo['duplicates']:
                with st.expander(f"⚠️ {len(pinfo['duplicates'])} 組 PO 的品項與數量完全相同，請確認是否重複下單（點擊展開）", expanded=False):
                    for g_ in pinfo['duplicates']:
                        st.write("• " + " / ".join(g_))
                all_warnings.append("品項與數量完全相同的 PO：" + "；".join(" / ".join(g_) for g_ in pinfo['duplicates']))
            # ── G5 PO 內部一致性自我驗證（Σ行小計=PO Total、數量×單價=行小計、Σ數量=Total Qty）──
            if pinfo['gate_problems']:
                with st.expander(f"⚠️ G5 PO 內部驗算：{len(pinfo['gate_problems'])} 張 PO 未通過，該 PO 的解析結果請人工複核（點擊展開）", expanded=True):
                    for gp in pinfo['gate_problems']:
                        st.warning(f"PO {gp['po']}（{gp['file']}）：" + "；".join(gp['issues']))
                all_warnings += [f"G5 驗算未過 PO {gp['po']}：" + "；".join(gp['issues']) for gp in pinfo['gate_problems']]

            if len(po_df) == 0:
                st.error("❌ 所有 PDF 均無法解析出訂單資料，請確認格式後再試。")
            else:
                n_gate_ok = pinfo['n_pos'] - len(pinfo['gate_problems'])
                st.success(f"✅ {pinfo['n_files']} 份 PDF → {pinfo['n_pos']} 張有效 PO、{len(po_df)} 筆品項列；"
                           f"{n_gate_ok}/{pinfo['n_pos']} 張通過 PO 內部驗算。")

                dispatch_arg = dispatch_df_global if len(dispatch_df_global) > 0 else None
                merged_df = run_validation(po_df, prod_df, asst_df, mode='standard', dispatch_df=dispatch_arg)

                summary_ctx = {k: pinfo[k] for k in ['duplicates', 'superseded', 'cancelled', 'gate_problems',
                                                      'skipped_other_program', 'warnings']}
                summary_ctx['title'] = re.sub(r'\.(xlsx|xls|csv)$', '', product_files[0].name, flags=re.I)
                summary_ctx['not_ordered'] = []
                # 主檔有計畫量但完全沒有 PO 的品項
                if 'Ent Ttl Rcpt U' in prod_df.columns and 'DPCI' in prod_df.columns:
                    ordered = set(merged_df.loc[merged_df['Row_Type'] != 'box', 'Final_DPCI'])
                    pm = prod_df.drop_duplicates(subset=['DPCI'])
                    pm_valid = pm[pm['DPCI'].astype(str).str.match(r'^\d{3}-\d{2}-\d{4}$')]
                    summary_ctx['plan_total'] = float(pm_valid['Ent Ttl Rcpt U'].fillna(0).sum())
                    not_ordered = pm[(pm['Ent Ttl Rcpt U'].fillna(0) > 0) & ~pm['DPCI'].isin(ordered)
                                     & pm['DPCI'].astype(str).str.match(r'^\d{3}-\d{2}-\d{4}$')]
                    if len(not_ordered) > 0:
                        all_warnings.append("尚無 PO 的品項：" + ", ".join(not_ordered['DPCI']))
                        dcol = 'Product Description' if 'Product Description' in not_ordered.columns else None
                        summary_ctx['not_ordered'] = [
                            {'DPCI': d, '品名': (str(n)[:60] if dcol else ''), 'PCN Commit Qty': int(q)}
                            for d, n, q in zip(not_ordered['DPCI'],
                                               not_ordered[dcol] if dcol else [''] * len(not_ordered),
                                               not_ordered['Ent Ttl Rcpt U'])]

                run_meta = {
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'input_files': f"PDFs: {pinfo['n_files']} files / {pinfo['n_pos']} POs | Products: {', '.join(f.name for f in product_files)}"
                }
                st.session_state['results'] = {'merged_df': merged_df, 'run_meta': run_meta, 'notes': all_warnings,
                                               'ctx': summary_ctx, 'excel': None}

                # ── PO GRID：有上傳既有 GRID → 更新；否則重建 ──
                grid_title = grid_title_input.strip() or summary_ctx['title']
                with st.spinner("PO GRID 產生中..."):
                    gres = run_grid_engine(pinfo['live_docs'], product_files, asst_files or [],
                                           grid_image_files or [], grid_existing_file, grid_title)
                    grebuild = None
                    if gres['ok'] and gres['mode'] == 'update':
                        no_row = [x for x in gres['report'].get('needs_attention', []) if x.get('dpci')]
                        if no_row:      # 既有 GRID 缺列 → 另外提供一份完整重建版
                            grebuild = run_grid_engine(pinfo['live_docs'], product_files, asst_files or [],
                                                       grid_image_files or [], None, grid_title)
                st.session_state['grid_out'] = {'main': gres, 'rebuild': grebuild}


def render_grid_section(go):
    """PO GRID 結果與下載（存在 session_state，按下載後不會消失）。"""
    gres, grebuild = go['main'], go.get('rebuild')
    st.markdown("---")
    st.markdown("### 📊 PO GRID")
    for n in gres['notes']:
        st.warning(n)
    if not gres['ok']:
        st.error("PO GRID 未能產生。")
        with st.expander("技術訊息"):
            st.code(gres['log'][-3000:])
        return
    rep_ = gres['report']
    xlsx_mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    if gres['mode'] == 'create':
        n_items = rep_.get('item_count', 0) or 0
        st.success(f"✅ 已依本次全部 PO 產生新的 PO GRID：{rep_.get('sheet_count', '?')} 間工廠（每間一個工作表）、"
                   f"{n_items:,} 個品項，產品圖片 {rep_.get('image_count', 0):,} / {n_items:,} 張。")
        if rep_.get('no_image') and rep_.get('images_supplied'):
            with st.expander(f"🖼️ {len(rep_['no_image'])} 個品項找不到對應圖片，PICTURE 欄留空（點擊展開）"):
                st.write(", ".join(rep_['no_image']))
        elif not rep_.get('images_supplied'):
            st.caption("PICTURE 欄目前是空的。要放產品圖，請在左側「步驟 3 › 產品圖片」上傳圖片 zip 或 SPK 檔，"
                       "再按一次「解析 PDF 並執行核對」。")
        if rep_.get('unmapped_dc'):
            st.warning("以下 DC 代碼尚無確認過的目的地，DES PORT 列直接顯示代碼，請人工確認：" +
                       "；".join(f"{dc}（PO {', '.join(pos)}）" for dc, pos in rep_['unmapped_dc'].items()))
        st.download_button("📥 下載 PO GRID", data=gres['data'], file_name=gres['filename'],
                           mime=xlsx_mime, key="dl_grid_main")
    else:
        ins = rep_.get('inserted', [])
        st.success(f"已更新既有 PO GRID：寫入 {len(rep_.get('written', []))} 格數量、新增 {len(ins)} 個 PO 欄；"
                   f"原有圖片 {rep_.get('media_before', 0)} 張{'完整保留' if rep_.get('image_integrity') else '請檢查'}。")
        if ins:
            with st.expander(f"➕ 新增的 PO 欄（{len(ins)}）", expanded=True):
                st.dataframe(pd.DataFrame([{'工作表': i.get('sheet'), 'PO': i.get('po'), '欄': i.get('column'),
                                            '類型': i.get('label'), 'Ship Window': i.get('window'),
                                            '目的地': i.get('dest'), '備註': i.get('note', '')} for i in ins]),
                             hide_index=True, use_container_width=True)
        att = rep_.get('needs_attention', [])
        if att:
            with st.expander(f"⚠️ {len(att)} 項需要人工處理（點擊展開）", expanded=True):
                st.dataframe(pd.DataFrame([{'工作表': x.get('sheet', ''), 'DPCI': x.get('dpci', ''),
                                            'PO': x.get('po') or ', '.join(x.get('pos', [])),
                                            '問題': x.get('issue', ''), '建議': x.get('next_step', '')} for x in att]),
                             hide_index=True, use_container_width=True)
        st.download_button("📥 下載 PO GRID（更新版，保留原檔圖片與手填內容）", data=gres['data'],
                           file_name=gres['filename'], mime=xlsx_mime, key="dl_grid_main")
        if grebuild and grebuild['ok']:
            st.info("既有 GRID 缺少部分品項的列，更新模式無法新增列。另外提供一份完整重建版（不含原檔的手填內容）。")
            st.download_button("📥 下載 PO GRID（完整重建版）", data=grebuild['data'],
                               file_name=grebuild['filename'], mime=xlsx_mime, key="dl_grid_rebuild")


# ── 結果區：放在按鈕區塊外、存在 session_state，按下載或展開明細後畫面不會消失 ──
if st.session_state.get('results'):
    _r = st.session_state['results']
    show_results(_r['merged_df'], 'PDF', run_meta=_r['run_meta'], validation_notes=_r['notes'], summary_ctx=_r['ctx'])
    st.download_button("📥 下載核對報告 (Excel：Validation Summary + 核對結果 + 執行摘要)", data=_r['excel'],
                       file_name=_r['excel_name'],
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key="dl_report")
if st.session_state.get('grid_out'):
    render_grid_section(st.session_state['grid_out'])
