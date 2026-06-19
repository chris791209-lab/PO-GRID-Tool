import streamlit as st
import pandas as pd
import io
import zipfile
import os
import re
import copy
import openpyxl
import xml.etree.ElementTree as ET
import unicodedata  
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Border, Side, Alignment  
from PIL import Image as PILImage

# ==========================================
# 1. 密碼保護機制定義
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

# ==========================================
# 2. 共用常數與輔助函數定義
# ==========================================
PORT_MAP = {
    '581': 'PSW', '3890': 'PNW', '584': 'ORF', '3891': 'SAV',
    '3851': 'NYC', '3850': 'OAK', '3887': 'HOU', '3758': 'CHARLESTON'
}

def resolve_zip_path(base_dir, relative_path):
    if relative_path.startswith('/'): return relative_path[1:]
    parts = [p for p in base_dir.split('/') if p]
    for part in relative_path.split('/'):
        if part == '..':
            if parts: parts.pop()
        elif part and part != '.':
            parts.append(part)
    return '/'.join(parts)

def extract_port_mapping(port_mapping_files):
    auto_port_dict = {}
    if not port_mapping_files: return auto_port_dict
    
    for port_file in port_mapping_files:
        try:
            raw_bytes = port_file.getvalue()
            try: content = raw_bytes.decode("utf-8").splitlines()
            except UnicodeDecodeError: content = raw_bytes.decode("big5", errors="ignore").splitlines()
                
            for line in content:
                line = line.strip()
                if not line: continue
                
                match = re.search(r'\d{3,4}-(\d+)-([A-Za-z0-9]+)', line)
                if match:
                    po_part = str(match.group(1)).strip().lstrip('0') 
                    port_part = str(match.group(2)).strip().lstrip('0').upper()
                    auto_port_dict[po_part] = port_part
                    continue
                    
                clean_line = re.sub(r'(?i)\b(po|port|no|num|number|code|港口|代碼|代號)\b|[:："\'#]', ' ', line)
                tokens = [t.strip() for t in re.split(r'[\s,;\t\-]+', clean_line) if t.strip()]
                
                if len(tokens) >= 2:
                    po_candidates = [t for t in tokens if t.isdigit() and len(t) >= 5]
                    if po_candidates:
                        po_part = str(po_candidates[0]).strip().lstrip('0')
                        port_candidates = [t for t in tokens if t != po_candidates[0]]
                        if port_candidates:
                            auto_port_dict[po_part] = str(port_candidates[0]).strip().lstrip('0').upper()
                    else:
                        auto_port_dict[str(tokens[0]).strip().lstrip('0')] = str(tokens[1]).strip().lstrip('0').upper()
        except: pass
    return auto_port_dict

def format_upc(val):
    if pd.isna(val) or val == '': return ''
    try: 
        v = str(int(float(val)))
        return v.zfill(12) if len(v) < 12 else v
    except: return str(val).strip()

st.set_page_config(page_title="PO GRID & 圖片萃取系統", layout="wide")

# ==========================================
# 3. 系統主程式
# ==========================================
if check_password():
    st.success("✅ 成功登入！歡迎使用 AE 部門專屬工具。")
    st.title("🎯 Target 季節性專案自動化系統")

    tab1, tab3, tab2 = st.tabs(["🎃 舊版引擎 (PO RAW DATA)", "🚀 新版引擎 (Modern PO Visibility)", "🖼️ 圖片自動萃取器"])

    # ------------------------------------------
    # 分頁一：舊版 PO GRID (PO RAW DATA)
    # ------------------------------------------
    with tab1:
        st.markdown("""
        此為 **舊版 PO RAW DATA** 專用通道。請依序上傳檔案。
        💡 支援 **Ship Window (SW) 篩選** 與 **多檔港口對照表自動解析**。
        """)

        col1, col2, col3, col4, col5 = st.columns(5)
        with col1: po_raw_file = st.file_uploader("📁 1. PO RAW DATA", type=['csv'], key="old_po1")
        with col2: po_list_file = st.file_uploader("📁 2. List of PO", type=['csv'], key="old_po2")
        with col3: prod_files = st.file_uploader("📁 3. 產品資料(PCN)", type=['xlsx', 'csv'], accept_multiple_files=True, key="old_pcn")
        with col4: image_zip_files = st.file_uploader("📁 4. 產品圖片包(ZIP)", type=['zip'], accept_multiple_files=True, key="old_zip")
        with col5: port_mapping_files = st.file_uploader("📁 5. 港口對照表", type=['csv', 'txt'], accept_multiple_files=True, key="old_port")

        if po_raw_file and prod_files and po_list_file:
            po_list = pd.read_csv(po_list_file)
            po_raw = pd.read_csv(po_raw_file)
            
            po_list['PO NUMBER'] = po_list['PO NUMBER'].astype(str).str.split('.').str[0].str.strip()
            po_raw['PO NUMBER'] = po_raw['PO NUMBER'].astype(str).str.split('.').str[0].str.strip()
            
            po_list = po_list[(po_list['PO NUMBER'] != 'nan') & (po_list['PO NUMBER'] != '')]
            po_raw = po_raw[(po_raw['PO NUMBER'] != 'nan') & (po_raw['PO NUMBER'] != '')]
            
            po_list['SHIP BEGIN DATE'] = pd.to_datetime(po_list['SHIP BEGIN DATE'], errors='coerce')
            po_list['SHIP END DATE'] = pd.to_datetime(po_list['SHIP END DATE'], errors='coerce')
            
            st.divider()
            st.subheader("📍 步驟 6: 篩選出貨期間 / Ship Window (選填)")
            use_sw_filter = st.checkbox("📅 啟用 SW 範圍篩選", key="old_sw")
            
            can_proceed = True  
            if use_sw_filter:
                sw_range = st.date_input("請選擇範圍", value=[], key="old_sw_date")
                if len(sw_range) == 2:
                    start_dt, end_dt = pd.to_datetime(sw_range[0]), pd.to_datetime(sw_range[1])
                    mask = (po_list['SHIP BEGIN DATE'] <= end_dt) & (po_list['SHIP END DATE'] >= start_dt)
                    valid_pos = po_list[mask]['PO NUMBER'].unique()
                    
                    po_list = po_list[po_list['PO NUMBER'].isin(valid_pos)]
                    po_raw = po_raw[po_raw['PO NUMBER'].isin(valid_pos)]
                    if len(valid_pos) > 0: st.success(f"🔍 篩選完成：保留 {len(valid_pos)} 筆 PO。")
                    else: 
                        st.error("❌ 找不到符合此範圍的訂單。")
                        can_proceed = False
                else:
                    st.info("👈 請選擇起始與結束日。")
                    can_proceed = False

            if can_proceed:
                po_list['SHIP_DATES'] = po_list['SHIP BEGIN DATE'].dt.strftime('%m/%d') + '-' + po_list['SHIP END DATE'].dt.strftime('%m/%d')
                po_info = po_list[['PO NUMBER', 'PURPOSE', 'SHIP_DATES']].drop_duplicates()
                active_pos = po_raw['PO NUMBER'].unique()
                po_info = po_info[po_info['PO NUMBER'].isin(active_pos)].copy()
                po_info['PO_CLEAN'] = po_info['PO NUMBER'].astype(str).str.strip().str.lstrip('0')
                
                auto_port_dict = extract_port_mapping(port_mapping_files)
                po_info['輸入港口代碼 (如:581)'] = po_info['PO_CLEAN'].map(auto_port_dict).fillna("")
                
                missing_ports_count = (po_info['輸入港口代碼 (如:581)'] == "").sum()
                st.divider()
                if missing_ports_count > 0:
                    st.warning(f"⚠️ 注意：有 **{missing_ports_count}** 筆 PO 找不到港口代碼！請在下方手動補齊。")
                    display_cols = ["PO NUMBER", "PURPOSE", "SHIP_DATES", "輸入港口代碼 (如:581)"]
                    edited_po_info = st.data_editor(po_info[display_cols].reset_index(drop=True), use_container_width=True, hide_index=True)
                    po_info['輸入港口代碼 (如:581)'] = edited_po_info['輸入港口代碼 (如:581)']
                else:
                    if port_mapping_files: st.success("🤖 完美！已自動填寫 100% 港口代碼。")
                
                st.divider()
                if st.button("🚀 開始自動生成 PO GRID (舊版引擎)", type="primary", key="btn_old"):
                    with st.spinner("舊版引擎運算與排版美化中，請稍候..."):
                        try:
                            image_dict = {}
                            if image_zip_files:
                                for zip_file_obj in image_zip_files:
                                    with zipfile.ZipFile(zip_file_obj, 'r') as z:
                                        for file_info in z.infolist():
                                            if file_info.filename.startswith('__MACOSX/') or file_info.filename.startswith('.'): continue
                                            if file_info.filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif')):
                                                base_name = os.path.basename(file_info.filename)
                                                clean_dpci = os.path.splitext(base_name)[0].strip().split('_')[0] 
                                                if clean_dpci not in image_dict: image_dict[clean_dpci] = z.read(file_info.filename)

                            prod_data_list = []
                            for p_file in prod_files:
                                df_temp = pd.read_csv(p_file) if p_file.name.lower().endswith('.csv') else pd.read_excel(p_file)
                                prod_data_list.append(df_temp)
                            prod_data = pd.concat(prod_data_list, ignore_index=True)

                            po_processed_records = []
                            parent_dpci_list = set()
                            child_assort_qty_dict = {}
                            parent_info_dict = {}
                            item_info_dict = {} 
                            parent_to_children = {}

                            for idx, row in po_raw.iterrows():
                                dept = str(int(row['DEPARTMENT'])) if pd.notna(row['DEPARTMENT']) else '0'
                                cls = str(int(row['CLASS'])).zfill(2) if pd.notna(row['CLASS']) else '00'
                                itm = str(int(row['ITEM'])).zfill(4) if pd.notna(row['ITEM']) else '0000'
                                dpci = f"{dept}-{cls}-{itm}"
                                try: qty = float(str(row['TOTAL ITEM QTY']).replace(',', ''))
                                except: qty = 0.0
                                desc = str(row['ITEM DESCRIPTION']).strip().upper()
                                po_num = row['PO NUMBER']
                                raw_style = str(row['VENDOR STYLE']).strip() if pd.notna(row['VENDOR STYLE']) else ''
                                raw_upc = str(row['ITEM BAR CODE']).strip() if pd.notna(row['ITEM BAR CODE']) else ''
                                
                                if dpci not in item_info_dict and raw_style: item_info_dict[dpci] = {'style': raw_style, 'upc': raw_upc}
                                
                                if desc.startswith('ASSORTMENT'):
                                    parent_dpci_list.add(dpci)
                                    style_val = raw_style
                                    if style_val and not style_val.upper().startswith('ASSORT'): style_val = f"ASSORTMENT-{style_val}"
                                    parent_info_dict[dpci] = {'style': style_val, 'upc': raw_upc}
                                    if dpci in item_info_dict: item_info_dict[dpci]['style'] = style_val
                                    po_processed_records.append({'PO NUMBER': po_num, 'DPCI_MERGE': dpci, 'QTY': qty, 'IS_PARENT': True})
                                    
                                    c_dept = str(int(row['COMPONENT DEPARTMENT'])) if pd.notna(row['COMPONENT DEPARTMENT']) else '0'
                                    c_cls = str(int(row['COMPONENT CLASS'])).zfill(2) if pd.notna(row['COMPONENT CLASS']) else '00'
                                    c_itm = str(int(row['COMPONENT ITEM'])).zfill(4) if pd.notna(row['COMPONENT ITEM']) else '0000'
                                    c_dpci = f"{c_dept}-{c_cls}-{c_itm}"
                                    c_style = str(row['COMPONENT STYLE']).strip() if 'COMPONENT STYLE' in row and pd.notna(row['COMPONENT STYLE']) else ''
                                    if c_dpci not in item_info_dict and c_style: item_info_dict[c_dpci] = {'style': c_style, 'upc': ''}
                                    try: c_qty = float(str(row['COMPONENT ITEM TOTAL QTY']).replace(',', ''))
                                    except: c_qty = 0.0
                                    try: c_assort = float(str(row['COMPONENT ASSORT QTY']).replace(',', ''))
                                    except: c_assort = 0.0
                                    child_assort_qty_dict[c_dpci] = c_assort
                                    if dpci not in parent_to_children: parent_to_children[dpci] = set()
                                    parent_to_children[dpci].add(c_dpci)
                                    po_processed_records.append({'PO NUMBER': po_num, 'DPCI_MERGE': c_dpci, 'QTY': c_qty, 'IS_PARENT': False})
                                else:
                                    po_processed_records.append({'PO NUMBER': po_num, 'DPCI_MERGE': dpci, 'QTY': qty, 'IS_PARENT': False})

                            po_processed = pd.DataFrame(po_processed_records)
                            parents = po_processed[po_processed['IS_PARENT']].drop_duplicates(subset=['PO NUMBER', 'DPCI_MERGE'])
                            children_and_regular = po_processed[~po_processed['IS_PARENT']]
                            po_processed_unique = pd.concat([parents, children_and_regular], ignore_index=True)

                            po_info['PORT_NAME'] = po_info['輸入港口代碼 (如:581)'].astype(str).str.strip()
                            po_info['PORT_NAME'] = po_info['PORT_NAME'].replace({'': '未指定港口', 'nan': '未指定港口'})
                            
                            po_raw_merged = po_processed_unique.merge(po_info[['PO NUMBER', 'PURPOSE', 'SHIP_DATES', 'PORT_NAME']], on='PO NUMBER', how='left')
                            
                            # 💡 強制轉換為 object，避免 float64 欄位填入空白字串時報錯
                            po_raw_merged['PURPOSE'] = po_raw_merged['PURPOSE'].astype(object).fillna('標籤遺失')
                            po_raw_merged['SHIP_DATES'] = po_raw_merged['SHIP_DATES'].astype(object).fillna('日期遺失')
                            po_raw_merged['PORT_NAME'] = po_raw_merged['PORT_NAME'].astype(object).fillna('未指定港口')
                            
                            pivot_df_temp = po_raw_merged.pivot_table(index='DPCI_MERGE', columns=['PURPOSE', 'PO NUMBER', 'SHIP_DATES', 'PORT_NAME'], values='QTY', aggfunc='sum').fillna(0)
                            
                            new_pivot_cols = [(col[0], '', col[1], col[2], col[3]) for col in pivot_df_temp.columns]
                            pivot_df = pd.DataFrame(pivot_df_temp.values, index=pivot_df_temp.index, columns=pd.MultiIndex.from_tuples(new_pivot_cols))
                            pivot_df[('', 'PO TOTAL', '', '', '')] = pivot_df.sum(axis=1)
                            
                            # 💡 強制轉換 pivot_df 為 object，確保可以填入字串與空白
                            pivot_df = pivot_df.astype(object)
                            
                            for parent_dpci in parent_dpci_list:
                                if parent_dpci in pivot_df.index:
                                    pivot_df.loc[parent_dpci, ('', 'PO TOTAL', '', '', '')] = '' 
                                    for col in pivot_df.columns:
                                        if col[1] !=
