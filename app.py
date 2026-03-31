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
                            po_raw_merged['PURPOSE'] = po_raw_merged['PURPOSE'].fillna('標籤遺失')
                            po_raw_merged['SHIP_DATES'] = po_raw_merged['SHIP_DATES'].fillna('日期遺失')
                            po_raw_merged['PORT_NAME'] = po_raw_merged['PORT_NAME'].fillna('未指定港口')
                            
                            pivot_df_temp = po_raw_merged.pivot_table(index='DPCI_MERGE', columns=['PURPOSE', 'PO NUMBER', 'SHIP_DATES', 'PORT_NAME'], values='QTY', aggfunc='sum').fillna(0)
                            
                            new_pivot_cols = [(col[0], '', col[1], col[2], col[3]) for col in pivot_df_temp.columns]
                            pivot_df = pd.DataFrame(pivot_df_temp.values, index=pivot_df_temp.index, columns=pd.MultiIndex.from_tuples(new_pivot_cols))
                            pivot_df[('', 'PO TOTAL', '', '', '')] = pivot_df.sum(axis=1)
                            
                            for parent_dpci in parent_dpci_list:
                                if parent_dpci in pivot_df.index:
                                    pivot_df.loc[parent_dpci, ('', 'PO TOTAL', '', '', '')] = '' 
                                    for col in pivot_df.columns:
                                        if col[1] != 'PO TOTAL': 
                                            val = pivot_df.loc[parent_dpci, col]
                                            if isinstance(val, (int, float)) and val > 0: pivot_df.loc[parent_dpci, col] = f"{parent_dpci}-({int(val):,})"
                            pivot_df = pivot_df.replace({0: '', 0.0: ''})

                            prod_data['DPCI_MERGE'] = prod_data['DPCI'].astype(str).str.strip()
                            if 'Manufacturer Style # *' not in prod_data.columns: prod_data['Manufacturer Style # *'] = ''
                            if 'Barcode' in prod_data.columns: prod_data['Barcode'] = prod_data['Barcode'].apply(format_upc)
                            else: prod_data['Barcode'] = ''

                            for dpci_key, info_dict in item_info_dict.items():
                                if dpci_key in prod_data['DPCI_MERGE'].values:
                                    idx_list = prod_data.index[prod_data['DPCI_MERGE'] == dpci_key].tolist()
                                    for i in idx_list:
                                        curr_style = str(prod_data.at[i, 'Manufacturer Style # *']).strip()
                                        if curr_style in ('', 'nan') and info_dict['style']: prod_data.at[i, 'Manufacturer Style # *'] = info_dict['style']
                                        if info_dict['upc']:
                                            curr_upc = str(prod_data.at[i, 'Barcode']).strip()
                                            if curr_upc in ('', 'nan'): prod_data.at[i, 'Barcode'] = format_upc(info_dict['upc'])

                            for parent_dpci, info in parent_info_dict.items():
                                vendor_name, factory_name, factory_id = '', '', ''
                                for c_dpci in parent_to_children.get(parent_dpci, []):
                                    if c_dpci in prod_data['DPCI_MERGE'].values:
                                        child_rows = prod_data[prod_data['DPCI_MERGE'] == c_dpci]
                                        vendor_name = child_rows.iloc[0].get('Import Vendor Name', '')
                                        factory_name = child_rows.iloc[0].get('Factory Name', '')
                                        factory_id = child_rows.iloc[0].get('Factory ID', '')
                                        if vendor_name and factory_name: break
                                if parent_dpci in prod_data['DPCI_MERGE'].values:
                                    idx = prod_data.index[prod_data['DPCI_MERGE'] == parent_dpci].tolist()
                                    for i in idx:
                                        prod_data.at[i, 'DPCI'] = parent_dpci 
                                        prod_data.at[i, 'Manufacturer Style # *'] = info['style']
                                        prod_data.at[i, 'Barcode'] = format_upc(info['upc'])
                                        prod_data.at[i, 'Product Description'] = '' 
                                        if vendor_name: prod_data.at[i, 'Import Vendor Name'] = vendor_name
                                        if factory_name: prod_data.at[i, 'Factory Name'] = factory_name
                                        if pd.notna(factory_id): prod_data.at[i, 'Factory ID'] = factory_id
                                else:
                                    new_row = {col: '' for col in prod_data.columns}
                                    new_row['DPCI'] = parent_dpci 
                                    new_row['DPCI_MERGE'] = parent_dpci
                                    new_row['Manufacturer Style # *'] = info['style']
                                    new_row['Barcode'] = format_upc(info['upc'])
                                    new_row['Product Description'] = '' 
                                    new_row['Import Vendor Name'] = vendor_name
                                    new_row['Factory Name'] = factory_name
                                    if pd.notna(factory_id): new_row['Factory ID'] = factory_id
                                    prod_data = pd.concat([prod_data, pd.DataFrame([new_row])], ignore_index=True)
                                    
                            if 'Assortment' not in prod_data.columns: prod_data['Assortment'] = ''
                            for parent_dpci, children in parent_to_children.items():
                                for child_dpci in children:
                                    assort_qty = child_assort_qty_dict.get(child_dpci, 0)
                                    if child_dpci in prod_data['DPCI_MERGE'].values:
                                        idx = prod_data.index[prod_data['DPCI_MERGE'] == child_dpci].tolist()
                                        for i in idx: prod_data.at[i, 'Assortment'] = int(assort_qty) if float(assort_qty).is_integer() else float(assort_qty)

                            if 'Factory Name' not in prod_data.columns: prod_data['Factory Name'] = '未提供工廠名稱'
                            if 'Factory ID' not in prod_data.columns: prod_data['Factory ID'] = ''
                            
                            def make_maker(row):
                                fid = str(row.get('Factory ID', '')).replace('.0', '').strip()
                                fname = str(row.get('Factory Name', '')).strip()
                                return f"{fid}-{fname}" if fid and fid != 'nan' else fname
                            prod_data['Maker'] = prod_data.apply(make_maker, axis=1)

                            prod_data['Packaging'] = prod_data['Retail Packaging Format (1) *'].fillna('') if 'Retail Packaging Format (1) *' in prod_data.columns else ''
                            prod_data['IMAGE'] = ''
                            prod_data['Age'] = ''

                            desired_left_columns = ['DPCI', 'Manufacturer Style # *', 'IMAGE', 'Product Description', 'Barcode', 'Primary Raw Material Type', 'Age', 'Maker', 'Packaging', 'Inner Pack Unit Quantity', 'Case Unit Quantity', 'Assortment', 'Import Vendor Name', 'Factory Name']
                            for col in desired_left_columns:
                                if col not in prod_data.columns: prod_data[col] = ''
                                    
                            left_data = prod_data[desired_left_columns + ['DPCI_MERGE']].drop_duplicates(subset=['DPCI_MERGE']).set_index('DPCI_MERGE')
                            
                            def get_left_tuple(col, idx):
                                spaces = " " * idx 
                                if col == 'DPCI': return ('Program Name', 'DPCI', '', '', '')
                                elif col == 'Barcode': return (spaces, 'UPC', '', '', '')
                                else: return (spaces, col, '', '', '')

                            left_data.columns = pd.MultiIndex.from_tuples([get_left_tuple(col, i+1) for i, col in enumerate(left_data.columns)])
                            final_df = left_data.join(pivot_df, how='inner')

                            zip_buffer = io.BytesIO()
                            
                            calibri_font = Font(name='Calibri', size=11)
                            calibri_bold = Font(name='Calibri', size=11, bold=True)
                            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                            
                            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                                actual_vendor_col = [c for c in final_df.columns if c[1] == 'Import Vendor Name'][0]
                                actual_factory_col = [c for c in final_df.columns if c[1] == 'Factory Name'][0]
                                final_df[actual_vendor_col] = final_df[actual_vendor_col].replace('', '未指定供應商')
                                final_df[actual_factory_col] = final_df[actual_factory_col].replace('', '未指定工廠')
                                
                                excel_buffer = io.BytesIO()
                                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                    grouped_factory = final_df.groupby(actual_factory_col)
                                    for factory_name, factory_data in grouped_factory:
                                        safe_factory_name = str(factory_name).replace('/', '_').replace('\\', '_').replace('[', '').replace(']', '').replace('*', '').replace(':', '').replace('?', '')[:31]
                                        export_data = factory_data.drop(columns=[actual_vendor_col, actual_factory_col])
                                        cols_to_keep = []
                                        for col in export_data.columns:
                                            if col[2] != '': 
                                                keep = any(isinstance(val, (int, float)) and val > 0 or isinstance(val, str) and val not in ('', '0', '0.0') for val in export_data[col])
                                                if keep: cols_to_keep.append(col)
                                            else: cols_to_keep.append(col)
                                                
                                        export_data = export_data[cols_to_keep]
                                        available_dpcis = export_data.index.tolist()
                                        ordered_dfs = []
                                        added_dpcis = set()
                                        
                                        for p_dpci in parent_dpci_list:
                                            if p_dpci in available_dpcis:
                                                ordered_dfs.append(export_data.loc[[p_dpci]])
                                                added_dpcis.add(p_dpci)
                                                for c_dpci in parent_to_children.get(p_dpci, set()):
                                                    if c_dpci in available_dpcis:
                                                        ordered_dfs.append(export_data.loc[[c_dpci]])
                                                        added_dpcis.add(c_dpci)
                                                blank = pd.DataFrame([[''] * len(export_data.columns)], columns=export_data.columns, index=[f'BLANK_{p_dpci}'])
                                                ordered_dfs.append(blank)
                                        
                                        regular_dpcis = [d for d in available_dpcis if d not in added_dpcis]
                                        if regular_dpcis: ordered_dfs.append(export_data.loc[regular_dpcis])
                                        if ordered_dfs: export_data = pd.concat(ordered_dfs)
                                        
                                        unmerged_columns = []
                                        po_idx = 0
                                        for col in export_data.columns:
                                            if col[2] != '': 
                                                new_purpose = str(col[0]) + (" " * po_idx) 
                                                unmerged_columns.append((new_purpose, col[1], col[2], col[3], col[4]))
                                                po_idx += 1
                                            else: unmerged_columns.append(col)
                                                
                                        export_data.columns = pd.MultiIndex.from_tuples(unmerged_columns)
                                        export_data_reset = export_data.reset_index(drop=True)
                                        export_data_reset.to_excel(writer, index=True, sheet_name=safe_factory_name)
                                        ws = writer.sheets[safe_factory_name]
                                        ws.delete_cols(1) 
                                        
                                        for row in ws.iter_rows():
                                            for cell in row:
                                                cell.border = thin_border
                                                if cell.row <= 5:  
                                                    cell.font = calibri_bold
                                                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                                                else:
                                                    cell.font = calibri_font
                                                    cell.alignment = Alignment(vertical='center')
                                                    if cell.column > 14 and isinstance(cell.value, (int, float)):
                                                        cell.number_format = '#,##0'
                                        
                                        img_col_letter = None
                                        if image_zip_files:
                                            image_dict = {}
                                            for zip_file_obj in image_zip_files:
                                                with zipfile.ZipFile(zip_file_obj, 'r') as z:
                                                    for file_info in z.infolist():
                                                        if file_info.filename.startswith('__MACOSX/') or file_info.filename.startswith('.'): continue
                                                        if file_info.filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif')):
                                                            base_name = os.path.basename(file_info.filename)
                                                            clean_dpci = os.path.splitext(base_name)[0].strip().split('_')[0] 
                                                            if clean_dpci not in image_dict: image_dict[clean_dpci] = z.read(file_info.filename)
                                            
                                            dpci_col_idx, img_col_idx = None, None
                                            for idx, col in enumerate(export_data_reset.columns):
                                                if col[1] == 'DPCI': dpci_col_idx = idx + 1
                                                if col[1] == 'IMAGE': img_col_idx = idx + 1
                                            if dpci_col_idx and img_col_idx:
                                                img_col_letter = get_column_letter(img_col_idx)
                                                ws.column_dimensions[img_col_letter].width = 13 
                                                for r_idx in range(6, ws.max_row + 1):
                                                    cell_dpci_val = str(ws.cell(row=r_idx, column=dpci_col_idx).value).strip()
                                                    if cell_dpci_val in image_dict:
                                                        try:
                                                            img_bytes = io.BytesIO(image_dict[cell_dpci_val])
                                                            with PILImage.open(img_bytes) as pil_img:
                                                                if pil_img.mode != 'RGB': pil_img = pil_img.convert('RGB')
                                                                clean_img_io = io.BytesIO()
                                                                pil_img.save(clean_img_io, format='JPEG')
                                                                clean_img_io.seek(0)
                                                            img_obj = OpenpyxlImage(clean_img_io)
                                                            img_obj.width = 90
                                                            img_obj.height = 90
                                                            ws.add_image(img_obj, f"{img_col_letter}{r_idx}")
                                                            ws.row_dimensions[r_idx].height = 70 
                                                        except: pass 
                                                        
                                        for col_idx in range(1, ws.max_column + 1):
                                            col_letter = get_column_letter(col_idx)
                                            if img_col_letter and col_letter == img_col_letter: continue  
                                            max_length = 0
                                            for row_idx in range(1, ws.max_row + 1):
                                                cell = ws.cell(row=row_idx, column=col_idx)
                                                if type(cell).__name__ != 'MergedCell' and cell.value is not None:
                                                    cell_val_str = str(cell.value)
                                                    for line in cell_val_str.split('\n'):
                                                        line_len = sum(2 if unicodedata.east_asian_width(c) in 'FWA' else 1 for c in line)
                                                        if line_len > max_length: max_length = line_len
                                            ws.column_dimensions[col_letter].width = max(8, min(max_length + 2, 50)) 
                                
                                zip_file.writestr("PO_GRID_Merged_Old.xlsx", excel_buffer.getvalue())
                            
                            st.success("✨ 處理完成！已為您產出舊版 PO GRID 表格。")
                            st.download_button(
                                label="📦 點擊下載合併版 PO GRID (ZIP)",
                                data=zip_buffer.getvalue(),
                                file_name="PO_GRIDs_Output_Old.zip",
                                mime="application/zip"
                            )
                        except Exception as e:
                            st.error(f"❌ 舊版處理過程中發生錯誤: {e}")

    # ------------------------------------------
    # 分頁二：新版 PO GRID (Modern PO Visibility)
    # ------------------------------------------
    with tab3:
        st.markdown("""
        此為 **全新 Modern PO** 專屬通道！
        🎯 **智慧偵測**：請將 `PO Level`, `Item Level`, `DC_Item Level` 這 3 份 Modern CSV **同時上傳**至第一個框框，系統會自動在背後為您縫合！
        💡 **混裝救星**：如果該專案有混裝商品(Assortment)，只要上傳【1.5 混裝明細表】，系統就會自動展開子商品並完美還原舊版計算邏輯！
        """)

        col1, col15, col2, col3 = st.columns(4)
        with col1:
            modern_po_files = st.file_uploader("📁 1. Modern PO 報表\n(請一次上傳 3 份 CSV)", type=['csv'], accept_multiple_files=True, key="m_po")
        with col15:
            m_assort_files = st.file_uploader("📁 1.5 混裝明細表\n(選傳，有混裝必傳)", type=['csv', 'xlsx'], accept_multiple_files=True, key="m_asst")
        with col2:
            m_prod_files = st.file_uploader("📁 2. 產品資料(PCN)\n(支援多檔)", type=['xlsx', 'csv'], accept_multiple_files=True, key="m_pcn")
        with col3:
            m_image_zip_files = st.file_uploader("📁 3. 產品圖片包(ZIP)\n(支援多檔)", type=['zip'], accept_multiple_files=True, key="m_zip")

        if modern_po_files and m_prod_files:
            df_po_level, df_item_level, df_dc_level = None, None, None
            
            for f in modern_po_files:
                try:
                    df_temp = pd.read_csv(f, dtype=str, nrows=5)
                    cols = df_temp.columns.tolist()
                    f.seek(0)
                    if 'PO PURPOSE' in cols or 'PO CREATE TYPE' in cols:
                        df_po_level = pd.read_csv(f, dtype=str)
                    elif 'MANUFACTURER STYLE' in cols:
                        df_item_level = pd.read_csv(f, dtype=str)
                    elif 'LOCATION' in cols and 'DPCI' in cols: 
                        df_dc_level = pd.read_csv(f, dtype=str)
                except Exception as e:
                    st.warning(f"檔案讀取失敗 {f.name}: {e}")

            if df_po_level is None or df_item_level is None or df_dc_level is None:
                st.error("❌ 系統未能集齊 3 份必要的 Modern 報表。請確認上傳了 PO Level、Item Level 與 DC_Item Level！")
            else:
                df_po_level['PO NUMBER'] = df_po_level['PO #'].astype(str).str.split('.').str[0].str.strip()
                df_item_level['PO NUMBER'] = df_item_level['PO #'].astype(str).str.split('.').str[0].str.strip()
                df_dc_level['PO NUMBER'] = df_dc_level['PO #'].astype(str).str.split('.').str[0].str.strip()

                df_po_level = df_po_level[(df_po_level['PO NUMBER'] != 'nan') & (df_po_level['PO NUMBER'] != '')]
                df_item_level = df_item_level[(df_item_level['PO NUMBER'] != 'nan') & (df_item_level['PO NUMBER'] != '')]
                df_dc_level = df_dc_level[(df_dc_level['PO NUMBER'] != 'nan') & (df_dc_level['PO NUMBER'] != '')]

                df_po_level['SHIP BEGIN DATE'] = pd.to_datetime(df_po_level['ORIG SHIP BEGIN'], errors='coerce')
                df_po_level['SHIP END DATE'] = pd.to_datetime(df_po_level['ORIG SHIP END'], errors='coerce')

                st.divider()
                st.subheader("📍 步驟 4: 篩選出貨期間 / Ship Window (選填)")
                m_use_sw_filter = st.checkbox("📅 啟用 SW 範圍篩選", key="m_sw")
                
                can_proceed_m = True
                if m_use_sw_filter:
                    sw_range_m = st.date_input("請選擇範圍", value=[], key="m_sw_date")
                    if len(sw_range_m) == 2:
                        start_dt, end_dt = pd.to_datetime(sw_range_m[0]), pd.to_datetime(sw_range_m[1])
                        mask = (df_po_level['SHIP BEGIN DATE'] <= end_dt) & (df_po_level['SHIP END DATE'] >= start_dt)
                        valid_pos = df_po_level[mask]['PO NUMBER'].unique()
                        
                        df_po_level = df_po_level[df_po_level['PO NUMBER'].isin(valid_pos)]
                        df_item_level = df_item_level[df_item_level['PO NUMBER'].isin(valid_pos)]
                        df_dc_level = df_dc_level[df_dc_level['PO NUMBER'].isin(valid_pos)]
                        
                        if len(valid_pos) > 0: st.success(f"🔍 篩選完成：保留 {len(valid_pos)} 筆 PO。")
                        else: 
                            st.error("❌ 找不到符合此範圍的訂單。")
                            can_proceed_m = False
                    else:
                        st.info("👈 請選擇起始與結束日。")
                        can_proceed_m = False

                if can_proceed_m:
                    df_po_level['SHIP_DATES'] = df_po_level['SHIP BEGIN DATE'].dt.strftime('%m/%d') + '-' + df_po_level['SHIP END DATE'].dt.strftime('%m/%d')
                    purp_col = 'PO PURPOSE' if 'PO PURPOSE' in df_po_level.columns else 'PURPOSE' if 'PURPOSE' in df_po_level.columns else None
                    if purp_col: po_info = df_po_level[['PO NUMBER', purp_col, 'SHIP_DATES']].copy().rename(columns={purp_col: 'PURPOSE'})
                    else: po_info = df_po_level[['PO NUMBER', 'SHIP_DATES']].copy().assign(PURPOSE='')
                    po_info.drop_duplicates(inplace=True)

                    item_info_dict = {}
                    parent_dpci_list = set()
                    df_item_level['DPCI_MERGE'] = df_item_level['DPCI'].astype(str).str.strip()
                    for _, row in df_item_level.iterrows():
                        dpci = str(row['DPCI_MERGE'])
                        style = str(row.get('MANUFACTURER STYLE', '')).strip()
                        upc = str(row.get('UPC', '')).split('.')[0].strip()
                        desc = str(row.get('ITEM DESCRIPTION', '')).strip().upper()
                        if desc.startswith('ASSORT'): parent_dpci_list.add(dpci)
                        if style == 'nan': style = ''
                        if upc == 'nan': upc = ''
                        if dpci not in item_info_dict: item_info_dict[dpci] = {'style': style, 'upc': upc}
                        else:
                            if style and not item_info_dict[dpci]['style']: item_info_dict[dpci]['style'] = style
                            if upc and not item_info_dict[dpci]['upc']: item_info_dict[dpci]['upc'] = upc

                    df_dc_level['DPCI_MERGE'] = df_dc_level['DPCI'].astype(str).str.strip()
                    df_dc_level['QTY'] = pd.to_numeric(df_dc_level['REVISED QUANTITY'].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0)
                    df_dc_level['NATIVE_PORT'] = df_dc_level['LOCATION'].astype(str).str.replace(r'\.0$', '', regex=True).replace({'nan': '', 'None': ''}).str.strip()
                    
                    unique_po_ports = df_dc_level[['PO NUMBER', 'NATIVE_PORT']].drop_duplicates(subset=['PO NUMBER']).copy()
                    unique_po_ports['輸入港口代碼 (如:581)'] = unique_po_ports['NATIVE_PORT']
                    
                    missing_ports_count_m = (unique_po_ports['輸入港口代碼 (如:581)'] == "").sum()
                    
                    st.divider()
                    st.subheader("📍 步驟 5: 最終港口確認")
                    if missing_ports_count_m > 0:
                        st.warning(f"⚠️ 注意：有 **{missing_ports_count_m}** 筆 PO 沒有標示港口 (LOCATION 空白)！請在下方手動補齊。")
                        display_cols = ["PO NUMBER", "輸入港口代碼 (如:581)"]
                        edited_unique = st.data_editor(unique_po_ports[display_cols].reset_index(drop=True), use_container_width=True, hide_index=True)
                        unique_po_ports['輸入港口代碼 (如:581)'] = edited_unique['輸入港口代碼 (如:581)'].values
                    else:
                        st.success("🤖 完美！系統已成功從 Modern PO 的 LOCATION 欄位直接載入 100% 的港口代碼。")

                    port_lookup = dict(zip(unique_po_ports['PO NUMBER'], unique_po_ports['輸入港口代碼 (如:581)']))
                    df_dc_level['PORT_NAME'] = df_dc_level['PO NUMBER'].map(port_lookup).fillna('未指定港口').replace({'': '未指定港口'})

                    st.divider()
                    if st.button("🚀 開始自動生成 PO GRID (新版引擎)", type="primary", key="btn_new"):
                        with st.spinner("新版引擎運算與排版美化中，請稍候..."):
                            try:
                                assort_dict = {}
                                if m_assort_files:
                                    for asst_file in m_assort_files:
                                        try:
                                            df_asst = pd.read_csv(asst_file) if asst_file.name.lower().endswith('.csv') else pd.read_excel(asst_file)
                                            header_idx = None
                                            for i in range(min(20, len(df_asst))):
                                                row_strs = [str(x).lower() for x in df_asst.iloc[i].values]
                                                if any('assortment dpci' in x or 'parent' in x for x in row_strs) and any('dpci' in x for x in row_strs):
                                                    header_idx = i
                                                    break
                                            if header_idx is not None:
                                                df_asst.columns = df_asst.iloc[header_idx]
                                                df_asst = df_asst.iloc[header_idx+1:].reset_index(drop=True)
                                                
                                            parent_col = next((c for c in df_asst.columns if 'assortment dpci' in str(c).lower() or 'parent' in str(c).lower()), None)
                                            child_col = next((c for c in df_asst.columns if str(c).lower() == 'dpci' or 'component dpci' in str(c).lower() or 'child dpci' in str(c).lower() or 'item dpci' in str(c).lower()), None)
                                            qty_col = next((c for c in df_asst.columns if 'units' in str(c).lower() or 'qty' in str(c).lower() or 'ratio' in str(c).lower() or 'pack' in str(c).lower()), None)
                                            
                                            if parent_col and child_col and qty_col:
                                                df_asst[parent_col] = df_asst[parent_col].ffill() 
                                                for _, row in df_asst.iterrows():
                                                    p_val = str(row[parent_col]).strip().replace('.0', '')
                                                    c_val = str(row[child_col]).strip()
                                                    q_val = pd.to_numeric(str(row[qty_col]), errors='coerce')
                                                    
                                                    if pd.isna(q_val) or not p_val or not c_val or p_val == 'nan' or c_val == 'nan': continue
                                                    
                                                    p_val = re.sub(r'\D', '', p_val)
                                                    if len(p_val) == 9: p_val = f"{p_val[:3]}-{p_val[3:5]}-{p_val[5:]}"
                                                    
                                                    c_val_match = re.search(r'\d{3}-\d{2}-\d{4}', c_val)
                                                    if c_val_match: c_val = c_val_match.group(0)
                                                    else:
                                                        c_val_num = re.sub(r'\D', '', c_val)
                                                        if len(c_val_num) == 9: c_val = f"{c_val_num[:3]}-{c_val_num[3:5]}-{c_val_num[5:]}"
                                                    
                                                    if p_val not in assort_dict: assort_dict[p_val] = []
                                                    assort_dict[p_val].append({'child_dpci': c_val, 'qty': float(q_val)})
                                        except Exception as e:
                                            st.warning(f"讀取混裝明細表失敗: {e}")

                                expanded_records = []
                                child_assort_qty_dict = {}
                                parent_to_children = {}
                                
                                for _, row in df_dc_level.iterrows():
                                    po_num = row['PO NUMBER']
                                    dpci = row['DPCI_MERGE']
                                    qty = row['QTY']
                                    port = row['PORT_NAME']
                                    
                                    is_parent = dpci in assort_dict or str(row.get('ITEM DESCRIPTION', '')).upper().startswith('ASSORT')
                                    if is_parent: parent_dpci_list.add(dpci)
                                    
                                    expanded_records.append({'PO NUMBER': po_num, 'DPCI_MERGE': dpci, 'PORT_NAME': port, 'QTY': qty, 'IS_PARENT': is_parent})
                                    
                                    if dpci in assort_dict:
                                        if dpci not in parent_to_children: parent_to_children[dpci] = set()
                                        for child in assort_dict[dpci]:
                                            c_dpci = child['child_dpci']
                                            c_ratio = child['qty']
                                            c_qty = qty * c_ratio  
                                            
                                            parent_to_children[dpci].add(c_dpci)
                                            child_assort_qty_dict[c_dpci] = c_ratio
                                            
                                            expanded_records.append({'PO NUMBER': po_num, 'DPCI_MERGE': c_dpci, 'PORT_NAME': port, 'QTY': c_qty, 'IS_PARENT': False})

                                df_expanded = pd.DataFrame(expanded_records)
                                po_raw_merged = df_expanded.merge(po_info, on='PO NUMBER', how='left')
                                po_raw_merged['PURPOSE'] = po_raw_merged['PURPOSE'].fillna('')
                                po_raw_merged['SHIP_DATES'] = po_raw_merged['SHIP_DATES'].fillna('日期遺失')

                                pivot_df_temp = po_raw_merged.pivot_table(index='DPCI_MERGE', columns=['PURPOSE', 'PO NUMBER', 'SHIP_DATES', 'PORT_NAME'], values='QTY', aggfunc='sum').fillna(0)
                                new_pivot_cols = [(col[0], '', col[1], col[2], col[3]) for col in pivot_df_temp.columns]
                                pivot_df = pd.DataFrame(pivot_df_temp.values, index=pivot_df_temp.index, columns=pd.MultiIndex.from_tuples(new_pivot_cols))
                                pivot_df[('', 'PO TOTAL', '', '', '')] = pivot_df.sum(axis=1)
                                
                                for parent_dpci in parent_dpci_list:
                                    if parent_dpci in pivot_df.index:
                                        pivot_df.loc[parent_dpci, ('', 'PO TOTAL', '', '', '')] = '' 
                                        for col in pivot_df.columns:
                                            if col[1] != 'PO TOTAL': 
                                                val = pivot_df.loc[parent_dpci, col]
                                                if isinstance(val, (int, float)) and val > 0: pivot_df.loc[parent_dpci, col] = f"{parent_dpci}-({int(val):,})"
                                pivot_df = pivot_df.replace({0: '', 0.0: ''})

                                prod_data_list = []
                                for p_file in m_prod_files:
                                    df_temp = pd.read_csv(p_file) if p_file.name.lower().endswith('.csv') else pd.read_excel(p_file)
                                    prod_data_list.append(df_temp)
                                prod_data = pd.concat(prod_data_list, ignore_index=True)
                                
                                if 'DPCI' not in prod_data.columns:
                                    st.error("❌ 產品資料(PCN) 缺少 'DPCI' 欄位。")
                                    st.stop()
                                    
                                prod_data['DPCI_MERGE'] = prod_data['DPCI'].astype(str).str.strip()
                                if 'Manufacturer Style # *' not in prod_data.columns: prod_data['Manufacturer Style # *'] = ''
                                if 'Barcode' in prod_data.columns: prod_data['Barcode'] = prod_data['Barcode'].apply(format_upc)
                                else: prod_data['Barcode'] = ''

                                for dpci_key, info_dict in item_info_dict.items():
                                    if dpci_key in prod_data['DPCI_MERGE'].values:
                                        idx_list = prod_data.index[prod_data['DPCI_MERGE'] == dpci_key].tolist()
                                        for i in idx_list:
                                            curr_style = str(prod_data.at[i, 'Manufacturer Style # *']).strip()
                                            if curr_style in ('', 'nan') and info_dict['style']: prod_data.at[i, 'Manufacturer Style # *'] = info_dict['style']
                                            if info_dict['upc']:
                                                curr_upc = str(prod_data.at[i, 'Barcode']).strip()
                                                if curr_upc in ('', 'nan'): prod_data.at[i, 'Barcode'] = format_upc(info_dict['upc'])

                                existing_pcn_dpcis = prod_data['DPCI_MERGE'].values
                                missing_dpcis = [d for d in pivot_df.index if d not in existing_pcn_dpcis]
                                new_rows = []
                                for miss_dpci in missing_dpcis:
                                    info = item_info_dict.get(miss_dpci, {'style': '', 'upc': ''})
                                    vendor_name = ""
                                    for p_dpci, children in parent_to_children.items():
                                        if miss_dpci in children:
                                            match_p = df_item_level[df_item_level['DPCI_MERGE'] == p_dpci]
                                            if not match_p.empty: vendor_name = str(match_p.iloc[0].get('VENDOR NAME', '')).strip()
                                            break
                                    if not vendor_name:
                                        match_item = df_item_level[df_item_level['DPCI_MERGE'] == miss_dpci]
                                        if not match_item.empty: vendor_name = str(match_item.iloc[0].get('VENDOR NAME', '')).strip()
                                        
                                    new_row = {c: '' for c in prod_data.columns}
                                    new_row['DPCI'] = miss_dpci
                                    new_row['DPCI_MERGE'] = miss_dpci
                                    new_row['Manufacturer Style # *'] = info['style']
                                    new_row['Barcode'] = format_upc(info['upc'])
                                    new_row['Import Vendor Name'] = vendor_name
                                    new_row['Factory Name'] = vendor_name 
                                    new_rows.append(new_row)
                                if new_rows: prod_data = pd.concat([prod_data, pd.DataFrame(new_rows)], ignore_index=True)

                                # 💡 【關鍵修復】母商品強制繼承子商品的工廠與供應商資訊
                                for p_dpci, children in parent_to_children.items():
                                    vendor_name, factory_name, factory_id = '', '', ''
                                    for c_dpci in children:
                                        if c_dpci in prod_data['DPCI_MERGE'].values:
                                            c_rows = prod_data[prod_data['DPCI_MERGE'] == c_dpci]
                                            vendor_name = c_rows.iloc[0].get('Import Vendor Name', '')
                                            factory_name = c_rows.iloc[0].get('Factory Name', '')
                                            factory_id = c_rows.iloc[0].get('Factory ID', '')
                                            if vendor_name and factory_name: break
                                            
                                    if vendor_name or factory_name:
                                        if p_dpci in prod_data['DPCI_MERGE'].values:
                                            idx = prod_data.index[prod_data['DPCI_MERGE'] == p_dpci].tolist()
                                            for i in idx:
                                                if vendor_name: prod_data.at[i, 'Import Vendor Name'] = vendor_name
                                                if factory_name: prod_data.at[i, 'Factory Name'] = factory_name
                                                if 'Factory ID' in prod_data.columns and pd.notna(factory_id): prod_data.at[i, 'Factory ID'] = factory_id

                                if 'Factory Name' not in prod_data.columns: prod_data['Factory Name'] = '未提供工廠名稱'
                                if 'Factory ID' not in prod_data.columns: prod_data['Factory ID'] = ''
                                def make_maker(row):
                                    fid = str(row.get('Factory ID', '')).replace('.0', '').strip()
                                    fname = str(row.get('Factory Name', '')).strip()
                                    return f"{fid}-{fname}" if fid and fid != 'nan' else fname
                                prod_data['Maker'] = prod_data.apply(make_maker, axis=1)

                                prod_data['Packaging'] = prod_data['Retail Packaging Format (1) *'].fillna('') if 'Retail Packaging Format (1) *' in prod_data.columns else ''
                                prod_data['Assortment'] = '' 
                                for parent_dpci, children in parent_to_children.items():
                                    for child_dpci in children:
                                        assort_qty = child_assort_qty_dict.get(child_dpci, 0)
                                        if child_dpci in prod_data['DPCI_MERGE'].values:
                                            idx = prod_data.index[prod_data['DPCI_MERGE'] == child_dpci].tolist()
                                            for i in idx: prod_data.at[i, 'Assortment'] = int(assort_qty) if float(assort_qty).is_integer() else float(assort_qty)
                                            
                                prod_data['IMAGE'] = ''
                                prod_data['Age'] = ''

                                desired_left_columns = ['DPCI', 'Manufacturer Style # *', 'IMAGE', 'Product Description', 'Barcode', 'Primary Raw Material Type', 'Age', 'Maker', 'Packaging', 'Inner Pack Unit Quantity', 'Case Unit Quantity', 'Assortment', 'Import Vendor Name', 'Factory Name']
                                for col in desired_left_columns:
                                    if col not in prod_data.columns: prod_data[col] = ''
                                        
                                left_data = prod_data[desired_left_columns + ['DPCI_MERGE']].drop_duplicates(subset=['DPCI_MERGE']).set_index('DPCI_MERGE')
                                def get_left_tuple(col, idx):
                                    spaces = " " * idx 
                                    if col == 'DPCI': return ('Program Name', 'DPCI', '', '', '')
                                    elif col == 'Barcode': return (spaces, 'UPC', '', '', '')
                                    else: return (spaces, col, '', '', '')

                                left_data.columns = pd.MultiIndex.from_tuples([get_left_tuple(col, i+1) for i, col in enumerate(left_data.columns)])
                                final_df = left_data.join(pivot_df, how='inner')

                                zip_buffer = io.BytesIO()
                                
                                calibri_font = Font(name='Calibri', size=11)
                                calibri_bold = Font(name='Calibri', size=11, bold=True)
                                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                                
                                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                                    actual_vendor_col = [c for c in final_df.columns if c[1] == 'Import Vendor Name'][0]
                                    actual_factory_col = [c for c in final_df.columns if c[1] == 'Factory Name'][0]
                                    final_df[actual_vendor_col] = final_df[actual_vendor_col].replace('', '未指定供應商')
                                    final_df[actual_factory_col] = final_df[actual_factory_col].replace('', '未指定工廠')
                                    
                                    excel_buffer = io.BytesIO()
                                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                        grouped_factory = final_df.groupby(actual_factory_col)
                                        for factory_name, factory_data in grouped_factory:
                                            safe_factory_name = str(factory_name).replace('/', '_').replace('\\', '_').replace('[', '').replace(']', '').replace('*', '').replace(':', '').replace('?', '')[:31]
                                            export_data = factory_data.drop(columns=[actual_vendor_col, actual_factory_col])
                                            cols_to_keep = []
                                            for col in export_data.columns:
                                                if col[2] != '': 
                                                    keep = any(isinstance(val, (int, float)) and val > 0 or isinstance(val, str) and val not in ('', '0', '0.0') for val in export_data[col])
                                                    if keep: cols_to_keep.append(col)
                                                else: cols_to_keep.append(col)
                                                    
                                            export_data = export_data[cols_to_keep]
                                            available_dpcis = export_data.index.tolist()
                                            ordered_dfs = []
                                            added_dpcis = set()
                                            
                                            for p_dpci in parent_dpci_list:
                                                if p_dpci in available_dpcis:
                                                    ordered_dfs.append(export_data.loc[[p_dpci]])
                                                    added_dpcis.add(p_dpci)
                                                    for c_dpci in parent_to_children.get(p_dpci, set()):
                                                        if c_dpci in available_dpcis:
                                                            ordered_dfs.append(export_data.loc[[c_dpci]])
                                                            added_dpcis.add(c_dpci)
                                                    blank = pd.DataFrame([[''] * len(export_data.columns)], columns=export_data.columns, index=[f'BLANK_{p_dpci}'])
                                                    ordered_dfs.append(blank)
                                            
                                            regular_dpcis = [d for d in available_dpcis if d not in added_dpcis]
                                            if regular_dpcis: ordered_dfs.append(export_data.loc[regular_dpcis])
                                            if ordered_dfs: export_data = pd.concat(ordered_dfs)
                                            
                                            unmerged_columns = []
                                            po_idx = 0
                                            for col in export_data.columns:
                                                if col[2] != '': 
                                                    new_purpose = str(col[0]) + (" " * po_idx) 
                                                    unmerged_columns.append((new_purpose, col[1], col[2], col[3], col[4]))
                                                    po_idx += 1
                                                else: unmerged_columns.append(col)
                                                    
                                            export_data.columns = pd.MultiIndex.from_tuples(unmerged_columns)
                                            export_data_reset = export_data.reset_index(drop=True)
                                            export_data_reset.to_excel(writer, index=True, sheet_name=safe_factory_name)
                                            ws = writer.sheets[safe_factory_name]
                                            ws.delete_cols(1) 
                                            
                                            for row in ws.iter_rows():
                                                for cell in row:
                                                    cell.border = thin_border
                                                    if cell.row <= 5:  
                                                        cell.font = calibri_bold
                                                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                                                    else:
                                                        cell.font = calibri_font
                                                        cell.alignment = Alignment(vertical='center')
                                                        if cell.column > 14 and isinstance(cell.value, (int, float)):
                                                            cell.number_format = '#,##0'
                                            
                                            img_col_letter = None
                                            if m_image_zip_files:
                                                image_dict = {}
                                                for zip_file_obj in m_image_zip_files:
                                                    with zipfile.ZipFile(zip_file_obj, 'r') as z:
                                                        for file_info in z.infolist():
                                                            if file_info.filename.startswith('__MACOSX/') or file_info.filename.startswith('.'): continue
                                                            if file_info.filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif')):
                                                                base_name = os.path.basename(file_info.filename)
                                                                clean_dpci = os.path.splitext(base_name)[0].strip().split('_')[0] 
                                                                if clean_dpci not in image_dict: image_dict[clean_dpci] = z.read(file_info.filename)
                                                
                                                dpci_col_idx, img_col_idx = None, None
                                                for idx, col in enumerate(export_data_reset.columns):
                                                    if col[1] == 'DPCI': dpci_col_idx = idx + 1
                                                    if col[1] == 'IMAGE': img_col_idx = idx + 1
                                                if dpci_col_idx and img_col_idx:
                                                    img_col_letter = get_column_letter(img_col_idx)
                                                    ws.column_dimensions[img_col_letter].width = 13 
                                                    for r_idx in range(6, ws.max_row + 1):
                                                        cell_dpci_val = str(ws.cell(row=r_idx, column=dpci_col_idx).value).strip()
                                                        if cell_dpci_val in image_dict:
                                                            try:
                                                                img_bytes = io.BytesIO(image_dict[cell_dpci_val])
                                                                with PILImage.open(img_bytes) as pil_img:
                                                                    if pil_img.mode != 'RGB': pil_img = pil_img.convert('RGB')
                                                                    clean_img_io = io.BytesIO()
                                                                    pil_img.save(clean_img_io, format='JPEG')
                                                                    clean_img_io.seek(0)
                                                                img_obj = OpenpyxlImage(clean_img_io)
                                                                img_obj.width = 90
                                                                img_obj.height = 90
                                                                ws.add_image(img_obj, f"{img_col_letter}{r_idx}")
                                                                ws.row_dimensions[r_idx].height = 70 
                                                            except: pass 
                                                            
                                            for col_idx in range(1, ws.max_column + 1):
                                                col_letter = get_column_letter(col_idx)
                                                if img_col_letter and col_letter == img_col_letter: continue  
                                                max_length = 0
                                                for row_idx in range(1, ws.max_row + 1):
                                                    cell = ws.cell(row=row_idx, column=col_idx)
                                                    if type(cell).__name__ != 'MergedCell' and cell.value is not None:
                                                        cell_val_str = str(cell.value)
                                                        for line in cell_val_str.split('\n'):
                                                            line_len = sum(2 if unicodedata.east_asian_width(c) in 'FWA' else 1 for c in line)
                                                            if line_len > max_length: max_length = line_len
                                                ws.column_dimensions[col_letter].width = max(8, min(max_length + 2, 50)) 
                                    
                                    zip_file.writestr("PO_GRID_Merged_Modern.xlsx", excel_buffer.getvalue())
                                
                                st.success("✨ 處理完成！已為您產出新版 PO GRID 表格。")
                                st.download_button(
                                    label="📦 點擊下載合併版 PO GRID (ZIP)",
                                    data=zip_buffer.getvalue(),
                                    file_name="PO_GRIDs_Output_Modern.zip",
                                    mime="application/zip"
                                )
                            except Exception as e:
                                st.error(f"❌ 新版處理過程中發生錯誤: {e}")

    # ------------------------------------------
    # 分頁三：Program Sheet 圖片自動萃取與命名工具
    # ------------------------------------------
    with tab2:
        st.markdown("""
        ### 🪄 圖片自動命名法寶 (無差別抓取版)
        此工具繞過了一般程式對「圖片群組化」的盲區，直接潛入 Excel 底層，將 **100% 所有的實體圖片** 硬抓出來。
        """)
        
        ps_file = st.file_uploader("📁 上傳 Program Sheet (包含圖片的 .xlsx)", type=['xlsx'], key="ps_uploader")
        
        if ps_file and st.button("🪄 開始自動萃取並命名圖片", type="primary"):
            with st.spinner("🕵️‍♂️ 正在深入 Excel 底層架構暴力抓取所有圖片..."):
                try:
                    ps_file.seek(0)
                    wb_source = openpyxl.load_workbook(ps_file, data_only=True)
                    dpci_pattern = re.compile(r'\d{3}-\d{2}-\d{4}')
                    
                    dpci_locations_by_sheet = {}
                    for sheet_name in wb_source.sheetnames:
                        ws_source = wb_source[sheet_name]
                        dpci_locations_by_sheet[sheet_name] = []
                        for r in range(1, ws_source.max_row + 1):
                            for c in range(1, ws_source.max_column + 1):
                                cell_val = ws_source.cell(row=r, column=c).value
                                if cell_val and isinstance(cell_val, str):
                                    match = dpci_pattern.search(cell_val)
                                    if match:
                                        dpci_locations_by_sheet[sheet_name].append({'dpci': match.group(0), 'row': r, 'col': c})
                    
                    ps_file.seek(0)
                    images_info = []
                    
                    with zipfile.ZipFile(ps_file, 'r') as z:
                        namelist = z.namelist()
                        media_files = {n: z.read(n) for n in namelist if n.startswith('xl/media/')}
                        wb_rels = {}
                        if 'xl/_rels/workbook.xml.rels' in namelist:
                            root = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
                            for rel in root.iter():
                                if rel.tag.endswith('}Relationship'): wb_rels[rel.attrib.get('Id')] = rel.attrib.get('Target')
                                
                        sheet_name_to_path = {}
                        if 'xl/workbook.xml' in namelist:
                            root = ET.fromstring(z.read('xl/workbook.xml'))
                            for sheet in root.iter():
                                if sheet.tag.endswith('}sheet'):
                                    rId = sheet.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                                    name = sheet.attrib.get('name')
                                    if rId and rId in wb_rels: sheet_name_to_path[name] = resolve_zip_path('xl', wb_rels[rId])

                        for sheet_name, sheet_path in sheet_name_to_path.items():
                            if sheet_path not in namelist: continue
                            sheet_xml = ET.fromstring(z.read(sheet_path))
                            for drawing in sheet_xml.iter():
                                if drawing.tag.endswith('}drawing'):
                                    draw_rId = drawing.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                                    if not draw_rId: continue
                                    
                                    sheet_rels_path = resolve_zip_path(os.path.dirname(sheet_path), '_rels/' + os.path.basename(sheet_path) + '.rels')
                                    drawing_path = None
                                    if sheet_rels_path in namelist:
                                        rels_root = ET.fromstring(z.read(sheet_rels_path))
                                        for rel in rels_root.iter():
                                            if rel.tag.endswith('}Relationship') and rel.attrib.get('Id') == draw_rId:
                                                drawing_path = resolve_zip_path(os.path.dirname(sheet_path), rel.attrib.get('Target'))
                                                break
                                                
                                    if not drawing_path or drawing_path not in namelist: continue
                                    
                                    drawing_rels_path = resolve_zip_path(os.path.dirname(drawing_path), '_rels/' + os.path.basename(drawing_path) + '.rels')
                                    draw_rels = {}
                                    if drawing_rels_path in namelist:
                                        d_rels_root = ET.fromstring(z.read(drawing_rels_path))
                                        for rel in d_rels_root.iter():
                                            if rel.tag.endswith('}Relationship'): draw_rels[rel.attrib.get('Id')] = rel.attrib.get('Target')
                                                
                                    draw_root = ET.fromstring(z.read(drawing_path))
                                    for anchor in draw_root:
                                        if 'Anchor' not in anchor.tag: continue
                                        row, col = 0, 0
                                        for from_marker in anchor.iter():
                                            if from_marker.tag.endswith('}from'):
                                                for child in from_marker.iter():
                                                    if child.tag.endswith('}col'): col = int(child.text) + 1 if child.text else 0
                                                    elif child.tag.endswith('}row'): row = int(child.text) + 1 if child.text else 0
                                                break
                                        
                                        for elem in anchor.iter():
                                            if elem.tag.endswith('}blip'):
                                                embed_id = elem.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                                                if embed_id and embed_id in draw_rels:
                                                    media_target = draw_rels[embed_id]
                                                    media_path = resolve_zip_path(os.path.dirname(drawing_path), media_target)
                                                    if media_path in media_files:
                                                        images_info.append({
                                                            'sheet': sheet_name, 'row': row, 'col': col,
                                                            'media_path': media_path, 'bytes': media_files[media_path], 'ext': media_path.split('.')[-1]
                                                        })

                    zip_buffer_images = io.BytesIO()
                    extracted_count = 0
                    extracted_dpcis_count = {}
                    matched_media = set()
                    
                    with zipfile.ZipFile(zip_buffer_images, "a", zipfile.ZIP_DEFLATED, False) as zip_file_img:
                        for img in images_info:
                            sheet_name, anchor_row, anchor_col, media_path = img['sheet'], img['row'], img['col'], img['media_path']
                            closest_dpci, min_dist = None, float('inf')
                            
                            if anchor_row > 0 and anchor_col > 0 and sheet_name in dpci_locations_by_sheet:
                                for loc in dpci_locations_by_sheet[sheet_name]:
                                    dist = abs(loc['row'] - anchor_row) + abs(loc['col'] - anchor_col)
                                    if dist < min_dist and dist < 40: min_dist, closest_dpci = dist, loc['dpci']
                            
                            if closest_dpci:
                                extracted_dpcis_count[closest_dpci] = extracted_dpcis_count.get(closest_dpci, 0) + 1
                                file_name = f"{closest_dpci}.{img['ext']}" if extracted_dpcis_count[closest_dpci] == 1 else f"{closest_dpci}_{extracted_dpcis_count[closest_dpci]}.{img['ext']}"
                                zip_file_img.writestr(file_name, img['bytes'])
                                extracted_count += 1
                                matched_media.add(media_path)
                                
                        unmatched_count = 0
                        for p, b in media_files.items():
                            if p not in matched_media and p.lower().endswith(('.png', '.jpg', '.jpeg', '.gif')):
                                unmatched_count += 1
                                zip_file_img.writestr(f"Unmatched_Image_{unmatched_count}.{p.split('.')[-1]}", b)
                                        
                    if extracted_count > 0 or unmatched_count > 0:
                        msg = f"✅ 大功告成！成功自動命名 **{extracted_count}** 張商品圖片！"
                        if unmatched_count > 0: msg += f"\n⚠️ 另外發現 **{unmatched_count}** 張因格式或距離太遠無法對位的圖片，已命名為 `Unmatched_Image` 一併匯出給您檢查。"
                        st.success(msg)
                        st.download_button(label="📦 點擊下載完整圖片包 (ZIP)", data=zip_buffer_images.getvalue(), file_name="Auto_Extracted_Images.zip", mime="application/zip")
                    else:
                        st.warning("⚠️ 檔案底層完全找不到任何圖片，請確認 Excel 檔案中是否含有實體插入的圖片。")
                except Exception as e:
                    st.error(f"❌ 萃取過程中發生底層錯誤: {e}")
