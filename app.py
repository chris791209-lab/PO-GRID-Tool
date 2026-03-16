import streamlit as st
import pandas as pd
import io
import zipfile
import os
import re
import copy
import openpyxl
import xml.etree.ElementTree as ET
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.utils import get_column_letter

# 港口代碼對照表
PORT_MAP = {
    '581': 'PSW',
    '3890': 'PNW',
    '584': 'ORF',
    '3891': 'SAV',
    '3851': 'NYC',
    '3850': 'OAK',
    '3887': 'HOU',
    '3758': 'CHARLESTON'
}

# --- 底層 ZIP 路徑解析輔助函數 ---
def resolve_zip_path(base_dir, relative_path):
    if relative_path.startswith('/'): return relative_path[1:]
    parts = [p for p in base_dir.split('/') if p]
    for part in relative_path.split('/'):
        if part == '..':
            if parts: parts.pop()
        elif part and part != '.':
            parts.append(part)
    return '/'.join(parts)

st.set_page_config(page_title="PO GRID & 圖片萃取系統", layout="wide")

st.title("🎯 Target 季節性專案自動化系統")

# 建立雙分頁 UI
tab1, tab2 = st.tabs(["🎃 PO GRID 自動生成器", "🖼️ Program Sheet 圖片自動萃取器"])

# ==========================================
# 分頁一：PO GRID 自動生成器 (ZIP精準對位版)
# ==========================================
with tab1:
    st.markdown("""
    請依序上傳 **PO RAW DATA**、**PO List (訂單清單)**、**產品資料(PCN)**，以及 **商品圖片包(ZIP)**。
    上傳後，請在下方的表格中輸入對應的「目的地港口代碼」(如 581, 3891 等)，系統將自動生成多層表頭與含圖片的報表。
    """)

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        po_raw_file = st.file_uploader("📁 1. PO RAW DATA (CSV)", type=['csv'])
    with col2:
        po_list_file = st.file_uploader("📁 2. List of PO (CSV)", type=['csv'])
    with col3:
        prod_file = st.file_uploader("📁 3. 產品資料(PCN)", type=['xlsx', 'csv'])
    with col4:
        image_zip_file = st.file_uploader("📁 4. 產品圖片包 (ZIP)\n(可透過分頁二工具產生)", type=['zip'])

    if po_raw_file and prod_file and po_list_file:
        po_list = pd.read_csv(po_list_file)
        po_raw = pd.read_csv(po_raw_file)
        
        po_list['PO NUMBER'] = po_list['PO NUMBER'].astype(str).str.split('.').str[0].str.strip()
        po_raw['PO NUMBER'] = po_raw['PO NUMBER'].astype(str).str.split('.').str[0].str.strip()
        
        po_list['SHIP BEGIN DATE'] = pd.to_datetime(po_list['SHIP BEGIN DATE'], errors='coerce')
        po_list['SHIP END DATE'] = pd.to_datetime(po_list['SHIP END DATE'], errors='coerce')
        po_list['SHIP_DATES'] = po_list['SHIP BEGIN DATE'].dt.strftime('%m/%d') + '-' + po_list['SHIP END DATE'].dt.strftime('%m/%d')
        
        po_info = po_list[['PO NUMBER', 'PURPOSE', 'SHIP_DATES']].drop_duplicates()
        active_pos = po_raw['PO NUMBER'].unique()
        po_info = po_info[po_info['PO NUMBER'].isin(active_pos)].copy()
        
        po_info['輸入港口代碼 (如:581)'] = ""
        
        st.divider()
        st.subheader("📍 步驟 5: 請為以下 PO 分配目的地港口代碼")
        st.info("✏️ 操作說明：請將滑鼠移到表格最右側「輸入港口代碼」的空白處【連點兩下】，即可直接打字輸入！")
        
        edited_po_info = st.data_editor(
            po_info.reset_index(drop=True),
            use_container_width=True,
            hide_index=True,
            disabled=["PO NUMBER", "PURPOSE", "SHIP_DATES"]
        )
        
        st.divider()
        if st.button("🚀 開始自動生成 PO GRID", type="primary"):
            with st.spinner("資料處理與圖片載入中，請稍候..."):
                try:
                    # 讀取 ZIP 圖片包
                    image_dict = {}
                    if image_zip_file:
                        with zipfile.ZipFile(image_zip_file, 'r') as z:
                            for file_info in z.infolist():
                                if file_info.filename.startswith('__MACOSX/') or file_info.filename.startswith('.'):
                                    continue
                                if file_info.filename.lower().endswith(('.png', '.jpg', '.jpeg')):
                                    base_name = os.path.basename(file_info.filename)
                                    dpci_name = os.path.splitext(base_name)[0].strip()
                                    clean_dpci = dpci_name.split('_')[0] 
                                    if clean_dpci not in image_dict:
                                        image_dict[clean_dpci] = z.read(file_info.filename)

                    # 讀取產品資料
                    if prod_file.name.endswith('.csv'):
                        prod_data = pd.read_csv(prod_file)
                    else:
                        prod_data = pd.read_excel(prod_file)

                    po_processed_records = []
                    parent_dpci_list = set()
                    child_assort_qty_dict = {}
                    parent_info_dict = {}
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
                        
                        if desc.startswith('ASSORTMENT'):
                            parent_dpci_list.add(dpci)
                            style_val = str(row['VENDOR STYLE']).strip() if pd.notna(row['VENDOR STYLE']) else ''
                            if style_val and not style_val.upper().startswith('ASSORT'):
                                style_val = f"ASSORTMENT-{style_val}"
                                
                            parent_info_dict[dpci] = {
                                'style': style_val,
                                'upc': str(row['ITEM BAR CODE']).strip() if pd.notna(row['ITEM BAR CODE']) else ''
                            }
                            po_processed_records.append({'PO NUMBER': po_num, 'DPCI_MERGE': dpci, 'QTY': qty, 'IS_PARENT': True})
                            
                            c_dept = str(int(row['COMPONENT DEPARTMENT'])) if pd.notna(row['COMPONENT DEPARTMENT']) else '0'
                            c_cls = str(int(row['COMPONENT CLASS'])).zfill(2) if pd.notna(row['COMPONENT CLASS']) else '00'
                            c_itm = str(int(row['COMPONENT ITEM'])).zfill(4) if pd.notna(row['COMPONENT ITEM']) else '0000'
                            c_dpci = f"{c_dept}-{c_cls}-{c_itm}"
                            
                            try: c_qty = float(str(row['COMPONENT ITEM TOTAL QTY']).replace(',', ''))
                            except: c_qty = 0.0
                            try: c_assort = float(str(row['COMPONENT ASSORT QTY']).replace(',', ''))
                            except: c_assort = 0.0
                                
                            child_assort_qty_dict[c_dpci] = c_assort
                            
                            if dpci not in parent_to_children:
                                parent_to_children[dpci] = set()
                            parent_to_children[dpci].add(c_dpci)
                            
                            po_processed_records.append({'PO NUMBER': po_num, 'DPCI_MERGE': c_dpci, 'QTY': c_qty, 'IS_PARENT': False})
                        else:
                            po_processed_records.append({'PO NUMBER': po_num, 'DPCI_MERGE': dpci, 'QTY': qty, 'IS_PARENT': False})

                    po_processed = pd.DataFrame(po_processed_records)
                    parents = po_processed[po_processed['IS_PARENT']].drop_duplicates(subset=['PO NUMBER', 'DPCI_MERGE'])
                    children_and_regular = po_processed[~po_processed['IS_PARENT']]
                    po_processed_unique = pd.concat([parents, children_and_regular], ignore_index=True)

                    edited_po_info['PORT_NAME'] = edited_po_info['輸入港口代碼 (如:581)'].astype(str).map(PORT_MAP).fillna(edited_po_info['輸入港口代碼 (如:581)'])
                    edited_po_info['PORT_NAME'] = edited_po_info['PORT_NAME'].replace({'': '未指定港口', 'nan': '未指定港口'})

                    po_raw_merged = po_processed_unique.merge(edited_po_info[['PO NUMBER', 'PURPOSE', 'SHIP_DATES', 'PORT_NAME']], on='PO NUMBER', how='left')
                    po_raw_merged['PURPOSE'] = po_raw_merged['PURPOSE'].fillna('標籤遺失')
                    po_raw_merged['SHIP_DATES'] = po_raw_merged['SHIP_DATES'].fillna('日期遺失')
                    po_raw_merged['PORT_NAME'] = po_raw_merged['PORT_NAME'].fillna('未指定港口')
                    
                    pivot_df_temp = po_raw_merged.pivot_table(
                        index='DPCI_MERGE', columns=['PURPOSE', 'PO NUMBER', 'SHIP_DATES', 'PORT_NAME'], 
                        values='QTY', aggfunc='sum'
                    ).fillna(0)
                    
                    new_pivot_cols = []
                    for col in pivot_df_temp.columns:
                        new_pivot_cols.append((col[0], '', col[1], col[2], col[3]))
                    pivot_df = pd.DataFrame(pivot_df_temp.values, index=pivot_df_temp.index, columns=pd.MultiIndex.from_tuples(new_pivot_cols))
                    
                    pivot_df[('', 'PO TOTAL', '', '', '')] = pivot_df.sum(axis=1)
                    
                    for parent_dpci in parent_dpci_list:
                        if parent_dpci in pivot_df.index:
                            pivot_df.loc[parent_dpci, ('', 'PO TOTAL', '', '', '')] = '' 
                            for col in pivot_df.columns:
                                if col[1] != 'PO TOTAL': 
                                    val = pivot_df.loc[parent_dpci, col]
                                    if isinstance(val, (int, float)) and val > 0:
                                        pivot_df.loc[parent_dpci, col] = f"{parent_dpci}-({int(val)})"
                                        
                    pivot_df = pivot_df.replace({0: '', 0.0: ''})

                    if 'DPCI' not in prod_data.columns:
                        st.error("❌ 產品資料(PCN) 中找不到必要的 'DPCI' 欄位。")
                        st.stop()
                        
                    prod_data['DPCI_MERGE'] = prod_data['DPCI'].astype(str).str.strip()
                    
                    def format_upc(val):
                        if pd.isna(val) or val == '': return ''
                        try: 
                            v = str(int(float(val)))
                            return v.zfill(12) if len(v) < 12 else v
                        except: return str(val).strip()
                    
                    if 'Barcode' in prod_data.columns: prod_data['Barcode'] = prod_data['Barcode'].apply(format_upc)
                    else: prod_data['Barcode'] = ''

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
                                for i in idx:
                                    prod_data.at[i, 'Assortment'] = int(assort_qty) if float(assort_qty).is_integer() else float(assort_qty)

                    if 'Factory Name' not in prod_data.columns: prod_data['Factory Name'] = '未提供工廠名稱'
                    if 'Factory ID' not in prod_data.columns: prod_data['Factory ID'] = ''
                    
                    def make_maker(row):
                        fid = str(row.get('Factory ID', '')).replace('.0', '').strip()
                        fname = str(row.get('Factory Name', '')).strip()
                        if fid and fid != 'nan': return f"{fid}-{fname}"
                        return fname
                    prod_data['Maker'] = prod_data.apply(make_maker, axis=1)

                    if 'Retail Packaging Format (1) *' in prod_data.columns:
                        prod_data['Packaging'] = prod_data['Retail Packaging Format (1) *'].fillna('')
                    else: prod_data['Packaging'] = ''

                    prod_data['IMAGE'] = ''
                    prod_data['Age'] = ''

                    desired_left_columns = [
                        'DPCI', 'Manufacturer Style # *', 'IMAGE', 'Product Description', 
                        'Barcode', 'Primary Raw Material Type', 'Age', 'Maker', 
                        'Packaging', 'Inner Pack Unit Quantity', 'Case Unit Quantity', 'Assortment',
                        'Import Vendor Name', 'Factory Name'
                    ]
                    
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
                                        keep = False
                                        for val in export_data[col]:
                                            if isinstance(val, (int, float)) and val > 0:
                                                keep = True
                                                break
                                            elif isinstance(val, str) and val not in ('', '0', '0.0'):
                                                keep = True
                                                break
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
                                
                                if image_zip_file and image_dict:
                                    dpci_col_idx = None
                                    img_col_idx = None
                                    for idx, col in enumerate(export_data_reset.columns):
                                        if col[1] == 'DPCI': dpci_col_idx = idx + 1
                                        if col[1] == 'IMAGE': img_col_idx = idx + 1
                                    
                                    if dpci_col_idx and img_col_idx:
                                        img_col_letter = get_column_letter(img_col_idx)
                                        ws.column_dimensions[img_col_letter].width = 13 
                                        
                                        for r_idx in range(6, ws.max_row + 1):
                                            cell_dpci_val = str(ws.cell(row=r_idx, column=dpci_col_idx).value).strip()
                                            
                                            if cell_dpci_val in image_dict:
                                                img_bytes = io.BytesIO(image_dict[cell_dpci_val])
                                                img_obj = OpenpyxlImage(img_bytes)
                                                img_obj.width = 90
                                                img_obj.height = 90
                                                ws.add_image(img_obj, f"{img_col_letter}{r_idx}")
                                                ws.row_dimensions[r_idx].height = 70 
                        
                        zip_file.writestr("PO_GRID_Merged_All.xlsx", excel_buffer.getvalue())
                    
                    st.success("✨ 處理完成！已透過 DPCI 精準匹配並填入商品圖片。")
                    st.download_button(
                        label="📦 點擊下載合併版 PO GRID (ZIP壓縮檔)",
                        data=zip_buffer.getvalue(),
                        file_name="PO_GRIDs_Output.zip",
                        mime="application/zip"
                    )
                    
                except Exception as e:
                    st.error(f"❌ 處理過程中發生錯誤: {e}")

# ==========================================
# 分頁二：Program Sheet 圖片自動萃取與命名工具 (究極 XML 暴力破解版)
# ==========================================
with tab2:
    st.markdown("""
    ### 🪄 圖片自動命名法寶 (底層暴力破解版)
    此工具繞過了一般程式對群組化圖片的盲區，直接潛入 Excel 底層檔案架構中，將 **100% 所有的商品圖片** 硬抓出來，並智慧對位 DPCI 重新命名！
    """)
    
    ps_file = st.file_uploader("📁 上傳 Program Sheet (包含圖片的 .xlsx)", type=['xlsx'], key="ps_uploader")
    
    if ps_file and st.button("🪄 開始自動萃取並命名圖片", type="primary"):
        with st.spinner("🕵️‍♂️ 正在深入 Excel 底層架構暴力抓取圖片..."):
            try:
                # ---------------------------------------------------------
                # 1. 為了取得 DPCI 的文字座標，我們依然用 openpyxl 掃一次文字
                # ---------------------------------------------------------
                ps_file.seek(0)
                wb_source = openpyxl.load_workbook(ps_file, data_only=True)
                dpci_pattern = re.compile(r'\d{3}-\d{2}-\d{4}')
                
                # 儲存每個分頁的 DPCI 位置
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
                                    dpci_locations_by_sheet[sheet_name].append({
                                        'dpci': match.group(0),
                                        'row': r,
                                        'col': c
                                    })
                
                # ---------------------------------------------------------
                # 2. XML 暴力破解：直接把 .xlsx 當作 ZIP 解壓縮提取所有媒體
                # ---------------------------------------------------------
                ps_file.seek(0)
                images_info = []
                
                with zipfile.ZipFile(ps_file, 'r') as z:
                    namelist = z.namelist()
                    # 抓出所有隱藏在媒體庫的實體圖片檔案
                    media_files = {n: z.read(n) for n in namelist if n.startswith('xl/media/')}
                    
                    # 分析 XML 關聯表，把 圖片 -> 座標 -> 分頁 串起來
                    wb_rels = {}
                    if 'xl/_rels/workbook.xml.rels' in namelist:
                        root = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
                        for rel in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                            wb_rels[rel.attrib['Id']] = rel.attrib['Target']
                            
                    sheet_name_to_path = {}
                    if 'xl/workbook.xml' in namelist:
                        root = ET.fromstring(z.read('xl/workbook.xml'))
                        for sheet in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
                            rId = sheet.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                            name = sheet.attrib.get('name')
                            if rId and rId in wb_rels:
                                target = wb_rels[rId]
                                sheet_name_to_path[name] = resolve_zip_path('xl', target)

                    for sheet_name, sheet_path in sheet_name_to_path.items():
                        if sheet_path not in namelist: continue
                        sheet_xml = ET.fromstring(z.read(sheet_path))
                        drawing = sheet_xml.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}drawing')
                        if drawing is None: continue
                        draw_rId = drawing.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')

                        sheet_rels_path = resolve_zip_path(os.path.dirname(sheet_path), '_rels/' + os.path.basename(sheet_path) + '.rels')
                        drawing_path = None
                        if sheet_rels_path in namelist:
                            rels_root = ET.fromstring(z.read(sheet_rels_path))
                            for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                                if rel.attrib['Id'] == draw_rId:
                                    drawing_path = resolve_zip_path(os.path.dirname(sheet_path), rel.attrib['Target'])
                                    break

                        if not drawing_path or drawing_path not in namelist: continue

                        drawing_rels_path = resolve_zip_path(os.path.dirname(drawing_path), '_rels/' + os.path.basename(drawing_path) + '.rels')
                        draw_rels = {}
                        if drawing_rels_path in namelist:
                            d_rels_root = ET.fromstring(z.read(drawing_rels_path))
                            for rel in d_rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                                draw_rels[rel.attrib['Id']] = rel.attrib['Target']

                        draw_root = ET.fromstring(z.read(drawing_path))
                        ns_xdr = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
                        ns_a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
                        ns_r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

                        for anchor in draw_root:
                            if 'Anchor' not in anchor.tag: continue
                            
                            row, col = 0, 0
                            from_marker = anchor.find(f'.//{{{ns_xdr}}}from')
                            if from_marker is not None:
                                col_elem = from_marker.find(f'{{{ns_xdr}}}col')
                                row_elem = from_marker.find(f'{{{ns_xdr}}}row')
                                if col_elem is not None: col = int(col_elem.text) + 1
                                if row_elem is not None: row = int(row_elem.text) + 1

                            blip = anchor.find(f'.//{{{ns_a}}}blip')
                            if blip is not None:
                                embed_id = blip.attrib.get(f'{{{ns_r}}}embed')
                                if embed_id and embed_id in draw_rels:
                                    media_target = draw_rels[embed_id]
                                    media_path = resolve_zip_path(os.path.dirname(drawing_path), media_target)
                                    if media_path in media_files:
                                        ext = media_path.split('.')[-1]
                                        if ext.lower() not in ['png', 'jpg', 'jpeg']: ext = 'png'
                                        images_info.append({
                                            'sheet': sheet_name,
                                            'row': row,
                                            'col': col,
                                            'bytes': media_files[media_path],
                                            'ext': ext
                                        })

                # ---------------------------------------------------------
                # 3. 執行智慧綁定與打包 ZIP
                # ---------------------------------------------------------
                zip_buffer_images = io.BytesIO()
                extracted_count = 0
                extracted_dpcis_count = {}
                
                with zipfile.ZipFile(zip_buffer_images, "a", zipfile.ZIP_DEFLATED, False) as zip_file_img:
                    for img in images_info:
                        sheet_name = img['sheet']
                        anchor_row = img['row']
                        anchor_col = img['col']
                        
                        closest_dpci = None
                        min_dist = float('inf')
                        
                        # 找出該分頁中距離這張圖片最近的 DPCI
                        if sheet_name in dpci_locations_by_sheet:
                            for loc in dpci_locations_by_sheet[sheet_name]:
                                dist = abs(loc['row'] - anchor_row) + abs(loc['col'] - anchor_col)
                                # 允許搜尋的範圍擴大，只要在同一個產品區塊內 (30格內)
                                if dist < min_dist and dist < 30:
                                    min_dist = dist
                                    closest_dpci = loc['dpci']
                        
                        if closest_dpci:
                            extracted_dpcis_count[closest_dpci] = extracted_dpcis_count.get(closest_dpci, 0) + 1
                            if extracted_dpcis_count[closest_dpci] == 1:
                                file_name = f"{closest_dpci}.{img['ext']}"
                            else:
                                file_name = f"{closest_dpci}_{extracted_dpcis_count[closest_dpci]}.{img['ext']}"
                                
                            zip_file_img.writestr(file_name, img['bytes'])
                            extracted_count += 1
                                    
                if extracted_count > 0:
                    st.success(f"✅ 大功告成！透過底層解析，成功萃取並命名了 **{extracted_count}** 張商品圖片！\n(已為您破解原有的群組化/格式限制)")
                    st.download_button(
                        label="📦 點擊下載自動命名圖片包 (ZIP)",
                        data=zip_buffer_images.getvalue(),
                        file_name="Auto_Extracted_Images.zip",
                        mime="application/zip"
                    )
                else:
                    st.warning("⚠️ 檔案底層真的找不到任何符合條件的圖片，請確認 Excel 檔案中是否含有實體插入的圖片。")
                    
            except Exception as e:
                st.error(f"❌ 萃取過程中發生錯誤: {e}")
