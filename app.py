import streamlit as st
import pandas as pd
import io
import zipfile
import os
import re
import copy
import openpyxl
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
                                    image_dict[dpci_name] = z.read(file_info.filename)

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
                                
                                # 💡 執行 100% 精準對位的圖片貼上邏輯
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
# 分頁二：Program Sheet 圖片自動萃取與命名工具
# ==========================================
with tab2:
    st.markdown("""
    ### 🪄 圖片自動命名法寶
    這個工具能幫你「解剖」工廠提供的 Program Sheet，自動把裡面的圖片抓出來，並**智慧對位附近的 DPCI 號碼重新命名**。
    你只需要把產出的 ZIP 檔，拿到【分頁一】上傳即可！
    """)
    
    ps_file = st.file_uploader("📁 上傳 Program Sheet (包含圖片的 .xlsx)", type=['xlsx'], key="ps_uploader")
    
    if ps_file and st.button("🪄 開始自動萃取並命名圖片", type="primary"):
        with st.spinner("🕵️‍♂️ 正在掃描 Excel 並自動比對 DPCI，請稍候..."):
            try:
                wb_source = openpyxl.load_workbook(ps_file, data_only=True)
                dpci_pattern = re.compile(r'\d{3}-\d{2}-\d{4}')
                
                zip_buffer_images = io.BytesIO()
                extracted_count = 0
                extracted_dpcis = set()
                
                with zipfile.ZipFile(zip_buffer_images, "a", zipfile.ZIP_DEFLATED, False) as zip_file_img:
                    for sheet_name in wb_source.sheetnames:
                        ws_source = wb_source[sheet_name]
                        
                        for img in ws_source._images:
                            anchor_row = img.anchor._from.row + 1
                            anchor_col = img.anchor._from.col + 1
                            
                            # 在圖片周圍尋找 DPCI (上2列~下15列，左2欄~右8欄)
                            start_r = max(1, anchor_row - 2)
                            end_r = anchor_row + 15
                            start_c = max(1, anchor_col - 2)
                            end_c = anchor_col + 8
                            
                            found_dpci = None
                            for r in range(start_r, end_r):
                                for c in range(start_c, end_c):
                                    cell_val = ws_source.cell(row=r, column=c).value
                                    if cell_val and isinstance(cell_val, str):
                                        match = dpci_pattern.search(cell_val)
                                        if match:
                                            found_dpci = match.group(0)
                                            break
                                if found_dpci: break
                                
                            if found_dpci and found_dpci not in extracted_dpcis:
                                # 取得圖片二進位資料
                                try:
                                    if hasattr(img.ref, 'getvalue'):
                                        img_bytes = img.ref.getvalue()
                                    else:
                                        img.ref.seek(0)
                                        img_bytes = img.ref.read()
                                except:
                                    img_bytes = img._data() if hasattr(img, '_data') else None
                                
                                if img_bytes:
                                    img_format = getattr(img, 'format', 'png').lower()
                                    # 防呆副檔名
                                    if img_format not in ['png', 'jpg', 'jpeg']: img_format = 'png'
                                    
                                    file_name = f"{found_dpci}.{img_format}"
                                    zip_file_img.writestr(file_name, img_bytes)
                                    extracted_dpcis.add(found_dpci)
                                    extracted_count += 1
                                    
                if extracted_count > 0:
                    st.success(f"✅ 大功告成！成功萃取並命名了 **{extracted_count}** 張商品圖片！")
                    st.download_button(
                        label="📦 點擊下載自動命名圖片包 (ZIP)",
                        data=zip_buffer_images.getvalue(),
                        file_name="Auto_Extracted_Images.zip",
                        mime="application/zip"
                    )
                else:
                    st.warning("⚠️ 系統沒有在檔案中找到任何與 DPCI 綁定的圖片，請確認 Excel 格式或是否有實體圖片。")
                    
            except Exception as e:
                st.error(f"❌ 萃取過程中發生錯誤: {e}")
