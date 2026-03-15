import streamlit as st
import pandas as pd
import io
import zipfile

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

st.set_page_config(page_title="PO GRID 自動生成系統", layout="wide")

st.title("🎃 季節性 PO GRID 自動生成器")

st.markdown("""
請依序上傳 **PO RAW DATA**、**PO List (訂單清單)** 與 **產品資料(PCN)**。
上傳後，請在下方的表格中輸入對應的「目的地港口代碼」(如 581, 3891 等)，系統將自動生成多層表頭。
""")

# ==========================================
# 1. 檔案上傳區
# ==========================================
col1, col2, col3 = st.columns(3)
with col1:
    po_raw_file = st.file_uploader("📁 1. PO RAW DATA (CSV)", type=['csv'])
with col2:
    po_list_file = st.file_uploader("📁 2. List of Purchase Orders (CSV)", type=['csv'])
with col3:
    prod_file = st.file_uploader("📁 3. 產品資料(PCN) (Excel/CSV)", type=['xlsx', 'csv'])

# ==========================================
# 2. 讀取並顯示「港口輸入互動表」
# ==========================================
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
    st.subheader("📍 步驟 4: 請為以下 PO 分配目的地港口代碼")
    st.info("✏️ 操作說明：請將滑鼠移到表格最右側「輸入港口代碼」的空白處【連點兩下】，即可直接打字輸入！\n\n支援代碼: 581(PSW), 3890(PNW), 584(ORF), 3891(SAV), 3851(NYC), 3850(OAK), 3887(HOU), 3758(CHARLESTON)")
    
    edited_po_info = st.data_editor(
        po_info.reset_index(drop=True),
        use_container_width=True,
        hide_index=True,
        disabled=["PO NUMBER", "PURPOSE", "SHIP_DATES"]
    )
    
    # ==========================================
    # 3. 執行產出按鈕
    # ==========================================
    st.divider()
    if st.button("🚀 開始自動生成 PO GRID", type="primary"):
        with st.spinner("資料處理中，請稍候..."):
            try:
                if prod_file.name.endswith('.csv'):
                    prod_data = pd.read_csv(prod_file)
                else:
                    prod_data = pd.read_excel(prod_file)

                # --- 💡 資料清洗與「混裝 Assortment」母子分離核心邏輯 ---
                po_processed_records = []
                parent_dpci_list = set()
                child_assort_qty_dict = {}
                parent_info_dict = {}
                parent_to_child = {}

                for idx, row in po_raw.iterrows():
                    dept = str(int(row['DEPARTMENT'])) if pd.notna(row['DEPARTMENT']) else '0'
                    cls = str(int(row['CLASS'])).zfill(2) if pd.notna(row['CLASS']) else '00'
                    itm = str(int(row['ITEM'])).zfill(4) if pd.notna(row['ITEM']) else '0000'
                    dpci = f"{dept}-{cls}-{itm}"
                    
                    try: qty = float(str(row['TOTAL ITEM QTY']).replace(',', ''))
                    except: qty = 0.0
                        
                    desc = str(row['ITEM DESCRIPTION']).strip().upper()
                    po_num = row['PO NUMBER']
                    
                    if desc == 'ASSORTMENT':
                        # 處理母 DPCI
                        parent_dpci_list.add(dpci)
                        parent_info_dict[dpci] = {
                            'style': str(row['VENDOR STYLE']) if pd.notna(row['VENDOR STYLE']) else '',
                            'upc': str(row['ITEM BAR CODE']) if pd.notna(row['ITEM BAR CODE']) else ''
                        }
                        
                        po_processed_records.append({'PO NUMBER': po_num, 'DPCI_MERGE': dpci, 'QTY': qty, 'IS_PARENT': True})
                        
                        # 處理子 ITEM
                        c_dept = str(int(row['COMPONENT DEPARTMENT'])) if pd.notna(row['COMPONENT DEPARTMENT']) else '0'
                        c_cls = str(int(row['COMPONENT CLASS'])).zfill(2) if pd.notna(row['COMPONENT CLASS']) else '00'
                        c_itm = str(int(row['COMPONENT ITEM'])).zfill(4) if pd.notna(row['COMPONENT ITEM']) else '0000'
                        c_dpci = f"{c_dept}-{c_cls}-{c_itm}"
                        
                        try: c_qty = float(str(row['COMPONENT ITEM TOTAL QTY']).replace(',', ''))
                        except: c_qty = 0.0
                            
                        try: c_assort = float(str(row['COMPONENT ASSORT QTY']).replace(',', ''))
                        except: c_assort = 0.0
                            
                        child_assort_qty_dict[c_dpci] = c_assort
                        parent_to_child[dpci] = c_dpci # 記憶母子關係
                        
                        po_processed_records.append({'PO NUMBER': po_num, 'DPCI_MERGE': c_dpci, 'QTY': c_qty, 'IS_PARENT': False})
                    else:
                        # 普通商品
                        po_processed_records.append({'PO NUMBER': po_num, 'DPCI_MERGE': dpci, 'QTY': qty, 'IS_PARENT': False})

                po_processed = pd.DataFrame(po_processed_records)
                # 去除母商品多餘的重複行 (因為多個子 ITEM 會重複產出同一個母 DPCI 的同一筆 PO 行)
                parents = po_processed[po_processed['IS_PARENT']].drop_duplicates(subset=['PO NUMBER', 'DPCI_MERGE'])
                children_and_regular = po_processed[~po_processed['IS_PARENT']]
                po_processed_unique = pd.concat([parents, children_and_regular], ignore_index=True)

                # --- 合併港口資訊 ---
                edited_po_info['PORT_NAME'] = edited_po_info['輸入港口代碼 (如:581)'].astype(str).map(PORT_MAP).fillna(edited_po_info['輸入港口代碼 (如:581)'])
                edited_po_info['PORT_NAME'] = edited_po_info['PORT_NAME'].replace({'': '未指定港口', 'nan': '未指定港口'})

                po_raw_merged = po_processed_unique.merge(edited_po_info[['PO NUMBER', 'PURPOSE', 'SHIP_DATES', 'PORT_NAME']], on='PO NUMBER', how='left')
                po_raw_merged['PURPOSE'] = po_raw_merged['PURPOSE'].fillna('標籤遺失')
                po_raw_merged['SHIP_DATES'] = po_raw_merged['SHIP_DATES'].fillna('日期遺失')
                po_raw_merged['PORT_NAME'] = po_raw_merged['PORT_NAME'].fillna('未指定港口')
                
                # --- 建立樞紐 (4層: PURPOSE, PO NUMBER, SHIP_DATES, PORT_NAME) ---
                pivot_df_temp = po_raw_merged.pivot_table(
                    index='DPCI_MERGE', 
                    columns=['PURPOSE', 'PO NUMBER', 'SHIP_DATES', 'PORT_NAME'], 
                    values='QTY', 
                    aggfunc='sum'
                ).fillna(0)
                
                # 擴展為 5 層表頭
                new_pivot_cols = []
                for col in pivot_df_temp.columns:
                    new_pivot_cols.append((col[0], '', col[1], col[2], col[3]))
                pivot_df = pd.DataFrame(pivot_df_temp.values, index=pivot_df_temp.index, columns=pd.MultiIndex.from_tuples(new_pivot_cols))
                
                # 💡 加入 PO TOTAL 計算
                pivot_df[('', 'PO TOTAL', '', '', '')] = pivot_df.sum(axis=1)
                
                # 💡 替換母 DPCI PO 的數值為文字 ("母DPCI")
                for parent_dpci in parent_dpci_list:
                    if parent_dpci in pivot_df.index:
                        for col in pivot_df.columns:
                            if col[1] != 'PO TOTAL': # 避開加總欄
                                val = pivot_df.loc[parent_dpci, col]
                                if isinstance(val, (int, float)) and val > 0:
                                    pivot_df.loc[parent_dpci, col] = parent_dpci

                # --- 準備左側靜態產品資料 ---
                if 'DPCI' not in prod_data.columns:
                    st.error("❌ 產品資料(PCN) 中找不到必要的 'DPCI' 欄位。")
                    st.stop()
                    
                prod_data['DPCI_MERGE'] = prod_data['DPCI'].astype(str).str.strip()
                
                def format_upc(val):
                    if pd.isna(val) or val == '': return ''
                    try: return str(int(float(val)))
                    except: return str(val).strip()
                
                if 'Barcode' in prod_data.columns:
                    prod_data['Barcode'] = prod_data['Barcode'].apply(format_upc)
                else: prod_data['Barcode'] = ''

                # 💡 將母 DPCI 建立並寫入左側資料庫
                for parent_dpci, info in parent_info_dict.items():
                    child_dpci = parent_to_child.get(parent_dpci)
                    vendor_name, factory_name = '', ''
                    
                    # 繼承子項目的供應商與廠區
                    if child_dpci and child_dpci in prod_data['DPCI_MERGE'].values:
                        child_rows = prod_data[prod_data['DPCI_MERGE'] == child_dpci]
                        vendor_name = child_rows.iloc[0].get('Import Vendor Name', '')
                        factory_name = child_rows.iloc[0].get('Factory Name', '')
                        
                    if parent_dpci in prod_data['DPCI_MERGE'].values:
                        idx = prod_data.index[prod_data['DPCI_MERGE'] == parent_dpci].tolist()
                        for i in idx:
                            prod_data.at[i, 'Manufacturer Style # *'] = info['style']
                            prod_data.at[i, 'Barcode'] = format_upc(info['upc'])
                            if vendor_name: prod_data.at[i, 'Import Vendor Name'] = vendor_name
                            if factory_name: prod_data.at[i, 'Factory Name'] = factory_name
                    else:
                        new_row = {col: '' for col in prod_data.columns}
                        new_row['DPCI'] = parent_dpci
                        new_row['DPCI_MERGE'] = parent_dpci
                        new_row['Manufacturer Style # *'] = info['style']
                        new_row['Barcode'] = format_upc(info['upc'])
                        new_row['Product Description'] = 'ASSORTMENT'
                        new_row['Import Vendor Name'] = vendor_name
                        new_row['Factory Name'] = factory_name
                        prod_data = pd.concat([prod_data, pd.DataFrame([new_row])], ignore_index=True)
                        
                # 💡 寫入子項目的 Assortment QTY
                if 'Assortment' not in prod_data.columns: prod_data['Assortment'] = ''
                for child_dpci, assort_qty in child_assort_qty_dict.items():
                    if child_dpci in prod_data['DPCI_MERGE'].values:
                        idx = prod_data.index[prod_data['DPCI_MERGE'] == child_dpci].tolist()
                        for i in idx:
                            prod_data.at[i, 'Assortment'] = int(assort_qty) if float(assort_qty).is_integer() else float(assort_qty)

                # 處理靜態欄位生成
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
                    if col not in prod_data.columns:
                        prod_data[col] = ''
                        
                left_data = prod_data[desired_left_columns + ['DPCI_MERGE']].drop_duplicates(subset=['DPCI_MERGE']).set_index('DPCI_MERGE')
                
                def get_left_tuple(col):
                    if col == 'DPCI': return ('Program Name', 'DPCI', '', '', '')
                    elif col == 'Barcode': return ('', 'UPC', '', '', '')
                    else: return ('', col, '', '', '')

                left_data.columns = pd.MultiIndex.from_tuples([get_left_tuple(col) for col in left_data.columns])
                
                final_df = left_data.join(pivot_df, how='inner')

                # --- 拆檔與寫入 ZIP ---
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    
                    vendor_col = ('', 'Import Vendor Name', '', '', '')
                    factory_col = ('', 'Factory Name', '', '', '')
                    
                    final_df[vendor_col] = final_df[vendor_col].replace('', '未指定供應商')
                    final_df[factory_col] = final_df[factory_col].replace('', '未指定工廠')
                    
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        
                        grouped_factory = final_df.groupby(factory_col)
                        
                        for factory_name, factory_data in grouped_factory:
                            safe_factory_name = str(factory_name).replace('/', '_').replace('\\', '_')
                            safe_factory_name = safe_factory_name.replace('[', '').replace(']', '').replace('*', '').replace(':', '').replace('?', '')[:31]
                            
                            export_data = factory_data.drop(columns=[vendor_col, factory_col])
                            
                            # 💡 升級版：過濾沒有任何數值或母代碼的欄位
                            cols_to_keep = []
                            for col in export_data.columns:
                                if col[2] != '': # 若為 PO 欄位
                                    keep = False
                                    for val in export_data[col]:
                                        if isinstance(val, (int, float)) and val > 0:
                                            keep = True
                                            break
                                        elif isinstance(val, str) and val not in ('', '0', '0.0'):
                                            keep = True
                                            break
                                    if keep:
                                        cols_to_keep.append(col)
                                else:
                                    cols_to_keep.append(col)
                                    
                            export_data = export_data[cols_to_keep]
                            
                            unmerged_columns = []
                            po_idx = 0
                            for col in export_data.columns:
                                if col[2] != '': 
                                    new_purpose = str(col[0]) + (" " * po_idx) 
                                    unmerged_columns.append((new_purpose, col[1], col[2], col[3], col[4]))
                                    po_idx += 1
                                else:
                                    unmerged_columns.append(col)
                                    
                            export_data.columns = pd.MultiIndex.from_tuples(unmerged_columns)
                            export_data_reset = export_data.reset_index(drop=True)
                            
                            export_data_reset.to_excel(writer, index=True, sheet_name=safe_factory_name)
                            writer.sheets[safe_factory_name].delete_cols(1) 
                    
                    zip_file.writestr("PO_GRID_Merged_All.xlsx", excel_buffer.getvalue())
                
                st.success("✨ 處理完成！已成功加入【Assortment 混裝邏輯】與所有的格式匹配。")
                st.download_button(
                    label="📦 點擊下載合併版 PO GRID (ZIP壓縮檔)",
                    data=zip_buffer.getvalue(),
                    file_name="PO_GRIDs_Output.zip",
                    mime="application/zip"
                )
                
            except Exception as e:
                st.error(f"❌ 處理過程中發生錯誤: {e}")
