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
    
    # 強制將 PO NUMBER 轉為乾淨的文字格式
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

                # --- 資料清洗 ---
                po_raw['DEPARTMENT'] = po_raw['DEPARTMENT'].fillna(0).astype(int).astype(str)
                po_raw['CLASS'] = po_raw['CLASS'].fillna(0).astype(int).astype(str).str.zfill(2)
                po_raw['ITEM'] = po_raw['ITEM'].fillna(0).astype(int).astype(str).str.zfill(4)
                po_raw['DPCI_MERGE'] = po_raw['DEPARTMENT'] + '-' + po_raw['CLASS'] + '-' + po_raw['ITEM']

                if po_raw['TOTAL ITEM QTY'].dtype == object:
                    po_raw['TOTAL ITEM QTY'] = po_raw['TOTAL ITEM QTY'].str.replace(',', '').astype(float)

                edited_po_info['PORT_NAME'] = edited_po_info['輸入港口代碼 (如:581)'].astype(str).map(PORT_MAP).fillna(edited_po_info['輸入港口代碼 (如:581)'])
                edited_po_info['PORT_NAME'] = edited_po_info['PORT_NAME'].replace({'': '未指定港口', 'nan': '未指定港口'})

                po_raw_merged = po_raw.merge(edited_po_info[['PO NUMBER', 'PURPOSE', 'SHIP_DATES', 'PORT_NAME']], on='PO NUMBER', how='left')
                
                po_raw_merged['PURPOSE'] = po_raw_merged['PURPOSE'].fillna('標籤遺失')
                po_raw_merged['SHIP_DATES'] = po_raw_merged['SHIP_DATES'].fillna('日期遺失')
                po_raw_merged['PORT_NAME'] = po_raw_merged['PORT_NAME'].fillna('未指定港口')
                
                # --- 建立樞紐 (4層: PURPOSE, PO NUMBER, SHIP_DATES, PORT_NAME) ---
                pivot_df_temp = po_raw_merged.pivot_table(
                    index='DPCI_MERGE', 
                    columns=['PURPOSE', 'PO NUMBER', 'SHIP_DATES', 'PORT_NAME'], 
                    values='TOTAL ITEM QTY', 
                    aggfunc='sum'
                ).fillna(0)
                
                # 擴展為 5 層，讓 PURPOSE 獨佔 Row 1，PO NUMBER 降至 Row 3 (Row 2 留空給靜態標題)
                new_pivot_cols = []
                for col in pivot_df_temp.columns:
                    new_pivot_cols.append((col[0], '', col[1], col[2], col[3]))
                pivot_df = pd.DataFrame(pivot_df_temp.values, index=pivot_df_temp.index, columns=pd.MultiIndex.from_tuples(new_pivot_cols))
                
                # 加入 PO TOTAL
                pivot_df[('', 'PO TOTAL', '', '', '')] = pivot_df.sum(axis=1)
                
                # --- 準備左側靜態產品資料 ---
                if 'DPCI' not in prod_data.columns:
                    st.error("❌ 產品資料(PCN) 中找不到必要的 'DPCI' 欄位。")
                    st.stop()
                    
                prod_data['DPCI_MERGE'] = prod_data['DPCI'].astype(str).str.strip()
                
                # 💡 處理 UPC 格式：解決科學記號 (1.9926E+11) 問題
                def format_upc(val):
                    if pd.isna(val): return ''
                    try:
                        return str(int(float(val)))
                    except:
                        return str(val).strip()
                
                if 'Barcode' in prod_data.columns:
                    prod_data['Barcode'] = prod_data['Barcode'].apply(format_upc)
                else:
                    prod_data['Barcode'] = ''

                # 💡 處理 Maker (Factory ID - Factory Name)
                if 'Factory Name' not in prod_data.columns: prod_data['Factory Name'] = '未提供工廠名稱'
                if 'Factory ID' not in prod_data.columns: prod_data['Factory ID'] = ''
                
                def make_maker(row):
                    fid = str(row.get('Factory ID', '')).replace('.0', '').strip()
                    fname = str(row.get('Factory Name', '')).strip()
                    if fid and fid != 'nan': return f"{fid}-{fname}"
                    return fname
                prod_data['Maker'] = prod_data.apply(make_maker, axis=1)

                # 💡 處理 Packaging (Retail Packaging Format (1) *)
                if 'Retail Packaging Format (1) *' in prod_data.columns:
                    prod_data['Packaging'] = prod_data['Retail Packaging Format (1) *'].fillna('')
                else:
                    prod_data['Packaging'] = ''

                # 新增預留空白欄位
                prod_data['IMAGE'] = ''
                prod_data['Age'] = ''
                prod_data['Assortment'] = ''

                # 定義最終左側的排序與所需欄位
                desired_left_columns = [
                    'DPCI', 'Manufacturer Style # *', 'IMAGE', 'Product Description', 
                    'Barcode', 'Primary Raw Material Type', 'Age', 'Maker', 
                    'Packaging', 'Inner Pack Unit Quantity', 'Case Unit Quantity', 'Assortment',
                    'Import Vendor Name', 'Factory Name'
                ]
                
                # 容錯機制補齊缺少的欄位
                for col in desired_left_columns:
                    if col not in prod_data.columns:
                        prod_data[col] = ''
                            
                left_data = prod_data[desired_left_columns + ['DPCI_MERGE']].drop_duplicates(subset=['DPCI_MERGE']).set_index('DPCI_MERGE')
                
                # 💡 建立 5 層靜態表頭 Tuple
                def get_left_tuple(col):
                    if col == 'DPCI': return ('Program Name', 'DPCI', '', '', '') # A1 為 Program Name
                    elif col == 'Barcode': return ('', 'UPC', '', '', '') # Barcode 改名為 UPC
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
                            
                            # 自動過濾數量為 0 的 PO 欄位 (PO NUMBER 在 Level 2)
                            cols_to_keep = []
                            for col in export_data.columns:
                                if col[2] != '': # 若為 PO 欄位
                                    if pd.to_numeric(export_data[col], errors='coerce').sum() > 0:
                                        cols_to_keep.append(col)
                                else:
                                    cols_to_keep.append(col)
                                    
                            export_data = export_data[cols_to_keep]
                            
                            # 加上隱形空白，強迫每個 PO 單獨顯示 PURPOSE 標籤
                            unmerged_columns = []
                            po_idx = 0
                            for col in export_data.columns:
                                if col[2] != '': # 若為 PO 欄位
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
                
                st.success("✨ 處理完成！版面已完全按照您的要求升級調整。")
                st.download_button(
                    label="📦 點擊下載合併版 PO GRID (ZIP壓縮檔)",
                    data=zip_buffer.getvalue(),
                    file_name="PO_GRIDs_Output.zip",
                    mime="application/zip"
                )
                
            except Exception as e:
                st.error(f"❌ 處理過程中發生錯誤: {e}")
