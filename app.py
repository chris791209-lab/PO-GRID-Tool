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
    
    # 💡 強制將 PO NUMBER 轉為乾淨的文字格式，避免 Unknown 匹配失敗
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

                # --- 合併港口並建立樞紐 ---
                edited_po_info['PORT_NAME'] = edited_po_info['輸入港口代碼 (如:581)'].astype(str).map(PORT_MAP).fillna(edited_po_info['輸入港口代碼 (如:581)'])
                
                # 若未輸入代碼，顯示未指定
                edited_po_info['PORT_NAME'] = edited_po_info['PORT_NAME'].replace({'': '未指定港口', 'nan': '未指定港口'})

                po_raw_merged = po_raw.merge(edited_po_info[['PO NUMBER', 'PURPOSE', 'SHIP_DATES', 'PORT_NAME']], on='PO NUMBER', how='left')
                
                po_raw_merged['PURPOSE'] = po_raw_merged['PURPOSE'].fillna('標籤遺失')
                po_raw_merged['SHIP_DATES'] = po_raw_merged['SHIP_DATES'].fillna('日期遺失')
                po_raw_merged['PORT_NAME'] = po_raw_merged['PORT_NAME'].fillna('未指定港口')
                
                pivot_df = po_raw_merged.pivot_table(
                    index='DPCI_MERGE', 
                    columns=['PURPOSE', 'PO NUMBER', 'SHIP_DATES', 'PORT_NAME'], 
                    values='TOTAL ITEM QTY', 
                    aggfunc='sum'
                ).fillna(0)
                
                pivot_df[('PO TOTAL', '', '', '')] = pivot_df.sum(axis=1)
                
                # --- 準備左側靜態產品資料 (加入 Factory Name) ---
                prod_data['DPCI_MERGE'] = prod_data['DPCI'].astype(str).str.strip()
                # 💡 在這裡加入了 'Factory Name' 欄位
                left_columns = ['DPCI', 'Manufacturer Style # *', 'Product Description', 'Barcode', 'Primary Raw Material Type', 'Import Vendor Name', 'Factory Name', 'Inner Pack Unit Quantity', 'Case Unit Quantity']
                
                # 確認 Factory Name 欄位存在，若無則建立空白欄位
                if 'Factory Name' not in prod_data.columns:
                    prod_data['Factory Name'] = '未提供工廠名稱'
                    
                left_data = prod_data[left_columns + ['DPCI_MERGE']].drop_duplicates(subset=['DPCI_MERGE']).set_index('DPCI_MERGE')
                left_data.columns = pd.MultiIndex.from_tuples([(col, '', '', '') for col in left_data.columns])
                
                final_df = left_data.join(pivot_df, how='inner')

                # --- 拆檔與寫入 ZIP (外層 Vendor，內層 Tab 工廠) ---
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    
                    vendor_col = ('Import Vendor Name', '', '', '')
                    factory_col = ('Factory Name', '', '', '')
                    
                    final_df[vendor_col] = final_df[vendor_col].fillna('未指定供應商')
                    final_df[factory_col] = final_df[factory_col].fillna('未指定工廠')
                    
                    # 第一層：依照供應商 (Vendor) 拆分 Excel 檔案
                    grouped_vendor = final_df.groupby(vendor_col)
                    
                    for vendor_name, vendor_data in grouped_vendor:
                        safe_vendor_name = str(vendor_name).replace('/', '_').replace('\\', '_')
                        
                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            
                            # 第二層：在同一個 Excel 檔案中，依照工廠 (Factory) 拆分 Tab 分頁
                            grouped_factory = vendor_data.groupby(factory_col)
                            
                            for factory_name, factory_data in grouped_factory:
                                # Excel 的 Tab 名稱限制最多 31 個字元
                                safe_factory_name = str(factory_name).replace('/', '_').replace('\\', '_')[:31] 
                                
                                # 匯出時隱藏作為分群依據的 Vendor 與 Factory 欄位，讓畫面更簡潔
                                export_data = factory_data.drop(columns=[vendor_col, factory_col])
                                export_data_reset = export_data.reset_index(drop=True)
                                
                                export_data_reset.to_excel(writer, index=True, sheet_name=safe_factory_name)
                                writer.sheets[safe_factory_name].delete_cols(1) # 刪除自動產生的 Index 欄位
                        
                        # 將完成的 Excel 檔寫入 ZIP
                        zip_file.writestr(f"PO_GRID_{safe_vendor_name}.xlsx", excel_buffer.getvalue())
                
                st.success("✨ 處理完成！已成功修復對齊問題，並完成「依供應商建檔、依工廠分子頁」的排版。")
                st.download_button(
                    label="📦 點擊下載全廠商 PO GRID (ZIP壓縮檔)",
                    data=zip_buffer.getvalue(),
                    file_name="PO_GRIDs_Output.zip",
                    mime="application/zip"
                )
                
            except Exception as e:
                st.error(f"❌ 處理過程中發生錯誤: {e}")
