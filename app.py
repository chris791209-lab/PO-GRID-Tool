import streamlit as st
import pandas as pd
import io
import zipfile

# 港口代碼對照表
PORT_MAP = {
    '581' 'PSW',
    '3890' 'PNW',
    '584' 'ORF',
    '3891' 'SAV',
    '3851' 'NYC',
    '3850' 'OAK',
    '3887' 'HOU',
    '3758' 'CHARLESTON'
}

st.set_page_config(page_title=PO GRID 自動生成系統, layout=wide)
st.title("PO GRID 自動生成器")

st.markdown(
請依序上傳 PO RAW DATA、產品資料 與 PO List (訂單清單)。
上傳後，請在下方的表格中輸入對應的「目的地港口代碼」(如 581, 3891 等)，系統將自動生成多層表頭。
)

# ==========================================
# 1. 檔案上傳區
# ==========================================
col1, col2, col3 = st.columns(3)
with col1
    po_raw_file = st.file_uploader(📁 1. PO RAW DATA (CSV), type=['csv'])
with col2
    prod_file = st.file_uploader(📁 2. 產品資料表 (ExcelCSV), type=['xlsx', 'csv'])
with col3
    po_list_file = st.file_uploader(📁 3. List of Purchase Orders (CSV), type=['csv'])

# ==========================================
# 2. 讀取並顯示「港口輸入互動表」
# ==========================================
if po_raw_file and prod_file and po_list_file
    # 讀取 PO List 以獲取標籤與日期
    po_list = pd.read_csv(po_list_file)
    po_raw = pd.read_csv(po_raw_file)
    
    # 處理 PO List 日期格式 (轉換為 MMDD-MMDD)
    po_list['SHIP BEGIN DATE'] = pd.to_datetime(po_list['SHIP BEGIN DATE'])
    po_list['SHIP END DATE'] = pd.to_datetime(po_list['SHIP END DATE'])
    po_list['SHIP_DATES'] = po_list['SHIP BEGIN DATE'].dt.strftime('%m%d') + '-' + po_list['SHIP END DATE'].dt.strftime('%m%d')
    
    # 擷取所需的 PO 資訊並去重複
    po_info = po_list[['PO NUMBER', 'PURPOSE', 'SHIP_DATES']].drop_duplicates()
    
    # 篩選出實際上存在於 PO RAW 中的訂單
    active_pos = po_raw['PO NUMBER'].unique()
    po_info = po_info[po_info['PO NUMBER'].isin(active_pos)].copy()
    
    # 新增一個空的欄位讓使用者輸入港口代碼
    po_info['輸入港口代碼 (如581)'] = 
    
    st.divider()
    st.subheader(📍 步驟 4 請為以下 PO 分配目的地港口代碼)
    st.info(提示 支援代碼 581(PSW), 3890(PNW), 584(ORF), 3891(SAV), 3851(NYC), 3850(OAK), 3887(HOU), 3758(CHARLESTON))
    
    # 顯示可編輯的 Dataframe 讓使用者輸入
    edited_po_info = st.data_editor(
        po_info.reset_index(drop=True),
        use_container_width=True,
        hide_index=True,
        disabled=[PO NUMBER, PURPOSE, SHIP_DATES] # 鎖定這三欄不給改，只能改代碼
    )
    
    # ==========================================
    # 3. 執行產出按鈕
    # ==========================================
    st.divider()
    if st.button(🚀 開始自動生成 PO GRID, type=primary)
        with st.spinner(資料處理中，請稍候...)
            try
                # 讀取產品資料
                if prod_file.name.endswith('.csv')
                    prod_data = pd.read_csv(prod_file)
                else
                    prod_data = pd.read_excel(prod_file)

                # --- 資料清洗與 DPCI 合併 ---
                po_raw['DEPARTMENT'] = po_raw['DEPARTMENT'].fillna(0).astype(int).astype(str)
                po_raw['CLASS'] = po_raw['CLASS'].fillna(0).astype(int).astype(str).str.zfill(2)
                po_raw['ITEM'] = po_raw['ITEM'].fillna(0).astype(int).astype(str).str.zfill(4)
                po_raw['DPCI_MERGE'] = po_raw['DEPARTMENT'] + '-' + po_raw['CLASS'] + '-' + po_raw['ITEM']

                if po_raw['TOTAL ITEM QTY'].dtype == object
                    po_raw['TOTAL ITEM QTY'] = po_raw['TOTAL ITEM QTY'].str.replace(',', '').astype(float)

                # --- 建立含「多層表頭」的樞紐分析表 ---
                # 1. 整理使用者輸入的港口資料
                edited_po_info['PORT_NAME'] = edited_po_info['輸入港口代碼 (如581)'].astype(str).map(PORT_MAP).fillna(edited_po_info['輸入港口代碼 (如581)'])
                
                # 將港口與日期資訊 Join 回 PO RAW
                po_raw_merged = po_raw.merge(edited_po_info[['PO NUMBER', 'PURPOSE', 'SHIP_DATES', 'PORT_NAME']], on='PO NUMBER', how='left')
                
                # 處理遺失值避免報錯
                po_raw_merged['PURPOSE'] = po_raw_merged['PURPOSE'].fillna('Unknown')
                po_raw_merged['SHIP_DATES'] = po_raw_merged['SHIP_DATES'].fillna('Unknown Date')
                po_raw_merged['PORT_NAME'] = po_raw_merged['PORT_NAME'].fillna('Unknown Port')
                
                # 建立 Pivot Table (列=DPCI, 欄=多層次(標籤PO日期港口))
                pivot_df = po_raw_merged.pivot_table(
                    index='DPCI_MERGE', 
                    columns=['PURPOSE', 'PO NUMBER', 'SHIP_DATES', 'PORT_NAME'], 
                    values='TOTAL ITEM QTY', 
                    aggfunc='sum'
                ).fillna(0)
                
                # 計算該品項的總 PO 數量 (加總所有欄位)
                pivot_df['PO TOTAL'] = pivot_df.sum(axis=1)
                
                # --- 準備左側靜態產品資料 ---
                prod_data['DPCI_MERGE'] = prod_data['DPCI'].astype(str).str.strip()
                left_columns = ['DPCI', 'Manufacturer Style # ', 'Product Description', 'Barcode', 'Primary Raw Material Type', 'Import Vendor Name', 'Inner Pack Unit Quantity', 'Case Unit Quantity']
                
                final_df = pd.merge(prod_data[left_columns + ['DPCI_MERGE']], pivot_df.reset_index(), on='DPCI_MERGE', how='inner')

                # --- 拆檔與寫入 ZIP ---
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, a, zipfile.ZIP_DEFLATED, False) as zip_file
                    
                    grouped = final_df.groupby('Import Vendor Name')
                    
                    for vendor_name, group_data in grouped
                        safe_vendor_name = str(vendor_name).replace('', '_').replace('', '_')
                        
                        # 刪除輔助用的合併欄位
                        export_data = group_data.drop(columns=['DPCI_MERGE', 'Import Vendor Name'])
                        
                        excel_buffer = io.BytesIO()
                        # 使用 ExcelWriter 匯出，Pandas 會自動將 MultiIndex 轉為漂亮的合併儲存格多層表頭
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer
                            export_data.to_excel(writer, index=False, sheet_name='PO GRID')
                        
                        zip_file.writestr(fPO_GRID_{safe_vendor_name}.xlsx, excel_buffer.getvalue())
                
                st.success(✨ 處理完成！多層次表頭與港口資訊已成功寫入。)
                st.download_button(
                    label=📦 點擊下載全廠商 PO GRID (ZIP壓縮檔),
                    data=zip_buffer.getvalue(),
                    file_name=PO_GRIDs_Output.zip,
                    mime=applicationzip
                )
                
            except Exception as e

                st.error(f❌ 處理過程中發生錯誤 {e})
