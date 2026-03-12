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
    # 讀取 PO List 以獲取標籤與日期
    po_list = pd.read_csv(po_list_file)
    po_raw = pd.read_csv(po_raw_file)
    
    # 處理 PO List 日期格式 (轉換為 MM/DD-MM/DD)
    po_list['SHIP BEGIN DATE'] = pd.to_datetime(po_list['SHIP BEGIN DATE'])
    po_list['SHIP END DATE'] = pd.to_datetime(po_list['SHIP END DATE'])
    po_list['SHIP_DATES'] = po_list['SHIP BEGIN DATE'].dt.strftime('%m/%d') + '-' + po_list['SHIP END DATE'].dt.strftime('%m/%d')
    
    # 擷取所需的 PO 資訊並去重複
    po_info = po_list[['PO NUMBER', 'PURPOSE', 'SHIP_DATES']].drop_duplicates()
    
    # 篩選出實際上存在於 PO RAW 中的訂單
    active_pos = po_raw['PO NUMBER'].unique()
    po_info = po_info[po_info['PO NUMBER'].isin(active_pos)].copy()
    
    # 新增一個空的欄位讓使用者輸入港口代碼
    po_info['輸入港口代碼 (如:581)'] = ""
    
    st.divider()
    st.subheader("📍 步驟 4: 請為以下 PO 分配目的地港口代碼")
    st.info("提示: 支援代碼 581(PSW), 3890(PNW), 584(ORF), 3891(SAV), 3851(NYC), 3850(OAK), 3887(HOU), 3758(CHARLESTON)")
    
    # 顯示可編輯的 Dataframe 讓使用者輸入
    edited_po_info = st.data_editor(
        po_info.reset_index(drop=True),
        use_container_width=True,
        hide_index=True,
        disabled=["PO NUMBER", "PURPOSE", "SHIP_DATES"] # 鎖定這三欄不給改，只能改代碼
    )
    
    # ==========================================
    # 3. 執行產出按鈕
    # ==========================================
    st.divider()
    if st.button("🚀 開始自動生成 PO GRID", type="primary"):
        with st.spinner("資料處理中，請稍候..."):
            try:
                # 讀取產品資料
                if prod_file.name.endswith('.csv'):
                    prod_data = pd.read_csv(prod_file)
                else:
                    prod_data = pd.read_excel(prod_file)

                # --- 資料清洗與 DPCI 合併 ---
                po_raw['DEPARTMENT'] = po_raw['DEPARTMENT'].fillna(0).astype(int).astype(str)
                po_raw['CLASS'] = po_raw['CLASS'].fillna(0).astype(int).astype(str).str.zfill(2)
                po_raw['ITEM'] = po_raw['ITEM'].fillna(0).astype(int).astype(str).str.zfill(4)
                po_raw['DPCI_MERGE'] = po_raw['DEPARTMENT'] + '-' + po_raw['CLASS'] + '-' + po_raw['ITEM']

                if po_raw['TOTAL ITEM QTY'].dtype == object:
                    po_raw['TOTAL ITEM QTY'] = po_raw['TOTAL ITEM QTY'].str.replace(',', '').astype(float)

                # --- 建立含「多層表頭」的樞紐分析表 ---
                edited_po_info['PORT_NAME'] = edited_po_info['輸入港口代碼 (如:581)'].astype(str).map(PORT_MAP).fillna(edited_po_info['輸入港口代碼 (如:581)'])
                
                po_raw_merged = po_raw.merge(edited_po_info[['PO NUMBER', 'PURPOSE', 'SHIP_DATES', 'PORT_NAME']], on='PO NUMBER', how='left')
                
                po_raw_merged['PURPOSE'] = po_raw_merged['PURPOSE'].fillna('Unknown')
                po_raw_merged['SHIP_DATES'] = po_raw_merged['SHIP_DATES'].fillna('Unknown Date')
                po_raw_merged['PORT_NAME'] = po_raw_merged['PORT_NAME'].fillna('Unknown Port')
                
                pivot_df = po_raw_merged.pivot_table(
                    index='DPCI_MERGE', 
                    columns=['PURPOSE', 'PO NUMBER', 'SHIP_DATES', 'PORT_NAME'], 
                    values='TOTAL ITEM QTY', 
                    aggfunc='sum'
                ).fillna(0)
                
                # 💡 修正點 1: 新增 PO TOTAL 時，必須配合 4 層表頭的格式 (Tuple)
                pivot_df[('PO TOTAL', '', '', '')] = pivot_df.sum(axis=1)
                
                # --- 準備左側靜態產品資料 ---
                prod_data['DPCI_MERGE'] = prod_data['DPCI'].astype(str).str.strip()
                left_columns = ['DPCI', 'Manufacturer Style # *', 'Product Description', 'Barcode', 'Primary Raw Material Type', 'Import Vendor Name', 'Inner Pack Unit Quantity', 'Case Unit Quantity']
                
                # 將 index 設為 DPCI_MERGE 以便後續無痛合併
                left_data = prod_data[left_columns + ['DPCI_MERGE']].drop_duplicates(subset=['DPCI_MERGE']).set_index('DPCI_MERGE')
                
                # 💡 修正點 2: 將左側資料的 1 層表頭擴充為 4 層表頭 (下方留空)
                left_data.columns = pd.MultiIndex.from_tuples([(col, '', '', '') for col in left_data.columns])
                
                # 💡 修正點 3: 使用 join 完美合併兩張多層表頭的表格
