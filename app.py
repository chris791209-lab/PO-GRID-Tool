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

                # --- 資料清洗與母子分離 ---
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
                        # 母 DPCI
                        parent_dpci_list.add(dpci)
                        
                        style_val = str(row['VENDOR STYLE']).strip() if pd.notna(row['VENDOR STYLE']) else ''
                        if style_val and not style_val.upper().startswith('ASSORT'):
                            style_val = f"ASSORTMENT-{style_val}"
                            
                        parent_info_dict[dpci] = {
                            'style': style_val,
                            'upc': str(row['ITEM BAR CODE']).strip() if pd.notna(row['ITEM BAR CODE']) else ''
