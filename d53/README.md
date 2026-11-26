# D53 PO GRID｜使用說明 SOP

在 App 側欄選 **部門 → D53**。D240 的流程不受影響。

## 新建 GRID｜Build a new GRID
1. 選 AE（只用在檔名與標題）。
2. 上傳 **PCN**：每份 PCN 就是一個 Program，可以一次上傳多份。
3. 上傳 **PO PDF**：整包丟進來就好，可多選或用 .zip，不用先挑。
4. （選填）上傳 **WRK** Workspace 匯出檔，用來放照片。
5. 確認 **本季起始日**，預設是 PCN 檔名年份的 1/1。
6. 選 **品項範圍**：
   - 整份 PCN（預設）
   - **只納入我負責的**：把自己的 DPCI 或 Vendor Style# 從 Excel 整欄複製貼上（換行、逗號、空白都可以），也可以勾整家工廠。找不到的編號會列出來提醒。之後 PCN 新版多出來的 DPCI 不會自動納入。
   - 排除少數品項／工廠
7. 按「產生 PO GRID」，下載檔案。

## 更新 GRID｜Update
上傳 App 產生的 GRID，再加上任一項：新的 PO PDF（新單、改單、CANCEL 都可以）、新的 PCN（加 Program，或同名 PCN 的新版本）、WRK（補照片）。
系統會依 PO 版本自動處理：新單會加一欄；改單會覆蓋數量；收到 CANCEL 的 PO 會排除。每次變動都會寫進「更新紀錄」分頁。

## 範圍規則｜Scope rules
| 關卡 | 規則 |
|---|---|
| 品項 | DPCI 在上傳的 PCN 裡，並符合品項範圍設定 |
| 季別 | Purpose 不是 BAS，而且 Ship Window 開始日 ≥ 本季起始日 |
| 版本 | 同一張 PO 取 SPS 列印時間最新的那份；收到 Import PO Change CANCEL 就整張排除 |

沒通過關卡的行全部列在「排除清單」分頁，並寫明原因。

## 核對項目｜Validation checklist
Validation Summary 第 0 節會列出每一項的「已檢查／差異」，逐筆明細在「Validation Detail」。
| 項目 | 規則 |
|---|---|
| Unit Cost$ | PO 單價 = PCN 成本（FCA／FOB） |
| Retail$ | PO Resale = PCN Suggested Unit Retail |
| Case Pack | PO # of Inners × Units per Inner = PCN Case；Units per Inner = PCN Inner |
| 數量｜整箱 | PO 數量是 Case Pack 的整數倍 |
| 數量｜vs Commit | 全部 PO 合計 ≤ Commit × 110% |
| Factory | PO 上的工廠 = PCN Factory Name |
| PO 解析驗算 | 每張 PO 行數量合計 = Total Qty |
| Image | WRK 有照片 |

## GRID 版面｜Layout
- 每個 Program（PCN）一個分頁，版面參照 Vanessa 的 PO Amount Grid。
- 第 1–8 列是 PO 表頭：PO No.、Factory、S/C（手填）、PO Date、Assigned by Sender（Purpose）、AC（手填）、Shipping Point、Shipping Window。
- 第 9 列是欄名：Product Description｜Factory｜Image｜DPCI｜STYLE｜Case/Inner｜Retail$｜Unit Cost$｜各張 PO｜Q'TY｜Forecasted Units（Commit）｜+-20%｜差異 %｜核對｜AE 備註。
- 品項依 Factory 分區塊，PO 欄也依同樣的工廠順序排列，工廠之間用粗線隔開。
- 舊版（每間工廠一頁）的 GRID 一樣可以拿來「更新」，產生的新檔會改成新版面，黃底資料會帶過去。

## 注意｜Notes
- 只有**黃底欄位**（S/C、AC、AE 備註）會在更新時保留，其他欄位請不要手寫。
- 不要刪除或改名隱藏頁 `_D53_STATE`、`_D53_PHOTOS`，更新要靠這兩頁。
- 成本以 PO 為準。PO 單價跟 PCN 成本不同的格子會標紅，英文明細在「Cost Discrepancy」分頁。

## 程式｜Code
- `d53/engine.py`：PDF 解析、PCN／WRK 讀取、範圍判斷、產生 GRID、讀回 GRID、合併更新
- `d53/page.py`：Streamlit 頁面
- 驗證：1120 → 1123 → 1124 分三批更新，結果和一次全跑相同（87 張 PO、764 行、3,788,954 pcs）
