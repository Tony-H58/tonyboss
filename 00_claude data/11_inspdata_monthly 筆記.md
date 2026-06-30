# 11_inspdata_monthly 筆記

## 路徑結構

| 項目 | 路徑/檔名 |
|------|------|
| 主資料夾 | `E:\88. Claude\11_inspdata_monthly\` |
| 原始資料 | `insprawdata_monthly\` |
| 已處理原始資料 | `insprawdata_monthly\completed rawdata\` |
| QF 輸出資料夾 | `insprecord_QF\` |
| **主腳本** | `run_monthly_full.ps1` 一鍵執行全流程 |
| 匯入腳本 | `import_monthly.ps1` |
| 後處理腳本 | `step4_fix_inspqty.ps1`, `step5_bl_br.ps1`, `step6_bs_cw.ps1`, `step7_format.ps1` |
| 對照表 | `00_plan_insprecord_QF_monthly.xlsx`（Plan、ChartTitle、NameList 頁） |
| QF 工作檔 | `insprecord_QF\insprecord_QF_current.xlsx`（永久保留，勿刪） |

---

## 執行方式（每月手動觸發）

> **慣例：說「執行」= 自動掃描原始檔、判斷月份、直接執行匯入，不詢問**
> **自動分析檔名日期決定 -month 參數；若有跨年資料，分別處理不同年度的 QF 檔**
> **若 insprawdata_monthly\ 根目錄無未處理檔案，告知使用者放入資料後再執行**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "E:\88. Claude\11_inspdata_monthly\run_monthly_full.ps1" -month "2606"
```

### 自動偵測月份邏輯
1. 掃描 `insprawdata_monthly\` 根目錄（排除子資料夾）未處理的 `.xlsx` 檔
2. 從檔名或最後修改日期推斷月份（YYMM 格式）
3. 若跨年（如同時有 2612 和 2701 資料）→ 先執行舊年，再處理新年 QF 檔
4. 跨年時需複製舊 QF 檔改名為 `insprecord_QF_2026.xlsx`，新 current.xlsx 清空資料後再匯入新年

---

## 效能優化（2026-06）
- import_monthly.ps1：字串欄位逐欄批次寫入（76 次範圍寫入取代 7萬+ 次單格寫入）
- Step 4~6：整批讀取陣列，記憶體計算，整批寫回（取代逐格讀寫）
- Step 4~7：只處理新增行（透過 startRow 參數，import_monthly 寫入 temp_startrow.txt）
- Step 8：改用 Where-Object 比對，解決中文檔名 -Filter 失效問題

## 主腳本流程（run_monthly_import.ps1 自動執行）
```
Step 0：將日期命名檔改回 insprecord_QF_current.xlsx
Step 1~3：執行 import9.ps1
          - 只處理根目錄沒有 "Imported_" 前綴的檔案（已移入 completed rawdata 者不處理）
          - 品管 + 工廠 資料寫入 QF，進行欄位對照
          - NameList 對照（客戶、廠商、供應地/QC）
          - 資料清理（EngOnly、數字、CleanFiber）
          - 依 B 欄日期範圍自動改名：insprecord_QF_260102-260430.xlsx
Step 4：  抽驗碼數 = A+B+C，重算 A/B/C% 及相關欄位
Step 5：  BL~BR 計數（單色結論為 Fail）
Step 6：  BS~CW 計數（單色結論為 Fail，Contains 比對名稱）
          　未分類項目統一計入 CW，並於執行時顯示清單
Step 7：  格式化（標題置中、文字靠左、數字靠右）
Step 8：  原始資料移入 completed rawdata（不加前綴，檔名不變）
```

**原始資料判斷規則**
- 無 `Imported_` 前綴 → 視為待處理，移入 completed rawdata
- 有 `Imported_` 前綴 → 舊格式，跳過不處理（理論上不應出現）

---

## QF 工作檔說明
- `insprecord_QF_current.xlsx` = 永久工作檔（含 101 欄標題），勿刪
- 輸出後改名為 `insprecord_QF_YYMMDD-YYMMDD.xlsx`
- 下次執行時自動將日期命名檔改回 current.xlsx 再追加
- **若需重建工作檔**：從 00_plan_insprecord_QF.xlsx 的 ChartTitle 頁建立

## 跨年處理步驟（手動）
1. 確認舊年度全年資料已匯入完成
2. 手動將日期命名檔改名為 `insprecord_QF_2026.xlsx`（年份檔）
3. **複製** `insprecord_QF_2026.xlsx`，改名為 `insprecord_QF_current.xlsx`
4. 清空 `insprecord_QF_current.xlsx` 的資料列（**保留第 1 行標題**）
5. 執行新年度第一個月匯入（如 `-month "2701"`）

---

## BL~BR 計數規則

條件：col25（Y）= "Fail"，AI(col35) = 百碼瑕疵點數

| 欄位 | 來源欄位 | 計算 |
|------|--------|------|
| BL(64) 外觀 | AP(42) 外觀不良 | 有值且 >33 加1，≤33 加0.5 |
| BM(65) 縮率 | AQ(43) 縮率不良 | 有值則 +1 |
| BN(66) 色差 | AR(44) 色差不良 | 有值則 +1 |
| BO(67) 手感 | AS(45) 手感不良 | 有值則 +1 |
| BP(68) 色牢 | AT(46) 色牢不良 | 有值則 +1 |
| BQ(69) 規格 | AU(47) 規格不良 | 有值則 +1 |
| BR(70) 其他 | AV(48) 其他不良 | 有值則 +1 |

---

## BS~CW 計數規則

條件：col25（Y）= "Fail"，各欄用 Contains 比對名稱
| 來源欄位 | 目標欄位 | 計算 |
|--------|--------|------|
| BS~CE (71~83) | AP(42) 外觀不良 | >33 加1，≤33 加0.5 |
| CF~CH (84~86) | AQ(43) 縮率不良 | +1 |
| CI (87) | AR(44) 色差不良 | +1 |
| CJ (88) | AS(45) 手感不良 | +1 |
| CK~CO (89~93) | AT(46) 色牢不良 | +1 |
| CP~CR (94~96) | AU(47) 規格不良 | +1 |
| CS~CV (97~100) | AV(48) 其他不良 | +1 |
| CW(101) 未分類 | AP~AV 均未命中 | 符合 Fail 條件 +1，執行時顯示清單 |

**注意事項：** 污(U+6C61)與汙(U+6C59)視為相同，腳本內自動替換

---

## 各月資料量

| 月份 | 筆數 | 狀態 |
|------|------|------|
| 1月（2601） | 5,939 | 已匯入 |
| 2月（2602） | 4,115 | 已匯入 |
| 3月（2603） | 4,920 | 已匯入 |
| 4月（2604） | 5,552 | 已匯入 |
| **合計** | **20,526** | 完成 |
| 5月（2605） | 5,745 | 已匯入 |
| **合計** | **26,271** | 完成 |
| 6月（2606） | - | 待執行 |

---

## 重要欄位對照

| QF 欄位 | 說明 |
|-------|------|
| A(1) | 來源（品管/工廠原文） |
| B(2) | 驗布日期（格式 yyyy/mm/dd，作為檔名依據） |
| H(8) | 客戶簡稱（NameList E~F 對照） |
| Y(25) | 單色結論（Fail/Pass/Discuss） |
| AI(35) | 百碼瑕疵點數 |

---

## 常見問題

**Q: 如何重建 QF 工作檔**
→ 從 00_plan_insprecord_QF.xlsx 的 ChartTitle 頁建立，需有 101 欄標題

**Q: BS~CV 計數全為 0**
→ 確認 QF 第 1 行有正確標題，需有 101 欄

**Q: 追加時起始行錯誤**
→ import9.ps1 掃描 A 欄找實際最後 lastDataRow

**Q: 客戶名稱未對照**
→ 確認 NameList E~F 有正確資料，loop 從 row 2 開始

---

## NameList 欄位對照

| 欄位 | 內容 |
|------|------|
| A~B | 客戶原名→簡稱 |
| C~D | 廠商原名→簡稱 |
| E~F | 供應地對照（如 CAMBODIA→CAM） |
| G~H | 工廠原名→簡稱 |
| I~J | QC 對照 |

---

## Skill 指令

`/task01-import-insp-record`
