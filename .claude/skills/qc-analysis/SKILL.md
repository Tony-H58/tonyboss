---
name: qc-analysis
description: 驗布品質分析與 QC 績效評估 — 執行前後端驗布比對（比對.py）與 QC 人員 KPI 計算（qc_kpi_final.py），並驗證產出的 HTML 報告。當使用者說「比對驗布」、「前後端比對」、「品質分析」、「QC 績效」、「跑 KPI」、「績效評估」，或提到 13_compare style inspquality、14_qc performance、比對.py、qc_kpi_final.py 時使用。
---

# 驗布品質分析與 QC 績效

兩支 Python 腳本，資料來源都是 `11_inspdata_monthly\insprecord_QF\` 的月度 QF 檔。
**跑之前請先確認月報已導入最新資料**（見 `monthly-import` skill）。

## 一、前後端驗布比對

比對同一款號/採購單號的**品管**與**工廠**驗布記錄差異。

```powershell
cd "E:\88. Claude\13_compare style inspquality"
python 比對.py "<款號或採購單號>" "<月份>"
```

範例：

```powershell
python 比對.py "MS6FK213R_FA26" "5月"
python 比對.py "TMKF-26-04485" "5月"
```

**不帶參數會進入互動模式**，逐項詢問款號與月份：

```powershell
python 比對.py
# 款號或採購單號 (如 MS6FK213R_FA26 或 TMKF-26-04485):
# 月份 (如 5月):
```

### 產出

| 檔案 | 說明 |
|------|------|
| `13_compare style inspquality\00_analysis data.xlsx` | 比對明細資料 |
| `13_compare style inspquality\compare_result.html` | 比對結果報告（腳本會自動開瀏覽器） |

### 比對欄位

前後端、實際驗布日、客戶、款號、採購人員、供應商、供應地、採購單號、成衣產區、
BL NO、成份、針平織、布種、染法、QC、顏色、收料碼數、抽驗碼數、A/B/C 級碼數、
總瑕疵點數、布面瑕疵、外觀問題、手感問題、緯斜問題、顏色問題、規格問題、其他問題、
C20需出貨原因、C20解決方案、簽核狀態

## 二、QC 績效 KPI

計算 QC 人員的績效評分。

```powershell
cd "E:\88. Claude\14_qc performance"
python qc_kpi_final.py "26-1~26-6"
```

**不帶參數會互動詢問統計期間**，直接按 Enter 用預設值 `26-1~26-6`：

```powershell
python qc_kpi_final.py
# 請輸入統計期間 (例如: 26-1~26-6):
```

期間格式是 `yy-M~yy-M`，例如 `26-1~26-6` 代表 2026 年 1 月到 6 月。

### 產出

```
E:\88. Claude\14_qc performance\qc_kpi_report.html
```

腳本另外會做一份備份檔。

### 計分邏輯

分別處理**工廠**記錄與**品管**記錄（QF 第 1 欄區分）：

| 指標 | 欄位（1-based） |
|------|----------------|
| Pass/Fail | 第 25 欄 |
| 收料碼數 | 第 26 欄（僅品管記錄） |
| 抽驗碼數 | 第 27 欄（僅品管記錄） |
| C 級 % | 第 33 欄 |
| 百碼瑕疵點數 | 第 35 欄 |
| QC 姓名 | 第 23 欄（僅工廠記錄有） |

## 三、驗證產出

跑完後開 HTML 確認：

- `compare_result.html` — 前後端筆數有對上、關鍵欄位沒整欄空白
- `qc_kpi_report.html` — QC 人數合理、評分沒有 NaN 或除以零的異常值

若數字明顯不對，先回頭確認 `11_inspdata_monthly\insprecord_QF\` 的 QF 檔是不是最新、
以及該期間的月報有沒有漏導入。

## 注意事項

- 兩支腳本都是**唯讀** QF 檔，不會改動月報資料。
- `比對.py` 會呼叫 `webbrowser` 自動開啟結果，在無 GUI 環境（如遠端 session）會失敗，
  但 HTML 檔仍會正常產生。
- 輸出檔每次執行都會**覆蓋**，需要保留就先另存。
