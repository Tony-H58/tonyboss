---
name: weekly-report
description: 驗布週報完整流程 — 下載 rawdata、格式轉換、導入 InspRecord_QF、計算週統計、寄送週報郵件。當使用者說「跑週報」、「週報」、「週報流程」、「週度導入」、「執行週報自動」、「寄週報」，或提到 12_inspdata_weekly、run_weekly_full.ps1、週報自動.ps1 時使用。
---

# 驗布週報完整流程

在 Windows 端以 PowerShell 執行。整條流程分兩段：**下載轉換**（03 資料夾）→ **導入統計**（12 資料夾）。

## 日期週期

週報週期是 **上週四 ～ 本週三**。

腳本會自己算日期，不需要手動指定：
- 若今天是**週四**，取的是「昨天剛結束」的那個週期（上週四 ～ 昨天/週三）
- 其他天，取的是本週三結束的週期

## 完整流程

### Phase 1：下載 + 轉換

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "E:\88. Claude\03_download rawdata\週報自動.ps1"
```

加 `-Fast` 可開快速模式（關閉 Excel 螢幕更新與自動計算，減少輸出）：

```powershell
... -File "E:\88. Claude\03_download rawdata\週報自動.ps1" -Fast
```

這支腳本會依序做：
1. 算出 `startDate` / `endDate`（yymmdd）
2. 呼叫 `auto_download_http.py <startDate> <endDate>` 下載報表
3. 呼叫 `run_download_and_convert.ps1 -type weekly -startDate .. -endDate ..` 轉換 + 改名 + 搬到 `12_inspdata_weekly\insprawdata_weekly\`

**下載失敗不會中斷**，腳本會印出提示後繼續做轉換。

### Phase 2：導入 + 統計

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "E:\88. Claude\12_inspdata_weekly\run_weekly_full.ps1" -Fast
```

這支腳本會：
1. 檢查 `insprawdata_weekly\` 有沒有 rawdata
   - 有 → 執行 `run_weekly_import.ps1 -Fast` 導入 InspRecord_QF
   - 沒有但已有 QF 檔 → 跳過導入，直接用現有 QF
   - 兩者都沒有 → 報錯離開
2. 執行 `calc_weekly_stats.ps1` 算統計到 AnalysisChart_Weekly
3. 印出最終的 QF 檔名與 Chart 檔名

### Phase 3：驗證

跑完後確認這兩個檔有更新（腳本結尾會印出檔名）：

- `E:\88. Claude\12_inspdata_weekly\insprecord_QF\insprecord_QF_yymmdd-yymmdd.xlsx`
- `E:\88. Claude\12_inspdata_weekly\analysis chart_weekly\analysis chart_weekly_*.xlsx`

檢查重點：日期範圍對不對、筆數合不合理、有沒有整欄空白。

### Phase 4：寄送郵件（選用）

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "E:\88. Claude\02_reminder\send_weekly_email.ps1"
```

寄給 `tonyhuang@makalot.com.tw`，附週統計圖表 + 品管儀表板。

內建兩道保護，手動跑也不會重複寄：
- 只在**週一~週三**執行，其他天直接跳過
- 用 `last_sent_week.txt` 記錄「本週週一日期」，同一週已寄過就跳過

紀錄寫在 `02_reminder\send_log.txt`。

## 單獨執行某一步

| 目的 | 指令 |
|------|------|
| 只下載 | `python "E:\88. Claude\03_download rawdata\auto_download_http.py" <yymmdd> <yymmdd>` |
| 只轉換 | `convert_rawdata.ps1 -startDate <yymmdd> -endDate <yymmdd> [-Fast]` |
| 下載後的搬檔改名 | `run_download_and_convert.ps1 -type weekly -startDate <yymmdd> -endDate <yymmdd> [-Fast]` |
| 只導入 | `12_inspdata_weekly\run_weekly_import.ps1 -Fast` |
| 只算統計 | `12_inspdata_weekly\calc_weekly_stats.ps1` |

## 注意事項

- **`calc_weekly_stats.ps1` 會強制關閉所有 Excel 程序**（`Stop-Process -Force`）。執行前先存檔關掉手邊的 Excel。
- 統計會同時讀**週 QF**（12 資料夾）和**月 QF**（11 資料夾），兩者都必須存在，否則報錯離開。
- 週 QF 檔名必須符合 `insprecord_QF_yyMMdd-yyMMdd.xlsx` 格式，`calc_weekly_stats.ps1` 要從檔名解析日期範圍。
- 已導入的 rawdata 會被加上 `Imported_` 前綴，後續掃描會跳過。

## 自動排程

`02_reminder\setup_schedules.ps1`（需以**管理員身份**執行）會建立三個工作排程：

| 排程名稱 | 時機 | 動作 |
|---------|------|------|
| CheckWeeklyDownload | 每週四 13:00 | 檢查週報 rawdata 是否已下載 |
| CheckMonthlyDownload | 每週一~三 13:00 | 檢查月報 rawdata 是否已下載 |
| SendWeeklyEmail | 每次登入 | 寄送週報郵件 |

另外 `process_pending_rawdata.ps1` 會檢查有無未處理的 rawdata，有就自動跑週報/月報導入（用 `weekly_auto.lock` 防重入）。
