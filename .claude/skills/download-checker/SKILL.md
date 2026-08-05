---
name: download-checker
description: 週報月報下載檢查與待處理監控 — 檢查下載狀態、自動處理待處理資料、設定定時檢查排程。當使用者說「檢查下載」、「檢查週報」、「檢查月報」、「待處理資料」、「設定排程」時使用。
---

# 下載檢查與待處理監控

自動監控週報與月報的下載狀態，並在有新資料時觸發自動導入流程。

## 一、快速檢查

### 檢查週報下載

檢查今天是否已下載週報 rawdata：

```powershell
cd "E:\88. Claude\02_reminder"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File check_weekly_download.ps1
```

**輸出**：

```
[OK] 週報已下載（今日新檔：Raw=3 Done=2）
```

或

```
[提醒] 今天（2026-08-05）尚未下載週報 rawdata，請執行桌面「週報一鍵.bat」
```

### 檢查月報下載

檢查本月是否已下載月報 rawdata：

```powershell
cd "E:\88. Claude\02_reminder"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File check_monthly_download.ps1
```

**輸出**：

```
[OK] 本月已有月報 rawdata（Raw=1）
```

或

```
[提醒] 本月（8月 / August）尚未下載月報 rawdata，請執行月報一鍵指令
```

## 二、待處理資料自動檢查

### 自動觸發原理

每天 **11:00** 和 **16:00** 系統會自動執行：

```powershell
E:\88. Claude\02_reminder\process_pending_rawdata.ps1
```

此腳本會：

1. 掃描週報與月報的 rawdata 目錄
2. 找出符合 `*_yymmdd-yymmdd.*` 命名的待處理檔案
3. 排除已標記的 `Imported_*` 檔案
4. 自動啟動 `run_weekly_full.ps1` 或 `auto_check_and_import.ps1`
5. 將結果寫入日誌

### 日誌查看

```powershell
# 查看自動檢查日誌
Get-Content "E:\88. Claude\02_reminder\process_pending_log.txt" -Tail 50

# 或用 PowerShell ISE 打開
notepad "E:\88. Claude\02_reminder\process_pending_log.txt"
```

**日誌範例**：

```
[2026-08-05 11:00:23] [OK] 週報與月報均無待處理 rawdata
[2026-08-05 11:05:45] [執行] 週報有 2 個待處理檔案，啟動 run_weekly_full.ps1
[2026-08-05 11:05:45]   - insprecord_QF_260729-260804.xlsx
[2026-08-05 11:05:45]   - insprecord_QF_260804-260811.xlsx
[2026-08-05 11:12:30] [完成] 週報處理結束
```

### 手動觸發待處理檢查

如需立即檢查而不等待定時排程：

```powershell
cd "E:\88. Claude\02_reminder"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File process_pending_rawdata.ps1
```

**可能的結果**：

| 結果 | 說明 |
|------|------|
| `[OK] 週報與月報均無待處理 rawdata` | 無新檔案，無需處理 |
| `[執行] 週報有 N 個待處理檔案...` | 自動啟動 `run_weekly_full.ps1` |
| `[執行] 月報有 N 個待處理檔案...` | 自動啟動 `auto_check_and_import.ps1` |
| `[略過] 週報處理中（lock 存在）...` | 前一個處理尚未完成，跳過本次 |
| `[錯誤] ...` | 處理過程中發生問題，見詳細訊息 |

## 三、定時排程設定

### 已預設的排程

系統已設定以下 Windows 工作排程（任務排程程式）：

| 排程名稱 | 觸發時機 | 動作 | 腳本 |
|---------|---------|------|------|
| **CheckWeeklyDownload** | 每週四 13:00 | 檢查週報是否下載 | `check_weekly_download.ps1` |
| **CheckMonthlyDownload** | 週一~三 13:00 | 檢查月報是否下載 | `check_monthly_download.ps1` |
| **ProcessPendingRawdata** | 每天 11:00 / 16:00 | 自動處理待檔案 | `process_pending_rawdata.ps1` |
| **SendWeeklyEmail** | 每次登入 | 寄送週報郵件 | `send_weekly_email.ps1` |

### 檢視現有排程

```powershell
# 打開 Windows 工作排程程式
taskmgr.exe

# 或用 PowerShell 列出所有任務
Get-ScheduledTask -TaskName "*Download*" -ErrorAction SilentlyContinue | Format-Table TaskName, State, Triggers
```

### 修改排程時機

**例**：將週報檢查改為每天 09:00 執行

```powershell
$trigger = New-ScheduledTaskTrigger -Daily -At 09:00AM
Set-ScheduledTask -TaskName "CheckWeeklyDownload" -Trigger $trigger
Write-Host "✓ 排程已更新為每天 09:00"
```

### 設定新排程

若 Windows 工作排程尚未建立，執行設定腳本：

```powershell
cd "E:\88. Claude\02_reminder"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File setup_schedules.ps1
```

**需求**：需要以**管理員身份**執行

**設定步驟**：

1. 右鍵點選 PowerShell
2. 選「以系統管理員身份執行」
3. 執行上述指令

**輸出**：

```
✓ 檢查週報下載排程已建立（每週四 13:00）
✓ 檢查月報下載排程已建立（週一~三 13:00）
✓ 自動處理待檔案排程已建立（每天 11:00、16:00）
✓ 寄送週報郵件排程已建立（每次登入）
```

## 四、下載檢查邏輯

### 週報下載檢查（check_weekly_download.ps1）

**掃描位置**：

```
E:\88. Claude\12_inspdata_weekly\insprawdata_weekly
└─ completed rawdata\  （已完成的檔案）
```

**檢查條件**：

- 副檔名必須是 `.xlsx` 或 `.xls`
- 修改日期必須是**今天**
- 至少找到 1 個檔案

**判定邏輯**：

```
如果 （Raw 目錄有今日檔） 或 （completed 目錄有今日檔）
  → [OK] 週報已下載
否則
  → [提醒] 尚未下載，請執行「週報一鍵.bat」
```

### 月報下載檢查（check_monthly_download.ps1）

**掃描位置**：

```
E:\88. Claude\11_inspdata_monthly\insprawdata_monthly
└─ completed rawdata\
```

**檢查條件**：

- 副檔名必須是 `.xlsx` 或 `.xls`
- 修改日期**任何時間**（本月即可）
- 至少找到 1 個檔案

**判定邏輯**：

```
如果 （本月份有任何 xlsx/xls 檔）
  → [OK] 本月已有月報 rawdata
否則
  → [提醒] 本月尚未下載，請執行月報流程
```

## 五、待處理檔案自動導入

### 檔案命名規則

系統只會自動處理符合以下名稱的檔案：

**格式**：`*_yymmdd-yymmdd.*`

**範例**：

```
✅ 會被處理
├─ insprecord_QF_260729-260804.xlsx
├─ weeklyreport_260801-260807.xlsx
└─ data_260801-260807.xlsx

❌ 不會被處理
├─ Imported_insprecord_QF_260729-260804.xlsx  （已標記為導入過）
├─ template.xlsx  （無日期）
├─ backup_260801.xlsx  （無日期範圍）
└─ 舊檔_260801-260807.xlsx  （非英文開頭？ — 自動代理會跳過）
```

### 自動導入流程

**週報**（檢測到新週報 rawdata）：

```
process_pending_rawdata.ps1 偵測到新檔案
    ↓
建立 lock 檔案（週報\weekly_auto.lock）防止重複處理
    ↓
啟動 run_weekly_full.ps1
    │
    ├─ run_weekly_import.ps1 → 導入週 QF
    ├─ calc_weekly_stats.ps1 → 計算統計
    └─ （可選）自動寄送郵件
    ↓
刪除 lock 檔案
    ↓
日誌記錄：[完成] 週報處理結束
```

**月報**（檢測到新月報 rawdata）：

```
process_pending_rawdata.ps1 偵測到新檔案
    ↓
啟動 auto_check_and_import.ps1
    │
    ├─ 重新命名 insprecord_QF_current.xlsx 為日期格式
    ├─ 導入到年度 insprecord_QF_yymmdd-yymmdd.xlsx
    ├─ 驗證導入筆數有增加
    └─ 搬移已處理檔案到 completed rawdata
    ↓
日誌記錄：[完成] 月報處理結束
```

### 鎖定機制（lock 檔案）

為防止週報重複處理，使用 `weekly_auto.lock` 檔案：

```
E:\88. Claude\12_inspdata_weekly\weekly_auto.lock
```

**工作原理**：

```
檢查時刻
  ↓
若 lock 檔存在 → 上次處理還未完成 → 跳過本次
若 lock 檔不存在 → 無進行中的處理 → 建立 lock、開始處理
  ↓
處理結束 → 刪除 lock 檔
  ↓
下次檢查時，可再次處理
```

## 六、手動導入

如需跳過自動檢查而直接導入，可手動執行：

### 週報手動導入

```powershell
cd "E:\88. Claude\12_inspdata_weekly"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File run_weekly_full.ps1 -Fast
```

### 月報手動導入

```powershell
cd "E:\88. Claude\11_inspdata_monthly"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File auto_check_and_import.ps1
```

## 七、故障排除

| 問題 | 排查方向 | 解決方案 |
|------|---------|---------|
| 檢查腳本報錯：「找不到檔案」 | rawdata 目錄路徑錯誤 | 確認 `E:\88. Claude\12_inspdata_weekly\insprawdata_weekly` 存在 |
| 自動導入一直不執行 | 排程未啟用 | 在工作排程程式檢查 ProcessPendingRawdata 狀態 |
| lock 檔案一直存在 | 前次處理掛起 | 手動刪除 `weekly_auto.lock`，重新執行 |
| 日誌檔案找不到 | 權限問題 | 確保 `02_reminder` 資料夾有寫入權限 |
| 檢查通過但檔案沒被處理 | 檔案名稱不符規則 | 檢查檔名是否包含 `_yymmdd-yymmdd` 且非 `Imported_*` 前綴 |

## 八、週程

| 時間 | 工作 | 腳本 |
|------|------|------|
| 每週四 13:00 | 檢查週報下載 | `check_weekly_download.ps1` |
| 週一~三 13:00 | 檢查月報下載 | `check_monthly_download.ps1` |
| 每天 11:00 / 16:00 | 自動處理待檔案 | `process_pending_rawdata.ps1` |
| 每週三 11:30 | 導入最新週報 | `run_weekly_full.ps1`（自動） |
| 月度指定日 | 導入月報 | `monthly-import` skill 觸發 |

## 九、進階 — 自訂檢查規則

### 修改檢查時機

編輯 `setup_schedules.ps1`，修改觸發時間：

```powershell
# 原設定
-At 13:00  # 下午 1 點

# 改為上午 9 點
-At 09:00AM
```

### 修改掃描條件

編輯 `process_pending_rawdata.ps1`，修改檔案篩選邏輯：

```powershell
# 原設定：僅掃描 xlsx / xls
-Match "\.xlsx?$"

# 新設定：也掃描 csv
-Match "\.(xlsx?|csv)$"
```

## 十、相關文件與 Skill

- `weekly-report` — 完整週報流程（含下載轉換）
- `monthly-import` — 完整月報流程（含下載轉換）
- `run` — 手動執行單一腳本
- `Skill使用指南.md` — 工作流程總覽

---

**常見場景**：

| 我想... | 說法 / 指令 |
|--------|-----------|
| 立即檢查週報下載 | 「檢查週報」→ 執行 check_weekly_download.ps1 |
| 立即檢查月報下載 | 「檢查月報」→ 執行 check_monthly_download.ps1 |
| 手動觸發待檔案處理 | 「檢查待處理」→ 執行 process_pending_rawdata.ps1 |
| 查看自動檢查日誌 | 開啟 `02_reminder\process_pending_log.txt` |
| 設定新排程 | 需要管理員權限執行 setup_schedules.ps1 |
| 修改檢查時機 | 編輯 setup_schedules.ps1 中的觸發時間 |
