# 監控-診斷-修復 三程序自動維運系統

## 🎯 系統概述

本系統實現了驗布報表下載的**自動監控、診斷、修復、知識積累**的完整維運循環。由三個 PowerShell 程序組成：

```
[程序 1: 監控]
     ↓
檢測下載失敗?
     ├─ 成功 → 結束
     └─ 失敗 ↓
          [程序 2: 自動修復]
               ↓
          嘗試自動修復
               ├─ 成功 → 結束
               └─ 失敗 ↓
                    [程序 3: 知識庫更新]
                         ↓
                    分析日誌 → 更新 skill
                    積累經驗
```

---

## 📋 程序說明

### 程序 1: 監控（Program 1: Monitoring）

**檔案**:
- `check_monthly_download.ps1` — 每週一、二、三 13:00 執行
- `check_weekly_download.ps1` — 每週四 13:00 執行

**功能**:
1. 檢查月報/週報是否按時下載
2. 若失敗，自動呼叫程序 2
3. 若成功或修復成功，退出
4. 呼叫程序 3 更新知識庫

**排程設定**:
```powershell
# 建立或更新排程，執行此指令
.\setup_schedules.ps1
```

**手動執行**:
```powershell
# 月報檢查
.\check_monthly_download.ps1

# 週報檢查
.\check_weekly_download.ps1
```

**退出代碼**:
- `0` = 成功（檔案已下載或已修復）
- `1` = 失敗（修復未成功）
- `2` = 錯誤（腳本缺失或其他嚴重問題）

---

### 程序 2: 自動修復（Program 2: Auto-Fix）

**檔案**: `auto_fix_downloads.ps1`

**功能**:
1. 接收程序 1 的失敗通知
2. 調用月報或週報自動下載腳本
3. 檢查修復是否成功
4. 記錄所有操作到 `fix_log.txt`

**使用方式**:
```powershell
# 修復月報下載
.\auto_fix_downloads.ps1 -Type "monthly"

# 修復週報下載
.\auto_fix_downloads.ps1 -Type "weekly"

# 快速模式（略過某些驗證）
.\auto_fix_downloads.ps1 -Type "monthly" -Fast
```

**日誌檔案**:
- `fix_log.txt` — 所有修復操作的記錄

**退出代碼**:
- `0` = 成功修復
- `1` = 修復失敗
- `2` = 發生例外錯誤

---

### 程序 3: 知識庫更新（Program 3: Knowledge Base Update）

**檔案**: `update_knowledge_skill.ps1`

**功能**:
1. 定期或被動讀取 `fix_log.txt`
2. 分析修復嘗試的統計數據
3. 更新或創建 skill 知識庫
4. 記錄常見問題與解決方案

**使用方式**:
```powershell
# 執行完整更新
.\update_knowledge_skill.ps1

# 僅分析，不寫入檔案
.\update_knowledge_skill.ps1 -AnalyzeOnly
```

**輸出檔案**:
- `.\.claude\skills\download-troubleshooting\SKILL.md` — 知識庫文件
- `fix_stats.json` — 統計數據（JSON 格式）

---

## 📊 系統流程

### 正常情況（成功下載）

```
1. check_monthly_download.ps1 執行
2. 檢查：月報檔案存在？
3. ✅ 存在 → 程序 1 結束 (exit 0)
4. 呼叫程序 3 更新知識庫（定期統計）
```

### 異常情況（下載失敗）

```
1. check_monthly_download.ps1 執行
2. 檢查：月報檔案存在？
3. ❌ 不存在 → 偵測失敗
4. 自動呼叫程序 2: auto_fix_downloads.ps1 -Type "monthly"
5. 程序 2 執行月報自動下載腳本
   ├─ ✅ 成功 → 記錄成功，exit 0
   └─ ❌ 失敗 → 記錄失敗，exit 1
6. 呼叫程序 3 分析日誌並更新知識庫
7. 將結果返回程序 1
8. 程序 1 根據結果決定是否繼續執行其他任務
```

---

## 🔧 安裝與設定

### 前置條件

- Windows PowerShell 5.0 或更新版本
- 系統管理員權限（用於建立排程任務）
- 月報/週報自動下載腳本已安裝（在 `03_download rawdata` 資料夾）

### 初始安裝

1. **驗證腳本存在**:
   ```bash
   ls -la 02_reminder/
   ls -la 03_download rawdata/
   ```

2. **建立排程任務**:
   ```powershell
   cd E:\88. Claude\02_reminder
   .\setup_schedules.ps1
   ```

3. **驗證排程已建立**:
   ```powershell
   Get-ScheduledTask -TaskPath "\Claude\02_reminder\"
   ```

4. **檢查日誌目錄**:
   - `fix_log.txt` — 程序 2 和 3 的操作日誌
   - `fix_stats.json` — 統計數據

---

## 📈 監控與調試

### 檢查系統狀態

```powershell
# 查看所有排程任務
Get-ScheduledTask -TaskPath "\Claude\02_reminder\" | Format-Table

# 查看排程執行歷史
Get-ScheduledTaskInfo -TaskPath "\Claude\02_reminder\" -TaskName "CheckMonthlyDownload"

# 查看修復日誌
Get-Content "E:\88. Claude\02_reminder\fix_log.txt" -Tail 20

# 查看統計數據
Get-Content "E:\88. Claude\02_reminder\fix_stats.json" | ConvertFrom-Json
```

### 手動測試

```powershell
# 測試月報檢查（會自動修復如果失敗）
.\check_monthly_download.ps1

# 測試週報檢查（會自動修復如果失敗）
.\check_weekly_download.ps1

# 測試程序 2（修復）
.\auto_fix_downloads.ps1 -Type "monthly"

# 測試程序 3（知識庫更新）
.\update_knowledge_skill.ps1 -AnalyzeOnly
```

### 常見問題排除

#### 問題 1: 排程任務無法執行

**症狀**: 排程未在指定時間執行

**排查步驟**:
1. 確認排程服務已啟動: `Get-Service Schedule | Select Status`
2. 檢查排程任務的「觸發器」設定
3. 查看事件檢視器 (Windows Logs > System)
4. 確認 PowerShell 執行原則: `Get-ExecutionPolicy`

#### 問題 2: 修復失敗

**症狀**: `fix_log.txt` 中有 `[ERROR] 修復失敗`

**排查步驟**:
1. 查看完整的修復日誌（含錯誤訊息）
2. 手動執行: `.\auto_fix_downloads.ps1 -Type "monthly"`
3. 檢查網路連線（月報/週報需從線上下載）
4. 檢查磁碟空間是否充足

#### 問題 3: 知識庫無法更新

**症狀**: `.claude\skills\download-troubleshooting\SKILL.md` 未更新

**排查步驟**:
1. 確認資料夾存在: `Test-Path "E:\88. Claude\.claude\skills\download-troubleshooting"`
2. 檢查檔案權限（是否可寫入）
3. 手動執行: `.\update_knowledge_skill.ps1`

---

## 📁 檔案結構

```
02_reminder/
├── README.md                           ← 本文件
├── setup_schedules.ps1                 ← 排程建立腳本
├── check_monthly_download.ps1          ← 程序 1 (月報)
├── check_weekly_download.ps1           ← 程序 1 (週報)
├── auto_fix_downloads.ps1              ← 程序 2 (自動修復)
├── update_knowledge_skill.ps1          ← 程序 3 (知識庫)
├── send_weekly_email.ps1               ← 寄送週報郵件
├── send_log.txt                        ← 郵件寄送日誌
├── fix_log.txt                         ← 修復操作日誌 (自動生成)
├── fix_stats.json                      ← 修復統計 (自動生成)
└── last_sent_week.txt                  ← 上次寄送週報的日期

.claude/skills/
└── download-troubleshooting/
    └── SKILL.md                        ← 知識庫 (自動更新)

03_download rawdata/
├── 月報自動.ps1                        ← 月報自動下載
├── 週報自動.ps1                        ← 週報自動下載
└── ...其他文件
```

---

## 🚀 最佳實踐

### 1. 定期檢查知識庫

每週檢查一次 SKILL.md，確保知識庫反映了最新的問題模式：

```powershell
# 查看最新知識庫
Get-Content "E:\88. Claude\.claude\skills\download-troubleshooting\SKILL.md"
```

### 2. 監控修復成功率

定期檢查修復統計，識別系統性問題：

```powershell
# 查看修復成功率
$stats = Get-Content "E:\88. Claude\02_reminder\fix_stats.json" | ConvertFrom-Json
$rate = ($stats.successful_fixes / $stats.total_attempts) * 100
Write-Host "修復成功率: $rate%"
```

### 3. 備份日誌檔案

建議每月備份 `fix_log.txt` 和 `fix_stats.json`：

```powershell
$date = Get-Date -Format "yyyy-MM"
Copy-Item "fix_log.txt" "fix_log_$date.bak"
Copy-Item "fix_stats.json" "fix_stats_$date.bak"
```

### 4. 自動通知（可選）

可在程序 2 失敗時發送郵件通知（需配置郵件設定）

---

## 📝 日誌格式

### fix_log.txt 範例

```
[2026-08-08 13:00:15] [INFO] 開始修復 monthly 報表下載失敗
[2026-08-08 13:00:16] [INFO] 執行 Monthly 自動下載腳本...
[2026-08-08 13:00:45] [SUCCESS] Monthly 下載修復成功
```

### fix_stats.json 範例

```json
{
  "total_attempts": 12,
  "successful_fixes": 10,
  "failed_fixes": 2,
  "monthly_attempts": 8,
  "monthly_success": 7,
  "weekly_attempts": 4,
  "weekly_success": 3,
  "last_updated": "2026-08-08 13:01:00",
  "errors": []
}
```

---

## 🔄 持續改進

### 程序 3 的自動化改進

程序 3 會根據修復日誌自動：
- 統計月報/週報的成功率
- 識別常見錯誤類型
- 建議改進方案
- 更新知識庫 SKILL.md

### 手動優化

若發現新的問題或解決方案：
1. 編輯 SKILL.md
2. 添加新的問題-解決方案對
3. 下次程序 3 執行時會保留手動編輯

---

## 📞 支持

如有問題，請檢查：
1. `fix_log.txt` — 修復操作日誌
2. `fix_stats.json` — 統計數據
3. SKILL.md — 已知問題與解決方案
4. Windows 事件檢視器 — 系統錯誤

---

**最後更新**: 2026-08-08  
**版本**: 1.0  
**系統架構**: 監控-診斷-修復三程序自動維運
