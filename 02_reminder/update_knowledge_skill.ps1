# update_knowledge_skill.ps1
# Program 3: 維護知識庫
# 定期讀取 Program 2 的日誌，分析常見問題模式，更新 skill 知識庫
#
# 使用方式:
#   .\update_knowledge_skill.ps1
#   .\update_knowledge_skill.ps1 -AnalyzeOnly

param(
    [switch]$AnalyzeOnly   # 只分析，不寫入
)

$ErrorActionPreference = "Stop"

$baseDir       = "E:\88. Claude\02_reminder"
$fixLogFile    = "$baseDir\fix_log.txt"
$skillDir      = "E:\88. Claude\.claude\skills\download-troubleshooting"
$skillFile     = "$skillDir\SKILL.md"
$statsFile     = "$baseDir\fix_stats.json"

# ============================================
# 日誌函數
# ============================================
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "SUCCESS", "ERROR", "WARN")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    Write-Host $logEntry -ForegroundColor $(
        switch ($Level) {
            "SUCCESS" { "Green" }
            "ERROR"   { "Red" }
            "WARN"    { "Yellow" }
            default   { "Cyan" }
        }
    )
}

# ============================================
# 確保 skill 目錄存在
# ============================================
if (-not (Test-Path $skillDir)) {
    Write-Log "建立 skill 目錄: $skillDir"
    New-Item -Path $skillDir -ItemType Directory -Force | Out-Null
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host " Update Knowledge Base (Program 3)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# ============================================
# 讀取並分析日誌
# ============================================
Write-Log "分析修復日誌..."

$stats = @{
    total_attempts = 0
    successful_fixes = 0
    failed_fixes = 0
    monthly_attempts = 0
    weekly_attempts = 0
    monthly_success = 0
    weekly_success = 0
    errors = @()
    last_updated = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
}

if (Test-Path $fixLogFile) {
    $logLines = @(Get-Content $fixLogFile -ErrorAction SilentlyContinue)

    foreach ($line in $logLines) {
        # 格式: [時間] [狀態] 訊息
        if ($line -match '\[.*?\] \[(.*?)\] (.*)') {
            $level = $matches[1]
            $message = $matches[2]

            # 統計嘗試次數
            if ($message -match '開始修復 (monthly|weekly)') {
                $type = $matches[1]
                $stats.total_attempts++
                if ($type -eq "monthly") { $stats.monthly_attempts++ }
                else { $stats.weekly_attempts++ }
            }

            # 統計成功修復
            if ($message -match '(Monthly|Weekly) 下載修復成功' -and $level -eq "SUCCESS") {
                $stats.successful_fixes++
                if ($message -match 'Monthly') { $stats.monthly_success++ }
                else { $stats.weekly_success++ }
            }

            # 記錄失敗
            if ($message -match '修復失敗|例外錯誤' -and $level -eq "ERROR") {
                $stats.failed_fixes++
                $stats.errors += @{
                    message = $message
                    time = Get-Date
                }
            }
        }
    }
}

# ============================================
# 計算統計
# ============================================
$successRate = if ($stats.total_attempts -gt 0) {
    [math]::Round(($stats.successful_fixes / $stats.total_attempts) * 100, 2)
} else {
    0
}

$monthlySuccessRate = if ($stats.monthly_attempts -gt 0) {
    [math]::Round(($stats.monthly_success / $stats.monthly_attempts) * 100, 2)
} else {
    0
}

$weeklySuccessRate = if ($stats.weekly_attempts -gt 0) {
    [math]::Round(($stats.weekly_success / $stats.weekly_attempts) * 100, 2)
} else {
    0
}

Write-Log "統計完成："
Write-Log "  總嘗試: $($stats.total_attempts) | 成功: $($stats.successful_fixes) | 失敗: $($stats.failed_fixes) | 成功率: $successRate%"
Write-Log "  月報: $($stats.monthly_attempts) 次, 成功率 $monthlySuccessRate%"
Write-Log "  週報: $($stats.weekly_attempts) 次, 成功率 $weeklySuccessRate%"

# ============================================
# 保存統計數據
# ============================================
$stats | ConvertTo-Json | Set-Content $statsFile -ErrorAction SilentlyContinue

# ============================================
# 生成或更新 SKILL.md
# ============================================
if (-not $AnalyzeOnly) {
    Write-Log "更新 skill 知識庫..."

    $skillContent = @"
---
name: download-troubleshooting
description: 驗布報表自動下載問題診斷與修復 — 自動檢測月報/週報下載失敗，並嘗試自動修復。由程序 1-3 自動維護。
---

# 驗布報表下載自動診斷與修復

三層自動維運系統：
- **程序 1 (監控)**: check_monthly_download.ps1 / check_weekly_download.ps1
- **程序 2 (修復)**: auto_fix_downloads.ps1
- **程序 3 (知識庫)**: update_knowledge_skill.ps1

## 系統健康狀況

📊 **最後更新**: $($stats.last_updated)

| 指標 | 值 |
|------|-----|
| 總修復嘗試 | $($stats.total_attempts) |
| 成功修復 | $($stats.successful_fixes) |
| 失敗修復 | $($stats.failed_fixes) |
| **總成功率** | **$successRate%** |

### 按報表類型

| 報表類型 | 嘗試次數 | 成功次數 | 成功率 |
|---------|---------|--------|-------|
| 月報 | $($stats.monthly_attempts) | $($stats.monthly_success) | $monthlySuccessRate% |
| 週報 | $($stats.weekly_attempts) | $($stats.weekly_success) | $weeklySuccessRate% |

## 監控排程

- **月報檢查**: 每週一、二、三 13:00 (檢查上上週四~上週三的月報)
- **週報檢查**: 每週四 13:00 (檢查上週的週報)
- **自動修復**: 檢測失敗時立即觸發
- **知識庫更新**: 定期分析修復日誌

## 常見問題與解決方案

### 問題 1: Python 下載失敗

**症狀**: `auto_download_http.py` 執行失敗

**可能原因**:
- 網路連線問題
- Python 環境未正確設定
- 目標伺服器無法連接

**解決方案**:
1. 檢查網路連線
2. 確認 Python 已安裝並在 PATH 中
3. 檢查 auto_download_http.py 的日誌輸出

### 問題 2: 檔案鎖定

**症狀**: 檔案移動失敗，"檔案正被其他程序使用"

**可能原因**:
- Excel 或其他程式開啟了該檔案
- 上次的轉換未完全結束

**解決方案**:
1. 關閉所有開啟 Excel 檔案的程式
2. 手動刪除暫存檔 (如有)
3. 重新執行修復

### 問題 3: 日期計算錯誤

**症狀**: 下載了錯誤日期範圍的報表

**可能原因**:
- 系統時間設定不正確
- 腳本中的日期邏輯錯誤

**解決方案**:
1. 確認系統時間正確
2. 檢查 check_*.ps1 中的日期計算邏輯
3. 如需修正，聯絡系統管理員

## 修復日誌位置

- **修復日誌**: `E:\88. Claude\02_reminder\fix_log.txt`
- **統計數據**: `E:\88. Claude\02_reminder\fix_stats.json`
- **月報自動下載**: `E:\88. Claude\03_download rawdata\月報自動.ps1`
- **週報自動下載**: `E:\88. Claude\03_download rawdata\週報自動.ps1`

## 手動觸發修復

若需手動觸發修復，可執行:

```powershell
# 月報修復
powershell -NoProfile -ExecutionPolicy Bypass -File "E:\88. Claude\02_reminder\auto_fix_downloads.ps1" -Type "monthly"

# 週報修復
powershell -NoProfile -ExecutionPolicy Bypass -File "E:\88. Claude\02_reminder\auto_fix_downloads.ps1" -Type "weekly"
```

## 系統架構

\`\`\`
程序 1 (監控)
    └─ 檢測下載失敗
         ↓
程序 2 (修復)
    └─ 自動執行月報/週報自動下載
         ↓
程序 3 (知識庫)
    └─ 分析日誌 → 更新本 skill
         ↓
    沉澱經驗，持續優化
\`\`\`

---

*最後由程序 3 自動更新於 $($stats.last_updated)*
"@

    $skillContent | Set-Content $skillFile -Encoding UTF8 -ErrorAction SilentlyContinue
    Write-Log "已更新 skill 檔案: $skillFile" "SUCCESS"
}

Write-Log "知識庫維護完成" "SUCCESS"
Write-Host "`n✅ 知識庫已更新！" -ForegroundColor Green

exit 0
