# process_pending_rawdata.ps1
# 每天 11:00、16:00 執行：檢查是否有未處理的 rawdata，有則自動處理

$weeklyRaw     = "E:\88. Claude\12_inspdata_weekly\insprawdata_weekly"
$monthlyRaw    = "E:\88. Claude\11_inspdata_monthly\insprawdata_monthly"
$weeklyScript  = "E:\88. Claude\12_inspdata_weekly\run_weekly_full.ps1"
$monthlyScript = "E:\88. Claude\11_inspdata_monthly\auto_check_and_import.ps1"
$logFile       = "E:\88. Claude\02_reminder\process_pending_log.txt"
$weeklyLock    = "E:\88. Claude\12_inspdata_weekly\weekly_auto.lock"

function Log($msg) {
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Add-Content -Path $logFile -Value $line -Encoding utf8
    Write-Host $line
}

# 過濾條件：符合 yymmdd-yymmdd 命名、非 Imported_*
$weeklyFiles  = @(Get-ChildItem $weeklyRaw  -File -ErrorAction SilentlyContinue | Where-Object {
    $_.Name -match "_\d{6}-\d{6}" -and $_.Extension -in ".xlsx",".xls" -and $_.Name -notlike "Imported_*"
})
$monthlyFiles = @(Get-ChildItem $monthlyRaw -File -ErrorAction SilentlyContinue | Where-Object {
    $_.Name -match "_\d{6}-\d{6}" -and $_.Extension -in ".xlsx",".xls" -and $_.Name -notlike "Imported_*"
})

if ($weeklyFiles.Count -eq 0 -and $monthlyFiles.Count -eq 0) {
    Log "[OK] 週報與月報均無待處理 rawdata"
    exit 0
}

if ($weeklyFiles.Count -gt 0) {
    if (Test-Path $weeklyLock) {
        Log "[略過] 週報處理中（lock 存在），跳過"
    } else {
        New-Item $weeklyLock -ItemType File -Force | Out-Null
        try {
            Log "[執行] 週報有 $($weeklyFiles.Count) 個待處理檔案，啟動 run_weekly_full.ps1"
            $weeklyFiles | ForEach-Object { Log "  - $($_.Name)" }
            powershell.exe -NoProfile -ExecutionPolicy Bypass -File $weeklyScript
            Log "[完成] 週報處理結束"
        } finally {
            Remove-Item $weeklyLock -Force -ErrorAction SilentlyContinue
        }
    }
}

if ($monthlyFiles.Count -gt 0) {
    Log "[執行] 月報有 $($monthlyFiles.Count) 個待處理檔案，啟動 auto_check_and_import.ps1"
    $monthlyFiles | ForEach-Object { Log "  - $($_.Name)" }
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File $monthlyScript
    Log "[完成] 月報處理結束"
}
