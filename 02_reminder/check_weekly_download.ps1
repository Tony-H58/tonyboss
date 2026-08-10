# check_weekly_download.ps1
# 每週四 13:00 執行：確認週報 rawdata 是否已下載
# 集成程序 1-3：監控 → 自動修復 → 知識庫更新

$weeklyRaw   = "E:\88. Claude\12_inspdata_weekly\insprawdata_weekly"
$weeklyDone  = "$weeklyRaw\completed rawdata"
$today       = (Get-Date).Date
$reminderDir = "E:\88. Claude\02_reminder"

# 檢查今天是否有新下載的週報檔 (Program 1: 監控)
Write-Host "`n[監控] 檢查週報 rawdata..."
$inRaw  = @(Get-ChildItem $weeklyRaw  -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime.Date -eq $today -and $_.Extension -in ".xlsx",".xls" })
$inDone = @(Get-ChildItem $weeklyDone -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime.Date -eq $today -and $_.Extension -in ".xlsx",".xls" })

if ($inRaw.Count -gt 0 -or $inDone.Count -gt 0) {
    Write-Host "[✅ OK] 週報已下載（今日新檔：Raw=$($inRaw.Count) Done=$($inDone.Count)）"
    exit 0
}

# 檢測到失敗，觸發程序 2: 自動修復
Write-Host "[⚠️  失敗] 今天（$((Get-Date).ToString('yyyy-MM-dd'))）尚未下載週報 rawdata"
Write-Host "`n[修復] 正在自動修復..."

$fixScript = "$reminderDir\auto_fix_downloads.ps1"
if (Test-Path $fixScript) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $fixScript -Type "weekly"
    $fixExitCode = $LASTEXITCODE

    if ($fixExitCode -eq 0) {
        Write-Host "[✅ 修復成功] 週報已自動下載並處理完成"
    } else {
        Write-Host "[❌ 修復失敗] 週報自動修復未能成功，請手動檢查"
    }
} else {
    Write-Host "[❌ 錯誤] 找不到修復腳本: $fixScript"
    $fixExitCode = 2
}

# 觸發程序 3: 更新知識庫
Write-Host "`n[更新] 更新知識庫..."
$updateScript = "$reminderDir\update_knowledge_skill.ps1"
if (Test-Path $updateScript) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $updateScript
} else {
    Write-Host "[⚠️  警告] 找不到知識庫更新腳本"
}

exit $fixExitCode
