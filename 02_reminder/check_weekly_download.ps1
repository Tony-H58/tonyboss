# check_weekly_download.ps1
# 每週四 13:00 執行：確認週報 rawdata 是否已下載

$weeklyRaw  = "E:\88. Claude\12_inspdata_weekly\insprawdata_weekly"
$weeklyDone = "$weeklyRaw\completed rawdata"
$today      = (Get-Date).Date

$inRaw  = @(Get-ChildItem $weeklyRaw  -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime.Date -eq $today -and $_.Extension -in ".xlsx",".xls" })
$inDone = @(Get-ChildItem $weeklyDone -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime.Date -eq $today -and $_.Extension -in ".xlsx",".xls" })

if ($inRaw.Count -gt 0 -or $inDone.Count -gt 0) {
    Write-Host "[OK] 週報已下載（今日新檔：Raw=$($inRaw.Count) Done=$($inDone.Count)）"
    exit 0
} else {
    Write-Host "[提醒] 今天（$((Get-Date).ToString('yyyy-MM-dd'))）尚未下載週報 rawdata，請執行桌面「週報一鍵.bat」"
    exit 1
}
