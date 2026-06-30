# check_monthly_download.ps1
# 每週一~三 13:00 執行：確認月報 rawdata 是否已下載

$monthlyRaw  = "E:\88. Claude\11_inspdata_monthly\insprawdata_monthly"
$monthlyDone = "$monthlyRaw\completed rawdata"
$today       = Get-Date

# 只在週一~三執行
if ($today.DayOfWeek -notin @([DayOfWeek]::Monday, [DayOfWeek]::Tuesday, [DayOfWeek]::Wednesday)) {
    Write-Host "[略過] 今天非週一~三，不執行月報檢查"
    exit 0
}

# 檢查今天是否有新下載的月報檔
$inRaw  = @(Get-ChildItem $monthlyRaw  -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime.Date -eq $today.Date -and $_.Extension -in ".xlsx",".xls" })
$inDone = @(Get-ChildItem $monthlyDone -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime.Date -eq $today.Date -and $_.Extension -in ".xlsx",".xls" })

if ($inRaw.Count -gt 0 -or $inDone.Count -gt 0) {
    Write-Host "[OK] 月報今日已下載（Raw=$($inRaw.Count) Done=$($inDone.Count)）"
    exit 0
} else {
    Write-Host "[提醒] $($today.ToString('yyyy-MM-dd')) 尚未下載月報 rawdata，請執行桌面「月報一鍵.bat」"
    exit 1
}
