# check_monthly_download.ps1
# 每週一~三 13:00 執行：確認月報 rawdata 是否已下載
# 集成程序 1-3：監控 → 自動修復 → 知識庫更新

$monthlyRaw  = "E:\88. Claude\11_inspdata_monthly\insprawdata_monthly"
$monthlyDone = "$monthlyRaw\completed rawdata"
$today       = Get-Date
$reminderDir = "E:\88. Claude\02_reminder"

# 只在週一~三執行
if ($today.DayOfWeek -notin @([DayOfWeek]::Monday, [DayOfWeek]::Tuesday, [DayOfWeek]::Wednesday)) {
    Write-Host "[略過] 今天非週一~三，不執行月報檢查"
    exit 0
}

# 檢查今天是否有新下載的月報檔 (Program 1: 監控)
Write-Host "`n[監控] 檢查月報 rawdata..."
$inRaw  = @(Get-ChildItem $monthlyRaw  -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime.Date -eq $today.Date -and $_.Extension -in ".xlsx",".xls" })
$inDone = @(Get-ChildItem $monthlyDone -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime.Date -eq $today.Date -and $_.Extension -in ".xlsx",".xls" })

if ($inRaw.Count -gt 0 -or $inDone.Count -gt 0) {
    Write-Host "[✅ OK] 月報今日已下載（Raw=$($inRaw.Count) Done=$($inDone.Count)）"
    exit 0
}

# 檢測到失敗，觸發程序 2: 自動修復
Write-Host "[⚠️  失敗] $($today.ToString('yyyy-MM-dd')) 尚未下載月報 rawdata"
Write-Host "`n[修復] 正在自動修復..."

$fixScript = "$reminderDir\auto_fix_downloads.ps1"
if (Test-Path $fixScript) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $fixScript -Type "monthly"
    $fixExitCode = $LASTEXITCODE

    if ($fixExitCode -eq 0) {
        Write-Host "[✅ 修復成功] 月報已自動下載並處理完成"
    } else {
        Write-Host "[❌ 修復失敗] 月報自動修復未能成功，請手動檢查"
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
