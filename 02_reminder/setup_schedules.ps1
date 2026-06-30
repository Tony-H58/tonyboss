# setup_schedules.ps1
# 建立 02_reminder 的三個排程任務
# 必須以管理員身份執行此腳本

$ErrorActionPreference = "Stop"

$taskPath = "\Claude\02_reminder\"

Write-Host "開始建立排程任務...`n" -ForegroundColor Green

# 1. 週報下載檢查排程 (每週四 13:00)
$script1 = "E:\88. Claude\02_reminder\check_weekly_download.ps1"
$trigger1 = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Thursday -At 13:00
$action1 = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$script1`""
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

try {
    if (Get-ScheduledTask -TaskPath $taskPath -TaskName "CheckWeeklyDownload" -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskPath $taskPath -TaskName "CheckWeeklyDownload" -Confirm:$false
    }
    Register-ScheduledTask -TaskPath $taskPath -TaskName "CheckWeeklyDownload" `
        -Action $action1 -Trigger $trigger1 -Settings $settings `
        -Description "每週四 13:00 檢查週報 rawdata 是否已下載" | Out-Null
    Write-Host "✅ 已建立排程：CheckWeeklyDownload" -ForegroundColor Green
} catch {
    Write-Host "❌ 失敗：$($_.Exception.Message)" -ForegroundColor Red
}

# 2. 月報下載檢查排程 (每週一~三 13:00)
$script2 = "E:\88. Claude\02_reminder\check_monthly_download.ps1"
$trigger2 = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday, Tuesday, Wednesday -At 13:00
$action2 = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$script2`""

try {
    if (Get-ScheduledTask -TaskPath $taskPath -TaskName "CheckMonthlyDownload" -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskPath $taskPath -TaskName "CheckMonthlyDownload" -Confirm:$false
    }
    Register-ScheduledTask -TaskPath $taskPath -TaskName "CheckMonthlyDownload" `
        -Action $action2 -Trigger $trigger2 -Settings $settings `
        -Description "每週一~三 13:00 檢查月報 rawdata 是否已下載" | Out-Null
    Write-Host "✅ 已建立排程：CheckMonthlyDownload" -ForegroundColor Green
} catch {
    Write-Host "❌ 失敗：$($_.Exception.Message)" -ForegroundColor Red
}

# 3. 寄週報郵件排程 (每天開機時執行)
$script3 = "E:\88. Claude\02_reminder\send_weekly_email.ps1"
$trigger3 = New-ScheduledTaskTrigger -AtLogon
$action3 = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$script3`""

try {
    if (Get-ScheduledTask -TaskPath $taskPath -TaskName "SendWeeklyEmail" -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskPath $taskPath -TaskName "SendWeeklyEmail" -Confirm:$false
    }
    Register-ScheduledTask -TaskPath $taskPath -TaskName "SendWeeklyEmail" `
        -Action $action3 -Trigger $trigger3 -Settings $settings `
        -Description "每天開機時寄送週報郵件" | Out-Null
    Write-Host "✅ 已建立排程：SendWeeklyEmail" -ForegroundColor Green
} catch {
    Write-Host "❌ 失敗：$($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n📋 排程建立完成。" -ForegroundColor Green
Write-Host "已建立排程列表：" -ForegroundColor Cyan
Get-ScheduledTask -TaskPath $taskPath | Select-Object TaskName, State | Format-Table -AutoSize

Write-Host "按任意鍵結束..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
