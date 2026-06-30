# 每週一~三開機時自動寄送週報統計+品管儀表板給 tonyhuang@makalot.com.tw
# 每週只寄一次：用「本週週一日期」當作標記寫入 last_sent_week.txt，已寄過則跳過

$ErrorActionPreference = "Stop"

$recipient   = "tonyhuang@makalot.com.tw"
$chartDir    = "E:\88. Claude\12_inspdata_weekly\analysis chart_weekly"
$dashboardDir= "C:\Users\tonyhuang\Desktop"
$markerPath  = "E:\88. Claude\02_reminder\last_sent_week.txt"
$logPath     = "E:\88. Claude\02_reminder\send_log.txt"

function Write-Log($msg) {
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Add-Content -Path $logPath -Value $line -Encoding utf8
    Write-Host $line
}

# 只在週一~三執行（避免手動誤跑或排程設定異常時寄到其他天）
$today = Get-Date
if ($today.DayOfWeek -notin @([DayOfWeek]::Monday, [DayOfWeek]::Tuesday, [DayOfWeek]::Wednesday)) {
    Write-Log "今天非週一~三，跳過。"
    exit 0
}

# 計算本週週一日期作為「本週」標記
$diff = ([int]$today.DayOfWeek + 6) % 7   # Sunday=0 -> 6, Monday=1 -> 0, ...
$monday = $today.Date.AddDays(-$diff)
$weekKey = $monday.ToString("yyyyMMdd")

if (Test-Path $markerPath) {
    $lastSent = (Get-Content $markerPath -Raw).Trim()
    if ($lastSent -eq $weekKey) {
        Write-Log "本週($weekKey)已寄送過，跳過。"
        exit 0
    }
}

# 找最新的 analysis chart_weekly_*.xlsx（排除無日期的範本檔）
$chartFile = Get-ChildItem -Path $chartDir -Filter "analysis chart_weekly_*.xlsx" -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1

if (-not $chartFile) {
    Write-Log "找不到最新的 analysis chart_weekly 檔案，中止寄送。"
    exit 1
}

# 品管儀表板三個檔案（總頁+A+B，需一起附上才能互相跳轉）
$dashboardFiles = @(
    (Join-Path $dashboardDir "品管儀表板_總頁.html"),
    (Join-Path $dashboardDir "品管儀表板_A.html"),
    (Join-Path $dashboardDir "品管儀表板_B.html")
) | Where-Object { Test-Path $_ }

if ($dashboardFiles.Count -eq 0) {
    Write-Log "找不到品管儀表板檔案，中止寄送。"
    exit 1
}

try {
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)  # olMailItem
    $mail.To = $recipient
    $mail.Subject = "每週驗布統計報表與品管儀表板（$($monday.ToString('yyyy/MM/dd'))週）"
    $mail.Body = @"
您好，

附上本週最新的驗布統計報表（$($chartFile.Name)）與品管儀表板（總頁/A/B），請查閱。

此信為自動發送，每週一寄送一次。
"@

    $mail.Attachments.Add($chartFile.FullName) | Out-Null
    foreach ($f in $dashboardFiles) {
        $mail.Attachments.Add($f) | Out-Null
    }

    $mail.Send()

    Set-Content -Path $markerPath -Value $weekKey -Encoding utf8
    Write-Log "已成功寄送週報郵件給 $recipient，附件：$($chartFile.Name) + $($dashboardFiles.Count)個儀表板檔案。"
}
catch {
    Write-Log "寄送失敗：$($_.Exception.Message)"
    exit 1
}


