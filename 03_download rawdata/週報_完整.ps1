param([switch]$Fast)

$base = "E:\88. Claude\03_download rawdata"
$downloads = "$env:USERPROFILE\Downloads"

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host " 週報完整流程 $(if($Fast){'⚡快速模式'}else{})" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# 計算週報日期
$today = Get-Date
$dow = [int]$today.DayOfWeek
$daysToWed = ((3 - $dow) + 7) % 7
$ed = $today.AddDays($daysToWed)
$sd = $ed.AddDays(-6)
$sdFmt = $sd.ToString("yyyy/MM/dd")
$edFmt = $ed.ToString("yyyy/MM/dd")

Write-Host "日期: $sdFmt → $edFmt" -ForegroundColor Gray

# Step 1: 檢查源檔案
Write-Host "`n[Step 1] 檢查源檔案..." -ForegroundColor Yellow
$qcFile = Get-ChildItem $downloads -Filter "*驗布結論報表*" -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -gt (Get-Date).AddDays(-1) } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$facFile = Get-ChildItem $downloads -Filter "*後端驗布報表*" -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -gt (Get-Date).AddDays(-1) } | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if ($qcFile -and $facFile) {
    Write-Host "✓ 找到源檔案" -ForegroundColor Green
} else {
    Write-Host "⚠ 缺少源檔案，開啟下載系統..." -ForegroundColor Yellow
    
    # 開啟瀏覽器
    Start-Process "http://nt-net2.makalot.com.tw/FabQC/ReportList.aspx?language=zh-TW"
    Start-Process "https://nt-net2.makalot.com.tw/FabricQuality/WebPage/FabricResultAnalysis"
    
    Write-Host "`n📍 已開啟兩個下載系統：" -ForegroundColor Cyan
    Write-Host "   1. 品管系統 - 選擇驗布日期, 設定 $sdFmt ~ $edFmt, 點擊匯出" -ForegroundColor Gray
    Write-Host "   2. 工廠系統 - 設定 $sdFmt ~ $edFmt, 點擊匯出" -ForegroundColor Gray
    
    $maxWait = if($Fast) { 180 } else { 600 }
    $checkInterval = if($Fast) { 1 } else { 3 }
    Write-Host "`n⏳ 監控下載檔案（最多等待 $($maxWait/60) 分鐘）..." -ForegroundColor Yellow
    
    $startTime = Get-Date
    
    while ((Get-Date) -lt $startTime.AddSeconds($maxWait)) {
        $qcFile = Get-ChildItem $downloads -Filter "*驗布結論報表*" -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -gt $startTime } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $facFile = Get-ChildItem $downloads -Filter "*後端驗布報表*" -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -gt $startTime } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        
        if ($qcFile -and $facFile) {
            Write-Host "✓ 兩份報表已下載完成！" -ForegroundColor Green
            break
        }
        
        Start-Sleep -Milliseconds ($checkInterval * 1000)
        $elapsed = [math]::Round(((Get-Date) - $startTime).TotalSeconds)
        Write-Host "⏳ 等待中... ($elapsed秒)" -ForegroundColor Gray -NoNewline
        Write-Host "`r" -NoNewline
    }
    
    # 最後驗證一次
    $qcFile = Get-ChildItem $downloads -Filter "*驗布結論報表*" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $facFile = Get-ChildItem $downloads -Filter "*後端驗布報表*" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    
    if (-not $qcFile -or -not $facFile) {
        Write-Host "`n❌ 下載超時，請手動重試" -ForegroundColor Red
        exit 1
    }
}

# Step 2: 複製到工作目錄
Write-Host "`n[Step 2] 複製到工作目錄..." -ForegroundColor Yellow
Copy-Item $qcFile.FullName "$base\$($qcFile.Name)" -Force
Copy-Item $facFile.FullName "$base\$($facFile.Name)" -Force
Write-Host "✓ 複製完成" -ForegroundColor Green

# Step 3: 執行轉檔
Write-Host "`n[Step 3] 執行轉檔..." -ForegroundColor Yellow
if ($Fast) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File "$base\run_download_and_convert.ps1" -type weekly -Fast
} else {
    & powershell -NoProfile -ExecutionPolicy Bypass -File "$base\run_download_and_convert.ps1" -type weekly
}

Write-Host "`n✓ 週報完整流程完成！" -ForegroundColor Green
Write-Host "📊 輸出: E:\88. Claude\12_inspdata_weekly\insprawdata_weekly" -ForegroundColor Cyan