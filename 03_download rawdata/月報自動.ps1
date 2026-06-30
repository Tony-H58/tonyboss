param([switch]$Fast)

$base      = "E:\88. Claude\03_download rawdata"
$pyScript  = "$base\auto_download_http.py"
$destDir   = "E:\88. Claude\11_inspdata_monthly\insprawdata_monthly"

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host " Monthly Report Auto (HTTP Download)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Step 1: date range (上上週四~上週三，從週一/二/三執行)
# DayOfWeek: Sun=0 Mon=1 Tue=2 Wed=3 Thu=4 Fri=5 Sat=6
$today = Get-Date
$dow = [int]$today.DayOfWeek
$ed = $today.AddDays(-($dow + 4))   # 上週三
$sd = $ed.AddDays(-6)               # 上上週四
$startDate = $sd.ToString("yyMMdd")
$endDate   = $ed.ToString("yyMMdd")
Write-Host " Date: $($sd.ToString('yyyy/MM/dd')) ~ $($ed.ToString('yyyy/MM/dd'))" -ForegroundColor Green

# Step 2: 檢查是否已下載（兩檔都存在則跳過）
$qcFile  = "$destDir\品管驗布報表_$startDate-$endDate.xls"
$facFile = "$destDir\工廠驗布報表_$startDate-$endDate.xls"
if ((Test-Path $qcFile) -and (Test-Path $facFile)) {
    Write-Host "`n Already downloaded: $startDate-$endDate, skip." -ForegroundColor Green
    exit 0
}

# Step 3: Python download
Write-Host "`n[Download] Auto downloading..." -ForegroundColor Yellow
$env:PYTHONIOENCODING = "utf-8"
python $pyScript $startDate $endDate
$downloadOK = ($LASTEXITCODE -eq 0)
if (-not $downloadOK) {
    Write-Host "[Download] Some failed, continuing..." -ForegroundColor Yellow
}

# Step 4: rename + move
Write-Host "`n[Process] Running..." -ForegroundColor Yellow
$convertArgs = @("-NoProfile", "-ExecutionPolicy", "Bypass",
                 "-File", "$base\run_download_and_convert.ps1",
                 "-type", "monthly",
                 "-startDate", $startDate, "-endDate", $endDate)
if ($Fast) { $convertArgs += "-Fast" }
& powershell @convertArgs

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n========================================"
    Write-Host " Monthly report done!" -ForegroundColor Green
    Write-Host "========================================"
} else {
    Write-Host "`n Error - please check" -ForegroundColor Yellow
}