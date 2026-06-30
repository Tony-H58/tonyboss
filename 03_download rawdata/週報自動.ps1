param([switch]$Fast)

$base     = "E:\88. Claude\03_download rawdata"
$pyScript = "$base\auto_download_http.py"

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host " Weekly Report Auto (HTTP Download)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Step 1: date range (last Thu -> this Wed)
# Special-case Thursday: today IS the start of a brand-new cycle, so
# "this week" (today..next Wed) has almost no data yet. On Thursday we
# want the cycle that JUST ENDED yesterday (last Thu -> yesterday/Wed).
$today = Get-Date
$dow = [int]$today.DayOfWeek
if ($dow -eq 4) {
    $ed = $today.AddDays(-1)
    $sd = $ed.AddDays(-6)
} else {
    $daysToWed = ((3 - $dow) + 7) % 7
    $ed = $today.AddDays($daysToWed)
    $sd = $ed.AddDays(-6)
}
$startDate = $sd.ToString("yyMMdd")
$endDate   = $ed.ToString("yyMMdd")
Write-Host " Date: $($sd.ToString('yyyy/MM/dd')) ~ $($ed.ToString('yyyy/MM/dd'))" -ForegroundColor Green

# Step 2: Python download
Write-Host "`n[Download] Auto downloading..." -ForegroundColor Yellow
$env:PYTHONIOENCODING = "utf-8"
python $pyScript $startDate $endDate
$downloadOK = ($LASTEXITCODE -eq 0)
if (-not $downloadOK) {
    Write-Host "[Download] Some failed, continuing convert..." -ForegroundColor Yellow
}

# Step 3: convert + move
Write-Host "`n[Convert] Running convert..." -ForegroundColor Yellow
$convertArgs = @("-NoProfile", "-ExecutionPolicy", "Bypass",
                 "-File", "$base\run_download_and_convert.ps1",
                 "-type", "weekly",
                 "-startDate", $startDate, "-endDate", $endDate)
if ($Fast) { $convertArgs += "-Fast" }
& powershell @convertArgs

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n========================================"
    Write-Host " Weekly report done!" -ForegroundColor Green
    Write-Host "========================================"
} else {
    Write-Host "`n Error - please check" -ForegroundColor Yellow
}
