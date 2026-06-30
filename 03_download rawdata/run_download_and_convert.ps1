<#
.SYNOPSIS
    驗布報表主流程：計算日期 + 歸檔 + 改名 + 移至指定資料夾

.PARAMETER type
    "weekly"（週報，預設）或 "monthly"（月報）

.PARAMETER startDate / endDate
    手動指定日期範圍 yymmdd
#>
param(
    [string]$type      = "weekly",
    [string]$startDate = "",
    [string]$endDate   = "",
    [switch]$Fast
)

$basePath     = "E:\88. Claude\03_download rawdata"
$downloadsDir = "$env:USERPROFILE\Downloads"
$archiveDir   = "$basePath\completed rawdata"

Write-Host "`n$('='*60)" -ForegroundColor Cyan
Write-Host " 驗布報表主流程" -ForegroundColor Cyan
Write-Host "$('='*60)" -ForegroundColor Cyan
Write-Host " 類型: $type | $(Get-Date -Format 'yyyy/MM/dd HH:mm')" -ForegroundColor Cyan

# ── Step 1：計算日期範圍 ──
if(-not $startDate -or -not $endDate){
    Write-Host "`n[1] 計算日期範圍 ($type)..." -ForegroundColor Yellow
    $today = Get-Date

    if($type -eq "monthly"){
        $firstOfThisMonth = Get-Date -Year $today.Year -Month $today.Month -Day 1
        $firstOfLastMonth = $firstOfThisMonth.AddMonths(-1)
        $lastOfLastMonth  = $firstOfThisMonth.AddDays(-1)
        $sd = $firstOfLastMonth; $ed = $lastOfLastMonth
    } else {
        $dow = [int]$today.DayOfWeek
        if($dow -eq 4){
            $ed = $today.AddDays(-1)
            $sd = $ed.AddDays(-6)
        } else {
            $daysToWed = ((3 - $dow) + 7) % 7
            $ed = $today.AddDays($daysToWed)
            $sd = $ed.AddDays(-6)
        }
    }

    $startDate = $sd.ToString("yyMMdd")
    $endDate   = $ed.ToString("yyMMdd")
    Write-Host "  日期: $($sd.ToString('yyyy/MM/dd')) -> $($ed.ToString('yyyy/MM/dd'))" -ForegroundColor Green
    Write-Host "  標記: $startDate - $endDate" -ForegroundColor Green
} else {
    Write-Host "`n[1] 使用手動日期: $startDate - $endDate" -ForegroundColor Yellow
}

# 目標目錄
if($type -eq "monthly") {
    $destDir    = "E:\88. Claude\11_inspdata_monthly\insprawdata_monthly"
    $reportType = "月報"
} else {
    $destDir    = "E:\88. Claude\12_inspdata_weekly\insprawdata_weekly"
    $reportType = "週報"
}

if(-not (Test-Path $archiveDir)) { New-Item -ItemType Directory -Path $archiveDir -Force | Out-Null }
if(-not (Test-Path $destDir))    { New-Item -ItemType Directory -Path $destDir    -Force | Out-Null }

# ── Step 2：掃描 Downloads ──
Write-Host "`n[2] 掃描 Downloads 資料夾..." -ForegroundColor Yellow
$cutoff = (Get-Date).AddDays(-3)
$qcFile = Get-ChildItem $downloadsDir -File | Where-Object {
    $_.Name -like "*驗布結論報表*" -and $_.Name -notlike "~`$*" -and $_.LastWriteTime -gt $cutoff
} | Sort-Object LastWriteTime -Descending | Select-Object -First 1

$facFile = Get-ChildItem $downloadsDir -File | Where-Object {
    $_.Name -like "*後端驗布報表*" -and $_.Name -notlike "~`$*" -and $_.LastWriteTime -gt $cutoff
} | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if (-not $qcFile -and -not $facFile) {
    Write-Host "  Warning: Downloads 中未找到新的源檔" -ForegroundColor Yellow
    exit 1
}
Write-Host "  找到 QC: $(if($qcFile){$qcFile.Name}else{'無'})" -ForegroundColor Green
Write-Host "  找到 FTY: $(if($facFile){$facFile.Name}else{'無'})" -ForegroundColor Green

# ── Step 3：歸檔到 completed rawdata（以實際日期範圍命名）──
Write-Host "`n[3] 歸檔原始檔..." -ForegroundColor Yellow
if ($qcFile) {
    Copy-Item $qcFile.FullName "$archiveDir\品管驗布報表_$startDate-$endDate.xls" -Force
    Write-Host "  OK 品管驗布報表_$startDate-$endDate.xls -> completed rawdata\" -ForegroundColor Gray
}
if ($facFile) {
    Copy-Item $facFile.FullName "$archiveDir\工廠驗布報表_$startDate-$endDate.xls" -Force
    Write-Host "  OK 工廠驗布報表_$startDate-$endDate.xls -> completed rawdata\" -ForegroundColor Gray
}

# ── Step 4：改名並從 Downloads 移至目標資料夾（刪除 Downloads 原始檔）──
Write-Host "`n[4] 改名並移至 $reportType 資料夾..." -ForegroundColor Yellow
if ($qcFile) {
    $dest = "$destDir\品管驗布報表_$startDate-$endDate.xls"
    Move-Item $qcFile.FullName $dest -Force
    Write-Host "  OK 品管驗布報表_$startDate-$endDate.xls -> $reportType 資料夾" -ForegroundColor Green
}
if ($facFile) {
    $dest = "$destDir\工廠驗布報表_$startDate-$endDate.xls"
    Move-Item $facFile.FullName $dest -Force
    Write-Host "  OK 工廠驗布報表_$startDate-$endDate.xls -> $reportType 資料夾" -ForegroundColor Green
}

Write-Host "`n$('='*60)" -ForegroundColor Cyan
Write-Host " Done!" -ForegroundColor Green
Write-Host " $reportType 輸出目錄: $destDir" -ForegroundColor Cyan
Write-Host "$('='*60)" -ForegroundColor Cyan