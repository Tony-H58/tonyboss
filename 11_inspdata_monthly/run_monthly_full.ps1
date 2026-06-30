param([string]$year, [switch]$NoKill)
# Usage: powershell -File run_monthly_full.ps1 -year 2026

$basePath  = "E:\88. Claude\11_inspdata_monthly"
$qfDir     = "$basePath\insprecord_QF"
$qfCurrent = "$qfDir\insprecord_QF_current.xlsx"

if(-not $year) { Write-Host "Please specify year, e.g. -year 2026"; exit 1 }

# Step 0: 將同年份 QF（yymmdd-yymmdd 命名）改回 current.xlsx 以供 import_monthly 追加
if(-not $NoKill){
  Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
  Start-Sleep -Seconds 2
}
$yy = $year.Substring(2)
$existing = Get-ChildItem $qfDir -Filter "insprecord_QF_${yy}*.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if($existing) {
  Rename-Item $existing.FullName $qfCurrent -Force
  Write-Host ("Using existing file: {0}" -f $existing.Name)
}
if(-not (Test-Path $qfCurrent)) {
  Write-Host "ERROR: QF working file not found: $qfCurrent"
  exit 1
}

# Step 1~3: run import_monthly.ps1
Write-Host "`n===== Step 1~3: Import ====="
$importScript = "$basePath\import_monthly.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File $importScript -year $year

# Find the QF file (now renamed by import_monthly.ps1 to yymmdd-yymmdd format)
$qfFile = (Get-ChildItem $qfDir -Filter "insprecord_QF_${yy}*.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
Write-Host ("QF file: {0}" -f (Split-Path $qfFile -Leaf))

# Clean up temp file if exists
$startRowFile = "$basePath\temp_startrow.txt"
if(Test-Path $startRowFile){ Remove-Item $startRowFile }

# Step 8: Move source files to processed folder
Write-Host "`n===== Step 8: Move source files to processed folder ====="
$rawPath = "$basePath\insprawdata_monthly"
$doneDir = "$rawPath\completed rawdata"
if(-not (Test-Path $doneDir)) { New-Item -ItemType Directory -Path $doneDir | Out-Null }
# 只移動結束年=目標年的檔（跨年rawdata等最後一年跑完才移走）
Get-ChildItem $rawPath -File | Where-Object { $_.Name -match "_\d{6}-${yy}\d{4}\." -and ($_.Name -like "*.xlsx" -or $_.Name -like "*.xls") -and $_.Name -notlike "Imported_*" } | ForEach-Object {
  Move-Item $_.FullName "$doneDir\$($_.Name)" -Force
  Write-Host ("  Moved: {0}" -f $_.Name)
}

Write-Host ("`n===== DONE: {0} =====" -f (Split-Path $qfFile -Leaf))


