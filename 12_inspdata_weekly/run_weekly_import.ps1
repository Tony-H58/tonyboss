param([switch]$Fast)
$basePath  = "E:\88. Claude\12_inspdata_weekly"
$qfDir     = "$basePath\insprecord_QF"
$qfCurrent = "$qfDir\insprecord_QF.xlsx"

# Step 0: kill Excel, 確認範本存在
$exProc = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if($exProc){ $exProc | Stop-Process -Force; Start-Sleep -Seconds 2 }

if(-not (Test-Path $qfCurrent)) {
  Write-Host "ERROR: QF 範本檔不存在: $qfCurrent"
  exit 1
}
Write-Host ("範本: insprecord_QF.xlsx")

# Step 1~3: run import_weekly.ps1
Write-Host "`n===== Step 1~3: Import ====="
& "$basePath\import_weekly.ps1" -Fast

# Find the QF file (SaveAs by import_weekly.ps1, date-named)
$qfFile = (Get-ChildItem $qfDir -Filter "insprecord_QF_??????-??????.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
Write-Host ("QF file: {0}" -f (Split-Path $qfFile -Leaf))

# Step 4~7: 合併單一 Excel 會話
Write-Host "`n===== Step 4~7: Fix / Count / Format ====="
& "$basePath\step4to7_weekly.ps1" -qfFile $qfFile -NoKill

# Step 8: Move source files to Completed
Write-Host "`n===== Step 8: Move source files to Completed ====="
$rawPath = "$basePath\insprawdata_weekly"
$doneDir = "$rawPath\completed rawdata"
if(-not (Test-Path $doneDir)) { New-Item -ItemType Directory -Path $doneDir | Out-Null }
$qfLeaf = Split-Path $qfFile -Leaf
$dateSuffix = if($qfLeaf -match "insprecord_QF_(\d{6}-\d{6})"){$Matches[1]}else{(Get-Date).ToString("yyMMdd")}
Get-ChildItem $rawPath | Where-Object { ($_.Extension -eq ".xlsx" -or $_.Extension -eq ".xls") -and $_.Name -notlike "~*" } | ForEach-Object {
  $baseName = [System.IO.Path]::GetFileNameWithoutExtension($_.Name) -replace "_\d{6}-\d{6}$",""
  $newName = $baseName + "_$dateSuffix" + $_.Extension
  Move-Item $_.FullName "$doneDir\$newName" -Force
  Write-Host ("  Moved: {0} -> {1}" -f $_.Name,$newName)
}

Write-Host ("`n===== DONE: {0} =====" -f (Split-Path $qfFile -Leaf))
exit 0
