param([switch]$Fast)
$basePath = "E:\88. Claude\12_inspdata_weekly"
$task01Path = "E:\88. Claude\11_inspdata_monthly"

Write-Host "========================================="
Write-Host " Task20 Weekly Full Process"
Write-Host "========================================="

# Check raw data files exist
$rawFiles = Get-ChildItem "$basePath\insprawdata_weekly" | Where-Object { ($_.Extension -eq ".xlsx" -or $_.Extension -eq ".xls") -and $_.Name -notlike "~*" }
$qfExists = (Get-ChildItem "$basePath\insprecord_QF" -Filter "insprecord_QF_??????-??????.xlsx" | Measure-Object).Count -gt 0

if($rawFiles.Count -eq 0 -and -not $qfExists){
    Write-Host "ERROR: No raw data files and no QF file found. Aborting."
    exit 1
}

if($rawFiles.Count -gt 0){
    Write-Host ("Raw files found: {0}" -f ($rawFiles.Name -join ", "))
    # Step 1~8: Import raw data into InspRecord_QF
    Write-Host "`n===== PHASE 1: Import InspRecord_QF ====="
    & "$basePath\run_weekly_import.ps1" -Fast
    if($LASTEXITCODE -ne 0){ Write-Host "ERROR in import phase. Aborting."; exit 1 }
} else {
    Write-Host "No raw files found — skipping Phase 1, using existing QF file."
}

# Step 9: Calculate statistics into AnalysisChart_Weekly
Write-Host "`n===== PHASE 2: Calculate Statistics ====="
& "$basePath\calc_weekly_stats.ps1"
if($LASTEXITCODE -ne 0){ Write-Host "ERROR in stats phase. Aborting."; exit 1 }

Write-Host "`n===== ALL DONE ====="
$qfFile = Get-ChildItem "$basePath\insprecord_QF" -Filter "insprecord_QF_*.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$chartFile = Get-ChildItem "$basePath\analysis chart_weekly" -Filter "analysis chart_weekly_*.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
Write-Host ("QF file : {0}" -f $qfFile.Name)
Write-Host ("Chart   : {0}" -f $chartFile.Name)
