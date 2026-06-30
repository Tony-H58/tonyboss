$basePath = "E:\88. Claude\11_inspdata_monthly"
$rawPath  = "$basePath\insprawdata_monthly"
$logFile  = "$basePath\auto_import_log.txt"
$lockFile = "$basePath\auto_import.lock"

function Log($msg){
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$ts  $msg" | Out-File $logFile -Append -Encoding utf8
    Write-Host "$ts  $msg"
}

if(Test-Path $lockFile){ Log "Skip: another instance running."; exit 0 }
New-Item $lockFile -ItemType File -Force | Out-Null

try {
    $pending = Get-ChildItem $rawPath -File | Where-Object {
        $_.Name -match "_(\d{2})\d{4}-(\d{2})\d{4}" -and
        ($_.Name -like "*.xlsx" -or $_.Name -like "*.xls") -and
        $_.Name -notlike "Imported_*"
    }

    if($pending.Count -eq 0){ Log "No pending rawdata."; exit 0 }

    Log ("Found {0} pending file(s):" -f $pending.Count)
    $pending | ForEach-Object { Log "  $($_.Name)" }

    $years = [System.Collections.Generic.SortedSet[string]]::new()
    foreach($f in $pending){
        if($f.Name -match "_(\d{2})\d{4}-(\d{2})\d{4}"){
            $years.Add("20$($Matches[1])") | Out-Null
            $years.Add("20$($Matches[2])") | Out-Null
        }
    }

    foreach($yr in $years){
        Log "===== Processing year $yr ====="
        powershell -NoProfile -ExecutionPolicy Bypass -File "$basePath\run_monthly_full.ps1" -year $yr
        Log "===== Year $yr done ====="
    }
} finally {
    Remove-Item $lockFile -Force -ErrorAction SilentlyContinue
}
