<#
.SYNOPSIS
    驗布報表轉換腳本 v2
    QC來源：驗布結論報表*.xls (Sheet1) → 品管驗布報表.xlsx (75欄)
    工廠來源：後端驗布報表*.xls (Sheet2=客戶-明細表, 47欄) → 工廠驗布報表.xlsx (45欄)

.PARAMETER startDate / endDate
    日期標記 yymmdd，用於輸出檔名
#>
param(
    [string]$startDate = "",
    [string]$endDate   = "",
    [switch]$Fast
)

$basePath     = "E:\88. Claude\03_download rawdata"
$planFile     = "$basePath\00_plan_download rawdata.xlsx"
$qcTemplate   = "$basePath\00_品管驗布報表.xlsx"
$facTemplate  = "$basePath\00_工廠驗布報表.xlsx"
$completedDir = "$basePath\completed rawdata"
if(-not (Test-Path $completedDir)){ New-Item -ItemType Directory -Path $completedDir | Out-Null }

function Log($msg, $color="White"){
    if(-not $Fast) { Write-Host "  $msg" -ForegroundColor $color }
    else { Write-Host "." -NoNewline -ForegroundColor $color }
}
Write-Host "`n$('='*60)" -ForegroundColor Cyan
Write-Host " 驗布報表轉換程序 v2 $(if($Fast){'⚡快速'}else{})" -ForegroundColor Cyan
Write-Host "$('='*60)" -ForegroundColor Cyan

# ── Step 1：讀取NameList / ChartTitle ──
Write-Host "`n[1] 讀取映射規則..." -ForegroundColor Yellow
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
if($Fast) {
    $excel.Calculation = -4135  # xlCalculationManual
    $excel.ScreenUpdating = $false
}

$wbPlan = $excel.Workbooks.Open($planFile, $false, $true)

# ── NameList → 5張對照表 ──
$wsNL   = $wbPlan.Sheets("NameList")
$nlData = $wsNL.UsedRange.Value2
$custMap=@{}; $vendorMap=@{}; $locMap=@{}; $factoryMap=@{}; $qcMap=@{}
for($r=2; $r -le $nlData.GetLength(0); $r++){
    $a=$nlData[$r,1];$b=$nlData[$r,2]
    $c=$nlData[$r,3];$d=$nlData[$r,4]
    $e=$nlData[$r,5];$f=$nlData[$r,6]
    $g=$nlData[$r,7];$h=$nlData[$r,8]
    $i=$nlData[$r,9];$j=$nlData[$r,10]
    if($a -and $b){$custMap[$a.ToString().Trim()]    = $b.ToString().Trim()}
    if($c -and $d){$vendorMap[$c.ToString().Trim()]  = $d.ToString().Trim()}
    if($e -and $f){$locMap[$e.ToString().Trim()]     = $f.ToString().Trim()}
    if($g -and $h){$factoryMap[$g.ToString().Trim()] = $h.ToString().Trim()}
    if($i -and $j){$qcMap[$i.ToString().Trim()]      = $j.ToString().Trim()}
}
Log "NameList: 客戶=$($custMap.Count) 廠商=$($vendorMap.Count) 地點=$($locMap.Count) 工廠=$($factoryMap.Count) QC=$($qcMap.Count)" "Green"

# ── ChartTitle → 建立映射表 ──
# QC:      Row2=target col#, Row6=source col#
# Factory: Row14=target col#, Row18=source col#
# QC NameList欄：target col 5→cust, 13→vendor, 14→loc, 28→qc, 75→factory
# Factory NameList欄：target col 5→cust, 11→vendor, 12→loc  (45→factory為計算欄，非映射)
$wsCT   = $wbPlan.Sheets("ChartTitle")
$ctData = $wsCT.UsedRange.Value2
$ctRows = $ctData.GetLength(0)
$ctCols = $ctData.GetLength(1)
$startR = $wsCT.UsedRange.Row

function GetCTRow($rowNum){
    $idx = $rowNum - $startR + 1
    if($idx -lt 1 -or $idx -gt $ctRows){ return @{} }
    $row = @{}
    for($c=1; $c -le $ctCols; $c++){
        $v = $ctData[$idx,$c]
        if($null -ne $v -and $v.ToString().Trim() -ne ""){ $row[$c]=$v }
    }
    return $row
}

$qcTgtRow  = GetCTRow 2;  $qcSrcRow  = GetCTRow 6
$facTgtRow = GetCTRow 14; $facSrcRow = GetCTRow 18

# QC：target欄的NameList類型
$qcNLCols  = @{5="cust";13="vendor";14="loc";28="qc";75="factory"}
# Factory：target欄的NameList類型（45為計算欄，另外處理）
$facNLCols = @{5="cust";11="vendor";12="loc"}

$qcColMap  = @{}   # srcColNum → @{tgt=tgtColNum; map=mapType}
$facColMap = @{}

for($c=1; $c -le $ctCols; $c++){
    $tgt=$qcTgtRow[$c]; $src=$qcSrcRow[$c]
    if($tgt -and $src -and "$tgt" -match '^\d+' -and "$src" -match '^\d+'){
        try{
            $ti=[int][double]$tgt; $si=[int][double]$src
            if($ti -gt 0 -and $si -gt 0){
                $mt=if($qcNLCols.ContainsKey($ti)){$qcNLCols[$ti]}else{""}
                $qcColMap[$si]=@{tgt=$ti;map=$mt}
            }
        }catch{}
    }
    $tgt=$facTgtRow[$c]; $src=$facSrcRow[$c]
    if($tgt -and $src -and "$tgt" -match '^\d+' -and "$src" -match '^\d+'){
        try{
            $ti=[int][double]$tgt; $si=[int][double]$src
            if($ti -gt 0 -and $si -gt 0){
                $mt=if($facNLCols.ContainsKey($ti)){$facNLCols[$ti]}else{""}
                $facColMap[$si]=@{tgt=$ti;map=$mt}
            }
        }catch{}
    }
}
Log "QC映射欄位: $($qcColMap.Count)" "Green"
Log "Factory映射欄位: $($facColMap.Count)" "Green"
$wbPlan.Close($false)

# ── 輔助函數 ──
function MapN($v, $mapType){
    if($null -eq $v -or $v.ToString().Trim() -eq ""){ return "" }
    $k = $v.ToString().Trim()
    switch($mapType){
        "cust"    { if($custMap.ContainsKey($k))    { return $custMap[$k] } }
        "vendor"  { if($vendorMap.ContainsKey($k))  { return $vendorMap[$k] } }
        "loc"     { if($locMap.ContainsKey($k))     { return $locMap[$k] } }
        "factory" { if($factoryMap.ContainsKey($k)) { return $factoryMap[$k] } }
        "qc"      { if($qcMap.ContainsKey($k))      { return $qcMap[$k] } }
    }
    return $k
}
function CleanStr($v){ if($null -eq $v){ return "" }; ($v.ToString() -replace "[\r\n]"," ").Trim() }
function DateVal($v){
    if($null -eq $v -or $v.ToString().Trim() -eq ""){ return [DBNull]::Value }
    if($v -is [double]){ return $v }
    $d=[DateTime]::MinValue
    if([DateTime]::TryParse($v.ToString(),[ref]$d)){ return $d.ToOADate() }
    return [DBNull]::Value
}
function AnyVal($v){
    if($null -eq $v -or $v.ToString().Trim() -eq ""){ return [DBNull]::Value }
    if($v -is [double]){ return $v }
    return CleanStr $v
}

# ── Step 2：掃描源數據檔案（.xls 和 .xlsx）──
Write-Host "`n[2] 掃描源數據檔案..." -ForegroundColor Yellow
$srcFiles = @(Get-ChildItem $basePath -Filter "*.xls") +
            @(Get-ChildItem $basePath -Filter "*.xlsx") |
    Where-Object { $_.Name -notlike "~$*" -and $_.Name -notlike "00_plan*" }

# 以檔名關鍵字區分（不含已輸出的模板檔）
$qcFiles  = $srcFiles | Where-Object { $_.Name -like "*驗布結論報表*" }
$facFiles = $srcFiles | Where-Object { $_.Name -like "*後端驗布報表*" }

Log "QC源檔($($qcFiles.Count)): $(($qcFiles | ForEach-Object { $_.Name }) -join ', ')" "Cyan"
Log "工廠源檔($($facFiles.Count)): $(($facFiles | ForEach-Object { $_.Name }) -join ', ')" "Cyan"

if($qcFiles.Count -eq 0 -and $facFiles.Count -eq 0){
    Write-Host "`n❌ 未找到任何源數據檔案！" -ForegroundColor Red
    $excel.Quit(); exit 1
}

# ── Step 3：轉換函數 ──
Write-Host "`n[3] 讀取並轉換數據..." -ForegroundColor Yellow

# 日期目標欄（QC）
$qcDateTgtCols  = @(2,16,17,68,69,70) | ForEach-Object { $_ }  # 申請日期, 預計/實際驗布日, 出口日
# 日期目標欄（Factory）
$facDateTgtCols = @(3) | ForEach-Object { $_ }                  # 實際驗布日

function ConvertRows($file, $colMap, $totalCols, $dateTgtCols, $sheetIdx){
    Log "讀取: $($file.Name) [Sheet$sheetIdx]" "Cyan"
    $wb = $excel.Workbooks.Open($file.FullName, $false, $true)
    $sheetCount = $wb.Sheets.Count
    $si = if($sheetIdx -le $sheetCount){ $sheetIdx } else { 1 }
    $ws = $wb.Sheets($si)
    $data = $ws.UsedRange.Value2
    $nR = $data.GetLength(0); $nC = $data.GetLength(1)
    $dateSet = @{}; $dateTgtCols | ForEach-Object { $dateSet[$_]=$true }

    $rows = [System.Collections.Generic.List[object[]]]::new()
    for($r=2; $r -le $nR; $r++){
        if($null -eq $data[$r,1] -or $data[$r,1].ToString().Trim() -eq ""){ continue }
        $row = New-Object object[] ($totalCols+1)  # 1-based
        for($ci=1; $ci -le $totalCols; $ci++){ $row[$ci]=[DBNull]::Value }

        foreach($srcCol in $colMap.Keys){
            if($srcCol -gt $nC){ continue }
            $info = $colMap[$srcCol]
            $tgt  = $info.tgt; $mt = $info.map
            $raw  = $data[$r,$srcCol]
            if($dateSet.ContainsKey($tgt)){
                $row[$tgt] = DateVal $raw
            } elseif($mt -ne ""){
                $mapped = MapN $raw $mt
                $row[$tgt] = if($mapped -ne ""){$mapped}else{AnyVal $raw}
            } else {
                $row[$tgt] = AnyVal $raw
            }
        }
        $rows.Add($row)
    }
    $wb.Close($false)
    Log "  → $($rows.Count) 筆" "Green"
    return ,$rows
}

$qcRows  = [System.Collections.Generic.List[object[]]]::new()
$facRows = [System.Collections.Generic.List[object[]]]::new()

# 轉換QC（Sheet1），並補算 target col75 = factoryMap(src col13 = 成衣產區)
foreach($f in $qcFiles){
    $rr = ConvertRows $f $qcColMap 75 $qcDateTgtCols 1
    # 補算 tgt75：重新讀 src col13 並套用 factoryMap
    $wb2 = $excel.Workbooks.Open($f.FullName,$false,$true)
    $ws2 = $wb2.Sheets(1)
    $data2 = $ws2.UsedRange.Value2
    $nR2   = $data2.GetLength(0)
    $nC2   = $data2.GetLength(1)
    $rowIdx = 0
    for($r=2; $r -le $nR2; $r++){
        if($null -eq $data2[$r,1] -or $data2[$r,1].ToString().Trim() -eq ""){ continue }
        if($rowIdx -lt $rr.Count -and $nC2 -ge 13){
            $area = $data2[$r,13]
            if($null -ne $area -and $area.ToString().Trim() -ne ""){
                # 取第一個工廠（空格分隔，如 "CAB-MH1 MH2" 只取 "CAB-MH1"）
                $firstArea = ($area.ToString().Trim() -split '\s+')[0]
                $mapped = MapN $firstArea "factory"
                if($mapped -ne ""){ $rr[$rowIdx][75] = $mapped }
            }
        }
        $rowIdx++
    }
    $wb2.Close($false)
    $rr | ForEach-Object { $qcRows.Add($_) }
}

# 轉換Factory（Sheet2 = 客戶-明細表），並補算 target col45 = factoryMap("産區-MAKER")
foreach($f in $facFiles){
    $rr = ConvertRows $f $facColMap 45 $facDateTgtCols 2
    # 補算 tgt45：重新讀 src col1(産區)、col2(MAKER) 並組合套 factoryMap
    $wb2 = $excel.Workbooks.Open($f.FullName,$false,$true)
    $si2 = if($wb2.Sheets.Count -ge 2){2}else{1}
    $ws2 = $wb2.Sheets($si2)
    $data2 = $ws2.UsedRange.Value2
    $nR2   = $data2.GetLength(0)
    $nC2   = $data2.GetLength(1)
    $rowIdx = 0
    for($r=2; $r -le $nR2; $r++){
        if($null -eq $data2[$r,1] -or $data2[$r,1].ToString().Trim() -eq ""){ continue }
        if($rowIdx -lt $rr.Count){
            $aStr = if($nC2 -ge 1 -and $null -ne $data2[$r,1]){$data2[$r,1].ToString().Trim()}else{""}
            $mStr = if($nC2 -ge 2 -and $null -ne $data2[$r,2]){$data2[$r,2].ToString().Trim()}else{""}
            $comb = if($aStr -ne "" -and $mStr -ne ""){"$aStr-$mStr"}elseif($aStr -ne ""){$aStr}else{""}
            if($comb -ne ""){
                $mapped = MapN $comb "factory"
                if($mapped -ne ""){ $rr[$rowIdx][45] = $mapped }
            }
        }
        $rowIdx++
    }
    $wb2.Close($false)
    $rr | ForEach-Object { $facRows.Add($_) }
}

# ── Step 4：日期標記 ──
$stamp = if($startDate -and $endDate){ "$startDate-$endDate" } else { Get-Date -Format "yyMMdd" }

# ── Step 5：寫入品管驗布報表 ──
Write-Host "`n[4] 寫入品管驗布報表_$stamp.xlsx..." -ForegroundColor Yellow
if($qcRows.Count -gt 0){
    $qcOut = "$basePath\品管驗布報表_$stamp.xlsx"
    Copy-Item $qcTemplate $qcOut -Force
    $wbQC = $excel.Workbooks.Open($qcOut)
    $wsQC = $wbQC.Sheets(1)
    $lastRow = 1
    for($r=$wsQC.UsedRange.Rows.Count; $r -ge 2; $r--){
        if($null -ne $wsQC.Cells($r,1).Value2){ $lastRow=$r; break }
    }
    $startRow = $lastRow + 1; $n = $qcRows.Count
    $arr = New-Object "object[,]" $n, 75
    for($r=0;$r -lt $n;$r++){
        for($c=1;$c -le 75;$c++){
            $v=$qcRows[$r][$c]
            $arr[$r,($c-1)]=if($null -eq $v -or $v -is [System.DBNull]){[System.DBNull]::Value}else{$v}
        }
    }
    $wsQC.Range($wsQC.Cells($startRow,1),$wsQC.Cells($startRow+$n-1,75)).Value2 = $arr
    # 日期欄格式
    foreach($c in @(2,16,17,68,69,70)){
        $wsQC.Range($wsQC.Cells($startRow,$c),$wsQC.Cells($startRow+$n-1,$c)).NumberFormat="yyyy/mm/dd"
    }
    $wsQC.UsedRange.HorizontalAlignment = -4131
    $wbQC.Save(); $wbQC.Close($false)
    Log "✓ 品管驗布報表_$stamp.xlsx ($n 行)" "Green"
} else { Log "⚠ 無品管數據" "Yellow" }

# ── Step 6：寫入工廠驗布報表 ──
Write-Host "`n[5] 寫入工廠驗布報表_$stamp.xlsx..." -ForegroundColor Yellow
if($facRows.Count -gt 0){
    $facOut = "$basePath\工廠驗布報表_$stamp.xlsx"
    Copy-Item $facTemplate $facOut -Force
    $wbFA = $excel.Workbooks.Open($facOut)
    $wsFA = $wbFA.Sheets(1)
    $lastRow = 1
    for($r=$wsFA.UsedRange.Rows.Count; $r -ge 2; $r--){
        if($null -ne $wsFA.Cells($r,1).Value2){ $lastRow=$r; break }
    }
    $startRow = $lastRow + 1; $n = $facRows.Count
    $arr = New-Object "object[,]" $n, 45
    for($r=0;$r -lt $n;$r++){
        for($c=1;$c -le 45;$c++){
            $v=$facRows[$r][$c]
            $arr[$r,($c-1)]=if($null -eq $v -or $v -is [System.DBNull]){[System.DBNull]::Value}else{$v}
        }
    }
    $wsFA.Range($wsFA.Cells($startRow,1),$wsFA.Cells($startRow+$n-1,45)).Value2 = $arr
    foreach($c in @(3)){
        $wsFA.Range($wsFA.Cells($startRow,$c),$wsFA.Cells($startRow+$n-1,$c)).NumberFormat="yyyy/mm/dd"
    }
    $wsFA.UsedRange.HorizontalAlignment = -4131
    $wbFA.Save(); $wbFA.Close($false)
    Log "✓ 工廠驗布報表_$stamp.xlsx ($n 行)" "Green"
} else { Log "⚠ 無工廠數據" "Yellow" }

# ── Step 7：移動源檔案（以固定中文名＋日期標記命名）──
Write-Host "`n[6] 移動源檔案到 completed rawdata..." -ForegroundColor Yellow
foreach($f in $qcFiles){
    $newName = "驗布結論報表_$stamp$($f.Extension)"
    Move-Item $f.FullName "$completedDir\$newName" -Force
    Log "移動: $newName" "Gray"
}
foreach($f in $facFiles){
    $newName = "後端驗布報表-客戶維度_$stamp$($f.Extension)"
    Move-Item $f.FullName "$completedDir\$newName" -Force
    Log "移動: $newName" "Gray"
}

if($Fast) {
    $excel.Calculation = -4106  # xlCalculationAutomatic
    $excel.ScreenUpdating = $true
}
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "`n$('='*60)" -ForegroundColor Cyan
Write-Host " ✓ 轉換完成$(if($Fast){' ⚡快速'})！" -ForegroundColor Green
if($qcRows.Count  -gt 0){ Write-Host "   品管驗布報表_$stamp.xlsx  ($($qcRows.Count) 筆)" -ForegroundColor Green }
if($facRows.Count -gt 0){ Write-Host "   工廠驗布報表_$stamp.xlsx  ($($facRows.Count) 筆)" -ForegroundColor Green }
Write-Host "$('='*60)" -ForegroundColor Cyan
