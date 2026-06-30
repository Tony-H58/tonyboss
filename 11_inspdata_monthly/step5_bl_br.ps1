param([string]$qfFile, [int]$startRow=2, [switch]$NoKill)
if(-not $NoKill){
    Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2
}
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($qfFile)
$ws = $wb.Sheets(1)
$maxRow = [int]$ws.UsedRange.Rows.Count
$n = $maxRow - $startRow + 1
Write-Host ("Step5: startRow={0}, maxRow={1}, n={2}" -f $startRow,$maxRow,$n)

# 清除新增行 BL~BR
$ws.Range($ws.Cells($startRow,64),$ws.Cells($maxRow,70)).ClearContents() | Out-Null

# 整批讀取 col25~48（24欄）
$data = $ws.Range($ws.Cells($startRow,25),$ws.Cells($maxRow,48)).Value2

# 欄位偏移（1-based）：col25=1, col35=11, col42=18, col43~48=19~24
$idx25=1; $idx35=11; $idx42=18

$outArr = New-Object "object[,]" $n,7
$blCount=0; $bmbrCount=0; $failCount=0

for($r=1;$r-le $n;$r++){
  $concl = if($n -eq 1){$data[$idx25]}else{$data[$r,$idx25]}
  if($concl -ne "Fail"){continue}
  $failCount++
  $aiRaw = if($n -eq 1){$data[$idx35]}else{$data[$r,$idx35]}
  $aiVal=0.0
  $aiOK=($null -ne $aiRaw) -and [double]::TryParse($aiRaw.ToString(),[ref]$aiVal)

  $v = if($n -eq 1){$data[$idx42]}else{$data[$r,$idx42]}
  if($null -ne $v -and $v.ToString().Trim() -ne ""){
    $outArr[($r-1),0]=if($aiOK -and $aiVal -gt 33){1.0}else{0.5}
    $blCount++
  }
  for($i=0;$i-lt 6;$i++){
    $colIdx = 19 + $i
    $v2 = if($n -eq 1){$data[$colIdx]}else{$data[$r,$colIdx]}
    if($null -ne $v2 -and $v2.ToString().Trim() -ne ""){
      $outArr[($r-1),($i+1)]=1.0; $bmbrCount++
    }
  }
}

$ws.Range($ws.Cells($startRow,64),$ws.Cells($maxRow,70)).Value2=$outArr
Write-Host ("Step5 done: Fail={0}, BL={1}, BM~BR={2}" -f $failCount,$blCount,$bmbrCount)
$wb.Save(); $wb.Close($false); $xl.Quit()
