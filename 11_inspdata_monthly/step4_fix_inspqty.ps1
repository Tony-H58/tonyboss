param([string]$qfFile, [int]$startRow=2, [switch]$NoKill)
if(-not $NoKill){
    Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2
}
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($qfFile)
$ws = $wb.Sheets(1)
$maxRow = $ws.UsedRange.Rows.Count
$n = $maxRow - $startRow + 1

# 整批讀取 col28~35（A=28,B=29,C=30 → 抽驗=27; A%,B%,C%=31~33; 總瑕疵=34; 百碼=35）
$data = $ws.Range($ws.Cells($startRow,28),$ws.Cells($maxRow,35)).Value2

# 輸出陣列
$col27    = [Array]::CreateInstance([object], $n, 1)
$col31_33 = [Array]::CreateInstance([object], $n, 3)
$col34    = [Array]::CreateInstance([object], $n, 1)
$fixCount = 0

for($r=1;$r-le $n;$r++){
  $ri = $r - 1
  $a  = if($n -eq 1){$data[1]}else{$data[$r,1]}
  $b  = if($n -eq 1){$data[2]}else{$data[$r,2]}
  $cc = if($n -eq 1){$data[3]}else{$data[$r,3]}
  if($null -eq $a -and $null -eq $b -and $null -eq $cc){continue}
  $av=if($null -ne $a){[double]$a}else{0.0}
  $bv=if($null -ne $b){[double]$b}else{0.0}
  $cv=if($null -ne $cc){[double]$cc}else{0.0}
  $sum=$av+$bv+$cv
  if($sum -le 0){continue}
  $col27[$ri,0]=$sum
  $col31_33[$ri,0]=[Math]::Round($av/$sum,4)
  $col31_33[$ri,1]=[Math]::Round($bv/$sum,4)
  $col31_33[$ri,2]=[Math]::Round($cv/$sum,4)
  $avgPts = if($n -eq 1){$data[8]}else{$data[$r,8]}
  if($null -ne $avgPts){$col34[$ri,0]=[Math]::Round([double]$avgPts*$sum/100)}
  $fixCount++
}

# 整批寫回
$ws.Range($ws.Cells($startRow,27),$ws.Cells($maxRow,27)).Value2=$col27
$ws.Range($ws.Cells($startRow,31),$ws.Cells($maxRow,33)).Value2=$col31_33
$ws.Range($ws.Cells($startRow,34),$ws.Cells($maxRow,34)).Value2=$col34

Write-Host ("Step4 done: {0} rows fixed" -f $fixCount)
$wb.Save(); $wb.Close($false); $xl.Quit()
