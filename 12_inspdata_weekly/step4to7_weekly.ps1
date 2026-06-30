param([string]$qfFile, [int]$startRow=2, [switch]$NoKill)
if(-not $NoKill){
    $ep=Get-Process -Name EXCEL -ErrorAction SilentlyContinue
    if($ep){$ep|Stop-Process -Force; Start-Sleep -Seconds 2}
}
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$xl.Calculation=-4135; $xl.ScreenUpdating=$false
$wb=$xl.Workbooks.Open($qfFile); $ws=$wb.Sheets(1)
$maxRow=[int]$ws.UsedRange.Rows.Count; $n=$maxRow-$startRow+1
Write-Host ("Step4~7: startRow={0}, n={1}" -f $startRow,$n)

# ===== Step 4: Fix InspQty =====
$d4=$ws.Range($ws.Cells($startRow,28),$ws.Cells($maxRow,35)).Value2
$col27=[Array]::CreateInstance([object],$n,1)
$col3133=[Array]::CreateInstance([object],$n,3)
$col34=[Array]::CreateInstance([object],$n,1)
$fix4=0
for($r=1;$r-le $n;$r++){
  $a=if($n-eq 1){$d4[1]}else{$d4[$r,1]}
  $b=if($n-eq 1){$d4[2]}else{$d4[$r,2]}
  $cc=if($n-eq 1){$d4[3]}else{$d4[$r,3]}
  if($null -eq $a -and $null -eq $b -and $null -eq $cc){continue}
  $av=if($null -ne $a){[double]$a}else{0.0}
  $bv=if($null -ne $b){[double]$b}else{0.0}
  $cv=if($null -ne $cc){[double]$cc}else{0.0}
  $sum=$av+$bv+$cv; if($sum -le 0){continue}
  $col27[($r-1),0]=$sum
  $col3133[($r-1),0]=[Math]::Round($av/$sum,4)
  $col3133[($r-1),1]=[Math]::Round($bv/$sum,4)
  $col3133[($r-1),2]=[Math]::Round($cv/$sum,4)
  $avgPts=if($n-eq 1){$d4[8]}else{$d4[$r,8]}
  if($null -ne $avgPts){$col34[($r-1),0]=[Math]::Round([double]$avgPts*$sum/100)}
  $fix4++
}
$ws.Range($ws.Cells($startRow,27),$ws.Cells($maxRow,27)).Value2=$col27
$ws.Range($ws.Cells($startRow,31),$ws.Cells($maxRow,33)).Value2=$col3133
$ws.Range($ws.Cells($startRow,34),$ws.Cells($maxRow,34)).Value2=$col34
Write-Host ("Step4 done: {0} rows fixed" -f $fix4)

# ===== Step 5 & 6: 共用同一批讀取 =====
$charWo=[System.Text.Encoding]::Unicode.GetString([byte[]](0x61,0x6C))
$charWu=[System.Text.Encoding]::Unicode.GetString([byte[]](0x59,0x6C))
$defNames=@{}
for($c=71;$c-le 101;$c++){$h=$ws.Cells(1,$c).Value2; $defNames[$c]=if($null -ne $h){$h.ToString().Replace($charWo,$charWu)}else{""}}
$d56=$ws.Range($ws.Cells($startRow,25),$ws.Cells($maxRow,48)).Value2

# Step 5 output: cols 64~70 (7 cols)
$ws.Range($ws.Cells($startRow,64),$ws.Cells($maxRow,70)).ClearContents()|Out-Null
$out5=New-Object "object[,]" $n,7
$bl5=0; $bmbr5=0; $fail5=0

# Step 6 output: cols 71~101 (31 cols)
$ws.Range($ws.Cells($startRow,71),$ws.Cells($maxRow,101)).ClearContents()|Out-Null
$out6=New-Object "object[,]" $n,31
$groups=@(
  @{srcIdx=18;dstStart=71;dstEnd=83;weighted=$true},
  @{srcIdx=19;dstStart=84;dstEnd=86;weighted=$false},
  @{srcIdx=20;dstStart=87;dstEnd=87;weighted=$false},
  @{srcIdx=21;dstStart=88;dstEnd=88;weighted=$false},
  @{srcIdx=22;dstStart=89;dstEnd=93;weighted=$false},
  @{srcIdx=23;dstStart=94;dstEnd=96;weighted=$false},
  @{srcIdx=24;dstStart=97;dstEnd=100;weighted=$false}
)
$bs6=0; $cw6=0; $unclassified=@{}

for($r=1;$r-le $n;$r++){
  $concl=if($n-eq 1){$d56[1]}else{$d56[$r,1]}
  if($concl -ne "Fail"){continue}
  $fail5++
  $aiRaw=if($n-eq 1){$d56[11]}else{$d56[$r,11]}
  $aiVal=0.0; $aiOK=($null -ne $aiRaw) -and [double]::TryParse($aiRaw.ToString(),[ref]$aiVal)
  # Step5: BL (col64, idx0) from C42 (srcIdx=18)
  $v=if($n-eq 1){$d56[18]}else{$d56[$r,18]}
  if($null -ne $v -and $v.ToString().Trim() -ne ""){
    $out5[($r-1),0]=if($aiOK -and $aiVal -gt 33){1.0}else{0.5}; $bl5++
  }
  # Step5: BM~BR (col65~70, idx1~6) from C43~C48 (srcIdx=19~24)
  for($i=0;$i-lt 6;$i++){
    $v2=if($n-eq 1){$d56[(19+$i)]}else{$d56[$r,(19+$i)]}
    if($null -ne $v2 -and $v2.ToString().Trim() -ne ""){$out5[($r-1),($i+1)]=1.0; $bmbr5++}
  }
  # Step6: BS~CW
  $cwAdd=0
  foreach($grp in $groups){
    $srcRaw=if($n-eq 1){$d56[$grp.srcIdx]}else{$d56[$r,$grp.srcIdx]}
    if($null -eq $srcRaw -or $srcRaw.ToString().Trim() -eq ""){continue}
    $srcText=$srcRaw.ToString().Replace($charWo,$charWu)
    $grpMatched=$false
    for($c=$grp.dstStart;$c-le $grp.dstEnd;$c++){
      $dname=$defNames[$c]; if($null -eq $dname -or $dname -eq ""){continue}
      if($srcText.Contains($dname)){
        $addVal=if($grp.weighted){if($aiOK -and $aiVal -gt 33){1.0}else{0.5}}else{1.0}
        $out6[($r-1),($c-71)]=$addVal; $bs6++; $grpMatched=$true
      }
    }
    if(-not $grpMatched){$cwAdd++; $key=$srcRaw.ToString().Trim(); if($key-ne ""){if(-not $unclassified.ContainsKey($key)){$unclassified[$key]=0}; $unclassified[$key]++}}
  }
  if($cwAdd-gt 0){$out6[($r-1),30]=[double]$cwAdd; $cw6++}
}
$ws.Range($ws.Cells($startRow,64),$ws.Cells($maxRow,70)).Value2=$out5
$ws.Range($ws.Cells($startRow,71),$ws.Cells($maxRow,101)).Value2=$out6
Write-Host ("Step5 done: Fail={0}, BL={1}, BM~BR={2}" -f $fail5,$bl5,$bmbr5)
Write-Host ("Step6 done: Fail={0}, BS~CV={1}, CW={2}" -f $fail5,$bs6,$cw6)
if($unclassified.Count-gt 0){
  Write-Host "--- Unclassified ---"
  $unclassified.GetEnumerator()|Sort-Object Value -Descending|ForEach-Object{Write-Host ("  [{0}] x{1}" -f $_.Key,$_.Value)}
}

# ===== Step 7: Format =====
$maxCol=$ws.UsedRange.Columns.Count
$numCols=@(2,18,19,20,21,26,27,28,29,30,31,32,33,34,35,36,37,59,60,61,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,100,101)
if($startRow -le 2){$ws.Rows(1).HorizontalAlignment=-4108}
$ws.Range($ws.Cells($startRow,1),$ws.Cells($maxRow,$maxCol)).HorizontalAlignment=-4131
foreach($c in $numCols){if($c -le $maxCol){$ws.Range($ws.Cells($startRow,$c),$ws.Cells($maxRow,$c)).HorizontalAlignment=-4152}}
Write-Host ("Step7 done: formatted {0} rows" -f $n)

$xl.Calculation=-4106; $xl.ScreenUpdating=$true
$wb.Save(); $wb.Close($false); $xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null
