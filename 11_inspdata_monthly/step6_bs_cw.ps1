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
Write-Host ("Step6: startRow={0}, maxRow={1}, n={2}" -f $startRow,$maxRow,$n)

$ws.Range($ws.Cells($startRow,71),$ws.Cells($maxRow,101)).ClearContents() | Out-Null

$charWo = [System.Text.Encoding]::Unicode.GetString([byte[]](0x61,0x6C))
$charWu = [System.Text.Encoding]::Unicode.GetString([byte[]](0x59,0x6C))

$defNames = @{}
for($c=71;$c-le 101;$c++){
  $h=$ws.Cells(1,$c).Value2
  $defNames[$c]=if($null -ne $h){$h.ToString().Replace($charWo,$charWu)}else{""}
}

$groups=@(
  @{srcIdx=18;dstStart=71;dstEnd=83;weighted=$true},
  @{srcIdx=19;dstStart=84;dstEnd=86;weighted=$false},
  @{srcIdx=20;dstStart=87;dstEnd=87;weighted=$false},
  @{srcIdx=21;dstStart=88;dstEnd=88;weighted=$false},
  @{srcIdx=22;dstStart=89;dstEnd=93;weighted=$false},
  @{srcIdx=23;dstStart=94;dstEnd=96;weighted=$false},
  @{srcIdx=24;dstStart=97;dstEnd=100;weighted=$false}
)

$data = $ws.Range($ws.Cells($startRow,25),$ws.Cells($maxRow,48)).Value2
$outArr = New-Object "object[,]" $n,31
$filledBS=0; $filledCW=0; $failCount=0
$unclassified=@{}

for($r=1;$r-le $n;$r++){
  $concl = if($n -eq 1){$data[1]}else{$data[$r,1]}
  if($concl -ne "Fail"){continue}
  $failCount++
  $aiRaw = if($n -eq 1){$data[11]}else{$data[$r,11]}
  $aiVal=0.0
  $aiOK=($null -ne $aiRaw) -and [double]::TryParse($aiRaw.ToString(),[ref]$aiVal)
  $cwAdd=0

  foreach($grp in $groups){
    $si = $grp.srcIdx
    $srcRaw = if($n -eq 1){$data[$si]}else{$data[$r,$si]}
    if($null -eq $srcRaw -or $srcRaw.ToString().Trim() -eq ""){continue}
    $srcText=$srcRaw.ToString().Replace($charWo,$charWu)
    $grpMatched=$false
    for($c=$grp.dstStart;$c-le $grp.dstEnd;$c++){
      $dname=$defNames[$c]
      if($null -eq $dname -or $dname -eq ""){continue}
      if($srcText.Contains($dname)){
        $addVal=if($grp.weighted){if($aiOK -and $aiVal -gt 33){1.0}else{0.5}}else{1.0}
        $outArr[($r-1),($c-71)]=$addVal
        $filledBS++; $grpMatched=$true
      }
    }
    if(-not $grpMatched){
      $cwAdd++
      $key=$srcRaw.ToString().Trim()
      if($key -ne "" -and -not $unclassified.ContainsKey($key)){$unclassified[$key]=0}
      if($key -ne ""){$unclassified[$key]++}
    }
  }
  if($cwAdd -gt 0){$outArr[($r-1),30]=[double]$cwAdd; $filledCW++}
}

$ws.Range($ws.Cells($startRow,71),$ws.Cells($maxRow,101)).Value2=$outArr
Write-Host ("Step6 done: Fail={0}, BS~CV={1}, CW={2}" -f $failCount,$filledBS,$filledCW)
if($unclassified.Count -gt 0){
  Write-Host "--- Unclassified ---"
  $unclassified.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
    Write-Host ("  [{0}] x{1}" -f $_.Key,$_.Value)
  }
}
$wb.Save(); $wb.Close($false); $xl.Quit()
