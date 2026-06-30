param([switch]$Fast)
$basePath = "E:\88. Claude\12_inspdata_weekly"
$rawPath = "$basePath\insprawdata_weekly"
$taskFile = "$basePath\00_plan_insprecord_QF_weekly.xlsx"
$qfPath = "$basePath\insprecord_QF\insprecord_QF.xlsx"  # 永久範本，只讀不改
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1
if($Fast){
    $excel.Calculation = -4135   # xlCalculationManual
    $excel.ScreenUpdating = $false
}
$wbT=$excel.Workbooks.Open($taskFile)
$nlD=$wbT.Sheets("NameList").UsedRange.Value2
$custMap=@{}; $vendorMap=@{}; $factoryMap=@{}; $qcMap=@{}; $locMap=@{}; $defNames=@{}; $defNameToCode=@{}
for($r=2;$r-le $nlD.GetLength(0);$r++){
  $o1=$nlD[$r,1];$s2=$nlD[$r,2]; if($o1 -and $s2){$custMap[$o1.ToString().Trim()]=$s2.ToString().Trim()}
  $o3=$nlD[$r,3];$s4=$nlD[$r,4]; if($o3 -and $s4){$vendorMap[$o3.ToString().Trim()]=$s4.ToString().Trim()}
  $o5=$nlD[$r,5];$s6=$nlD[$r,6]; if($o5 -and $s6){$locMap[$o5.ToString().Trim()]=$s6.ToString().Trim()}
  $o7=$nlD[$r,7];$s8=$nlD[$r,8]; if($o7 -and $s8){$factoryMap[$o7.ToString().Trim()]=$s8.ToString().Trim()}
  $qc9=$nlD[$r,9];$qc10=$nlD[$r,10]; if($qc9 -and $qc10){$qcMap[$qc9.ToString().Trim()]=$qc10.ToString().Trim()}
  $dc=$nlD[$r,12]; $dn=$nlD[$r,14]
  if($dc -and $dn){$defNames[$dc.ToString().Trim()]=$dn.ToString().Trim(); $defNameToCode[$dn.ToString().Trim()]=$dc.ToString().Trim()}
}
$wbT.Close($false)
Write-Host ("NameList: cust="+$custMap.Count+" vendor="+$vendorMap.Count+" factory="+$factoryMap.Count+" qc="+$qcMap.Count+" loc="+$locMap.Count+" defects="+$defNames.Count)
$defCodeToQFIdx=@{"D103"=70;"D104"=71;"D105"=72;"D106"=73;"D107"=74;"D108"=75;"D109"=76;"D110"=77;"D111"=78;"D112"=79;"D114"=80;"D115"=81;"D116"=82;"D201"=83;"D202"=84;"D209"=85;"D301"=86;"D401"=87;"D502"=88;"D503"=89;"D504"=90;"D505"=91;"D506"=92;"D601"=93;"D605"=94;"D609"=95;"D610"=96;"D703"=97;"D705"=98;"D706"=99}
function MapN($v,$m){if($null -eq $v -or $v.ToString().Trim() -eq ""){return ""}; $k=$v.ToString().Trim(); if($m.ContainsKey($k)){return $m[$k]}; return $k}
function CleanStr($v){if($null -eq $v){return ""}; return ($v.ToString() -replace "[\r\n]"," ").Trim()}
function EngOnly($v){if($null -eq $v){return ""}; $s=($v.ToString() -replace "[^\x00-\x7F]","") -replace ","," "; return ($s -replace "\s+"," ").Trim()}
function DateVal($v){if($null -eq $v -or $v.ToString().Trim() -eq ""){return [DBNull]::Value}; if($v -is [double]){return [double]$v}; $d=[DateTime]::MinValue; if([DateTime]::TryParse($v.ToString(),[ref]$d)){return $d.ToOADate()}; return [DBNull]::Value}
function SafeDiv($a,$b){if($b -isnot [double] -or $b -eq 0){return [DBNull]::Value}; try{return [double]$a/[double]$b}catch{return [DBNull]::Value}}
function HasDef($t,$d){if($null -eq $t -or $t.ToString().Trim() -eq "" -or $null -eq $d -or $d -eq ""){return [DBNull]::Value}; if($t.ToString().Contains($d)){return 1}; return [DBNull]::Value}
function NonEmpty($v){if($v -isnot [string] -or $v.Trim() -eq ""){return [DBNull]::Value}; return 1}
function NumD($v){if($null -eq $v -or $v.ToString().Trim() -eq ""){return [DBNull]::Value}; try{return [double]$v}catch{return [DBNull]::Value}}
function Concl($avg,$b,$c){if($null -eq $avg -or $avg.ToString().Trim() -eq ""){return ""}; try{$a=[double]$avg; $bv=if($b -and "$b"-ne ""){[double]$b}else{0}; $cv=if($c -and "$c"-ne ""){[double]$c}else{0}; if($cv-gt 0 -or $a-gt 33){return "Fail"}; if($a-le 10 -and $bv-eq 0 -and $cv-eq 0){return "Pass"}; return "Discuss"}catch{}; return ""}
function First1($v){if($null -eq $v -or $v.ToString().Trim() -eq ""){return ""}; $parts=$v.ToString() -split "[,;/ ]"; foreach($p in $parts){$t=$p.Trim(); if($t-ne ""){return $t}}; return ""}
function CleanFiber($v){if($null -eq $v){return ""}; return ($v.ToString() -replace "^/+","" -replace " /+"," ").Trim()}
function GetCatSrcIdx($code){try{$n=[int]$code.Substring(1); if($n-ge 101-and $n-le 120){return 41}; if($n-ge 201-and $n-le 209){return 42}; if($n-eq 301){return 43}; if($n-eq 401){return 44}; if($n-ge 501-and $n-le 506){return 45}; if($n-ge 601-and $n-le 610){return 47}; if($n-ge 701-and $n-le 707){return 46}}catch{}; return 41}
if(!(Test-Path $qfPath)){
  Write-Host "ERROR: QF 範本檔不存在 $qfPath"; exit 1
}
# 只讀取標頭，確認欄位結構
$wbQ0=$excel.Workbooks.Open($qfPath,$false,$true); $wsQ0=$wbQ0.Sheets(1)
$qfH=@{}; for($c=1;$c-le 101;$c++){$h=$wsQ0.Cells(1,$c).Value2; if($h){$qfH[$c]=$h.ToString()}}
$wbQ0.Close($false)
Write-Host "範本欄位確認完成"
$QA=@{DATE=15;CUST=4;BUYER=5;DEPT=6;STYLE=7;PO=8;OTYPE=9;VENDOR=11;AREA=12;GAREA=13;FIBER=16;KNIT=17;FABTYPE=18;DYE=19;WIDTH=22;GSM=23;OQTY=24;OUNIT=25;QC=26;COLOR=27;RCVQTY=30;INSPQTY=31;AVGPTS=32;AQTY=33;BQTY=34;CQTY=35;BUZI=40;WAIG=41;SHOUF=42;WEIXIE=43;YANSE=44;QITA=45;GUIGE=46;CHUZI=47;C20ZT=49;C20YJ=50;C20XC=51;C20JJ=52;BZYL=54;FSYJG=55;QMSG=56;BEIZHU=57;EMAIL=58;EXPDATE=59;CUTDATE=60;GMTEXP=61;VMYJ=62;SIGNST=63}
$FA=@{AREA=1;MAKER=2;DATE=4;CUST=6;STYLE=7;DEPT=8;BUYER=9;VMAP=12;LOC=13;FIBER=16;KNIT=14;FABTYPE=18;DYE=15;STDW=44;COLOR=19;PO=20;BLNO=21;RCVQTY=22;INSPQTY=23;AQTY=25;BQTY=26;CQTY=27;TOTDEF=28;AVGPTS=29;D1=31;D2=32;D3=33;D4=34;D5=35;WAIG=36;SHOUF=37;WEIXIE=38;YANSE=39;QITA=40;GUIGE=41;GMM2=42;REALW=43;SHORTRATE1=45;SHORTRATE2=46}
$qaStr=[string][char]0x54C1+[string][char]0x7BA1
$facStr=[string][char]0x5DE5+[string][char]0x5EE0
$allRows=[System.Collections.Generic.List[object[]]]::new()
$srcFiles=Get-ChildItem $rawPath | Where-Object { ($_.Extension -eq ".xlsx" -or $_.Extension -eq ".xls") -and $_.Name -notlike "~*" } | Sort-Object Name
Write-Host ("Source files: "+$srcFiles.Count)
foreach($file in $srcFiles){
  $fn=$file.Name
  $isQA=$fn.Contains([string][char]0x54C1)
  $isFac=$fn.Contains([string][char]0x5DE5)
  if(-not $isQA -and -not $isFac){continue}
  $type=if($isQA){"QA"}else{"Factory"}; Write-Host ("  "+$type+": "+$fn)
  $wb=$excel.Workbooks.Open($file.FullName,$false,$true)
  # 品管和工廠都找頁名含"明細"的頁
  $ws=$null
  $mdStr=[string][char]0x660E+[string][char]0x7D30   # 明細
  for($si=1;$si-le $wb.Sheets.Count;$si++){
    $sh=$wb.Sheets($si)
    if($sh.Name.Contains($mdStr)){$ws=$sh; break}
  }
  if($null -eq $ws){Write-Host "    WARNING: target sheet not found, using sheet 1"; $ws=$wb.Sheets(1)}
  $data=$ws.UsedRange.Value2; $nR=$data.GetLength(0); $rc=0
  for($r=2;$r-le $nR;$r++){
    $kv=if($isQA){$data[$r,$QA.STYLE]}else{$data[$r,$FA.STYLE]}
    if($null -eq $kv -or $kv.ToString().Trim() -eq ""){continue}
    $row=New-Object object[] 101
    for($ci=0;$ci-le 100;$ci++){$row[$ci]=[DBNull]::Value}
    if($isQA){
      $row[0]=$qaStr
      $row[1]=DateVal $data[$r,$QA.DATE]
      $row[2]=MapN $data[$r,$QA.CUST] $custMap
      $row[3]=CleanStr $data[$r,$QA.STYLE]
      $row[4]=CleanStr $data[$r,$QA.DEPT]
      $row[5]=EngOnly $data[$r,$QA.BUYER]
      $row[6]=MapN $data[$r,$QA.VENDOR] $vendorMap
      $row[7]=MapN $data[$r,$QA.AREA] $locMap
      $row[8]=CleanStr $data[$r,$QA.PO]
      $row[9]=CleanStr $data[$r,$QA.OTYPE]
      $row[10]=CleanStr $data[$r,$QA.GAREA]
      $row[11]=MapN (First1 $data[$r,$QA.GAREA]) $factoryMap
      $row[12]=[DBNull]::Value
      $row[13]=CleanFiber $data[$r,$QA.FIBER]
      $row[14]=CleanStr $data[$r,$QA.KNIT]
      $row[15]=CleanStr $data[$r,$QA.FABTYPE]
      $row[16]=CleanStr $data[$r,$QA.DYE]
      $width=NumD $data[$r,$QA.WIDTH]; $row[17]=$width
      $gsm=NumD $data[$r,$QA.GSM]; $row[18]=$gsm
      $insp=NumD $data[$r,$QA.INSPQTY]
      if($gsm -is [double] -and $width -is [double]){try{$row[19]=[Math]::Round([double]$gsm*([double]$width+2)/43,1)}catch{}}
      $row[20]=NumD $data[$r,$QA.OQTY]
      $row[21]=CleanStr $data[$r,$QA.OUNIT]
      $row[22]=MapN (EngOnly $data[$r,$QA.QC]) $qcMap
      $row[23]=CleanStr $data[$r,$QA.COLOR]
      $avgPts=NumD $data[$r,$QA.AVGPTS]
      $row[24]=Concl $avgPts $data[$r,$QA.BQTY] $data[$r,$QA.CQTY]
      $rcv=NumD $data[$r,$QA.RCVQTY]; $row[25]=if($rcv -is [double]){[Math]::Round($rcv)}else{[DBNull]::Value}
      $row[26]=if($insp -is [double]){[Math]::Round($insp)}else{[DBNull]::Value}
      $aq=NumD $data[$r,$QA.AQTY]; $bq=NumD $data[$r,$QA.BQTY]; $cq=NumD $data[$r,$QA.CQTY]
      $row[27]=if($aq -is [double]){[Math]::Round($aq)}else{[DBNull]::Value}
      $row[28]=if($bq -is [double]){[Math]::Round($bq)}else{[DBNull]::Value}
      $row[29]=if($cq -is [double]){[Math]::Round($cq)}else{[DBNull]::Value}
      $row[30]=SafeDiv $aq $insp; $row[31]=SafeDiv $bq $insp; $row[32]=SafeDiv $cq $insp
      $totDef=if($avgPts -is [double] -and $insp -is [double]){[double]$avgPts*[double]$insp/100}else{$null}
      $row[33]=if($null -ne $totDef){[Math]::Round($totDef)}else{[DBNull]::Value}
      $row[34]=if($avgPts -is [double]){[Math]::Round($avgPts,1)}else{[DBNull]::Value}
      if($null -ne $totDef -and $width -is [double] -and $insp -is [double] -and [double]$width-ne 0 -and [double]$insp-ne 0){try{$row[35]=[Math]::Round($totDef*3600/[double]$width/[double]$insp,1)}catch{}}
      if($avgPts -is [double]){try{$row[36]=[double]$avgPts*0.3/100}catch{}}
      $row[41]=CleanStr $data[$r,$QA.BUZI]
      $row[42]=CleanStr $data[$r,$QA.WAIG]
      $row[43]=CleanStr $data[$r,$QA.SHOUF]
      $row[44]=CleanStr $data[$r,$QA.WEIXIE]
      $row[45]=CleanStr $data[$r,$QA.YANSE]
      $row[46]=CleanStr $data[$r,$QA.GUIGE]
      $row[47]=CleanStr $data[$r,$QA.QITA]
      $row[48]=CleanStr $data[$r,$QA.CHUZI]
      $row[49]=CleanStr $data[$r,$QA.C20ZT]
      $row[50]=CleanStr $data[$r,$QA.C20YJ]
      $row[51]=CleanStr $data[$r,$QA.C20XC]
      $row[52]=CleanStr $data[$r,$QA.C20JJ]
      $row[53]=CleanStr $data[$r,$QA.BZYL]
      $row[54]=CleanStr $data[$r,$QA.FSYJG]
      $row[55]=CleanStr $data[$r,$QA.QMSG]
      $row[56]=CleanStr $data[$r,$QA.BEIZHU]
      $row[57]=CleanStr $data[$r,$QA.EMAIL]
      $row[58]=DateVal $data[$r,$QA.EXPDATE]
      $row[59]=DateVal $data[$r,$QA.CUTDATE]
      $row[60]=DateVal $data[$r,$QA.GMTEXP]
      $row[61]=CleanStr $data[$r,$QA.VMYJ]
      $row[62]=CleanStr $data[$r,$QA.SIGNST]
    } else {
      $row[0]=$facStr
      $row[1]=DateVal $data[$r,$FA.DATE]
      $row[2]=MapN $data[$r,$FA.CUST] $custMap
      $row[3]=CleanStr $data[$r,$FA.STYLE]
      $row[4]=CleanStr $data[$r,$FA.DEPT]
      $row[5]=EngOnly $data[$r,$FA.BUYER]
      $row[6]=MapN $data[$r,$FA.VMAP] $vendorMap
      $row[7]=MapN (CleanStr $data[$r,$FA.LOC]) $locMap
      $row[8]=CleanStr $data[$r,$FA.PO]
      $row[9]=[DBNull]::Value
      $ar2=CleanStr $data[$r,$FA.AREA]; $mk2=CleanStr $data[$r,$FA.MAKER]
      $comb=if($ar2-ne "" -and $mk2-ne ""){"$ar2-$mk2"}elseif($ar2-ne ""){$ar2}else{""}
      $row[10]=$comb
      $row[11]=$comb
      $row[12]=CleanStr $data[$r,$FA.BLNO]
      $row[13]=CleanFiber $data[$r,$FA.FIBER]
      $row[14]=CleanStr $data[$r,$FA.KNIT]
      $row[15]=CleanStr $data[$r,$FA.FABTYPE]
      $row[16]=CleanStr $data[$r,$FA.DYE]
      $width=NumD $data[$r,$FA.STDW]; $row[17]=$width
      $row[23]=CleanStr $data[$r,$FA.COLOR]
      $avgPts=NumD $data[$r,$FA.AVGPTS]
      $row[24]=Concl $avgPts $data[$r,$FA.BQTY] $data[$r,$FA.CQTY]
      $insp=NumD $data[$r,$FA.INSPQTY]
      $aq=NumD $data[$r,$FA.AQTY]; $bq=NumD $data[$r,$FA.BQTY]; $cq=NumD $data[$r,$FA.CQTY]
      $rcv=NumD $data[$r,$FA.RCVQTY]; $row[25]=if($rcv -is [double]){[Math]::Round($rcv)}else{[DBNull]::Value}
      $row[26]=if($insp -is [double]){[Math]::Round($insp)}else{[DBNull]::Value}
      $row[27]=if($aq -is [double]){[Math]::Round($aq)}else{[DBNull]::Value}
      $row[28]=if($bq -is [double]){[Math]::Round($bq)}else{[DBNull]::Value}
      $row[29]=if($cq -is [double]){[Math]::Round($cq)}else{[DBNull]::Value}
      $row[30]=SafeDiv $aq $insp; $row[31]=SafeDiv $bq $insp; $row[32]=SafeDiv $cq $insp
      $totDef=NumD $data[$r,$FA.TOTDEF]
      $row[33]=if($totDef -is [double]){[Math]::Round($totDef)}else{[DBNull]::Value}
      $row[34]=if($avgPts -is [double]){[Math]::Round($avgPts,1)}else{[DBNull]::Value}
      if($totDef -is [double] -and $width -is [double] -and $insp -is [double] -and [double]$width-ne 0 -and [double]$insp-ne 0){try{$row[35]=[Math]::Round([double]$totDef*3600/[double]$width/[double]$insp,1)}catch{}}
      if($avgPts -is [double]){try{$row[36]=[double]$avgPts*0.3/100}catch{}}
      $row[37]=CleanStr $data[$r,$FA.GMM2]
      $row[38]=CleanStr $data[$r,$FA.REALW]
      $row[39]=CleanStr $data[$r,$FA.SHORTRATE1]
      $row[40]=CleanStr $data[$r,$FA.SHORTRATE2]
      $wv=CleanStr $data[$r,$FA.WAIG];   if($wv-ne ""){$row[42]=$wv}
      $sv=CleanStr $data[$r,$FA.SHOUF];  if($sv-ne ""){$row[43]=$sv}
      $xv=CleanStr $data[$r,$FA.WEIXIE]; if($xv-ne ""){$row[44]=$xv}
      $yv=CleanStr $data[$r,$FA.YANSE];  if($yv-ne ""){$row[45]=$yv}
      $gv=CleanStr $data[$r,$FA.GUIGE];  if($gv-ne ""){$row[46]=$gv}
      $qv=CleanStr $data[$r,$FA.QITA];   if($qv-ne ""){$row[47]=$qv}
      for($d=$FA.D1;$d-le $FA.D5;$d++){
        $dv=CleanStr $data[$r,$d]; if($dv -eq ""){continue}
        $catIdx=41
        if($defNameToCode.ContainsKey($dv)){$catIdx=GetCatSrcIdx $defNameToCode[$dv]}
        $cur=$row[$catIdx]; if($cur -isnot [string] -or $cur -eq ""){$row[$catIdx]=$dv}else{$row[$catIdx]+=",$dv"}
      }
    }
    $row[63]=NonEmpty $row[41]; $row[64]=NonEmpty $row[42]
    $row[65]=NonEmpty $row[43]; $row[66]=NonEmpty $row[44]
    $row[67]=NonEmpty $row[45]; $row[68]=NonEmpty $row[46]
    $row[69]=NonEmpty $row[47]
    foreach($code in $defCodeToQFIdx.Keys){
      $qfIdx=$defCodeToQFIdx[$code]
      $srcIdx=GetCatSrcIdx $code
      $dn=if($defNames.ContainsKey($code)){$defNames[$code]}else{""}
      if($dn-ne ""){$row[$qfIdx]=HasDef $row[$srcIdx] $dn}
    }
    $allRows.Add([object[]]$row.Clone()); $rc++
  }
  $wb.Close($false)
  Write-Host ("    Added "+$rc+" rows")
}
$n=$allRows.Count; Write-Host ("Total: "+$n+" rows")
$arr=New-Object "object[,]" $n,101
for($r=0;$r-lt $n;$r++){for($c=0;$c-lt 101;$c++){$arr[$r,$c]=$allRows[$r][$c]}}
Write-Host "Writing to QF..."
# 複製範本到暫存檔，開啟複本寫入（不動原範本）
$tempWrite = "$basePath\insprecord_QF\insprecord_QF_writing.xlsx"
Copy-Item $qfPath $tempWrite -Force
$wbQF=$excel.Workbooks.Open($tempWrite); $wsQF=$wbQF.Sheets(1)
# 清空複本的舊資料（保留標頭）
$tmplRows=$wsQF.UsedRange.Rows.Count
if($tmplRows -ge 2){
    try{ $wsQF.Range($wsQF.Cells(2,1),$wsQF.Cells($tmplRows,101)).Delete() | Out-Null }catch{}
}
$startRow = 2; $eRow=$startRow+$n-1
# 先格式化字串欄為 "@"，再一次寫入（省掉第二次逐欄覆寫）
$strCols=@(3,4,5,6,7,8,9,10,13,14,15,16,17,18,23,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,100)
Write-Host "Pre-formatting string columns..."
foreach($c in $strCols){ $wsQF.Range($wsQF.Cells($startRow,$c),$wsQF.Cells($eRow,$c)).NumberFormat="@" }
Write-Host ("Writing all data (single pass)...")
$wsQF.Range($wsQF.Cells($startRow,1),$wsQF.Cells($eRow,101)).Value2=$arr
Write-Host "Applying formats..."
foreach($c in @(26,27,28,29,30)){$wsQF.Range($wsQF.Cells($startRow,$c),$wsQF.Cells($eRow,$c)).NumberFormat="#,##0"}
$wsQF.Range($wsQF.Cells($startRow,34),$wsQF.Cells($eRow,34)).NumberFormat="#,##0"
foreach($c in @(35,36)){$wsQF.Range($wsQF.Cells($startRow,$c),$wsQF.Cells($eRow,$c)).NumberFormat="0.0"}
foreach($c in @(31,32,33,37)){$wsQF.Range($wsQF.Cells($startRow,$c),$wsQF.Cells($eRow,$c)).NumberFormat="0.0%"}
foreach($c in @(2,59,60,61)){$wsQF.Range($wsQF.Cells($startRow,$c),$wsQF.Cells($eRow,$c)).NumberFormat="yyyy/mm/dd"}
$wsQF.UsedRange.HorizontalAlignment=-4131
$numCols=@(2,18,19,20,21,26,27,28,29,30,31,32,33,34,35,36,37,59,60,61,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,100,101)
foreach($c in $numCols){$wsQF.Range($wsQF.Cells($startRow,$c),$wsQF.Cells($eRow,$c)).HorizontalAlignment=-4152}
Write-Host "Saving..."
# 從記憶體陣列取日期範圍（$arr 第2欄=index 1 = C2驗布日）
$minD=$null; $maxD=$null
for($r=0;$r-lt $n;$r++){
  $d=$arr[$r,1]
  if($null -ne $d -and $d -ne [DBNull]::Value){
    $dv=[double]0
    if([double]::TryParse($d.ToString(),[ref]$dv) -and $dv -gt 0){
      if($null -eq $minD -or $dv -lt $minD){$minD=$dv}
      if($null -eq $maxD -or $dv -gt $maxD){$maxD=$dv}
    }
  }
}
if($null -ne $minD -and $null -ne $maxD){
  $s=[DateTime]::FromOADate($minD).ToString("yyMMdd")
  $e=[DateTime]::FromOADate($maxD).ToString("yyMMdd")
  $saveName = "$basePath\insprecord_QF\insprecord_QF_$s-$e.xlsx"
} else {
  $saveName = "$basePath\insprecord_QF\insprecord_QF_unknown.xlsx"
}
$xlFmt = 51  # xlOpenXMLWorkbook (.xlsx)
$wbQF.SaveAs($saveName, $xlFmt)
$wbQF.Close($false)
if($Fast){
    $excel.Calculation = -4106   # xlCalculationAutomatic
    $excel.ScreenUpdating = $true
}
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)|Out-Null
# 刪除暫存複本
Remove-Item $tempWrite -Force -ErrorAction SilentlyContinue
Write-Host ("DONE! {0} rows. → {1}" -f $n,(Split-Path $saveName -Leaf))
