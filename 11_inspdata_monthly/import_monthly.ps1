param([string]$year)
$basePath = "E:\88. Claude\11_inspdata_monthly"
$rawPath = "$basePath\insprawdata_monthly"
$taskFile = "$basePath\00_plan_insprecord_QF_monthly.xlsx"
# QF 檔路徑：固定暫用名，導入後依實際日期改名
$qfPath = "$basePath\insprecord_QF\insprecord_QF_current.xlsx"
$yy = $year.Substring(2)  # e.g. "2026" -> "26"
# 找同年份的 QF 檔（yymmdd-yymmdd 命名）
$existingQF = Get-ChildItem "$basePath\insprecord_QF" -Filter "insprecord_QF_${yy}*.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if($existingQF -and $existingQF.FullName -ne $qfPath) {
  Rename-Item $existingQF.FullName $qfPath -Force
  Write-Host "使用現有檔案: $($existingQF.Name)"
}
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
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
$grps6=@(@{s=18;d1=71;d2=83;w=$true},@{s=19;d1=84;d2=86;w=$false},@{s=20;d1=87;d2=87;w=$false},@{s=21;d1=88;d2=88;w=$false},@{s=22;d1=89;d2=93;w=$false},@{s=23;d1=94;d2=96;w=$false},@{s=24;d1=97;d2=100;w=$false})
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
# 開啟 QF 檔，讀取標題，計算追加起始行
if(!(Test-Path $qfPath)){
  Write-Host "ERROR: QF 範本檔不存在 $qfPath"; exit 1
}
$wbQ0=$excel.Workbooks.Open($qfPath,$false,$false); $wsQ0=$wbQ0.Sheets(1)
$qfH=@{}; for($c=1;$c-le 101;$c++){$h=$wsQ0.Cells(1,$c).Value2; if($h){$qfH[$c]=$h.ToString()}}
# 用 End(xlUp) 一次定位最後一行（避免逐格掃描）
$lastDataRow = $wsQ0.Cells($wsQ0.Rows.Count,1).End(-4162).Row
if($lastDataRow -lt 2){ $lastDataRow = 1 }
Write-Host ("現有資料末行: {0}" -f $lastDataRow)
# 讀取 step6 所需的瑕疵名稱標題（cols 71~101）
$charWo=[System.Text.Encoding]::Unicode.GetString([byte[]](0x61,0x6C))
$charWu=[System.Text.Encoding]::Unicode.GetString([byte[]](0x59,0x6C))
$defN6=@{}; for($c6h=71;$c6h-le 101;$c6h++){$h=$wsQ0.Cells(1,$c6h).Value2; $defN6[$c6h]=if($null-ne $h){$h.ToString().Replace($charWo,$charWu)}else{""}}
$wbQ0.Close($false)
$QA=@{DATE=15;CUST=4;STYLE=7;DEPT=6;BUYER=5;VENDOR=11;AREA=12;GAREA=13;FIBER=16;KNIT=17;FABTYPE=18;DYE=19;WIDTH=22;GSM=23;OQTY=24;OUNIT=25;QC=26;COLOR=27;RCVQTY=30;INSPQTY=31;AVGPTS=32;AQTY=33;BQTY=34;CQTY=35;BUZI=40;WAIG=41;SHOUF=42;WEIXIE=43;YANSE=44;GUIGE=46;QITA=45;CHUZI=47;C20ZT=49;C20YJ=50;C20XC=51;C20JJ=52;BZYL=54;FSYJG=55;QMSG=56;BEIZHU=57;EMAIL=58;EXPDATE=59;CUTDATE=60;GMTEXP=61;VMYJ=62;SIGNST=63}
$FA=@{AREA=1;MAKER=2;DATE=4;CUST=6;STYLE=7;DEPT=8;BUYER=9;VMAP=12;LOC=13;PO=20;BLNO=21;FIBER=16;KNIT=14;FABTYPE=18;DYE=15;STDW=44;COLOR=19;RCVQTY=22;INSPQTY=23;AQTY=25;BQTY=26;CQTY=27;TOTDEF=28;AVGPTS=29;D1=31;D2=32;D3=33;D4=34;D5=35;WAIG=36;SHOUF=37;WEIXIE=38;YANSE=39;GUIGE=41;QITA=40;GMM2=42;REALW=43;SHORTRATE1=45;SHORTRATE2=46}
$qaStr=[string][char]0x54C1+[string][char]0x7BA1
$facStr=[string][char]0x5DE5+[string][char]0x5EE0
$allRows=[System.Collections.Generic.List[object[]]]::new()
$maxOADate=$null  # 記錄新增列最大日期，供改名使用（不需再開 xl2）
$blC=0;$bmbrC=0;$fail56=0;$bs6=0;$cw6=0;$uncl6=@{}
$yy = $year.Substring(2)  # e.g. "2026" -> "26"
# 匹配起始年或結束年含目標年（支援跨年rawdata）
$srcFiles=Get-ChildItem $rawPath -File | Where-Object {
    $_.Name -match "_(\d{2})\d{4}-(\d{2})\d{4}" -and
    ($_.Name -like "*.xlsx" -or $_.Name -like "*.xls") -and
    $_.Name -notlike "Imported_*" -and
    ("20$($Matches[1])" -eq $year -or "20$($Matches[2])" -eq $year)
} | Sort-Object Name
Write-Host ("Source files: "+$srcFiles.Count)
foreach($file in $srcFiles){
  $fn=$file.Name
  $isQA=$fn.Contains([string][char]0x54C1)
  $isFac=$fn.Contains([string][char]0x5DE5)
  if(-not $isQA -and -not $isFac){continue}
  $type=if($isQA){"QA"}else{"Factory"}; Write-Host ("  "+$type+": "+$fn)
  $wb=$excel.Workbooks.Open($file.FullName,$false,$true)
  # 品管和工廠都優先使用含「明細」的工作表，找不到就選列數最多的工作表
  $ws=$wb.Sheets(1)
  $found=$false
  foreach($s in $wb.Sheets){ if($s.Name -like "*明細*"){ $ws=$s; $found=$true; break } }
  if(-not $found){
    $maxR=0
    foreach($s in $wb.Sheets){ $r=$s.UsedRange.Rows.Count; if($r -gt $maxR){ $maxR=$r; $ws=$s } }
  }
  Write-Host ("    Using sheet: {0}" -f $ws.Name)
  $data=$ws.UsedRange.Value2; $nR=$data.GetLength(0); $rc=0
  for($r=2;$r-le $nR;$r++){
    $kv=if($isQA){$data[$r,$QA.STYLE]}else{$data[$r,$FA.STYLE]}
    if($null -eq $kv -or $kv.ToString().Trim() -eq ""){continue}
    # 跨年rawdata：只寫入目標年的列
    $dateRaw=if($isQA){$data[$r,$QA.DATE]}else{$data[$r,$FA.DATE]}
    $dateOA=DateVal $dateRaw
    if($dateOA -isnot [double]){continue}
    if([DateTime]::FromOADate($dateOA).Year -ne [int]$year){continue}
    if($null -eq $maxOADate -or $dateOA -gt $maxOADate){$maxOADate=$dateOA}
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
      $gareaRaw=CleanStr $data[$r,$QA.GAREA]; $row[11]=($gareaRaw -split '\s+')[0]
      $row[12]=[DBNull]::Value
      $row[13]=CleanFiber $data[$r,$QA.FIBER]
      $row[14]=CleanStr $data[$r,$QA.KNIT]
      $row[15]=CleanStr $data[$r,$QA.FABTYPE]
      $row[16]=CleanStr $data[$r,$QA.DYE]
      $width=NumD $data[$r,$QA.WIDTH]; $row[17]=$width
      $gsm=NumD $data[$r,$QA.GSM]; $row[18]=$gsm
      $insp=NumD $data[$r,$QA.INSPQTY]
      if($gsm -is [double] -and $width -is [double] -and $insp -is [double]){try{$row[19]=[Math]::Round([double]$gsm*([double]$width+2)*[double]$insp/43000,2)}catch{}}  # 碼重=克重*(幅寬+2)*抽驗數/43000
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
      $locRaw=$data[$r,$FA.LOC]; $row[7]=MapN (CleanStr $locRaw) $locMap
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
      for($d=$FA.D1;$d-le $FA.D5;$d++){
        $dv=CleanStr $data[$r,$d]; if($dv -eq ""){continue}
        $catIdx=41
        if($defNameToCode.ContainsKey($dv)){$catIdx=GetCatSrcIdx $defNameToCode[$dv]}
        $cur=$row[$catIdx]; if($cur -isnot [string] -or $cur -eq ""){$row[$catIdx]=$dv}else{$row[$catIdx]+=",$dv"}
      }
      # C36~C41 有值才追加（不覆蓋 D1~D5 迴圈填入的內容）
      foreach($pair in @(@(42,$FA.WAIG),@(43,$FA.SHOUF),@(44,$FA.WEIXIE),@(45,$FA.YANSE),@(46,$FA.GUIGE),@(47,$FA.QITA))){
        $v=CleanStr $data[$r,$pair[1]]; if($v -eq ""){continue}
        $idx=$pair[0]
        if($row[$idx] -isnot [string] -or $row[$idx] -eq ""){$row[$idx]=$v}else{$row[$idx]+=",$v"}
      }
    }
    # Step4 inline：修正抽驗數(col27)、A%/B%/C%(col31~33)、總瑕疵(col34)
    $aqV=if($aq-is[double]){[double]$aq}else{0.0}
    $bqV=if($bq-is[double]){[double]$bq}else{0.0}
    $cqV=if($cq-is[double]){[double]$cq}else{0.0}
    $sum4=$aqV+$bqV+$cqV
    if($sum4-gt 0){
      $row[26]=$sum4
      $row[30]=[Math]::Round($aqV/$sum4,4); $row[31]=[Math]::Round($bqV/$sum4,4); $row[32]=[Math]::Round($cqV/$sum4,4)
      if($avgPts-is[double]){$row[33]=[Math]::Round([double]$avgPts*$sum4/100)}
    }
    # Step5/6 inline：只在 Fail 列計算 BL~CW
    if($row[24]-eq "Fail"){
      $fail56++
      $aiV5=if($row[34]-is[double]){[double]$row[34]}else{0.0}; $aiOK5=$row[34]-is[double]
      # BL(63)
      $v41=$row[41]; if($v41-is[string]-and $v41.Trim()-ne ""){$row[63]=if($aiOK5-and $aiV5-gt 33){1.0}else{0.5}; $blC++}
      # BM~BR(64~69)
      for($i5=0;$i5-lt 6;$i5++){$v2=$row[42+$i5]; if($v2-is[string]-and $v2.Trim()-ne ""){$row[64+$i5]=1.0; $bmbrC++}}
      # BS~CW(70~100)
      $cwA6=0
      foreach($g6 in $grps6){
        $srcR6=$row[$g6.s+23]; if($srcR6-isnot[string]-or $srcR6.Trim()-eq ""){continue}
        $srcT6=$srcR6.Replace($charWo,$charWu); $gm6=$false
        for($c6d=$g6.d1;$c6d-le $g6.d2;$c6d++){
          $dn6=$defN6[$c6d]; if($dn6-eq ""){continue}
          if($srcT6.Contains($dn6)){$row[$c6d-1]=if($g6.w){if($aiOK5-and $aiV5-gt 33){1.0}else{0.5}}else{1.0}; $bs6++; $gm6=$true}
        }
        if(-not $gm6){$cwA6++; $k6=$srcR6.Trim(); if($k6-ne ""){if(-not $uncl6.ContainsKey($k6)){$uncl6[$k6]=0}; $uncl6[$k6]++}}
      }
      if($cwA6-gt 0){$row[100]=[double]$cwA6; $cw6++}
    }
    $allRows.Add([object[]]$row.Clone()); $rc++
  }
  $wb.Close($false)
  Write-Host ("    Added "+$rc+" rows")
}
$n=$allRows.Count; Write-Host ("Total: {0} rows | Fail={1}, BL={2}, BM~BR={3}, BS~CV={4}, CW={5}" -f $n,$fail56,$blC,$bmbrC,$bs6,$cw6)
if($uncl6.Count-gt 0){ Write-Host "--- Unclassified ---"; $uncl6.GetEnumerator()|Sort-Object Value -Descending|ForEach-Object{Write-Host ("  [{0}] x{1}" -f $_.Key,$_.Value)} }
Write-Host "Building 2D array..."
$arr=New-Object "object[,]" $n,101
for($r=0;$r-lt $n;$r++){for($c=0;$c-lt 101;$c++){$arr[$r,$c]=$allRows[$r][$c]}}
Write-Host "Writing to QF..."
$wbQF=$excel.Workbooks.Open($qfPath); $wsQF=$wbQF.Sheets(1)
$startRow = $lastDataRow + 1
Write-Host ("追加起始行: {0}" -f $startRow)
$startRow | Out-File "$basePath\temp_startrow.txt" -Encoding ASCII
$wsQF.Range($wsQF.Cells($startRow,1),$wsQF.Cells($startRow+$n-1,101)).Value2=$arr
# 字串欄分 5 個連續區塊批次寫入（76 次 COM 呼叫 → 5 次）
Write-Host "Writing string columns (block mode)..."
$strGroups=@(@(3,10),@(13,18),@(23,23),@(39,60),@(62,100))
foreach($g in $strGroups){
  $c1=$g[0]; $c2=$g[1]; $w=$c2-$c1+1
  $blk=New-Object "object[,]" $n,$w
  for($r=0;$r-lt $n;$r++){
    for($ci=0;$ci-lt $w;$ci++){
      $v=$allRows[$r][$c1+$ci-1]
      if($null -ne $v -and $v -ne [DBNull]::Value){$s=$v.ToString(); if($s-ne ""){$blk[$r,$ci]=$s}}
    }
  }
  $wsQF.Range($wsQF.Cells($startRow,$c1),$wsQF.Cells($startRow+$n-1,$c2)).Value=$blk
}
Write-Host "String columns done."
Write-Host "Applying formats..."
$eRow=$startRow+$n-1
# NumberFormat：合併相鄰欄為區塊（16次→7次 COM 呼叫）
$wsQF.Range($wsQF.Cells($startRow,26),$wsQF.Cells($eRow,30)).NumberFormat="#,##0"
$wsQF.Range($wsQF.Cells($startRow,34),$wsQF.Cells($eRow,34)).NumberFormat="#,##0"
$wsQF.Range($wsQF.Cells($startRow,35),$wsQF.Cells($eRow,36)).NumberFormat="0.0"
$wsQF.Range($wsQF.Cells($startRow,31),$wsQF.Cells($eRow,33)).NumberFormat="0.0%"
$wsQF.Range($wsQF.Cells($startRow,37),$wsQF.Cells($eRow,37)).NumberFormat="0.0%"
$wsQF.Range($wsQF.Cells($startRow,2),$wsQF.Cells($eRow,2)).NumberFormat="yyyy/mm/dd"
$wsQF.Range($wsQF.Cells($startRow,59),$wsQF.Cells($eRow,61)).NumberFormat="yyyy/mm/dd"
$wsQF.UsedRange.HorizontalAlignment=-4131
Write-Host ("DONE! "+$n+" rows.")

# ===== Step 7: Format（同一 session）=====
Write-Host "Step 7: Format..."
$maxRow7=$wsQF.UsedRange.Rows.Count; $maxCol7=$wsQF.UsedRange.Columns.Count
$numC7=@(2,18,19,20,21,26,27,28,29,30,31,32,33,34,35,36,37,59,60,61,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,100,101)
if($startRow -le 2){$wsQF.Rows(1).HorizontalAlignment=-4108}
$wsQF.Range($wsQF.Cells($startRow,1),$wsQF.Cells($maxRow7,$maxCol7)).HorizontalAlignment=-4131
# 數值欄右對齊：5個連續區塊（58次→5次 COM 呼叫）
foreach($rng7 in @(@(2,2),@(18,21),@(26,37),@(59,61),@(64,101))){
  $c1=$rng7[0]; $c2=[Math]::Min($rng7[1],$maxCol7)
  if($c1-le $maxCol7){$wsQF.Range($wsQF.Cells($startRow,$c1),$wsQF.Cells($maxRow7,$c2)).HorizontalAlignment=-4152}
}
Write-Host ("Step7 done: formatted {0} rows" -f ($maxRow7-$startRow+1))

$wbQF.Save(); $wbQF.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)|Out-Null

# 依 QF 日期範圍改名（yymmdd-yymmdd）：起始日期從現有檔名取，終止日期用處理時記錄的 maxOADate
$startDateStr = $null
if($existingQF -and $existingQF.Name -match "insprecord_QF_(\d{6})-\d{6}\.xlsx"){
  $startDateStr = $Matches[1]
} elseif($existingQF -and $existingQF.Name -match "insprecord_QF_(\d{6})\.xlsx"){
  $startDateStr = $Matches[1]
}
# 首次建檔（無現有 QF）：用 maxOADate 當起始（因為沒有之前的記錄）
if($null -eq $startDateStr -and $null -ne $maxOADate){
  $startDateStr = [DateTime]::FromOADate($maxOADate).ToString("yyMMdd")
}
if($null -ne $maxOADate -and $null -ne $startDateStr){
  $e=[DateTime]::FromOADate($maxOADate).ToString("yyMMdd")
  $newName="$basePath\insprecord_QF\insprecord_QF_${startDateStr}-${e}.xlsx"
  if($qfPath -ne $newName){ Rename-Item $qfPath $newName -Force }
  Write-Host ("檔名: insprecord_QF_${startDateStr}-${e}.xlsx")
}












