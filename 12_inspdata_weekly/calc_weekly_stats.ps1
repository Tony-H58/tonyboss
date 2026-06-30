$exProc = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if($exProc){ $exProc | Stop-Process -Force; Start-Sleep -Seconds 2 }

$xl = New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$xl.Calculation = -4135   # xlCalculationManual
$xl.ScreenUpdating = $false

# 動態找最新週 QF 和月 QF 檔
$weeklyQF = Get-ChildItem "E:\88. Claude\12_inspdata_weekly\insprecord_QF" -Filter "insprecord_QF_??????-??????.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$monthlyQF = Get-ChildItem "E:\88. Claude\11_inspdata_monthly\insprecord_QF" -Filter "insprecord_QF_*.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if(-not $weeklyQF){ Write-Host "ERROR: 找不到週 QF 檔"; $xl.Quit(); exit 1 }
if(-not $monthlyQF){ Write-Host "ERROR: 找不到月 QF 檔"; $xl.Quit(); exit 1 }
Write-Host "週 QF: $($weeklyQF.Name)"
Write-Host "月 QF: $($monthlyQF.Name)"

# Load QF data
$wbW=$xl.Workbooks.Open($weeklyQF.FullName,$false,$true)
$dW=$wbW.Sheets(1).UsedRange.Value2; $wbW.Close($false)
$wbY=$xl.Workbooks.Open($monthlyQF.FullName,$false,$true)
$dY=$wbY.Sheets(1).UsedRange.Value2; $wbY.Close($false)
# 從已載入的 $dW 第1列取瑕疵標頭（C71~C100），不需重開檔
$defHdr=@(); for($c=71;$c-le 100;$c++){$defHdr+=$dW[1,$c]}

$qaStr=[string][char]0x54C1+[string][char]0x7BA1
$facStr=[string][char]0x5DE5+[string][char]0x5EE0

# 從週 QF 檔名解析日期範圍（格式 insprecord_QF_yyMMdd-yyMMdd.xlsx）
$fnMatch = $weeklyQF.BaseName -match 'insprecord_QF_(\d{6})-(\d{6})'
if(-not $fnMatch){ Write-Host "ERROR: 無法解析週 QF 檔名日期: $($weeklyQF.Name)"; $xl.Quit(); exit 1 }
$sdStr = $Matches[1]; $edStr = $Matches[2]
$wkMin = [DateTime]::ParseExact("20$sdStr","yyyyMMdd",$null).ToOADate()
$wkMax = [DateTime]::ParseExact("20$edStr","yyyyMMdd",$null).ToOADate()
$yrMin = [DateTime]::Parse("$([DateTime]::FromOADate($wkMin).Year)/01/01").ToOADate()
$dateHeader = "$([DateTime]::FromOADate($wkMin).ToString('yyyy/M/d'))~$([DateTime]::FromOADate($wkMax).ToString('yyyy/M/d'))"
Write-Host "日期範圍: $dateHeader"

# Short/full name maps
$shortMap=@{CAB="CAMBODIA";CHN="CHINA";IND="INDONESIA";TWN="TAIWAN";VIN="VIETNAM"}
$fullToShort=@{CAMBODIA="CAB";CHINA="CHN";INDONESIA="IND";TAIWAN="TWN";VIETNAM="VIN"}
$locOrder=@("CHINA","TAIWAN","INDONESIA","VIETNAM","CAMBODIA")

# Char strings for C63 matching
$agree1=[string][char]0x540C+[string][char]0x610F+[string][char]0x51FA+[string][char]0x8CA8
$agree2=[string][char]0x81EA+[string][char]0x6AA2+[string][char]0x5F8C+[string][char]0x51FA+[string][char]0x8CA8
$rej1=[string][char]0x4E0D+[string][char]0x540C+[string][char]0x610F+[string][char]0x51FA+[string][char]0x8CA8
$rej2=[string][char]0x8986+[string][char]0x9A57  # 覆驗（含覆驗中/覆驗完成等）
$rej3=[string][char]0x4E0D+[string][char]0x518D+[string][char]0x5B89+[string][char]0x6392+[string][char]0x51FA+[string][char]0x8CA8
$rejectTxt=[string][char]0x6253+[string][char]0x4E0B+[string][char]0xFF0C+[string][char]0x5B89+[string][char]0x6392+[string][char]0x8986+[string][char]0x9A57  # 打下，安排覆驗
$signingTxt=[string][char]0x7C3D+[string][char]0x6838+[string][char]0x4E2D  # 簽核中


function GetRows($data,$typeStr,$locCode,$vendorStr,$dMin,$dMax){
    $rows=[System.Collections.Generic.List[int]]::new()
    for($r=2;$r-le $data.GetLength(0);$r++){
        $t=$data[$r,1]; if($null -eq $t -or $t -ne $typeStr){continue}
        if($null -ne $locCode -and $data[$r,8] -ne $locCode){continue}
        if($null -ne $vendorStr -and $data[$r,7] -ne $vendorStr){continue}
        if($null -ne $dMin){$d=$data[$r,2]; if($null -eq $d -or $d -isnot [double] -or [double]$d -lt $dMin){continue}}
        if($null -ne $dMax){$d=$data[$r,2]; if($null -eq $d -or $d -isnot [double] -or [double]$d -gt $dMax){continue}}
        $rows.Add($r)
    }; return $rows
}

function Stats($data,$rows){
    $rcv=0.0;$insp=0.0;$cqty=0.0;$totDef=0.0
    foreach($r in $rows){
        $v=$data[$r,26]; if($v -is [double]){$rcv+=$v}
        $v=$data[$r,27]; if($v -is [double]){$insp+=$v}
        $v=$data[$r,30]; if($v -is [double]){$cqty+=$v}
        $v=$data[$r,34]; if($v -is [double]){$totDef+=$v}
    }
    return @{rcv=[Math]::Round($rcv);insp=[Math]::Round($insp);
             cPct=if($insp-gt 0){[Math]::Round($cqty/$insp,4)}else{$null}
             avgPts=if($insp-gt 0){[Math]::Round($totDef/$insp*100,2)}else{$null}}
}

function YtdStats($dA,$rA,$dB,$rB){
    $insp=0.0;$cqty=0.0;$totDef=0.0
    foreach($r in $rA){$v=$dA[$r,27];if($v -is [double]){$insp+=$v};$v=$dA[$r,30];if($v -is [double]){$cqty+=$v};$v=$dA[$r,34];if($v -is [double]){$totDef+=$v}}
    foreach($r in $rB){$v=$dB[$r,27];if($v -is [double]){$insp+=$v};$v=$dB[$r,30];if($v -is [double]){$cqty+=$v};$v=$dB[$r,34];if($v -is [double]){$totDef+=$v}}
    return @{cPct=if($insp-gt 0){[Math]::Round($cqty/$insp,4)}else{$null};avgPts=if($insp-gt 0){[Math]::Round($totDef/$insp*100,2)}else{$null}}
}

function TopByCPct($data,$rows,$gIdx){
    $g=@{}; foreach($r in $rows){$k=$data[$r,$gIdx]; if($null -eq $k -or "$k" -eq ""){continue}; $k="$k".Trim()
        if(-not $g[$k]){$g[$k]=@{i=0.0;c=0.0}}; $v=$data[$r,27];if($v -is [double]){$g[$k].i+=$v}; $v=$data[$r,30];if($v -is [double]){$g[$k].c+=$v}}
    $best=$null;$bestP=-1; foreach($k in $g.Keys){$p=if($g[$k].i-gt 0){$g[$k].c/$g[$k].i}else{0}; if($p-gt $bestP){$bestP=$p;$best=$k}}
    return $best,$bestP
}

function TopCase($data,$rows){
    # 找單筆最大C%(C30/C27)，同C%取最大百碼pts(C34/C27*100)
    # 回傳 hashtable: cust/fab/def/color/cPct/avgPts
    $bestR=-1; $bestCP=-1.0; $bestPts=-1.0
    foreach($r in $rows){
        $insp=$data[$r,27]; if($insp -isnot [double] -or $insp -le 0){continue}
        $cqty=$data[$r,30]; $def=$data[$r,34]
        $cp=if($cqty -is [double]){$cqty/$insp}else{0.0}
        $pts=if($def -is [double]){$def/$insp*100}else{0.0}
        if($cp -gt $bestCP -or ($cp -eq $bestCP -and $pts -gt $bestPts)){$bestCP=$cp;$bestPts=$pts;$bestR=$r}
    }
    if($bestR -lt 0){return @{cust=$null;fab=$null;def=$null;color=$null;cPct=$null;avgPts=$null}}
    $cust="$($data[$bestR,3])".Trim(); if($cust -eq ""){$cust=$null}
    $fab="$($data[$bestR,16])".Trim(); if($fab -eq ""){$fab=$null}
    $color="$($data[$bestR,24])".Trim(); if($color -eq "" -or $color -eq $null){$color=$null}
    $fabStr=if($null -ne $fab -and $bestCP -gt 0){"{0}: {1:P1}" -f $fab,$bestCP}elseif($null -ne $fab){$fab}else{$null}
    $defFlags=@(); for($i=0;$i-lt 30;$i++){$ci=71+$i;$v=$data[$bestR,$ci];if($null -ne $v -and $v -is [double] -and $v-gt 0){$defFlags+=$defHdr[$i]}}
    $defTxt=if($defFlags.Count-gt 0){$defFlags -join "、"}else{$null}
    $insp2=$data[$bestR,27]; $cqty2=$data[$bestR,30]; $def2=$data[$bestR,34]
    $rowCP=if($insp2 -is [double] -and $insp2-gt 0 -and $cqty2 -is [double]){[Math]::Round($cqty2/$insp2,4)}else{$null}
    $rowPts=if($insp2 -is [double] -and $insp2-gt 0 -and $def2 -is [double]){[Math]::Round($def2/$insp2*100,2)}else{$null}
    return @{cust=$cust;fab=$fabStr;def=$defTxt;color=$color;cPct=$rowCP;avgPts=$rowPts}
}

function C20s($data,$rows){
    # C20 = rows where C% (C30/C27) > 20%
    $n=[double]0;$a=[double]0;$rej=[double]0;$sig=[double]0;$agTxt="";$rejTxt=""
    foreach($r in $rows){
        $insp=$data[$r,27]; $cqty=$data[$r,30]
        if($insp -is [double] -and $insp-gt 0 -and $cqty -is [double]){
            if(($cqty/$insp) -gt 0.20){ $n++ }else{ continue }
        } else { continue }
        # For QC: check 簽核狀態 (C63) + C20狀態 (C50) for disposition
        $v=$data[$r,63]; $c20zt="$($data[$r,50])".Trim()
        $hasC63=($null -ne $v -and "$v".Trim() -ne "")
        if(-not $hasC63){
            if($c20zt.Contains($rej2)){$rej++}else{$sig++}
            continue
        }
        $s="$v".Trim()
        if($s.Contains($agree1)-or $s.Contains($agree2)){$a++;if($agTxt -eq ""){$agTxt=$s}}
        elseif($s.Contains($rej1)-or $s.Contains($rej2)-or $s.Contains($rej3)-or $c20zt.Contains($rej2)){$rej++;if($rejTxt -eq ""){$rejTxt=$s}}
        else{$sig++}
    }
    $prog=""
    if($a-gt 0){$prog=$agTxt} elseif($rej-gt 0){$prog=$rejectTxt} elseif($sig-gt 0){$prog=$signingTxt}
    return @{tot=[double]$rows.Count;n=$n;agree=$a;reject=$rej;signing=$sig;prog=$prog}
}

function GetVendors($data,$typeStr,$locCode,$dMin,$dMax){
    $v=@{}; for($r=2;$r-le $data.GetLength(0);$r++){
        $t=$data[$r,1]; if($null -eq $t -or $t -ne $typeStr){continue}
        if($null -ne $locCode -and $data[$r,8] -ne $locCode){continue}
        if($null -ne $dMin){$d=$data[$r,2];if($null -eq $d -or $d -isnot [double] -or [double]$d -lt $dMin){continue}}
        if($null -ne $dMax){$d=$data[$r,2];if($null -eq $d -or $d -isnot [double] -or [double]$d -gt $dMax){continue}}
        $vn=$data[$r,7]; if($null -eq $vn -or "$vn" -eq ""){continue}; $vn="$vn".Trim()
        if(-not $v[$vn]){$v[$vn]=1}
    }; return @($v.Keys)
}

function WriteRow($ws,$row,$locLabel,$vendorName,$st,$ytd,$c20,$cust,$custP,$fab,$fabP,$defTxt,$prog,$isSummary){
    if($null -ne $locLabel){$ws.Cells($row,2).Value=$locLabel}
    if($null -ne $vendorName){$ws.Cells($row,3).Value=$vendorName}
    if($st.rcv-gt 0){$ws.Cells($row,4).Value2=$st.rcv}
    if($st.insp-gt 0){$ws.Cells($row,5).Value2=$st.insp}
    if($null -ne $st.cPct){$ws.Cells($row,6).Value2=$st.cPct}
    if($null -ne $st.avgPts){$ws.Cells($row,7).Value2=$st.avgPts}
    if($null -ne $ytd.cPct){$ws.Cells($row,8).Value2=$ytd.cPct}
    if($null -ne $ytd.avgPts){$ws.Cells($row,9).Value2=$ytd.avgPts}
    return $row+1
}

# Excel color = B*65536 + G*256 + R (BGR format)
function ExcelRGB($r,$g,$b){ return [int]($b*65536 + $g*256 + $r) }
$totalColor  = ExcelRGB 189 215 238   # light blue
$agreeColor  = ExcelRGB 255 255   0   # yellow
$regionColors=@{
    CHINA     = ExcelRGB 255 255 153  # light yellow
    TAIWAN    = ExcelRGB 198 239 206  # light green
    INDONESIA = ExcelRGB 252 228 214  # light peach
    VIETNAM   = ExcelRGB 255 199 206  # light pink
    CAMBODIA  = ExcelRGB 217 217 217  # light gray
}

$wbC=$xl.Workbooks.Open("E:\88. Claude\12_inspdata_weekly\analysis chart_weekly\analysis chart_weekly.xlsx")

foreach($sheetDef in @(
    @{name="QC";type=$qaStr;qcMode=$true;custCol=15;fabCol=16;defCol=17;progCol=18;c20Col=10},
    @{name="FTY";type=$facStr;qcMode=$false;custCol=12;fabCol=13;defCol=14;progCol=15;c20Col=10})){
    $ws=$wbC.Sheets($sheetDef.name)
    $typeStr=$sheetDef.type
    # Update date header
    $ws.Cells(2,4).Value=$dateHeader
    # Clear existing data rows: unmerge → ClearFormats → re-apply font
    $lastUsed=$ws.UsedRange.Rows.Count
    if($lastUsed-ge 4){
        $clearRng=$ws.Range($ws.Cells(4,1),$ws.Cells($lastUsed+50,20))
        $clearRng.UnMerge()
        $clearRng.ClearFormats()   # fully wipe old colors/borders/font
        $clearRng.ClearContents()
        $clearRng.Font.Name=[string][char]0x5FAE+[string][char]0x8EDF+[string][char]0x6B63+[string][char]0x9ED1+[string][char]0x9AD4
        $clearRng.Font.Size=12
        $clearRng.Font.Bold=$false
    }

    Write-Host ("=== {0} ===" -f $sheetDef.name)
    $currentRow=4
    $summaryRows=[System.Collections.Generic.List[int]]::new()

    # ---- TOTAL ROW ----
    $wkAllW=GetRows $dW $typeStr $null $null $wkMin $wkMax
    $st=Stats $dW $wkAllW
    $ytdW=GetRows $dW $typeStr $null $null $yrMin $null
    $ytdY=GetRows $dY $typeStr $null $null $yrMin $null
    $ytd=YtdStats $dW $ytdW $dY $ytdY
    $c20=C20s $dW $wkAllW
    $ws.Cells($currentRow,2).Value=[string][char]0x7E3D+[string][char]0x8A08  # 總計
    if($st.rcv-gt 0){$ws.Cells($currentRow,4).Value2=$st.rcv}
    if($st.insp-gt 0){$ws.Cells($currentRow,5).Value2=$st.insp}
    if($null -ne $st.cPct -and $st.cPct -ne 0){$ws.Cells($currentRow,6).Value2=$st.cPct}
    if($null -ne $st.avgPts -and $st.avgPts -ne 0){$ws.Cells($currentRow,7).Value2=$st.avgPts}
    if($null -ne $ytd.cPct -and $ytd.cPct -ne 0){$ws.Cells($currentRow,8).Value2=$ytd.cPct}
    if($null -ne $ytd.avgPts -and $ytd.avgPts -ne 0){$ws.Cells($currentRow,9).Value2=$ytd.avgPts}
    if($c20.tot-gt 0){$ws.Cells($currentRow,$sheetDef.c20Col).Value2=$c20.tot}
    if($c20.n-gt 0){$ws.Cells($currentRow,$sheetDef.c20Col+1).Value2=$c20.n}
    if($sheetDef.qcMode){
        if($c20.agree-gt 0){$ws.Cells($currentRow,$sheetDef.c20Col+2).Value2=$c20.agree}
        if($c20.reject-gt 0){$ws.Cells($currentRow,$sheetDef.c20Col+3).Value2=$c20.reject}
        if($c20.signing-gt 0){$ws.Cells($currentRow,$sheetDef.c20Col+4).Value2=$c20.signing}
    }
    # Total row: NO case data
    $ws.Range($ws.Cells($currentRow,2),$ws.Cells($currentRow,$sheetDef.progCol)).Interior.Color=$totalColor
    $summaryRows.Add($currentRow)
    Write-Host ("  TOTAL: rcv={0} insp={1} C%={2:P1} pts={3}" -f $st.rcv,$st.insp,$st.cPct,$st.avgPts)
    $currentRow++

    # ---- PER LOCATION ----
    foreach($locFull in $locOrder){
        $locShort=$fullToShort[$locFull]
        $wkLocW=GetRows $dW $typeStr $locShort $null $wkMin $wkMax
        # 本週無資料則跳過整個地區
        if($wkLocW.Count -eq 0){continue}
        $stLoc=Stats $dW $wkLocW
        $ytdLocW=GetRows $dW $typeStr $locShort $null $yrMin $null
        $ytdLocY=GetRows $dY $typeStr $locShort $null $yrMin $null
        $ytdLoc=YtdStats $dW $ytdLocW $dY $ytdLocY

        # Region summary row (use short name)
        $ws.Cells($currentRow,2).Value=$locShort
        if($stLoc.rcv-gt 0){$ws.Cells($currentRow,4).Value2=$stLoc.rcv}
        if($stLoc.insp-gt 0){$ws.Cells($currentRow,5).Value2=$stLoc.insp}
        if($null -ne $stLoc.cPct -and $stLoc.cPct -ne 0){$ws.Cells($currentRow,6).Value2=$stLoc.cPct}
        if($null -ne $stLoc.avgPts -and $stLoc.avgPts -ne 0){$ws.Cells($currentRow,7).Value2=$stLoc.avgPts}
        if($null -ne $ytdLoc.cPct -and $ytdLoc.cPct -ne 0){$ws.Cells($currentRow,8).Value2=$ytdLoc.cPct}
        if($null -ne $ytdLoc.avgPts -and $ytdLoc.avgPts -ne 0){$ws.Cells($currentRow,9).Value2=$ytdLoc.avgPts}
        if($wkLocW.Count-gt 0){
            $c20Loc=C20s $dW $wkLocW
            if($c20Loc.tot-gt 0){$ws.Cells($currentRow,$sheetDef.c20Col).Value2=$c20Loc.tot}
            if($c20Loc.n-gt 0){$ws.Cells($currentRow,$sheetDef.c20Col+1).Value2=$c20Loc.n}
            if($sheetDef.qcMode){
                if($c20Loc.agree-gt 0){$ws.Cells($currentRow,$sheetDef.c20Col+2).Value2=$c20Loc.agree}
                if($c20Loc.reject-gt 0){$ws.Cells($currentRow,$sheetDef.c20Col+3).Value2=$c20Loc.reject}
                if($c20Loc.signing-gt 0){$ws.Cells($currentRow,$sheetDef.c20Col+4).Value2=$c20Loc.signing}
            }
            # Region summary rows: NO case data
        }
        $locColor=if($regionColors.ContainsKey($locFull)){$regionColors[$locFull]}else{0xD9E1F2}
        $ws.Range($ws.Cells($currentRow,2),$ws.Cells($currentRow,$sheetDef.progCol)).Interior.Color=$locColor
        $summaryRows.Add($currentRow)
        $currentRow++

        # Vendor rows - get all vendors in this location this week
        $vendors=GetVendors $dW $typeStr $locShort $wkMin $wkMax
        # Calculate stats per vendor, then sort by C% desc
        $vendorStats=[System.Collections.Generic.List[object]]::new()
        foreach($vn in $vendors){
            $wkVW=GetRows $dW $typeStr $locShort $vn $wkMin $wkMax
            $sv=Stats $dW $wkVW
            $ytdVW=GetRows $dW $typeStr $locShort $vn $yrMin $null
            $ytdVY=GetRows $dY $typeStr $locShort $vn $yrMin $null
            $ytdV=YtdStats $dW $ytdVW $dY $ytdVY
            $c20V=C20s $dW $wkVW
            $tc=TopCase $dW $wkVW
            $vendorStats.Add(@{
                vn=$vn; st=$sv; ytd=$ytdV; c20=$c20V; tc=$tc
                cPctVal=if($null -ne $sv.cPct){$sv.cPct}else{-1.0}
            })
        }
        $sorted=$vendorStats | Sort-Object {$_.cPctVal} -Descending
        foreach($vd in $sorted){
            # Vendor row: fill entire row with region color
            $ws.Range($ws.Cells($currentRow,2),$ws.Cells($currentRow,$sheetDef.progCol)).Interior.Color=$locColor
            $ws.Cells($currentRow,2).Value=$locShort
            $ws.Cells($currentRow,3).Value=$vd.vn
            if($vd.st.rcv-gt 0){$ws.Cells($currentRow,4).Value2=$vd.st.rcv}
            if($vd.st.insp-gt 0){$ws.Cells($currentRow,5).Value2=$vd.st.insp}
            if($null -ne $vd.st.cPct -and $vd.st.cPct -ne 0){$ws.Cells($currentRow,6).Value2=$vd.st.cPct}
            if($null -ne $vd.st.avgPts -and $vd.st.avgPts -ne 0){$ws.Cells($currentRow,7).Value2=$vd.st.avgPts}
            if($null -ne $vd.ytd.cPct -and $vd.ytd.cPct -ne 0){$ws.Cells($currentRow,8).Value2=$vd.ytd.cPct}
            if($null -ne $vd.ytd.avgPts -and $vd.ytd.avgPts -ne 0){$ws.Cells($currentRow,9).Value2=$vd.ytd.avgPts}
            if($vd.c20.tot-gt 0){$ws.Cells($currentRow,$sheetDef.c20Col).Value2=$vd.c20.tot}
            if($vd.c20.n-gt 0){$ws.Cells($currentRow,$sheetDef.c20Col+1).Value2=$vd.c20.n}
            if($vd.c20.n-gt 0){
                if($sheetDef.qcMode){
                    if($vd.c20.agree-gt 0){$ws.Cells($currentRow,$sheetDef.c20Col+2).Value2=$vd.c20.agree}
                    if($vd.c20.reject-gt 0){$ws.Cells($currentRow,$sheetDef.c20Col+3).Value2=$vd.c20.reject}
                    if($vd.c20.signing-gt 0){$ws.Cells($currentRow,$sheetDef.c20Col+4).Value2=$vd.c20.signing}
                    if($vd.c20.prog -ne ""){$ws.Cells($currentRow,$sheetDef.progCol).Value=$vd.c20.prog}
                }
            }
            # 主要case欄位只在有C20時填值（取TopCase單筆數據）
            if($vd.c20.n-gt 0){
                if($null -ne $vd.tc.cust){$ws.Cells($currentRow,$sheetDef.custCol).Value=$vd.tc.cust}
                if($null -ne $vd.tc.fab){$ws.Cells($currentRow,$sheetDef.fabCol).Value=$vd.tc.fab}
                if($null -ne $vd.tc.def){$ws.Cells($currentRow,$sheetDef.defCol).Value=$vd.tc.def}
            }
            # FTY 處理進度追蹤：只有C20>0才填入 "顏色: xxx，C級%: xx.x%，平均扣點: xx.x"（取TopCase單筆數據）
            if(-not $sheetDef.qcMode -and $vd.c20.n-gt 0){
                $progParts=@()
                if($null -ne $vd.tc.color){$progParts+=("{0}: {1}" -f ([string][char]0x984F+[string][char]0x8272),$vd.tc.color)}
                if($null -ne $vd.tc.cPct -and $vd.tc.cPct -gt 0){$progParts+=("C{0}%: {1:P1}" -f [string][char]0x7D1A,$vd.tc.cPct)}
                if($null -ne $vd.tc.avgPts -and $vd.tc.avgPts -gt 0){$progParts+=("{0}: {1:F1}" -f ([string][char]0x5E73+[string][char]0x5747+[string][char]0x6263+[string][char]0x9EDE),$vd.tc.avgPts)}
                if($progParts.Count-gt 0){$ws.Cells($currentRow,$sheetDef.progCol).Value=$progParts -join "，"}
            }
            $currentRow++
        }
        Write-Host ("  {0}: {1} vendors, loc rows written" -f $locFull,$vendors.Count)
    }

    # Apply number formats for all data rows
    $endRow=$currentRow-1
    if($endRow-ge 4){
        $ws.Range($ws.Cells(4,4),$ws.Cells($endRow,5)).NumberFormat="#,##0"
        $ws.Range($ws.Cells(4,6),$ws.Cells($endRow,6)).NumberFormat="0.0%"
        $ws.Range($ws.Cells(4,7),$ws.Cells($endRow,7)).NumberFormat="0.0"
        $ws.Range($ws.Cells(4,8),$ws.Cells($endRow,8)).NumberFormat="0.0%"
        $ws.Range($ws.Cells(4,9),$ws.Cells($endRow,9)).NumberFormat="0.0"
        $c20End=if($sheetDef.qcMode){$sheetDef.c20Col+4}else{$sheetDef.c20Col+1}
        $lastCol=$sheetDef.progCol
        $ws.Range($ws.Cells(4,$sheetDef.c20Col),$ws.Cells($endRow,$c20End)).NumberFormat="#,##0"
        # Alignment: text left, numbers right
        $ws.Range($ws.Cells(4,2),$ws.Cells($endRow,3)).HorizontalAlignment=-4131
        $ws.Range($ws.Cells(4,4),$ws.Cells($endRow,9)).HorizontalAlignment=-4152
        $ws.Range($ws.Cells(4,$sheetDef.c20Col),$ws.Cells($endRow,$c20End)).HorizontalAlignment=-4152
        $ws.Range($ws.Cells(4,$sheetDef.custCol),$ws.Cells($endRow,$lastCol)).HorizontalAlignment=-4131

        # Borders: thin all cells only
        $dataRange=$ws.Range($ws.Cells(4,2),$ws.Cells($endRow,$lastCol))
        $dataRange.Borders.LineStyle=1; $dataRange.Borders.Weight=2  # xlContinuous, xlThin

        # Summary rows: bold font
        foreach($sumRow in $summaryRows){
            if($sumRow-ge 4 -and $sumRow-le $endRow){
                $ws.Range($ws.Cells($sumRow,2),$ws.Cells($sumRow,$lastCol)).Font.Bold=$true
            }
        }


        # Column widths (only set once, on QC sheet)
        if($sheetDef.name -eq "QC"){
            $ws.Columns(2).ColumnWidth=12   # 地區
            $ws.Columns(3).ColumnWidth=20   # 供應商
            $ws.Columns(4).ColumnWidth=11   # 收料
            $ws.Columns(5).ColumnWidth=10   # 抽驗
            $ws.Columns(6).ColumnWidth=8    # C%
            $ws.Columns(7).ColumnWidth=8    # 扣點
            $ws.Columns(8).ColumnWidth=8    # 2026 C%
            $ws.Columns(9).ColumnWidth=8    # 2026 扣點
            $ws.Columns(10).ColumnWidth=8   # 驗布筆數
            $ws.Columns(11).ColumnWidth=8   # C20筆數
            $ws.Columns(12).ColumnWidth=8   # 同意
            $ws.Columns(13).ColumnWidth=8   # 打下
            $ws.Columns(14).ColumnWidth=8   # 簽核
            $ws.Columns(15).ColumnWidth=10  # 客戶
            $ws.Columns(16).ColumnWidth=18  # 布種C%
            $ws.Columns(17).ColumnWidth=22  # 主要瑕疵
            $ws.Columns(18).ColumnWidth=35  # 處理進度追蹤
        }
        if($sheetDef.name -eq "FTY"){
            $ws.Columns(2).ColumnWidth=12
            $ws.Columns(3).ColumnWidth=20
            $ws.Columns(4).ColumnWidth=11
            $ws.Columns(5).ColumnWidth=10
            $ws.Columns(6).ColumnWidth=8
            $ws.Columns(7).ColumnWidth=8
            $ws.Columns(8).ColumnWidth=8
            $ws.Columns(9).ColumnWidth=8
            $ws.Columns(10).ColumnWidth=8
            $ws.Columns(11).ColumnWidth=8
            $ws.Columns(12).ColumnWidth=10
            $ws.Columns(13).ColumnWidth=18
            $ws.Columns(14).ColumnWidth=22
            $ws.Columns(15).ColumnWidth=35
        }
        # Row height
        $ws.Range($ws.Cells(4,2),$ws.Cells($endRow,2)).RowHeight=15
    }
    Write-Host ("  {0} total data rows written" -f ($endRow-3))
}

$sDate=[DateTime]::FromOADate($wkMin).ToString("yyMMdd")
$eDate=[DateTime]::FromOADate($wkMax).ToString("yyMMdd")
$chartSave="E:\88. Claude\12_inspdata_weekly\analysis chart_weekly\analysis chart_weekly_$sDate-$eDate.xlsx"
$xlFmt=51  # xlOpenXMLWorkbook (.xlsx)
$wbC.SaveAs($chartSave,$xlFmt); $wbC.Close($false)
$xl.Calculation = -4106   # xlCalculationAutomatic
$xl.ScreenUpdating = $true
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null
Write-Host "DONE."
exit 0
