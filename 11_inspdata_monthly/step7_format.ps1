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
$maxCol = $ws.UsedRange.Columns.Count

$numCols = @(2,18,19,20,21,26,27,28,29,30,31,32,33,34,35,36,37,59,60,61,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,100,101)

# Row 1: center（只在 startRow=2 時執行，避免重複）
if($startRow -le 2){ $ws.Rows(1).HorizontalAlignment = -4108 }

# 新增行：文字靠左
$ws.Range($ws.Cells($startRow,1),$ws.Cells($maxRow,$maxCol)).HorizontalAlignment = -4131

# 新增行數字欄：靠右
foreach($c in $numCols){
  if($c -le $maxCol){
    $ws.Range($ws.Cells($startRow,$c),$ws.Cells($maxRow,$c)).HorizontalAlignment = -4152
  }
}

Write-Host ("Step7 done: formatted {0} rows" -f ($maxRow-$startRow+1))
$wb.Save(); $wb.Close($false); $xl.Quit()
