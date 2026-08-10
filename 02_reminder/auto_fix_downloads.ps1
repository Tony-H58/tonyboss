# auto_fix_downloads.ps1
# Program 2: 自動修復下載失敗
# 由 Program 1 (check_*.ps1) 呼叫當檢測到下載失敗時
#
# 使用方式:
#   .\auto_fix_downloads.ps1 -Type "monthly"
#   .\auto_fix_downloads.ps1 -Type "weekly"

param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("monthly", "weekly")]
    [string]$Type,
    [switch]$Fast
)

$ErrorActionPreference = "Stop"

$baseDir   = "E:\88. Claude\02_reminder"
$logFile   = "$baseDir\fix_log.txt"
$downloadDir = "E:\88. Claude\03_download rawdata"

# ============================================
# 日誌函數
# ============================================
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "SUCCESS", "ERROR", "WARN")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    Write-Host $logEntry -ForegroundColor $(
        switch ($Level) {
            "SUCCESS" { "Green" }
            "ERROR"   { "Red" }
            "WARN"    { "Yellow" }
            default   { "Cyan" }
        }
    )

    Add-Content -Path $logFile -Value $logEntry -ErrorAction SilentlyContinue
}

# ============================================
# 主邏輯
# ============================================
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host " Auto Fix Download (Program 2)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

Write-Log "開始修復 $Type 報表下載失敗"

try {
    # Step 1: 選擇對應的自動下載腳本
    if ($Type -eq "monthly") {
        $autoScript = "$downloadDir\月報自動.ps1"
        $reportName = "Monthly"
    } else {
        $autoScript = "$downloadDir\週報自動.ps1"
        $reportName = "Weekly"
    }

    if (-not (Test-Path $autoScript)) {
        Write-Log "錯誤：找不到自動下載腳本 $autoScript" "ERROR"
        exit 1
    }

    # Step 2: 執行自動下載腳本
    Write-Log "執行 $reportName 自動下載腳本..."
    $psArgs = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", $autoScript
    )
    if ($Fast) { $psArgs += "-Fast" }

    & powershell @psArgs
    $scriptExitCode = $LASTEXITCODE

    # Step 3: 檢查結果
    if ($scriptExitCode -eq 0) {
        Write-Log "$reportName 下載修復成功" "SUCCESS"
        Write-Host "`n✅ 修復完成！" -ForegroundColor Green
        exit 0
    } else {
        Write-Log "$reportName 下載修復失敗 (ExitCode: $scriptExitCode)" "ERROR"
        Write-Host "`n❌ 修復失敗，請檢查日誌" -ForegroundColor Red
        exit 1
    }
}
catch {
    $errorMsg = $_.Exception.Message
    Write-Log "例外錯誤：$errorMsg" "ERROR"
    Write-Host "`n❌ 發生例外錯誤：$errorMsg" -ForegroundColor Red
    exit 2
}
