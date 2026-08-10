# 快速開始（Quick Start）

**使用者**: 開機後立即執行此清單  
**預計時間**: 30 分鐘  
**日期**: __________

---

## ⚡ 30 秒快速檢查

```powershell
# 在管理員 PowerShell 中執行：
cd "E:\88. Claude\02_reminder"

# 檢查排程是否存在
Get-ScheduledTask -TaskPath "\Claude\02_reminder\" | Select-Object TaskName, State

# 執行程序 3 分析（檢查一切是否正常）
.\update_knowledge_skill.ps1 -AnalyzeOnly
```

**預期結果**：
- 看到 4 個排程任務 (CheckMonthlyDownload, CheckWeeklyDownload, SendWeeklyEmail, ProcessPendingRawdata)
- 程序 3 顯示統計資訊

---

## 📋 如果排程不存在

執行建立排程的腳本：

```powershell
.\setup_schedules.ps1
```

**預期結果**：
```
✅ 已建立排程：CheckMonthlyDownload
✅ 已建立排程：CheckWeeklyDownload
✅ 已建立排程：SendWeeklyEmail
✅ 已建立排程：ProcessPendingRawdata
```

---

## ✅ 驗證系統正常

執行這三個命令測試三個程序：

```powershell
# 測試程序 1: 監控（月報）
.\check_monthly_download.ps1

# 測試程序 2: 修復（如果上面檢測到失敗）
# .\auto_fix_downloads.ps1 -Type "monthly" -Fast

# 測試程序 3: 知識庫更新
.\update_knowledge_skill.ps1 -AnalyzeOnly
```

---

## 📊 檢查系統日誌

```powershell
# 查看最新 20 行日誌
Get-Content "fix_log.txt" -Tail 20

# 查看統計數據
$stats = Get-Content "fix_stats.json" | ConvertFrom-Json
Write-Host "總修復: $($stats.total_attempts) | 成功: $($stats.successful_fixes) | 失敗: $($stats.failed_fixes)"
```

---

## ⚙️ 排程執行時間表

| 時間 | 任務 | 腳本 |
|------|------|------|
| 每週 Mon/Tue/Wed 13:00 | 檢查月報 | check_monthly_download.ps1 |
| 每週 Thu 13:00 | 檢查週報 | check_weekly_download.ps1 |
| 每天 11:00、16:00 | 檢查待轉換 rawdata | process_pending_rawdata.ps1 |
| 每天開機時 | 寄送週報郵件 | send_weekly_email.ps1 |

---

## 🆘 問題排查

| 問題 | 解決方案 |
|------|--------|
| 排程不存在 | 執行 `.\setup_schedules.ps1` |
| 排程不執行 | 1. 檢查 Get-Service Schedule<br>2. 查看 Windows 事件檢視器 |
| 修復失敗 | 檢查 `fix_log.txt` 中的錯誤訊息 |
| 知識庫無法更新 | 確認資料夾權限，或手動執行 `.\update_knowledge_skill.ps1` |

---

## 📞 需要幫助？

詳細步驟請見：
- **完整驗證**: `BOOT_VERIFICATION_CHECKLIST.md`
- **使用手冊**: `README.md`
- **常見問題**: README.md 中的「常見問題排除」章節
- **部署概況**: `DEPLOYMENT_SUMMARY.md`

---

**系統版本**: 1.0  
**最後更新**: 2026-08-09
