# 開機後驗證清單（Boot Verification Checklist）

**說明**: 在開機後（約 09:15），按照以下步驟驗證監控-診斷-修復系統是否正確設定。

**預計時間**: 30-45 分鐘  
**執行日期**: ________________  
**執行者**: ________________

---

## 📋 開機前準備

- [ ] 已開機並登入 Windows 
- [ ] 開啟 PowerShell（以系統管理員身份執行）
- [ ] 導航到 `E:\88. Claude\02_reminder` 目錄

```powershell
cd "E:\88. Claude\02_reminder"
```

---

## 第一步：系統審計（約 5-10 分鐘）

### Step 1.1: 列出所有現有排程

執行以下命令以查看所有 Claude 相關的排程任務：

```powershell
# 方法 1: 查看 \Claude 路徑下的所有任務
Get-ScheduledTask -TaskPath "\Claude*" | Select-Object TaskName, State, Triggers | Format-Table -AutoSize

# 方法 2: 詳細清單（含更多資訊）
Get-ScheduledTask -TaskPath "\Claude*" | Format-List TaskName, Description, State, Triggers
```

**記錄結果**:

實際存在的排程任務：
```
任務名稱                            | 狀態     | 路徑
------------------------------------|---------|------------------
                                   |         |
                                   |         |
                                   |         |
                                   |         |
```

**發現是否有漏掉的排程？**
- [ ] 否，所有已知排程都存在
- [ ] 是，發現額外排程：_________________________
- [ ] 是，發現遺漏排程：_________________________

---

### Step 1.2: 導出排程清單（用於紀錄）

```powershell
# 導出為 CSV 檔案，便於追蹤
$fileName = "scheduled_tasks_audit_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
Get-ScheduledTask -TaskPath "\Claude*" | Export-Csv $fileName -Encoding UTF8
Write-Host "✅ 已導出至: $fileName"
```

---

## 第二步：驗證 9 個預期排程（約 10 分鐘）

根據下表檢查每個預期的排程是否存在。填寫「實際狀態」欄位。

| # | 排程名稱 | 預期觸發時間 | 預期存在 | 實際存在 | 狀態 | 備註 |
|---|---------|-----------|---------|--------|------|------|
| 1 | CheckMonthlyDownload | 每週一、二、三 13:00 | ✅ | [ ] | | |
| 2 | CheckWeeklyDownload | 每週四 13:00 | ✅ | [ ] | | |
| 3 | SendWeeklyEmail | 每天開機時 (AtLogon) | ✅ | [ ] | | |
| 4 | ProcessPendingRawdata | 每天 11:00、16:00 | ✅ | [ ] | | |
| 5 | (其他已知排程 1) | ? | ❓ | [ ] | | |
| 6 | (其他已知排程 2) | ? | ❓ | [ ] | | |
| 7 | (其他已知排程 3) | ? | ❓ | [ ] | | |
| 8 | (其他已知排程 4) | ? | ❓ | [ ] | | |
| 9 | (其他已知排程 5) | ? | ❓ | [ ] | | |

**統計**:
- 預期 4 個排程中，實際存在：_____ / 4 ✅
- 發現的其他排程：_____ 個
- **遺漏排程總數**：_____ 個

---

## 第三步：執行排程設定腳本（如有遺漏，約 10 分鐘）

**如果發現遺漏排程，執行以下命令以建立它們**：

```powershell
# 使用 setup_schedules.ps1 建立或更新所有排程
# （此腳本會檢查並跳過已存在的排程，只建立遺漏的排程）

.\setup_schedules.ps1
```

**預期輸出**:
```
✅ 已建立排程：CheckMonthlyDownload
✅ 已建立排程：CheckWeeklyDownload
✅ 已建立排程：SendWeeklyEmail
✅ 已建立排程：ProcessPendingRawdata
```

**驗證結果**:
- [ ] 所有排程已成功建立
- [ ] 沒有錯誤訊息
- [ ] 新排程已立即顯示在排程清單中

---

## 第四步：功能測試（約 10-15 分鐘）

### Test 4.1: 測試程序 1（監控）

**測試月報檢查**:
```powershell
# 執行月報檢查（不會自動修復，僅檢查狀態）
.\check_monthly_download.ps1

# 預期輸出示例：
# [2026-08-09 09:30:15] [INFO] 檢查月報下載狀態...
# [2026-08-09 09:30:20] [INFO] 月報檔案已存在，檢查完成
```

**執行結果**:
- [ ] 成功執行 (exit code: 0)
- [ ] 檢測到下載失敗 (exit code: 1)
- [ ] 發生錯誤 (exit code: 2)

**測試週報檢查**:
```powershell
# 執行週報檢查
.\check_weekly_download.ps1
```

**執行結果**:
- [ ] 成功執行 (exit code: 0)
- [ ] 檢測到下載失敗 (exit code: 1)
- [ ] 發生錯誤 (exit code: 2)

---

### Test 4.2: 測試程序 2（自動修復）

**測試修復機制（分析模式，不實際修復）**:
```powershell
# 測試月報修復（可選，如果前面檢測到失敗）
.\auto_fix_downloads.ps1 -Type "monthly" -Fast

# 或測試週報修復
.\auto_fix_downloads.ps1 -Type "weekly" -Fast
```

**預期結果**:
- [ ] 修復成功 (exit code: 0)
- [ ] 修復失敗 (exit code: 1)
- [ ] 發生異常 (exit code: 2)

**注意**: 若修復失敗，檢查 `fix_log.txt` 以瞭解失敗原因。

---

### Test 4.3: 測試程序 3（知識庫）

**測試知識庫分析（分析模式，不寫入）**:
```powershell
# 分析修復日誌（不更新檔案）
.\update_knowledge_skill.ps1 -AnalyzeOnly
```

**預期輸出示例**:
```
統計完成：
  總嘗試: X | 成功: Y | 失敗: Z | 成功率: A%
  月報: M 次, 成功率 B%
  週報: W 次, 成功率 C%
```

**驗證結果**:
- [ ] 腳本執行成功
- [ ] 顯示正確的統計數據
- [ ] 無錯誤訊息

---

### Test 4.4: 檢查日誌檔案

```powershell
# 查看修復日誌（最後 20 行）
Get-Content "fix_log.txt" -Tail 20

# 查看統計數據
Get-Content "fix_stats.json" | ConvertFrom-Json | Format-Table
```

**驗證檢查清單**:
- [ ] `fix_log.txt` 存在且有內容
- [ ] `fix_stats.json` 存在且格式正確
- [ ] 日誌中沒有異常錯誤
- [ ] 統計數據結構完整

---

## 第五步：SKILL 知識庫驗證（約 5 分鐘）

```powershell
# 檢查 SKILL.md 是否存在
Test-Path "E:\88. Claude\.claude\skills\download-troubleshooting\SKILL.md"

# 查看 SKILL.md 內容（前 50 行）
Get-Content "E:\88. Claude\.claude\skills\download-troubleshooting\SKILL.md" -Head 50
```

**驗證檢查清單**:
- [ ] SKILL.md 檔案存在
- [ ] 包含系統健康狀況表格
- [ ] 包含常見問題與解決方案
- [ ] 包含監控排程資訊

---

## 第六步：最終驗證（約 5 分鐘）

### 排程服務狀態

```powershell
# 確認 Windows 排程服務已啟動
Get-Service Schedule | Select-Object Name, Status

# 預期輸出：
# Name     Status
# ----     ------
# Schedule Running
```

**驗證結果**:
- [ ] 排程服務正在運行
- [ ] 無異常狀態

### PowerShell 執行原則

```powershell
# 查看目前的執行原則
Get-ExecutionPolicy

# 預期輸出: RemoteSigned 或更寬鬆的原則
```

**驗證結果**:
- [ ] 執行原則允許本地 PowerShell 腳本執行
- [ ] 若為 Restricted，需要運行管理員 PowerShell 並執行：
  ```powershell
  Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```

---

## 第七步：完整系統測試（可選，約 10 分鐘）

**完整運行一次知識庫更新**（如果 fix_log.txt 中有足夠的資料）：

```powershell
# 執行完整的知識庫更新（會寫入 SKILL.md）
.\update_knowledge_skill.ps1
```

**驗證檢查清單**:
- [ ] 腳本執行成功
- [ ] 不見任何 ERROR 訊息
- [ ] SKILL.md 已更新（檢查「最後更新」時間戳）
- [ ] 統計數據在 SKILL.md 中正確顯示

---

## 驗證結果總結

**日期**: ________  
**執行者**: ________  
**執行時間**: ________ 分鐘

### 排程系統
- [ ] ✅ 所有 4 個預期排程已存在且啟用
- [ ] ⚠️  發現 _____ 個額外排程
- [ ] ❌ 遺漏 _____ 個排程（已通過 setup_schedules.ps1 建立）

### 功能測試
- [ ] ✅ 程序 1（監控）執行正常
- [ ] ✅ 程序 2（修復）執行正常
- [ ] ✅ 程序 3（知識庫）執行正常
- [ ] ❌ 發現問題（請描述）：_________________________

### 日誌與數據
- [ ] ✅ 所有日誌檔案已建立
- [ ] ✅ 統計數據結構完整
- [ ] ✅ SKILL.md 知識庫已初始化

### 整體狀態
- [ ] ✅ 系統已準備好投入運作
- [ ] ⚠️  需要微調（問題描述）：_________________________
- [ ] ❌ 無法運作（問題描述）：_________________________

---

## 後續行動

### 如果一切正常 ✅
1. 提交此驗證清單到 Git
2. 提交任何額外發現的排程資訊
3. 系統已準備好開始監控並自動修復

### 如果發現問題 ⚠️
1. 檢查 `fix_log.txt` 中的詳細錯誤訊息
2. 查閱 README.md 中的「常見問題排除」章節
3. 如有必要，手動執行個別腳本進行調試

### 如果系統無法正常運作 ❌
1. 收集所有錯誤訊息和日誌
2. 檢查 Windows 事件檢視器 (Application 和 System 日誌)
3. 驗證所有必要檔案和路徑是否正確
4. 尋求進一步協助

---

## 附件：快速參考命令

### 快速檢查排程
```powershell
Get-ScheduledTask -TaskPath "\Claude\02_reminder\" | Format-Table TaskName, State
```

### 快速查看日誌
```powershell
# 最後 20 行修復日誌
tail -f fix_log.txt
```

### 快速執行測試
```powershell
# 執行所有程序的快速測試
.\check_monthly_download.ps1
.\check_weekly_download.ps1
.\update_knowledge_skill.ps1 -AnalyzeOnly
```

---

**文件版本**: 1.0  
**建立日期**: 2026-08-09  
**適用系統**: Windows 10/11 + PowerShell 5.0+
