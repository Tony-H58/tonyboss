---
name: download-troubleshooting
description: 驗布報表自動下載問題診斷與修復 — 自動檢測月報/週報下載失敗，並嘗試自動修復。由程序 1-3 自動維護。
---

# 驗布報表下載自動診斷與修復

三層自動維運系統：
- **程序 1 (監控)**: check_monthly_download.ps1 / check_weekly_download.ps1
- **程序 2 (修復)**: auto_fix_downloads.ps1
- **程序 3 (知識庫)**: update_knowledge_skill.ps1

## 系統健康狀況

📊 **最後更新**: (初始化中)

| 指標 | 值 |
|------|-----|
| 總修復嘗試 | 0 |
| 成功修復 | 0 |
| 失敗修復 | 0 |
| **總成功率** | **0%** |

### 按報表類型

| 報表類型 | 嘗試次數 | 成功次數 | 成功率 |
|---------|---------|--------|-------|
| 月報 | 0 | 0 | 0% |
| 週報 | 0 | 0 | 0% |

## 監控排程

- **月報檢查**: 每週一、二、三 13:00 (檢查上上週四~上週三的月報)
- **週報檢查**: 每週四 13:00 (檢查上週的週報)
- **待處理rawdata**: 每天 11:00 和 16:00 (檢查並自動處理待轉換檔案)
- **自動修復**: 檢測失敗時立即觸發
- **知識庫更新**: 定期分析修復日誌

## 常見問題與解決方案

### 問題 1: Python 下載失敗

**症狀**: `auto_download_http.py` 執行失敗

**可能原因**:
- 網路連線問題
- Python 環境未正確設定
- 目標伺服器無法連接

**解決方案**:
1. 檢查網路連線
2. 確認 Python 已安裝並在 PATH 中
3. 檢查 auto_download_http.py 的日誌輸出

### 問題 2: 檔案鎖定

**症狀**: 檔案移動失敗，"檔案正被其他程序使用"

**可能原因**:
- Excel 或其他程式開啟了該檔案
- 上次的轉換未完全結束

**解決方案**:
1. 關閉所有開啟 Excel 檔案的程式
2. 手動刪除暫存檔 (如有)
3. 重新執行修復

### 問題 3: 日期計算錯誤

**症狀**: 下載了錯誤日期範圍的報表

**可能原因**:
- 系統時間設定不正確
- 腳本中的日期邏輯錯誤

**解決方案**:
1. 確認系統時間正確
2. 檢查 check_*.ps1 中的日期計算邏輯
3. 如需修正，聯絡系統管理員

## 修復日誌位置

- **修復日誌**: `E:\88. Claude\02_reminder\fix_log.txt`
- **統計數據**: `E:\88. Claude\02_reminder\fix_stats.json`
- **待處理日誌**: `E:\88. Claude\02_reminder\process_pending_log.txt`
- **月報自動下載**: `E:\88. Claude\03_download rawdata\月報自動.ps1`
- **週報自動下載**: `E:\88. Claude\03_download rawdata\週報自動.ps1`

## 手動觸發修復

若需手動觸發修復，可執行:

```powershell
# 月報修復
powershell -NoProfile -ExecutionPolicy Bypass -File "E:\88. Claude\02_reminder\auto_fix_downloads.ps1" -Type "monthly"

# 週報修復
powershell -NoProfile -ExecutionPolicy Bypass -File "E:\88. Claude\02_reminder\auto_fix_downloads.ps1" -Type "weekly"

# 待處理rawdata檢查
powershell -NoProfile -ExecutionPolicy Bypass -File "E:\88. Claude\02_reminder\process_pending_rawdata.ps1"
```

## 系統架構

```
程序 1 (監控)
    └─ 檢測下載失敗
         ↓
程序 2 (修復)
    └─ 自動執行月報/週報自動下載
         ↓
程序 3 (知識庫)
    └─ 分析日誌 → 更新本 skill
         ↓
    沉澱經驗，持續優化
```

---

*由程序 3 自動維護，初始化於 2026-08-08*
