---
name: monthly-import
description: 驗布月報完整流程 — 下載 rawdata、轉換歸檔、導入年度 InspRecord_QF、歸檔來源檔。當使用者說「跑月報」、「月報」、「月報流程」、「月度導入」、「執行月報自動」，或提到 11_inspdata_monthly、run_monthly_full.ps1、月報自動.ps1、import_monthly.ps1 時使用。
---

# 驗布月報完整流程

在 Windows 端以 PowerShell 執行。分兩段：**下載轉換**（03 資料夾）→ **導入**（11 資料夾）。

月報的 QF 是**整年一個檔案逐週追加**，不是每月一個檔。

## 日期週期

月報下載腳本抓的是 **上上週四 ～ 上週三**，設計成在**週一/二/三**執行。

腳本會自己算日期。若該週期的兩個檔案都已存在，會直接跳過下載。

## 完整流程

### Phase 1：下載 + 轉換歸檔

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "E:\88. Claude\03_download rawdata\月報自動.ps1"
```

加 `-Fast` 開快速模式。

這支腳本會：
1. 算出 `startDate` / `endDate`（yymmdd）
2. **檢查是否已下載** — 若 `品管驗布報表_<sd>-<ed>.xls` 與 `工廠驗布報表_<sd>-<ed>.xls` 都已存在於
   `11_inspdata_monthly\insprawdata_monthly\`，直接 `exit 0` 跳過
3. 呼叫 `auto_download_http.py <startDate> <endDate>` 下載
4. 呼叫 `run_download_and_convert.ps1 -type monthly -startDate .. -endDate ..` 改名 + 搬到
   `11_inspdata_monthly\insprawdata_monthly\`

### Phase 2：導入 QF

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "E:\88. Claude\11_inspdata_monthly\run_monthly_full.ps1" -year 2026
```

**`-year` 是必填**，沒給會直接 `exit 1`。填西元年（如 `2026`），不是月份。

若不想讓腳本強制關閉 Excel，加 `-NoKill`：

```powershell
... -File "...\run_monthly_full.ps1" -year 2026 -NoKill
```

這支腳本會：
1. **Step 0** — （除非 `-NoKill`）強制關閉所有 Excel 程序，等 2 秒。
   然後找同年份的 `insprecord_QF_<yy>*.xlsx`，改名回 `insprecord_QF_current.xlsx` 供追加。
   找不到工作檔就報錯離開。
2. **Step 1~3** — 呼叫 `import_monthly.ps1 -year <year>` 導入（讀 `00_plan_insprecord_QF_monthly.xlsx`
   的 `NameList` 做客戶/廠商/地點/工廠/QC/瑕疵代碼映射）
3. 導入完 QF 會被改成 `insprecord_QF_yymmdd-yymmdd.xlsx` 格式
4. 清掉 `temp_startrow.txt` 暫存
5. **Step 8** — 把**結束年 = 目標年**的來源檔移到 `insprawdata_monthly\completed rawdata\`
   （跨年的 rawdata 會等最後一年跑完才搬走）

### Phase 3：驗證

確認 QF 檔有更新（腳本結尾會印出檔名）：

```
E:\88. Claude\11_inspdata_monthly\insprecord_QF\insprecord_QF_yymmdd-yymmdd.xlsx
```

檢查重點：
- 日期範圍涵蓋這次導入的週期
- 筆數有增加（是**追加**不是覆蓋）
- NameList 映射有生效（客戶/廠商/QC 欄是簡稱不是原始字串）
- 結論欄（Pass / Discuss / Fail）有算出來

## 單獨執行某一步

| 目的 | 指令 |
|------|------|
| 只下載 | `python "E:\88. Claude\03_download rawdata\auto_download_http.py" <yymmdd> <yymmdd>` |
| 只搬檔改名 | `run_download_and_convert.ps1 -type monthly -startDate <yymmdd> -endDate <yymmdd> [-Fast]` |
| 只導入（步驟 4-7） | `11_inspdata_monthly\import_monthly.ps1 -year <YYYY>` |
| 自動檢查後導入 | `11_inspdata_monthly\auto_check_and_import.ps1` |

拆開的單步腳本：`step4_fix_inspqty.ps1`、`step5_bl_br.ps1`、`step6_bs_cw.ps1`、`step7_format.ps1`。

## 結論判定邏輯

`import_monthly.ps1` 依百碼瑕疵點數平均值（avg）、B、C 值判定：

```
C > 0  或 avg > 33                → Fail
avg <= 10 且 B = 0 且 C = 0       → Pass
其餘                              → Discuss
```

## 注意事項

- **預設會強制關閉所有 Excel 程序**。執行前先存檔關掉手邊的 Excel，或用 `-NoKill`。
- QF 工作檔在導入期間叫 `insprecord_QF_current.xlsx`，導入完才改成日期命名。
  若上次導入中斷，可能會留下 `current` 檔，下次跑會直接沿用。
- 月報 QF 是**整年累加**的，不要手動刪檔重跑，會掉整年資料。
- `12_inspdata_weekly\calc_weekly_stats.ps1` 會讀月 QF，所以月報導入壞掉會連帶影響週報統計。
