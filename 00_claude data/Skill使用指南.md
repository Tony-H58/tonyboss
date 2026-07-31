# Claude Skills 使用指南

**最後更新**: 2026-07-31  
**維護者**: Tony Huang

> 本指南只列**實際存在**的 skill 與**實際存在**的腳本。
> 自訂 skill 定義在 `E:\88. Claude\.claude\skills\<名稱>\SKILL.md`。

---

## ⚠️ 全域規則：儲存行為

執行任何 skill 或工作時，**預設只寫入/更新主檔**（工作的主要輸出檔案，如記錄檔、QF 檔、報告檔）。

**副檔、備份檔、進度筆記、其他文件筆記本**（如 `發票整理筆記.md`、`11_inspdata_monthly 筆記.md`、
`tony筆記本` 等）**不主動更新**——只有使用者在對話中明確說「**儲存**」時，才連同這些檔案一起寫入/更新。

適用於所有自訂 skill（weekly-report、monthly-import、qc-analysis、invoice-transfer）與一般操作。

**範例（qc-analysis）**：`比對.py`／`qc_kpi_final.py` 預設只寫主檔 HTML
（`compare_result.html`、`qc_kpi_report.html`），複製到 `00_html file\` 的副本
是主檔 HTML 的備份，只有加 `--save` 參數（使用者說「儲存」時才加）才會複製。

---

## 📋 目錄
1. [自訂 Skill 一覽](#自訂-skill-一覽)
2. [系統工作流程對應](#系統工作流程對應)
3. [各 Skill 詳細說明](#各-skill-詳細說明)
4. [常見工作場景](#常見工作場景)
5. [快速參考](#快速參考)

---

## 自訂 Skill 一覽

| Skill | 觸發詞 | 涵蓋範圍 |
|-------|--------|---------|
| **weekly-report** | 跑週報、週報流程、週度導入、寄週報 | 03 下載轉換 → 12 導入統計 → 02 寄信 |
| **monthly-import** | 跑月報、月報流程、月度導入 | 03 下載轉換 → 11 導入年度 QF |
| **qc-analysis** | 比對驗布、前後端比對、QC 績效、跑 KPI | 13 比對分析 + 14 績效評估 |
| **invoice-transfer** | 整理發票、辨識發票、轉入發票、執行 | 31 平台發票圖片辨識轉入 |

定義檔位置：`E:\88. Claude\.claude\skills\`

---

## 系統工作流程對應

### 🔄 驗布報表下載轉換（03_download rawdata）

| 工作 | 使用 Skill | 實際指令 |
|------|-----------|---------|
| 週報下載轉換 | **weekly-report** | `週報自動.ps1 [-Fast]` |
| 月報下載轉換 | **monthly-import** | `月報自動.ps1 [-Fast]` |
| 只下載 | **run** | `python auto_download_http.py <yymmdd> <yymmdd>` |
| 只轉換 | **run** | `convert_rawdata.ps1 -startDate <yymmdd> -endDate <yymmdd> [-Fast]` |
| 搬檔改名 | **run** | `run_download_and_convert.ps1 -type weekly\|monthly -startDate .. -endDate .. [-Fast]` |

---

### 📁 驗布資料導入系統

#### 月度導入（11_inspdata_monthly）

| 工作 | 使用 Skill | 實際指令 |
|------|-----------|---------|
| 完整月度流程 | **monthly-import** | `run_monthly_full.ps1 -year 2026` |
| 不強制關 Excel | **monthly-import** | `run_monthly_full.ps1 -year 2026 -NoKill` |
| 只跑導入 | **run** | `import_monthly.ps1 -year 2026` |
| 自動檢查後導入 | **run** | `auto_check_and_import.ps1` |

> ⚠️ **`-year` 是必填且是西元年**（如 `2026`），不是 `-month "YYMM"`。
> `run_monthly_full.ps1` **沒有** `-Fast` 參數，只有 `-NoKill`。

拆開的單步腳本：`step4_fix_inspqty.ps1`、`step5_bl_br.ps1`、`step6_bs_cw.ps1`、`step7_format.ps1`

#### 週度導入（12_inspdata_weekly）

| 工作 | 使用 Skill | 實際指令 |
|------|-----------|---------|
| 完整週度流程 | **weekly-report** | `run_weekly_full.ps1 [-Fast]` |
| 只跑導入 | **run** | `run_weekly_import.ps1 -Fast` |
| 只算統計 | **run** | `calc_weekly_stats.ps1` |

> ⚠️ `calc_weekly_stats.ps1` 會**強制關閉所有 Excel 程序**，執行前先存檔。
> 統計需要**週 QF 和月 QF 都存在**，缺一報錯。

---

### 📊 分析與績效

#### 前後端驗布比對（13_compare style inspquality）

| 工作 | 使用 Skill | 實際指令 |
|------|-----------|---------|
| 指定款號比對 | **qc-analysis** | `python 比對.py "<款號或採購單號>" "<月份>"` |
| 互動模式 | **qc-analysis** | `python 比對.py`（會逐項詢問） |

產出：`00_analysis data.xlsx`、`compare_result.html`

#### QC 績效評估（14_qc performance）

| 工作 | 使用 Skill | 實際指令 |
|------|-----------|---------|
| 計算 KPI | **qc-analysis** | `python qc_kpi_final.py "26-1~26-6"` |
| 互動模式 | **qc-analysis** | `python qc_kpi_final.py` |

產出：`qc_kpi_report.html`

---

### 📈 品管儀表板（15_dashboard）

目前是**靜態 HTML 檔**，沒有產生器腳本：

- `品管儀表板_A.html`
- `品管儀表板_B.html`
- `品管儀表板_總頁.html`

直接開檔檢視即可。儀表板會由 `send_weekly_email.ps1` 一併寄出。

---

### 📧 郵件與提醒系統（02_reminder）

| 工作 | 使用 Skill | 實際指令 |
|------|-----------|---------|
| 發送週報郵件 | **weekly-report** | `send_weekly_email.ps1` |
| 檢查週報下載 | **run** | `check_weekly_download.ps1` |
| 檢查月報下載 | **run** | `check_monthly_download.ps1` |
| 處理待處理資料 | **run** | `process_pending_rawdata.ps1` |
| 建立自動排程 | **run** | `setup_schedules.ps1`（需管理員權限） |

#### 已建立的 Windows 工作排程

| 排程名稱 | 時機 | 動作 |
|---------|------|------|
| CheckWeeklyDownload | 每週四 13:00 | 檢查週報 rawdata 是否已下載 |
| CheckMonthlyDownload | 每週一~三 13:00 | 檢查月報 rawdata 是否已下載 |
| SendWeeklyEmail | 每次登入 | 寄送週報郵件（每週只寄一次） |

---

### 🧵 布料檢驗系統（24_fabric inspsystem）

| 工作 | 使用 Skill | 實際指令 |
|------|-----------|---------|
| 啟動驗布系統 | **run** | `cd "24_fabric inspsystem\webapp"` 後 `python main.py` |
| 查看 Web 介面 | — | 瀏覽器開 `http://localhost:8000` |

> ⚠️ 後端是 **FastAPI + uvicorn**（不是 Flask），埠號 `8000`，開啟 reload 模式。
> 靜態頁面在 `webapp\static\fabric inspsystem.html`。

---

### 💰 發票整理（31_electronic invoice）

| 工作 | 使用 Skill | 說明 |
|------|-----------|------|
| 平台發票轉入 | **invoice-transfer** | 說「執行」即一次性完成所有工作表 |

詳細規則見 `00_claude data\發票整理筆記.md` 與 skill 定義檔。

---

## 各 Skill 詳細說明

### 🎯 自訂 Skill

#### **weekly-report** — 週報完整流程
- **週期**：上週四 ～ 本週三（腳本自動計算，週四當天取剛結束的週期）
- **兩段式**：Phase 1 下載轉換（03）→ Phase 2 導入統計（12）
- **產出**：`insprecord_QF_yymmdd-yymmdd.xlsx`、`analysis chart_weekly_*.xlsx`

#### **monthly-import** — 月報完整流程
- **週期**：上上週四 ～ 上週三，設計成週一/二/三執行
- **QF 是整年累加**，不要刪檔重跑
- **必填 `-year <西元年>`**

#### **qc-analysis** — 品質分析與績效
- 資料來源是月度 QF，**跑之前先確認月報已導入最新資料**
- 兩支腳本都唯讀 QF，不會改動資料
- 輸出 HTML 每次執行會覆蓋

#### **invoice-transfer** — 平台發票辨識轉入
- 發票圖片內嵌在 xlsx 工作表中（xlsx 即 zip，圖片在 `xl/media/`）
- **必須二次讀圖驗證**，有差異修正後再比對
- Uber 費用未稅用差額回推，確保合計正確

### 🛠 內建 Skill（Claude Code 官方）

| Skill | 用途 |
|-------|------|
| **run** | 啟動並執行腳本、服務、應用程式 |
| **loop** | 固定間隔重複執行某指令（如每 5 分鐘檢查） |
| **xlsx** | 建立/讀取/編輯 Excel、資料清理、公式圖表 |
| **pdf** | PDF 讀取、合併拆分、OCR、表單 |
| **docx** | Word 文件建立與編輯 |
| **pptx** | 簡報建立與編輯 |
| **dataviz** | 圖表與儀表板設計規範 |
| **skill-creator** | 建立/優化新的 skill |
| **update-config** | 設定 hook、權限、環境變數 |
| **review** | 審查 GitHub Pull Request |
| **security-review** | 對當前 branch 變更做安全審查 |
| **simplify** | 程式碼品質、重用性、效率優化 |
| **init** | 建立 CLAUDE.md 專案說明 |
| **claude-api** | Claude API / Anthropic SDK 開發參考 |
| **fewer-permission-prompts** | 掃描記錄，自動加入常用指令允許清單 |
| **keybindings-help** | 自訂鍵盤快捷鍵 |

---

## 常見工作場景

### 📅 週報流程
```
1. 說「跑週報」→ weekly-report skill
   ↓
2. Phase 1：週報自動.ps1
   - auto_download_http.py 下載
   - run_download_and_convert.ps1 轉換搬檔
   ↓
3. Phase 2：run_weekly_full.ps1 -Fast
   - run_weekly_import.ps1 導入 QF
   - calc_weekly_stats.ps1 算統計
   ↓
4. Phase 3：驗證 QF 檔與 Chart 檔的日期範圍、筆數
   ↓
5. 確認無誤 → send_weekly_email.ps1 寄出
```

### 📊 月報流程
```
1. 說「跑月報」→ monthly-import skill
   ↓
2. Phase 1：月報自動.ps1
   （若該週期兩個檔都已存在會自動跳過下載）
   ↓
3. Phase 2：run_monthly_full.ps1 -year 2026
   - 找同年 QF 改回 current.xlsx
   - import_monthly.ps1 追加導入
   - 改名回 yymmdd-yymmdd 格式
   - 搬走已處理的 rawdata
   ↓
4. Phase 3：驗證筆數有「增加」（追加不是覆蓋）
```

### 🔍 品質分析流程
```
1. 先確認月報已導入最新資料
   ↓
2. 說「比對驗布」→ qc-analysis skill
   python 比對.py "<款號>" "<月份>"
   ↓
3. 開 compare_result.html 檢查前後端筆數有對上
```

### 📈 績效評估流程
```
1. 說「跑 QC 績效」→ qc-analysis skill
   python qc_kpi_final.py "26-1~26-6"
   ↓
2. 開 qc_kpi_report.html
   ↓
3. 檢查 QC 人數合理、無 NaN 或除零異常
```

### 💰 發票整理流程
```
1. 說「整理發票」或「執行」→ invoice-transfer skill
   ↓
2. 解壓 xl/media/ 取圖 → 批量辨識 → 計算未稅
   ↓
3. 一次性寫入發票整理記錄（主檔）
   ↓
4. 二次讀圖驗證，有差異修正到正確
   ↓
5. 套格式 → 清暫存 → 全工作表完成才搬到 completed rawdata
   ↓
6. （使用者說「儲存」時才執行）更新 發票整理筆記.md 進度表
```

---

## 快速參考

### 🚀 一鍵快速命令

| 需求 | 說法 / 指令 | 資料夾 |
|------|------------|--------|
| 週報完整流程 | 「跑週報」 | 03 → 12 |
| 月報完整流程 | 「跑月報」 | 03 → 11 |
| 前後端比對 | 「比對驗布」 | 13_compare style inspquality |
| QC 績效計算 | 「跑 QC 績效」 | 14_qc performance |
| 發票轉入 | 「執行」 | 31_electronic invoice |
| 啟動驗布系統 | `python main.py` | 24_fabric inspsystem\webapp |

### 📝 使用頻率

**高頻**（每週）
- `weekly-report` — 週報流程
- `run` — 單獨執行腳本

**定期**（每月）
- `monthly-import` — 月報導入
- `invoice-transfer` — 發票整理

**按需**
- `qc-analysis` — 品質分析與績效
- `xlsx` / `pdf` / `docx` / `pptx` — 文件處理

---

## 📞 故障排除

| 問題 | 檢查方向 |
|------|---------|
| 週報統計報錯「找不到月 QF 檔」 | `11_inspdata_monthly\insprecord_QF\` 是否有檔，先跑月報導入 |
| 統計報錯「無法解析週 QF 檔名日期」 | 週 QF 檔名必須是 `insprecord_QF_yyMMdd-yyMMdd.xlsx` |
| 月報導入 `exit 1` | `-year` 沒填，或找不到 `insprecord_QF_current.xlsx` |
| 腳本卡住 / Excel 鎖檔 | 腳本會強制關 Excel；先手動存檔關閉，或用 `-NoKill` |
| rawdata 沒被處理 | 檢查檔名是否符合 `_yymmdd-yymmdd` 且沒有 `Imported_` 前綴 |
| 月報 QF 少了整年資料 | 不要刪 QF 檔重跑，QF 是整年累加 |
| 週報郵件沒寄出 | 只在週一~三執行；`last_sent_week.txt` 記錄本週已寄過 |
| 發票合計對不上 | Uber 費用未稅必須用 `I − 其他品項` 差額回推 |
| 分析數字異常 | 先確認該期間月報有沒有漏導入 |

---

## 📁 路徑

一律以 `E:\88. Claude\` 為準（所有 ps1 / py 腳本、`.claude\settings.local.json` 皆已統一）。

**例外**：`.claude_sync\projects\C--88--Claude\` 這個資料夾名保留 `C--` 前綴，
因為它是資料夾還在 `C:\88. Claude` 時期同步下來的歷史記錄，改名反而會失去對應關係。

---

**📌 備註**: 本指南與記憶系統配合使用，確保工作流程一致性。
新增或修改 skill 後請同步更新本檔與 `Claude Skills 總表.md`。
