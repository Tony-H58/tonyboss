# Claude Skills 總表

> 更新日期：2026-08-17
> 觸發方式：在對話中說出對應的關鍵詞，Claude 會自動使用對應 skill。
> 本表只列**實際可用**的 skill。自訂 skill 定義在 `E:\88. Claude\.claude\skills\`。

---

## ⭐ 自訂 Skill（本專案）

### download-tasks — 下載任務管理 ⭐ 新增
**觸發**：「下載」、「批量下載」、「下載報表」、「下載 rawdata」、「啟動下載任務」
**功能**：管理和執行下載任務隊列；支持批量下載、斷點續傳、自動分類存檔
**涵蓋**：03_download rawdata、31_electronic invoice
**特性**：並發下載、進度跟踪、完整性驗證、NTLM 認證

### image-to-html — 圖片處理轉 HTML ⭐ 新增
**觸發**：「生成相冊」、「整理照片」、「圖片轉 HTML」、「批量縮圖」、「製作照片展示」
**功能**：用戶選擇本機圖片，自動縮圖並生成響應式 HTML 相冊
**涵蓋**：00_html file
**特性**：本機選圖、自動縮圖、Base64 嵌入、燈箱查看、多設備適配

### data-converter — 數據格式轉換 ⭐ 新增
**觸發**：「轉換格式」、「CSV 轉 Excel」、「Excel 導出 HTML」、「數據轉換」、「格式化導出」
**功能**：CSV ↔ Excel ↔ HTML 多向格式轉換；自動檢測源格式、智能表頭識別、樣式保留
**涵蓋**：所有涉及數據格式轉換的工作流程
**特性**：智能類型檢測、格式保留、批量轉換、交互式 HTML 表格

### weekly-report — 驗布週報完整流程 ⭐ 自訂
**觸發**：「跑週報」、「週報」、「週報流程」、「週度導入」、「寄週報」
**功能**：下載 rawdata → 格式轉換 → 導入 InspRecord_QF → 計算週統計 → 寄送週報郵件
**涵蓋**：03_download rawdata → 12_inspdata_weekly → 02_reminder

### monthly-import — 驗布月報完整流程 ⭐ 自訂
**觸發**：「跑月報」、「月報」、「月報流程」、「月度導入」
**功能**：下載 rawdata → 轉換歸檔 → 導入年度 InspRecord_QF → 歸檔來源檔
**涵蓋**：03_download rawdata → 11_inspdata_monthly
**注意**：`-year` 必填且為西元年（如 2026）

### qc-analysis — 品質分析與 QC 績效 ⭐ 自訂
**觸發**：「比對驗布」、「前後端比對」、「品質分析」、「QC 績效」、「跑 KPI」
**功能**：執行前後端驗布比對（比對.py）與 QC 人員 KPI 計算（qc_kpi_final.py），驗證 HTML 報告
**涵蓋**：13_compare style inspquality、14_qc performance

### invoice-transfer — 平台發票辨識轉入 ⭐ 自訂
**觸發**：「整理發票」、「辨識發票」、「轉入發票」、「執行 XX月 XX分店」、「繼續發票整理」
**功能**：從 Uber Eats / FoodPanda 平台發票 Excel 提取內嵌圖片，辨識發票號碼、日期、費用明細、金額，二次驗證後寫入整理記錄
**適用檔案**：平台發票-114-X.xlsx（工作表內含發票圖片，非儲存格資料）
**涵蓋**：31_electronic invoice

---

## 文件處理類

### xlsx — Excel 試算表
**觸發**：提到 `.xlsx`、`.csv`、試算表、欄位計算、格式整理
**功能**：建立、讀取、編輯 Excel；資料清理；公式計算；圖表；格式化輸出

### pdf — PDF 文件
**觸發**：提到 `.pdf`、PDF 合併/拆分、OCR、表格提取
**功能**：讀取文字/表格、合併/拆分 PDF、加浮水印、填表單、加密、OCR 掃描檔

### docx — Word 文件
**觸發**：提到 Word doc、`.docx`、報告、備忘錄、信件、目錄
**功能**：建立/編輯 Word 文件；插入圖片；目錄/頁碼；追蹤修訂；格式排版

### pptx — PowerPoint 簡報
**觸發**：提到 deck、slides、簡報、`.pptx`
**功能**：建立/編輯投影片；讀取/提取文字；版面、備忘稿、範本

### dataviz — 圖表與視覺化
**觸發**：要做圖表、儀表板、資料視覺化前
**功能**：配色系統、圖型選擇、座標軸/圖例/提示規範；深淺色模式一致性

---

## 執行與自動化類

### run — 執行應用程式
**觸發**：「執行」、「啟動」、「跑一下」某個腳本或服務
**功能**：啟動並執行 PowerShell 腳本、Python 腳本、後端服務

### loop — 循環執行
**觸發**：「每 5 分鐘檢查」、「持續監控」、「重複執行」
**功能**：以固定間隔重複執行某個指令或任務

> 📌 **定時排程**沒有對應 skill。Windows 工作排程用 `02_reminder\setup_schedules.ps1` 建立（需管理員權限）。

---

## 開發類

### claude-api — Claude API 開發
**觸發**：程式碼中有 `import anthropic`、問 Claude API 用法、Anthropic SDK
**功能**：建立/除錯 Claude API 應用；Prompt caching；工具使用；模型版本遷移

### init — 初始化 CLAUDE.md
**觸發**：「初始化專案」、「建立 CLAUDE.md」
**功能**：為程式碼庫建立 CLAUDE.md 說明文件

### review — PR 審查
**觸發**：「審查 PR」、「review pull request」
**功能**：審查 GitHub Pull Request

### security-review — 安全審查
**觸發**：「安全檢查」、「security review」
**功能**：對當前 branch 的變更進行安全性審查

### simplify — 程式碼簡化
**觸發**：程式碼修改後要求優化、簡化
**功能**：審查修改的程式碼品質、重用性、效率，修復問題

### session-start-hook — 啟動 hook
**觸發**：「設定 web session」、「建立 SessionStart hook」
**功能**：讓專案在 Claude Code web session 中能跑測試與 linter

---

## 設定類

### update-config — 設定 Claude Code
**觸發**：「從現在起每次都要…」、「允許 XXX 指令」、「設定 hook」、「修改 settings.json」
**功能**：設定自動行為 hook；新增/修改權限；環境變數

### keybindings-help — 自訂快捷鍵
**觸發**：「重新綁定快捷鍵」、「修改 keybindings」
**功能**：自訂 Claude Code 鍵盤快捷鍵

### fewer-permission-prompts — 減少權限提示
**觸發**：「減少確認提示」、「常用指令自動允許」
**功能**：掃描對話記錄，自動加入常用指令的允許清單

---

## Skill 管理

### skill-creator — 建立新 Skill
**觸發**：「建立 skill」、「把這個流程做成 skill」、「新增自訂 skill」
**功能**：設計、撰寫、測試、優化新的 skill；迭代改善現有 skill

---

## 快速參考

| 我想做的事 | 用哪個 Skill |
| --------- | ----------- |
| 下載報表/rawdata | download-tasks ⭐ 新 |
| 圖片處理轉 HTML 相冊 | image-to-html ⭐ 新 |
| 數據格式轉換（CSV/Excel/HTML） | data-converter ⭐ 新 |
| 跑週報全流程 | weekly-report |
| 跑月報全流程 | monthly-import |
| 前後端比對 / QC 績效 | qc-analysis |
| 整理平台發票 | invoice-transfer |
| 建立/編輯 Excel | xlsx |
| 處理 PDF | pdf |
| 做簡報 | pptx |
| 寫 Word 報告 | docx |
| 做圖表/儀表板 | dataviz |
| 執行單一腳本 | run |
| 定間隔重複執行 | loop |
| 開發 Claude API | claude-api |
| 設定自動行為 | update-config |
| 建立新的 skill | skill-creator |

---

## 📌 更新記錄

### 2026-08-17 新增三個 Skill
- ✅ **download-tasks** — 下載任務管理（批量下載、斷點續傳、自動分類）
- ✅ **image-to-html** — 圖片處理（本機選圖→縮圖→生成相冊）
- ✅ **data-converter** — 數據格式轉換（CSV ↔ Excel ↔ HTML）

### 2026-07-29 修正記錄

先前版本列出但**實際不存在**的 skill，已從本表移除：

`verify`、`schedule`、`engineering:standup`、`engineering:debug`、
`engineering:tech-debt`、`testing-strategy`、`code-review`、
`consolidate-memory`、`setup-cowork`

原「快速參考」表誤寫的 `invoice-extractor` 已更正為 `invoice-transfer`。
