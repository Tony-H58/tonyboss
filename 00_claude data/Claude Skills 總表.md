# Claude Skills 總表

> 更新日期：2026-07-29
> 觸發方式：在對話中說出對應的關鍵詞，Claude 會自動使用對應 skill。
> 本表只列**實際可用**的 skill。自訂 skill 定義在 `E:\88. Claude\.claude\skills\`。

---

## ⭐ 自訂 Skill（本專案）— 9 個

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

### dashboard-qc — 品管儀表板查看與管理 ⭐ 新增
**觸發**：「開儀表板」、「品管儀表板」、「查看儀表板」、「儀表板總頁」、「品質看板」、「績效看板」
**功能**：打開三頁 HTML 儀表板、檢查週度績效指標、更新資料、自動寄送
**涵蓋**：15_dashboard
**內容**：快速訪問方法、資料更新流程、常見錯誤排查、鍵盤快捷鍵、效能最佳化

### download-checker — 週報月報下載監控 ⭐ 新增
**觸發**：「檢查下載」、「檢查週報」、「檢查月報」、「待處理資料」、「設定排程」
**功能**：檢查下載狀態、自動處理待處理檔案、設定定時檢查排程、監控日誌
**涵蓋**：02_reminder
**內容**：快速檢查指令、自動處理邏輯、排程設定、常見陷阱、進階監控

### fabric-inspection-system — 布料檢驗 Web 系統 ⭐ 新增
**觸發**：「啟動驗布系統」、「開驗布系統」、「記錄驗布」、「發送驗報」
**功能**：啟動 FastAPI Web 應用、查詢訂單、記錄驗布資料、發送驗報、查看已發送記錄
**涵蓋**：24_fabric inspsystem
**內容**：啟動方法、使用流程、API 參考、資料管理、常見錯誤、實際案例

### thumbnail-html-embed — HTML 縮圖最佳實踐 ⭐ 新增
**觸發**：「嵌入縮圖」、「HTML 縮圖」、「圖片縮圖」、「響應式縮圖」、「優化圖片」
**功能**：提供響應式縮圖技術、圖片最佳化方法、CSS 樣式範本、自動化工具
**涵蓋**：通用（HTML 縮圖指南）
**內容**：技術方案、性能優化、實際範例、線上工具、快速故障排除

### qc-expense-management — QC 費用與報銷管理 ⭐ 新增
**觸發**：「整理出差費」、「QC 報銷」、「差旅費管理」、「報銷單」
**功能**：管理 QC 人員出差費用、整理報銷發票、統計費用數據、編制支出報表
**涵蓋**：21_qc expense chart、22_import qc expense
**內容**：月度流程、Excel 表單、費用分類、對帳方法、常見問題

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
| 跑週報全流程 | weekly-report |
| 跑月報全流程 | monthly-import |
| 前後端比對 / QC 績效 | qc-analysis |
| 整理平台發票 | invoice-transfer |
| 查看品管儀表板 | dashboard-qc |
| 檢查週報月報下載 | download-checker |
| 啟動驗布系統 | fabric-inspection-system |
| HTML 縮圖嵌入 | thumbnail-html-embed |
| 整理 QC 出差費用 | qc-expense-management |
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

## 📌 版本記錄

### 2026-08-05 更新（當前版本）

**新增 5 個自訂 Skill**：
- `dashboard-qc` — 品管儀表板查看與管理
- `download-checker` — 週報月報下載監控
- `fabric-inspection-system` — 布料檢驗 Web 系統
- `thumbnail-html-embed` — HTML 縮圖最佳實踐
- `qc-expense-management` — QC 費用與報銷管理

總計自訂 Skill 數：**9 個**（原 4 個 + 新增 5 個）

### 2026-07-29 修正記錄

先前版本列出但**實際不存在**的 skill，已從本表移除：

`verify`、`schedule`、`engineering:standup`、`engineering:debug`、
`engineering:tech-debt`、`testing-strategy`、`code-review`、
`consolidate-memory`、`setup-cowork`

原「快速參考」表誤寫的 `invoice-extractor` 已更正為 `invoice-transfer`。
