# Claude Skills 總表

> 更新日期：2026-06-09
> 觸發方式：在對話中說出對應的關鍵詞，Claude 會自動使用對應 skill。

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

---

## 發票處理類

### invoice-transfer — 平台發票辨識轉入 ⭐ 自訂
**觸發**：「整理發票」、「辨識發票」、「轉入發票」、「執行 XX月 XX分店」、「繼續發票整理」
**功能**：從 Uber Eats / FoodPanda 平台發票 Excel 提取圖片，辨識發票號碼、日期、費用明細、金額，二次驗證後寫入整理記錄
**適用檔案**：平台發票-114-X.xlsx（工作表內含發票圖片，非儲存格資料）

---

## 排程與自動化類

### schedule — 排程任務
**觸發**：「每天定時」、「排程執行」、「定期提醒」、「cron job」
**功能**：建立、管理、執行定期遠端 Agent 任務；一次性排程

### loop — 循環執行
**觸發**：「每 5 分鐘檢查」、「持續監控」、「重複執行」
**功能**：以固定間隔重複執行某個指令或任務

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

## 記憶與協作類

### consolidate-memory — 整理記憶
**觸發**：「整理記憶」、「清理舊記憶」
**功能**：合併重複記憶、修正過時資料、精簡記憶索引

### setup-cowork — 設定協作模式
**觸發**：「設定 Cowork」、「安裝協作插件」
**功能**：引導安裝角色對應的插件、連接工具、試用 skill

---

## Skill 管理

### skill-creator — 建立新 Skill
**觸發**：「建立 skill」、「把這個流程做成 skill」、「新增自訂 skill」
**功能**：設計、撰寫、測試、優化新的 skill；迭代改善現有 skill

---

## 快速參考

| 我想做的事         | 用哪個 Skill         |
| ------------- | ----------------- |
| 整理發票 Excel 圖片 | invoice-extractor |
| 建立/編輯 Excel   | xlsx              |
| 處理 PDF        | pdf               |
| 做簡報           | pptx              |
| 寫 Word 報告     | docx              |
| 定時自動執行        | schedule / loop   |
| 開發 Claude API | claude-api        |
| 設定自動行為        | update-config     |
| 建立新的 skill    | skill-creator     |
