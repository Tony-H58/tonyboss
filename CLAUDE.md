# CLAUDE.md

本專案是 Tony Huang 的驗布 / QC / 發票自動化工作流程集合（原路徑 `E:\88. Claude\`）。
多數腳本以 PowerShell（Windows 端）與 Python 執行，路徑與 skill 定義見
`00_claude data/INDEX.md`、`00_claude data/Skill使用指南.md`、`00_claude data/Claude Skills 總表.md`。

## 開發流程規則：先查 skill、寫完評估優化、明確說「儲存」才寫入

### 1. 寫程式前，先檢查是否有可用的 skill

開始撰寫任何腳本或程式之前，先確認 `.claude/skills/` 底下、以及
`00_claude data/Skill使用指南.md`、`00_claude data/Claude Skills 總表.md`
是否已有涵蓋這個需求的 skill，避免重複造輪子。目前的自訂 skill：

- `invoice-transfer` — 平台發票辨識與轉入整理記錄
- `monthly-import` — 驗布月報下載、轉換、導入
- `weekly-report` — 驗布週報下載、轉換、導入、寄送
- `qc-analysis` — 前後端驗布比對、QC 績效 KPI

找不到符合的 skill 才開始寫新程式；若有現成 skill，優先參考或直接沿用它的流程與路徑。

### 2. 程式完成後，評估是否該新增或更新 skill

程式驗證可用之後，評估：

- 這是不是一個之後還會重複執行的流程？如果是，考慮用 `skill-creator` 把它包裝成新的 skill。
- 這次的修改是否讓某個既有 skill 的 `SKILL.md` 內容過時（路徑、指令、步驟改變）？如果是，準備好對應的更新內容。

### 3. 儲存規則：使用者說「儲存」才真正寫入

評估後若決定要新增或更新 skill，**先在對話中提出草稿**，不要主動寫入
`.claude/skills/`。只有在使用者明確說「儲存」時，才：

1. 把新增/更新的 skill 內容寫入 `.claude/skills/<name>/SKILL.md`
2. 同步更新 `00_claude data/Claude Skills 總表.md` 與 `00_claude data/Skill使用指南.md`
3. 若在 git 儲存庫中工作，commit 對應變更（是否 push 依當下情境判斷）

其他一般程式修改（非 skill 定義本身）不受此限制，仍照平常流程處理。
