# 📚 00_claude data 資訊導航

**最後更新**: 2026-08-04  
**管理者**: Tony Huang

---

## 📖 文檔導航

### 🎯 **Skill使用指南.md** ⭐ 必讀
核心工作流程指南，包含：
- 開發流程規則（先查 skill、寫完評估優化、說「儲存」才寫入 — 同步寫在專案根目錄 `CLAUDE.md`）
- 各系統與 Skill 的對應關係
- 常見工作場景的完整步驟
- 快速命令參考表
- 故障排除指南

**用途**: 執行任何系統工作時的第一參考  
**更新頻率**: 每月  
**相關系統**: 所有 (品管、驗布、儀表板、分析)

---

### 📋 **Claude Skills 總表.md**
Skill 總覽（⭐ 自訂 + 內建），包含：
- 4 個本專案自訂 Skill 的觸發詞與涵蓋範圍
- Claude Code 內建 Skill 清單與功能說明
- 使用觸發條件

**用途**: 查詢某件事該用哪個 Skill  
**更新頻率**: 新增/修改 Skill 時，或 Claude 更新時  
**相關系統**: 通用

---

### 📊 **11_inspdata_monthly 筆記.md**
月度驗布資料導入系統的筆記，包含：
- 月度流程說明
- 常見問題與解決方案
- 導入檢查清單

**用途**: 執行月度導入前的準備與檢查  
**更新頻率**: 每月一次  
**相關系統**: 11_inspdata_monthly 資料夾

---

### 💰 **發票整理筆記.md**
發票與出差費用管理筆記，包含：
- 發票分類規則
- 出差費用報銷流程
- 整理步驟與檢查清單

**用途**: 處理發票與費用報銷時參考  
**更新頻率**: 每季  
**相關系統**: 31_electronic invoice, 21_qc expense chart, 22_import qc expense

---

## 🗂️ 目錄結構說明

```
E:\88. Claude\00_claude data\
├── INDEX.md                           ← 本檔案（導航入口）
├── Skill使用指南.md                    ← ⭐ 核心工作流程指南
├── Claude Skills 總表.md              ← Skill 總覽（自訂 + 內建）
├── 11_inspdata_monthly 筆記.md        ← 月度導入說明
├── 發票整理筆記.md                     ← 發票與費用管理
└── tony筆記本\                        ← 個人筆記
```

### ⭐ 自訂 Skill 定義

```
E:\88. Claude\.claude\skills\
├── weekly-report\SKILL.md      ← 週報完整流程
├── monthly-import\SKILL.md     ← 月報完整流程
├── qc-analysis\SKILL.md        ← 品質分析與 QC 績效
└── invoice-transfer\SKILL.md   ← 平台發票辨識轉入
```

---

## 🚀 快速開始

### 新手入門
1. 閱讀 **Skill使用指南.md**
2. 根據需求找到對應章節
3. 按流程步驟執行

### 常見任務
| 任務 | 查看文檔 | 系統資料夾 | Skill |
|------|--------|---------|-------|
| 週報自動化 | Skill使用指南.md | 03 → 12 | weekly-report |
| 月度驗布導入 | 11_inspdata_monthly 筆記.md | 03 → 11 | monthly-import |
| 發票管理 | 發票整理筆記.md | 31_electronic invoice | invoice-transfer |
| 出差費用 | 發票整理筆記.md | 22_import qc expense | — |
| 品管儀表板 | Skill使用指南.md | 15_dashboard | — |
| 品質分析 | Skill使用指南.md | 13_compare style inspquality | qc-analysis |
| QC 績效 | Skill使用指南.md | 14_qc performance | qc-analysis |

---

## 📝 文檔管理規則

### 何時更新
- ✅ 發現新的工作方式
- ✅ 系統流程改變
- ✅ 新增 Skill 或工具
- ❌ 不需要逐次記錄瑣碎變化

### 更新方式
1. 在對應文檔中新增內容
2. 更新本 INDEX.md 的日期
3. 存檔並備份

### 版本控制
- 核心文檔 (Skill使用指南.md): 月度審查
- 流程筆記: 按需更新
- 總表文檔: 每季檢查

---

## 🔗 相關資訊

### 整體系統地圖
```
E:\88. Claude\
├── .claude/skills/               ← ⭐ 自訂 Skill 定義
├── 00_claude data/               ← 你在這裡 📍
├── 00_html file/                 ← HTML 產出
├── 02_reminder/                  ← 郵件與提醒、排程
├── 03_download rawdata/          ← 週報月報下載轉換
├── 11_inspdata_monthly/          ← 月度驗布導入
├── 12_inspdata_weekly/           ← 週度驗布導入
├── 13_compare style inspquality/ ← 前後端比對分析
├── 14_qc performance/            ← QC 績效評估
├── 15_dashboard/                 ← 品管儀表板
├── 21_qc expense chart/          ← 出差報銷單
├── 22_import qc expense/         ← 差旅費導入
├── 23_textile knowledge/         ← 紡織知識
├── 24_fabric inspsystem/         ← 布料檢驗系統（FastAPI）
├── 31_electronic invoice/        ← 平台發票管理
└── 32_business trip report/      ← 訪廠簡報
```

### 外部參考
- 📌 記憶系統: `C:\Users\tonyhuang\.claude\projects\...\memory\MEMORY.md`
- 📌 Claude 官方文檔: https://claude.com
- 📌 工作檔案: `E:\88. WorkFile\`

---

## 💬 使用建議

**1. 首次查閱**
→ 從 **Skill使用指南.md** 開始

**2. 特定系統問題**
→ 查閱對應的筆記檔案

**3. 需要全面回顧**
→ 按順序閱讀所有文檔

**4. 想要快速查詢**
→ 使用本 INDEX.md 的快速開始表

---

**最後更新**: 2026-07-29 by Claude Code

