---
name: dashboard-qc
description: 品管儀表板查看與管理 — 開啟三頁互動儀表板、檢查關鍵績效指標、建立儀表板副本、設定自動寄送。當使用者說「開儀表板」、「品管儀表板」、「查看儀表板」、「儀表板總頁」、「品質看板」、「績效看板」時使用。
---

# 品管儀表板

三份靜態 HTML 儀表板，展示週度檢驗績效、不合格數據與品質趨勢。

## 一、儀表板結構

| 檔案 | 用途 | 連結方式 |
|------|------|---------|
| **品管儀表板_總頁.html** | 導航頁面，跳轉至 A / B 頁 | 點選分店連結 |
| **品管儀表板_A.html** | 分店 A 績效指標 | 表格、圖表、返回總頁 |
| **品管儀表板_B.html** | 分店 B 績效指標 | 表格、圖表、返回總頁 |

### 📍 檔案位置

```
E:\88. Claude\15_dashboard\
├─ 品管儀表板_總頁.html
├─ 品管儀表板_A.html
├─ 品管儀表板_B.html
└─ MAKALOT GROUP.png （品牌標誌）
```

## 二、快速查看

### 方式 1：雙擊開啟（最快）

直接在資源管理器開啟任一 HTML 檔，瀏覽器會自動開啟。

```powershell
# 從 PowerShell 開啟總頁
Invoke-Item "E:\88. Claude\15_dashboard\品管儀表板_總頁.html"
```

### 方式 2：瀏覽器書籤

將儀表板加入瀏覽器書籤，下次快速訪問：
1. 開啟 `品管儀表板_總頁.html`
2. 按 `Ctrl+D` 加入書籤
3. 命名為「品管儀表板」

### 方式 3：桌面快捷方式

建立快捷方式到桌面：

```powershell
$source = "E:\88. Claude\15_dashboard\品管儀表板_總頁.html"
$desktop = [Environment]::GetFolderPath("Desktop")
$link = (New-Object -ComObject WScript.Shell).CreateShortCut("$desktop\品管儀表板.lnk")
$link.TargetPath = $source
$link.Save()
Write-Host "✓ 快捷方式已建立到桌面"
```

## 三、儀表板內容

### 總頁（品管儀表板_總頁.html）

```
┌─────────────────────────────────────┐
│   MAKALOT GROUP — 品管儀表板       │
│                                     │
│  選擇檢查分店：                      │
│  ┌─────────────┐  ┌─────────────┐  │
│  │   分店 A    │  │   分店 B    │  │
│  │ 點此檢查    │  │ 點此檢查    │  │
│  └─────────────┘  └─────────────┘  │
│                                     │
│  更新時間：yyyy-MM-dd HH:mm:ss     │
└─────────────────────────────────────┘
```

### 分店儀表板（A / B 頁）

#### 上方 — 關鍵指標

- **本週檢驗批數** — 已完成的檢驗件數
- **合格率** — 通過檢驗的百分比
- **平均缺陷點** — 每百碼平均瑕疵點數
- **C 級比例** — 不合格品佔比

#### 中方 — 每日趨勢圖

- X 軸：日期（過去 7 天）
- Y 軸：合格率 (%)、缺陷點數
- 藍線：合格率趨勢
- 紅線：缺陷點數趨勢

#### 下方 — 明細表

| 日期 | 檢驗數 | 合格 | 不合格 | 合格率 | 平均缺陷點 |
|------|--------|------|--------|--------|-----------|
| ... | ... | ... | ... | ... | ... |

**返回總頁**：點選左上「← 返回」按鈕

## 四、資料更新

### 更新流程

儀表板資料來自 **週報統計 Excel**：

```
12_inspdata_weekly\
└─ analysis chart_weekly_yymmdd-yymmdd.xlsx
    └─ Sheet 內含圖表與明細
```

**更新時機**：
- 每週三 (UTC+8 11:30)：`weekly-report` skill 執行
- 或手動執行 `週報自動.ps1`

### 手動更新儀表板

如需緊急更新（如修正檢驗數據），聯絡系統維護人員：

```powershell
# 1. 執行週報流程
cd "E:\88. Claude"
./週報自動.ps1 -Fast

# 2. 導入新資料（已包含在上一步）
cd "12_inspdata_weekly"
./run_weekly_full.ps1 -Fast

# 3. 刷新瀏覽器（Ctrl+F5 强制重載）
```

## 五、常見問題

### Q: 儀表板顯示的是舊資料怎麼辦？

**A:** 瀏覽器快取導致。按 `Ctrl+F5`（或 `Cmd+Shift+R` Mac）強制重新載入。

### Q: 只有兩個分店嗎？能不能新增分店 C / D？

**A:** 目前固定為 A / B 兩個分店。如需新增，需修改 HTML 和週報統計邏輯。
聯絡系統維護人員。

### Q: 能否將儀表板分享給其他人？

**A:** 可以。直接傳送三個 HTML 檔案（含 MAKALOT GROUP.png）給他人。
他人開啟時無需設定，直接雙擊即可使用。

### Q: 儀表板在公司網路能否網頁瀏覽（不用下載檔案）？

**A:** 需要架設靜態 Web 伺服器。目前沒有部署，只能本機開啟。

### Q: 圖表為什麼不能互動（無法篩選日期範圍）？

**A:** 目前儀表板是**靜態 HTML**。若需互動功能，需改用 JavaScript 圖表庫
（如 Chart.js、Plotly）重寫。

## 六、技術細節

### HTML 結構

```
品管儀表板_總頁.html
├─ <meta charset="utf-8">
├─ <style> 內嵌 CSS（無外部依賴）
├─ <body>
│  ├─ Logo（MAKALOT GROUP.png）
│  ├─ 分店按鈕（連結到 A / B 頁）
│  └─ 更新時間戳
└─ <script> 簡單的導航邏輯

品管儀表板_A.html
品管儀表板_B.html
├─ 表格（<table> 靜態資料）
├─ 圖表（嵌入式 SVG 或 base64 image）
└─ 返回總頁按鈕
```

### 資料綁定

表格與圖表的資料來自：

```
12_inspdata_weekly\analysis chart_weekly_yymmdd-yymmdd.xlsx
│
├─ Sheet1（週統計明細）
│  └─ 日期 | 分店 | 檢驗數 | 合格 | 不合格 | 缺陷點 | ...
│
└─ Charts（內嵌圖表物件）
   └─ 趨勢線圖、柱狀圖等
```

更新資料時需同時更新 HTML 中的表格與圖表 base64 image。

## 七、進階 — 自訂儀表板

### 複製並修改模板

1. 複製 `品管儀表板_A.html` 為 `品管儀表板_C.html`
2. 編輯文字與資料（用文字編輯器開啟）
3. 修改表格數值與圖表 base64 編碼

### 編輯步驟

```html
<!-- 修改分店名稱 -->
<h1>分店 C 週度績效</h1>

<!-- 修改表格資料 -->
<tr>
  <td>2026-08-04</td>
  <td>45</td>
  <td>44</td>
  <td>1</td>
  <td>97.8%</td>
  <td>2.1</td>
</tr>

<!-- 修改圖表（需更新 base64） -->
<img src="data:image/png;base64,iVBORw..." />
```

### 圖表 Base64 轉換

如需更新圖表：

1. 在 Excel 中製作新圖表
2. 複製圖表 → 貼至 Paint
3. 儲存為 PNG
4. 使用線上轉換工具：https://www.base64-image.de/
5. 將 base64 碼貼回 HTML 的 `<img src="data:image/png;base64,...">`

## 八、週程

| 時間 | 工作 | 涵蓋 Skill |
|------|------|-----------|
| 每週三 11:30 | 自動更新儀表板資料 | **weekly-report** |
| 每週三 12:00 | 寄送儀表板至信箱 | **weekly-report** → `send_weekly_email.ps1` |
| 隨時 | 手動查看儀表板 | **dashboard-qc** |

## 九、相關文件

- `12_inspdata_weekly\analysis chart_weekly_*.xlsx` — 資料來源
- `02_reminder\send_weekly_email.ps1` — 自動寄送週報 + 儀表板
- `00_claude data\Skill使用指南.md` — 完整工作流程

---

**需要更新儀表板或新增功能？**  
說「更新儀表板」→ **run** skill 手動執行週報流程  
說「新增分店」→ 聯絡系統維護人員建立 C / D 頁面
