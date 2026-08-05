---
name: fabric-inspection-system
description: 布料檢驗 Web 系統啟動與操作 — 啟動驗布 FastAPI 應用、查詢訂單、記錄檢驗數據、發送驗報、查看檢驗記錄。當使用者說「啟動驗布系統」、「開驗布系統」、「記錄驗布」、「發送驗報」時使用。
---

# 布料檢驗 Web 系統（FabQC）

FastAPI 驅動的驗布員現場輸入系統，支援線上查詢、離線記錄、驗報發送。

## 一、快速啟動

### 方式 1：使用 PowerShell（推薦）

```powershell
cd "E:\88. Claude\24_fabric inspsystem\webapp"
python main.py
```

**輸出**：

```
INFO:     Uvicorn running on http://127.0.0.1:8000
INFO:     Press CTRL+C to quit
```

### 方式 2：使用 Visual Studio Code

1. 在 VS Code 打開 `E:\88. Claude\24_fabric inspsystem\` 資料夾
2. 點選「執行」→「開始除錯」（或按 `F5`）
3. 選擇「Python」環境
4. 應用會在 `http://localhost:8000` 啟動

### 方式 3：使用 `run` Skill

```powershell
# Claude Code 中說：
# 「啟動驗布系統」或「執行驗布系統」
```

### 確認啟動成功

打開瀏覽器，前往：

```
http://localhost:8000
```

若看到 **FabQC 驗布系統** 界面，表示啟動成功。

## 二、系統架構

### 後端 — FastAPI (main.py)

```
E:\88. Claude\24_fabric inspsystem\webapp\main.py
│
├─ 讀取資料來源
│  └─ 驗布結論報表*.xls  （Excel 資料庫）
│
├─ 主要功能
│  ├─ GET /  ─ 首頁 HTML
│  ├─ GET /api/records  ─ 查詢驗布記錄（支援篩選）
│  ├─ POST /api/records  ─ 新增或修改驗布記錄
│  ├─ GET /api/po-list  ─ 查詢未開 PO 訂單
│  ├─ POST /api/submit-record  ─ 發送驗報
│  └─ GET /api/submitted  ─ 查看已發送驗報
│
└─ 資料儲存
   └─ submitted_records.json  （已回傳驗報記錄）
```

### 前端 — HTML + JavaScript

```
webapp/static/fabric inspsystem.html
│
├─ 訂單查詢區（上方）
│  ├─ 搜尋欄位
│  ├─ 篩選按鈕
│  └─ 查詢結果表
│
├─ 驗布輸入區（中央）
│  ├─ 驗布信息表單
│  ├─ 缺陷記錄
│  └─ 意見欄位
│
└─ 控制按鈕（右側）
   ├─ 新增
   ├─ 編輯
   ├─ 提交
   └─ 查看已發送
```

## 三、使用流程

### 查詢訂單

```
1. 啟動系統 → http://localhost:8000
   ↓
2. 頂部「搜尋欄位」輸入：
   - PO 號
   - 款號（Style No）
   - 客戶名稱
   ↓
3. 點「查詢」按鈕
   ↓
4. 表格顯示符合的訂單
   - 若標記「未開PO」❎ → 不能驗布
   - 若已發送驗報 ❎ → 已驗報
   ↓
5. 點選訂單列 → 顯示完整訂單信息
```

### 記錄驗布

```
1. 選擇訂單（觸碰行）
   ↓
2. 中央區顯示驗布表單
   - 驗布日期
   - 檢驗人員
   - 驗布結論（Pass / Fail）
   - 收料碼數 / 抽驗碼數
   - A / B / C 級碼數（或缺陷點數）
   ↓
3. 填寫表單
   - 合法值：見下方「欄位規則」
   ↓
4. 點「新增驗布」或「編輯驗布」
   - 新增：新建驗布記錄
   - 編輯：修改既有驗布
   ↓
5. 點「提交」發送驗報
```

### 發送驗報

```
1. 驗布記錄完成
   ↓
2. 點「驗報發送」按鈕
   ↓
3. 系統檢查：
   - ✅ 所有必填欄位都有值
   - ✅ 數值範圍合法
   - ✅ 尚未發送過
   ↓
4. 發送成功 → 記錄存入 submitted_records.json
   ↓
5. 該訂單標記「✅ 已發送」
```

### 查看已發送記錄

```
1. 點「查看已發送」按鈕
   ↓
2. 顯示已成功發送的驗報列表
   - 發送時間
   - 訂單編號
   - 驗布結論
   ↓
3. 點選記錄查看詳細內容
```

## 四、欄位規則

### 必填欄位

| 欄位 | 類型 | 範例 | 說明 |
|------|------|------|------|
| **訂單編號** | 文字 | A202608001 | 系統自動帶出 |
| **款號** | 文字 | MS6FK213R_FA26 | 系統自動帶出 |
| **驗布日期** | 日期 | 2026-08-05 | 必填 |
| **驗布人員** | 文字 | 王小明 | 必填 |
| **驗布結論** | 選單 | Pass / Fail | 必填 |

### 驗布數據欄位

| 欄位 | 單位 | 類型 | 範圍 | 說明 |
|------|------|------|------|------|
| **收料碼數** | yards | 數字 | > 0 | 總收料 |
| **抽驗碼數** | yards | 數字 | > 0, ≤ 收料碼數 | 實際驗布碼數 |
| **A 級碼數** | yards | 數字 | ≥ 0 | 合格品 |
| **B 級碼數** | yards | 數字 | ≥ 0 | 輕微缺陷 |
| **C 級碼數** | yards | 數字 | ≥ 0 | 嚴重缺陷 |
| **平均缺陷點** | pts/100yd | 數字 | ≥ 0 | 百碼瑕疵點 |

### 驗算邏輯（系統自動檢查）

```
✅ A級 + B級 + C級 = 抽驗碼數
✅ 抽驗碼數 ≤ 收料碼數
✅ 驗布結論 = "Pass" 若 C級 = 0，否則 = "Fail"
```

### 缺陷分類（可複選）

| 分類 | 說明 |
|------|------|
| **布面瑕疵** | 織造、染色、印花缺陷 |
| **外觀問題** | 布邊、摺痕、污漬 |
| **手感問題** | 硬度、光澤度異常 |
| **緯斜問題** | 布料歪斜 |
| **顏色問題** | 色差、色牢度 |
| **規格問題** | 寬度、克重不符 |
| **其他問題** | 上述未列項目 |

## 五、API 參考

### 查詢驗布記錄

```bash
GET http://localhost:8000/api/records?po_no=TMKF-26-04485&limit=50

# 回應
{
  "total": 3,
  "records": [
    {
      "id": "1",
      "po_no": "TMKF-26-04485",
      "style_no": "MS6FK213R_FA26",
      "insp_date": "2026-08-05",
      "inspector": "王小明",
      "result": "Pass",
      "insp_qty_yds": 500,
      "a_qty": 480,
      "b_qty": 15,
      "c_qty": 0,
      ...
    },
    ...
  ]
}
```

### 新增或修改驗布記錄

```bash
POST http://localhost:8000/api/records
Content-Type: application/json

{
  "po_no": "TMKF-26-04485",
  "insp_date": "2026-08-05",
  "inspector": "王小明",
  "result": "Pass",
  "bulk_yds": 1000,
  "insp_qty_yds": 500,
  "a_qty": 480,
  "b_qty": 15,
  "c_qty": 0,
  "avg_pts": 2.1,
  "defect_types": ["布面瑕疵", "緯斜問題"],
  "disposition": "OK",
  "remark": "品質良好"
}

# 成功回應 (201)
{
  "id": "123456",
  "status": "success",
  "message": "驗布記錄已儲存"
}
```

### 查詢未開 PO

```bash
GET http://localhost:8000/api/po-list

# 回應
{
  "total": 15,
  "po_list": [
    {
      "po_no": "TMKF-26-04485",
      "style_no": "MS6FK213R_FA26",
      "customer": "Uniqlo",
      "supplier": "帛冠紡織",
      "order_qty": 1000,
      "order_unit": "yds"
    },
    ...
  ]
}
```

### 發送驗報

```bash
POST http://localhost:8000/api/submit-record
Content-Type: application/json

{
  "record_id": "123456"
}

# 成功回應 (200)
{
  "status": "success",
  "message": "驗報已發送",
  "sent_at": "2026-08-05T14:30:00+08:00"
}
```

### 查看已發送驗報

```bash
GET http://localhost:8000/api/submitted?days=7

# 回應
{
  "total": 12,
  "submitted": [
    {
      "record_id": "123456",
      "po_no": "TMKF-26-04485",
      "submitted_at": "2026-08-05T14:30:00+08:00",
      "result": "Pass"
    },
    ...
  ]
}
```

## 六、資料管理

### 資料來源

系統從 Excel 檔案讀取驗布資料：

```
E:\88. Claude\24_fabric inspsystem\驗布結論報表*.xls
```

**檔案名稱格式**：`驗布結論報表_yymmdd.xls`

**系統邏輯**：
- 每次啟動時讀取**最新的** `驗布結論報表*.xls`
- 若多個檔案存在，選擇最新修改時間的檔案

### 欄位對應

系統內部 63 個欄位對應 Excel 列（1-based）：

| 欄位名稱 | Excel 列 | 說明 |
|---------|---------|------|
| id | 1 | 訂單編號 |
| po_no | 8 | PO 號 |
| style_no | 7 | 款號 |
| customer | 4 | 客戶 |
| supplier | 11 | 供應商 |
| insp_date | 15 | 實際驗布日 |
| inspector | 26 | 驗布人員 |
| result | Derived | 驗布結論 (Pass/Fail) |
| insp_qty_yds | 30 | 抽驗碼數 |
| a_qty | 33 | A 級 |
| b_qty | 34 | B 級 |
| c_qty | 35 | C 級 |
| avg_pts | 32 | 平均缺陷點 |
| disposition | 47 | 處置方式 |
| approval_status | 63 | 核准狀態 |

完整欄位對應見 `main.py` 中的 `COLUMNS` dict（共 63 欄）。

### 已發送驗報記錄

所有成功發送的驗報會記錄在：

```
E:\88. Claude\24_fabric inspsystem\webapp\submitted_records.json
```

**格式**：

```json
{
  "submitted": [
    {
      "record_id": "123456",
      "po_no": "TMKF-26-04485",
      "submitted_at": "2026-08-05T14:30:00",
      "result": "Pass"
    }
  ]
}
```

## 七、常見操作

### 查詢特定廠商的驗布記錄

```
URL: http://localhost:8000
1. 搜尋欄輸入廠商名稱
2. 點「搜尋」
3. 表格顯示該廠商所有訂單
```

### 修改已有的驗布記錄

```
1. 查詢訂單 → 點選列
2. 若已有驗布記錄，點「編輯驗布」
3. 修改欄位
4. 點「提交」
```

### 新建多 PO 驗布

```
系統支援單一驗布記錄覆蓋多個 PO 號
（需要配置支援，見後續「進階設定」）
```

### 導出驗報

```
1. 查看已發送記錄
2. 點「下載」（若支援）或 右鍵 → 另存為 PDF
```

## 八、故障排除

| 問題 | 原因 | 解決方案 |
|------|------|---------|
| 啟動失敗：`ModuleNotFoundError: No module named 'fastapi'` | 套件未安裝 | `pip install fastapi uvicorn openpyxl xlrd` |
| 啟動成功但無法訪問 http://localhost:8000 | 防火牆阻擋 | 允許 Python 使用 8000 埠；或在「Windows 防火牆」中新增例外 |
| 表格顯示「找不到驗布結論報表」 | Excel 檔案不存在或路徑錯誤 | 確認 `24_fabric inspsystem` 目錄下有 `驗布結論報表*.xls` 檔 |
| 提交驗報失敗：「欄位驗算不符」 | 數字邏輯錯誤 | 檢查 A+B+C 是否等於抽驗碼數 |
| 查詢結果為空 | 搜尋條件過嚴 | 嘗試搜尋更寬鬆的條件（如只輸入客戶名首字） |
| 系統執行緩慢 | Excel 檔案太大（超過 10MB） | 將舊資料歸檔；只保留當月或當年的 Excel 檔 |

## 九、進階設定

### 修改埠號

預設埠號為 `8000`。若需更改，編輯 `main.py`：

```python
# 最後一行，改為：
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=9000)  # 改成 9000
```

重啟應用後，訪問 `http://localhost:9000`。

### 連接遠端 Excel

若 Excel 檔放在網路共享資料夾：

```python
EXCEL_DIR = Path(r"\\server_ip\shared_folder\fabric_data")
```

### 啟用 CORS（跨域請求）

若需從其他 Web 應用調用 API：

```python
from fastapi.middleware.cors import CORSMiddleware

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
```

### 自訂驗證規則

編輯 `main.py` 中的驗算邏輯（`validate_record()` 函數）。

## 十、相關檔案與技術

| 檔案 | 用途 |
|------|------|
| `main.py` | FastAPI 後端（254 行） |
| `fabric inspsystem.html` | Web 前端介面 |
| `submitted_records.json` | 已發送驗報日誌 |
| `驗布結論報表*.xls` | Excel 資料庫 |
| `驗布系統功能.pptx` | 功能說明簡報 |

### 技術棧

- **後端**：FastAPI + Uvicorn + Python 3.8+
- **前端**：HTML5 + CSS3 + Vanilla JavaScript
- **資料庫**：Excel（openpyxl 讀取 .xlsx；xlrd 讀取 .xls）
- **伺服器**：Uvicorn（ASGI）

### 環境要求

```bash
Python >= 3.8
fastapi >= 0.100
uvicorn >= 0.23
openpyxl >= 3.0
xlrd >= 2.0
```

## 十一、週程與集成

| 場景 | 觸發 | 系統操作 |
|------|------|---------|
| 日常驗布 | 驗布員進廠 | 啟動系統 → 查詢訂單 → 記錄驗布 → 發送驗報 |
| 週報統計 | 週三 11:30 | 讀取已發送驗報 JSON → 彙總至週報 |
| 月度審計 | 月末 | 導出所有驗報記錄 → 檔案存檔 |

---

**需要協助**？

| 任務 | 說法 |
|------|------|
| 啟動系統 | 「啟動驗布系統」→ `run` skill |
| 查詢訂單 | 直接用系統界面搜尋 |
| API 集成 | 見上方 API 參考章節 |
| 修改欄位 | 編輯 `main.py` 中的 `COLUMNS` dict |
| 備份資料 | 定期複製 `submitted_records.json` 與 Excel 檔 |
