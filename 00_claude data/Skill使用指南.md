# Claude Skills 使用指南

**最後更新**: 2026-06-24  
**維護者**: Tony Huang

---

## 📋 目錄
1. [系統工作流程對應](#系統工作流程對應)
2. [各 Skill 詳細說明](#各skill詳細說明)
3. [常見工作場景](#常見工作場景)
4. [快速參考](#快速參考)

---

## 系統工作流程對應

### 🔄 驗布報表自動化系統（03_download rawdata）

| 工作 | 使用 Skill | 檔案位置 | 說明 |
|------|-----------|--------|------|
| 週報下載轉換 | **run** | 03_download rawdata | 執行 `週報自動.ps1` 一鍵自動 |
| 月報下載轉換 | **run** | 03_download rawdata | 執行 `月報自動.ps1` 一鍵自動 |
| HTTP 自動下載 | **run** | 03_download rawdata | 執行 `auto_download_http.py` |
| 格式轉換 | **run** | 03_download rawdata | 執行 `convert_rawdata.ps1` |
| 完整流程 | **run** | 03_download rawdata | 執行 `run_download_and_convert.ps1` |

---

### 📊 儀表板與分析系統

#### 品管儀表板生成（15_dashboard）

| 工作 | 使用 Skill | 檔案位置 | 說明 |
|------|-----------|--------|------|
| 生成儀表板 A | **run** | 15_dashboard | 執行品管儀表板_A.html 相關 Python |
| 生成儀表板 B | **run** | 15_dashboard | 執行品管儀表板_B.html 相關 Python |
| 檢查更新 | **engineering:standup** | 02_reminder | 執行 `check_dashboard_update.ps1` |

#### 驗布比對分析（13_比對前後驗布記錄）

| 工作 | 使用 Skill | 檔案位置 | 說明 |
|------|-----------|--------|------|
| 年度比對分析 | **run** | 13_比對前後驗布記錄 | 執行 `比對.py` |
| 個別供應商分析 | **run** | 13_比對前後驗布記錄 | 執行 `analyze_supplier.py` |
| 查看分析結果 | **verify** | 13_比對前後驗布記錄 | 檢查 `比對結果.html` |

#### QC績效評估（14_qc performance）

| 工作 | 使用 Skill | 檔案位置 | 說明 |
|------|-----------|--------|------|
| 計算 KPI | **run** | 14_qc performance | 執行 `qc_kpi_final.py` |
| 驗證報告 | **verify** | 14_qc performance | 檢查 `qc_kpi_report.html` |

---

### 📁 驗布資料導入系統

#### 月度導入（11_inspdata_monthly）

| 工作 | 使用 Skill | 檔案位置 | 說明 |
|------|-----------|--------|------|
| 完整月度流程 | **run** | 11_inspdata_monthly | 執行 `run_monthly_full.ps1 -month "YYMM"` |
| 快速導入 | **run** | 11_inspdata_monthly | 執行 `run_monthly_full.ps1 -month "YYMM" -Fast` |
| 無等待導入 | **run** | 11_inspdata_monthly | 執行 `run_monthly_full.ps1 -month "YYMM" -NoKill` |
| 步驟 4-7 | **run** | 11_inspdata_monthly | 執行 `import_monthly.ps1` 單獨步驟 |

#### 週度導入（12_inspdata_weekly）

| 工作 | 使用 Skill | 檔案位置 | 說明 |
|------|-----------|--------|------|
| 完整週度流程 | **run** | 12_inspdata_weekly | 執行 `run_weekly_full.ps1` |
| 快速週度 | **run** | 12_inspdata_weekly | 執行 `run_weekly_full.ps1 -Fast` |
| 週度統計 | **run** | 12_inspdata_weekly | 執行 `calc_weekly_stats.ps1` |
| 檢查與處理 | **run** | 12_inspdata_weekly | 執行 `check_and_process.ps1` |

---

### 📧 郵件與提醒系統（02_reminder）

| 工作 | 使用 Skill | 檔案位置 | 說明 |
|------|-----------|--------|------|
| 發送週報郵件 | **schedule** | 02_reminder | 執行 `send_weekly_email.ps1` |
| 檢查週報下載 | **engineering:standup** | 02_reminder | 執行 `check_weekly_download.ps1` |
| 檢查月報下載 | **engineering:standup** | 02_reminder | 執行 `check_monthly_download.ps1` |
| 處理待處理資料 | **schedule** | 02_reminder | 執行 `process_pending_rawdata.ps1` |

---

### 🧵 布料檢驗系統（23_fabric inspsystem）

| 工作 | 使用 Skill | 檔案位置 | 說明 |
|------|-----------|--------|------|
| 啟動驗布系統 | **run** | 23_fabric inspsystem | 執行 `main.py` 啟動 Flask |
| 查看 Web 介面 | **verify** | 23_fabric inspsystem | 開啟 `fabric inspsystem.html` |

---

## 各Skill詳細說明

### 🎯 常用 Skill

#### **run** - 執行應用程式
- **用途**: 啟動並運行專案、腳本或應用程式
- **使用場景**: 
  - 執行 PowerShell 自動化腳本
  - 執行 Python 分析腳本
  - 啟動 Flask 後端服務
  - 執行數據導入程序
- **常見命令**:
  ```bash
  # PowerShell 腳本
  ./週報自動.ps1
  ./run_weekly_full.ps1
  
  # Python 腳本
  python qc_kpi_final.py
  python 比對.py
  
  # 啟動服務
  python main.py
  ```

#### **verify** - 驗證代碼或功能
- **用途**: 檢查代碼變更、測試功能、驗證輸出
- **使用場景**:
  - 驗證 HTML 儀表板生成正確
  - 檢查 KPI 報告計算無誤
  - 驗證分析結果輸出格式
  - 測試 Web 系統功能
- **常見命令**:
  ```bash
  # 檢查 HTML 輸出
  預覽 品管儀表板_A.html
  
  # 驗證 Excel 導出
  檢查 比對結果.html
  
  # 測試 API 回應
  驗證 Flask 後端 /api/... 端點
  ```

#### **engineering:standup** - 工程站會與狀態檢查
- **用途**: 快速概覽系統健康狀況、進度檢查
- **使用場景**:
  - 檢查下載任務狀態
  - 驗證儀表板更新
  - 查詢待處理資料數量
  - 確認導入進度
- **命令範例**:
  ```bash
  # 檢查各系統狀態
  檢查下載進度
  驗證儀表板更新
  列出待處理項目
  ```

#### **schedule** - 定時任務排程
- **用途**: 建立定期執行的自動化任務
- **使用場景**:
  - 每週一自動發送統計郵件
  - 每日自動檢查與處理
  - 定時下載報表
  - 定期更新儀表板
- **排程設定**:
  ```bash
  # 每週一上午 9 點發送郵件
  schedule weekly Monday 09:00 send_weekly_email.ps1
  
  # 每日下午 2 點檢查更新
  schedule daily 14:00 check_dashboard_update.ps1
  ```

---

## 常見工作場景

### 📅 週報流程
```
1. 說「週報」→ 執行 run 技能
   ↓
2. 系統自動:
   - 下載最新驗布報表
   - 進行格式轉換
   - 更新 insprecord_QF
   - 生成統計圖表
   - 更新儀表板
   ↓
3. 說「檢查週報」→ 執行 verify 技能
   ↓
4. 系統自動:
   - 驗證檔案完整性
   - 檢查數據準確性
   - 驗證儀表板更新
   ↓
5. 確認無誤後,說「發送」→ 執行 schedule 技能
```

### 📊 月報流程
```
1. 說「月報」→ 執行 run 技能
   ↓
2. 同週報流程,但使用月度腳本
   ↓
3. 執行 verify 驗證月度統計
   ↓
4. 執行 schedule 設定定時發送
```

### 🔍 品質分析流程
```
1. 說「分析」→ 列表選擇分析類型
   ↓
2. 年度比對分析:
   - 執行 run 技能 → 比對.py
   - 執行 verify 技能 → 檢查 HTML 結果
   ↓
3. 個別供應商分析:
   - 執行 run 技能 → analyze_supplier.py <廠商名>
   - 執行 verify 技能 → 驗證統計結果
```

### 📈 績效評估流程
```
1. 說「執行績效評估」→ 執行 run 技能
   ↓
2. 系統自動:
   - 執行 qc_kpi_final.py
   - 計算 23 位 QC 人員評分
   - 生成績效報告
   ↓
3. 說「驗證績效」→ 執行 verify 技能
   ↓
4. 檢查 qc_kpi_report.html 報告
```

---

## 快速參考

### 🚀 一鍵快速命令

| 需求 | 命令 | 對應檔案 |
|------|------|--------|
| 週報完整流程 | `/run` (選擇週報自動) | 03_download rawdata |
| 月報完整流程 | `/run` (選擇月報自動) | 03_download rawdata |
| 品管儀表板更新 | `/run` (選擇相關腳本) | 15_dashboard |
| 驗布比對分析 | `/run` (選擇比對.py) | 13_比對前後驗布記錄 |
| QC績效計算 | `python qc_kpi_final.py` | 14_qc performance |
| 週度導入 | `./run_weekly_full.ps1` | 12_inspdata_weekly |
| 月度導入 | `./run_monthly_full.ps1 -month "YYMM"` | 11_inspdata_monthly |

### 📝 Skill 優先級建議

**高頻使用** (每週 3+ 次):
- ✅ `/run` - 執行自動化腳本
- ✅ `/verify` - 驗證輸出結果
- ✅ `/engineering:standup` - 檢查狀態

**定期使用** (每週 1-2 次):
- ✅ `/schedule` - 排程任務
- ✅ `/code-review` - 檢查腳本品質

**按需使用** (每月 1-2 次):
- ✅ `/engineering:debug` - 調試問題
- ✅ `/testing-strategy` - 驗證測試覆蓋

---

## 📞 故障排除

| 問題 | 對應 Skill | 建議動作 |
|------|-----------|--------|
| 腳本執行失敗 | `/engineering:debug` | 檢查日誌,分析錯誤 |
| 輸出結果異常 | `/verify` | 驗證數據來源和計算邏輯 |
| 自動化卡住 | `/engineering:standup` | 檢查各系統狀態 |
| 需要優化效能 | `/engineering:tech-debt` | 識別瓶頸並優化 |
| 記錄不完整 | `/code-review` | 審視代碼品質 |

---

**📌 備註**: 本指南與記憶系統配合使用,以確保工作流程的一致性和效率最大化。

