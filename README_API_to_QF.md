# API → QF JSON 直接转换工作流

## 🎯 工作流

```
1. 下载 rawdata
   python auto_download_http.py
   或手动从浏览器下载 rawdata
   ↓
   文件保存到: ~/Downloads/

2. 直接转成 QF JSON
   python convert_rawdata_to_qf.py
   ↓
   生成: ./11_inspdata_monthly/insprecord_QF/insprecord_QF_YYMMDD-YYMMDD.json

3. Python 分析（自动读 JSON）
   python qc_kpi_final.py
   ↓
   输出: HTML 报表 + 数据分析
```

---

## 📌 关键脚本说明

### 1. convert_rawdata_to_qf.py（新增）

**功能**: 
- 从 Downloads 读取下载的 rawdata
- 应用列映射和值映射
- 生成标准格式 JSON

**使用**:
```bash
python convert_rawdata_to_qf.py
```

**输出**:
```
./11_inspdata_monthly/insprecord_QF/insprecord_QF_260102-260531.json
```

**特点**:
- ✅ 自动检测最新的 rawdata 文件
- ✅ 自动生成日期范围
- ✅ 输出 JSON 格式
- ✅ 跳过 convert_rawdata.ps1 和 import_monthly.ps1

---

### 2. qc_kpi_final.py（改造）

**改进**:
- ✅ 优先读 JSON 文件
- ✅ 如果没有 JSON，自动降级到 Excel
- ✅ 支持 `use_json` 参数

**使用**:
```python
# 自动读 JSON（如果有）
df = load_data('26-1~26-6')

# 强制读 JSON
df = load_data('26-1~26-6', use_json=True)

# 强制读 Excel
df = load_data('26-1~26-6', use_json=False)
```

---

## 🚀 快速开始

### Step 1: 下载 rawdata
```bash
# 方式 A: 自动下载（需要 NTLM 认证）
python 03_download\ rawdata/auto_download_http.py

# 方式 B: 手动下载
# 打开浏览器登录系统，导出 rawdata 到 ~/Downloads/
```

### Step 2: 转成 QF JSON
```bash
python convert_rawdata_to_qf.py

# 输出:
# 📖 读取名称映射规则...
# 📋 读取列映射规则...
# 🔍 查找下载的 rawdata 文件...
# 📥 读取 rawdata...
# 🔄 转换数据...
# 💾 导出 JSON...
# ✅ 完成！
```

### Step 3: 运行分析
```bash
# 方式 A: 生成 QC KPI 报表
python 14_qc\ performance/qc_kpi_final.py

# 方式 B: 其他分析脚本
# 都会自动读取 JSON 文件
```

---

## 📊 性能对比

| 步骤 | 原有流程 | 新流程 |
|------|--------|-------|
| 下载 | auto_download_http.py | 同上 |
| 转换 | convert_rawdata.ps1 (PS) | convert_rawdata_to_qf.py (Python) |
| 导入 | import_monthly.ps1 (PS) | 已集成 |
| **生成格式** | Excel (45-75列) | **JSON 混合** |
| **分析** | qc_kpi_final.py (读 Excel) | **qc_kpi_final.py (读 JSON)** |

**优势**:
- ✅ 少一步转换（no import_monthly.ps1）
- ✅ 直接输出 JSON（易于重复分析）
- ✅ Python 性能更好
- ✅ 代码更容易维护

---

## 🔧 配置文件

脚本依赖的配置文件:
```
03_download rawdata/00_plan_download rawdata.xlsx
  ├─ NameList (客户、供应商、工厂、QC 映射)
  └─ ChartTitle (列映射规则)
```

确保这些文件存在且正确。

---

## 📝 输出文件

### 生成的 JSON 结构
```json
{
  "metadata": {
    "source": "convert_rawdata_to_qf.py",
    "created_at": "2026-08-22T10:30:00",
    "period": "260102-260531",
    "rows": 26272,
    "columns": 101
  },
  "headers": ["前後端", "實際驗布日", "客戶", ...],
  "data": [
    {"前後端": "品管", "實際驗布日": "2026-01-15", ...},
    ...
  ]
}
```

### JSON 文件位置
```
./11_inspdata_monthly/insprecord_QF/
├── insprecord_QF_260102-260531.json  (新)
├── insprecord_QF_260102-260531.xlsx  (可选，原有)
└── ...
```

---

## ⚠️ 注意事项

1. **rawdata 位置**: 确保 rawdata 文件在 `~/Downloads/` 目录
2. **映射规则**: 确保 `00_plan_download rawdata.xlsx` 是最新的
3. **路径配置**: 脚本中的路径使用 Windows 格式，需按需调整
4. **字符编码**: JSON 使用 UTF-8 编码，确保中文正确

---

## 🐛 故障排除

### 找不到 rawdata 文件
```
❌ 找不到 rawdata 文件
   QC 文件: []
   Factory 文件: []
```

**解决**:
- 确认已下载 rawdata 到 `~/Downloads/`
- 文件名应包含 "驗布結論報表" 和 "後端驗布報表"
- 检查文件是否被其他程序占用

### JSON 读取失败
```
if not dfs:
    raise FileNotFoundError(f"No files found matching period")
```

**解决**:
- 确认 JSON 文件已生成
- 检查日期范围是否正确
- 用 `use_json=False` 强制读 Excel 验证

---

## 🎯 下一步

- [ ] 测试 convert_rawdata_to_qf.py 的正确性
- [ ] 验证 JSON 和 Excel 版本的数据一致性
- [ ] 集成自动化流程（定时运行）
- [ ] 考虑迁移到真正的 API（如有需求）

---

## 📚 相关文件

- `convert_rawdata_to_qf.py` - 转换脚本（新）
- `qc_kpi_final.py` - 分析脚本（改造）
- `03_download rawdata/00_plan_download rawdata.xlsx` - 配置文件
- `11_inspdata_monthly/insprecord_QF/` - 输出目录
