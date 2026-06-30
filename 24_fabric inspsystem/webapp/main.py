"""
FabQC Web System - 驗布員現場輸入系統
Mock DB: 驗布結論報表_*.xls (openpyxl)
"""
import json
import os
import glob
import shutil
from datetime import datetime
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import openpyxl
import xlrd

app = FastAPI(title="FabQC Web System")

BASE_DIR = Path(__file__).parent
EXCEL_DIR = BASE_DIR.parent  # E:\88. Claude\31_fabric inspsystem\
UPLOAD_DIR = BASE_DIR / "uploads"
SUBMIT_FILE = BASE_DIR / "submitted_records.json"  # 已回傳的驗布記錄

app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")


# ── 找最新的驗布結論報表 ──────────────────────────────
def get_excel_path():
    files = sorted(glob.glob(str(EXCEL_DIR / "驗布結論報表*.xls*")), reverse=True)
    if not files:
        raise FileNotFoundError("找不到驗布結論報表")
    return files[0]


# ── 讀取 Excel 欄位定義 ───────────────────────────────
COLUMNS = {
    "id": 1, "apply_date": 2, "applicant": 3, "customer": 4,
    "purchaser": 5, "dept": 6, "style_no": 7, "po_no": 8,
    "order_type": 9, "priority_a": 10, "supplier": 11,
    "insp_region": 12, "garment_region": 13,
    "plan_insp_date": 14, "actual_insp_date": 15,
    "composition": 16, "knit_woven": 17, "fabric_type": 18,
    "dyeing": 19, "print": 20, "finishing": 21,
    "width_inch": 22, "weight_gm": 23,
    "order_qty": 24, "order_unit": 25,
    "inspector": 26, "color": 27, "color_result": 28,
    "bulk_qty": 29, "bulk_yds": 30,
    "insp_qty_yds": 31, "avg_pts": 32,
    "a_qty": 33, "b_qty": 34, "c_qty": 35,
    "a_pct": 36, "b_pct": 37, "c_pct": 38,
    "loss_pct": 39,
    "defect_fabric": 40, "defect_appearance": 41,
    "defect_hand": 42, "defect_skew": 43,
    "defect_color": 44, "defect_other": 45, "defect_spec": 46,
    "disposition": 47, "partial_lot": 48,
    "c20_status": 49, "c20_inspector_opinion": 50,
    "c20_ship_reason": 51, "c20_solution": 52, "c20_qc_suggest": 53,
    "has_std_sample": 54, "result_sent": 55,
    "qc_mgr_opinion": 56, "remark": 57, "mail_desc": 58,
    "plan_export_date": 59, "cut_date": 60, "garment_export_date": 61,
    "c20_vm_opinion": 62, "approval_status": 63,
}
COL_REV = {v: k for k, v in COLUMNS.items()}


def read_excel_rows():
    path = get_excel_path()
    rows = []
    if path.endswith(".xls"):
        wb = xlrd.open_workbook(path)
        ws = wb.sheet_by_index(0)
        for r_idx in range(1, ws.nrows):  # skip header row
            r = ws.row_values(r_idx)
            if not r[0]:
                continue
            row = {}
            for i, val in enumerate(r):
                col_name = COL_REV.get(i + 1)
                if col_name:
                    # xlrd 日期型態轉字串
                    if ws.cell_type(r_idx, i) == xlrd.XL_CELL_DATE:
                        try:
                            dt = xlrd.xldate_as_datetime(val, wb.datemode)
                            val = dt.strftime("%Y/%m/%d")
                        except:
                            pass
                    row[col_name] = str(val).strip() if val is not None else ""
                    # 清除 xlrd 對數字的 .0 尾綴
                    if row[col_name].endswith(".0"):
                        row[col_name] = row[col_name][:-2]
            rows.append(row)
    else:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r[0] is None:
                continue
            row = {}
            for i, val in enumerate(r):
                col_name = COL_REV.get(i + 1)
                if col_name:
                    row[col_name] = str(val).strip() if val is not None else ""
            rows.append(row)
    return rows


# ── 載入已提交記錄 ─────────────────────────────────────
def load_submitted():
    if SUBMIT_FILE.exists():
        return json.loads(SUBMIT_FILE.read_text(encoding="utf-8"))
    return []


def save_submitted(records):
    SUBMIT_FILE.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")


# ══════════════════════════════════════════════════════
# API 路由
# ══════════════════════════════════════════════════════

@app.get("/", response_class=HTMLResponse)
async def index():
    html = (BASE_DIR / "static" / "index.html").read_text(encoding="utf-8")
    return HTMLResponse(html)


@app.get("/api/search")
async def search(q: str = "", field: str = "po_no"):
    """查詢訂單（PO 或 style）—— 模擬連線查詢"""
    if not q or len(q) < 3:
        raise HTTPException(400, "請輸入至少3個字元")
    rows = read_excel_rows()
    q_lower = q.lower()
    results = []
    seen = set()
    for row in rows:
        val = row.get(field, "").lower()
        if q_lower in val:
            key = (row["po_no"], row["style_no"], row["color"])
            if key not in seen:
                seen.add(key)
                results.append({
                    "id": row["id"],
                    "po_no": row["po_no"],
                    "style_no": row["style_no"],
                    "customer": row["customer"],
                    "supplier": row["supplier"],
                    "color": row["color"],
                    "order_qty": row["order_qty"],
                    "order_unit": row["order_unit"],
                    "composition": row["composition"],
                    "fabric_type": row["fabric_type"],
                    "knit_woven": row["knit_woven"],
                    "dyeing": row["dyeing"],
                    "width_inch": row["width_inch"],
                    "weight_gm": row["weight_gm"],
                    "plan_insp_date": row["plan_insp_date"],
                    "insp_region": row["insp_region"],
                    "purchaser": row["purchaser"],
                    "dept": row["dept"],
                })
    return {"count": len(results), "results": results[:50]}


@app.get("/api/order/{po_no}")
async def get_order_detail(po_no: str):
    """下載特定 PO 的所有顏色資訊"""
    rows = read_excel_rows()
    colors = [r for r in rows if r["po_no"].strip() == po_no.strip()]
    if not colors:
        raise HTTPException(404, f"找不到 PO: {po_no}")
    # 取第一筆的訂單資訊
    base = colors[0]
    return {
        "po_no": po_no,
        "style_no": base["style_no"],
        "customer": base["customer"],
        "supplier": base["supplier"],
        "purchaser": base["purchaser"],
        "dept": base["dept"],
        "composition": base["composition"],
        "fabric_type": base["fabric_type"],
        "knit_woven": base["knit_woven"],
        "dyeing": base["dyeing"],
        "print": base["print"],
        "finishing": base["finishing"],
        "width_inch": base["width_inch"],
        "weight_gm": base["weight_gm"],
        "order_type": base["order_type"],
        "insp_region": base["insp_region"],
        "garment_region": base["garment_region"],
        "plan_insp_date": base["plan_insp_date"],
        "colors": [
            {
                "color": r["color"],
                "order_qty": r["order_qty"],
                "order_unit": r["order_unit"],
                "id": r["id"],
            }
            for r in colors
        ],
    }


class InspectionRecord(BaseModel):
    po_no: str
    style_no: str
    customer: str
    supplier: str
    inspector: str
    actual_insp_date: str
    insp_region: str
    colors: list  # list of color inspection data
    remark: Optional[str] = ""
    submitted_at: Optional[str] = ""


@app.post("/api/submit")
async def submit_inspection(record: InspectionRecord):
    """驗布完成後回傳資料"""
    records = load_submitted()
    data = record.model_dump()
    data["submitted_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    data["status"] = "submitted"
    records.append(data)
    save_submitted(records)
    return {"ok": True, "message": f"驗布記錄已儲存，共 {len(data['colors'])} 個顏色"}


@app.get("/api/submitted")
async def get_submitted():
    """查詢已提交記錄"""
    records = load_submitted()
    return {"count": len(records), "records": records}


@app.post("/api/upload-photo")
async def upload_photo(file: UploadFile = File(...), po_no: str = Form(...), color: str = Form(...)):
    """上傳驗布照片"""
    ext = Path(file.filename).suffix
    filename = f"{po_no}_{color}_{datetime.now().strftime('%H%M%S')}{ext}".replace("/", "-").replace(" ", "_")
    dest = UPLOAD_DIR / filename
    with open(dest, "wb") as f:
        shutil.copyfileobj(file.file, f)
    return {"ok": True, "filename": filename, "url": f"/static/uploads/{filename}"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
