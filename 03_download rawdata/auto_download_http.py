# -*- coding: utf-8 -*-
"""
驗布報表自動下載腳本
用途：自動下載品管(QC)和工廠(Factory)驗布報表 Excel 檔案
認證：Windows NTLM + ASP.NET 表單登入
"""

import requests
import urllib3
import re
import os
import sys
import io
import urllib.parse
import concurrent.futures

# 確保 stdout 支援 UTF-8（PowerShell cp950 環境）
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except Exception:
    pass

try:
    from requests_ntlm import HttpNtlmAuth
except ImportError:
    print("ERROR: 需要安裝 requests_ntlm: pip install requests_ntlm")
    sys.exit(1)

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ── 設定 ──────────────────────────────────────────────────────────
USERNAME = 'tonyhuang'
PASSWORD = '@Mk13579'
DOWNLOADS = os.path.expanduser('~') + '\\Downloads'

QC_BASE  = 'https://nt-net2.makalot.com.tw/FabQC'
FAC_BASE = 'https://nt-net2.makalot.com.tw/FabricQuality'

# ── 工具函數 ──────────────────────────────────────────────────────
def parse_date(yymmdd):
    """將 YYMMDD (260604) 或 YYYYMMDD (20260604) 轉成 YYYY/MM/DD"""
    s = str(yymmdd).strip()
    if len(s) == 6:
        return f"20{s[0:2]}/{s[2:4]}/{s[4:6]}"
    elif len(s) == 8:
        return f"{s[0:4]}/{s[4:6]}/{s[6:8]}"
    else:
        raise ValueError(f"日期格式錯誤: {yymmdd}")

def save_file(response, default_name, downloads=DOWNLOADS):
    """從 response 的 Content-Disposition 取檔名並儲存"""
    cd = response.headers.get('Content-Disposition', '')
    fname_m = re.search(r"filename\*=UTF-8''([^\s;\"]+)", cd, re.IGNORECASE)
    if fname_m:
        fname = urllib.parse.unquote(fname_m.group(1))
    else:
        fname_m2 = re.search(r'filename="?([^";]+)"?', cd, re.IGNORECASE)
        fname = urllib.parse.unquote(fname_m2.group(1).strip()) if fname_m2 else default_name

    save_path = os.path.join(downloads, fname)
    content = b''.join(response.iter_content(8192))
    with open(save_path, 'wb') as f:
        f.write(content)
    size_kb = len(content) // 1024
    return save_path, fname, size_kb

# ── QC 品管系統下載 ───────────────────────────────────────────────
def download_qc(sd_fmt, ed_fmt):
    """
    下載品管驗布結論報表
    sd_fmt/ed_fmt: 'YYYY/MM/DD' 格式
    回傳: (save_path, filename) 或 raise Exception
    """
    print(f"  [QC] 連線 {QC_BASE} ...")
    s = requests.Session()
    s.verify = False
    s.auth = HttpNtlmAuth(USERNAME, PASSWORD)

    # Step 1: 取得首頁 ViewState
    resp = s.get(f'{QC_BASE}/', timeout=20)
    resp.raise_for_status()
    vs  = re.search(r'id="__VIEWSTATE"\s+value="([^"]+)"', resp.text).group(1)
    vsg = re.search(r'id="__VIEWSTATEGENERATOR"\s+value="([^"]+)"', resp.text).group(1)
    ev  = re.search(r'id="__EVENTVALIDATION"\s+value="([^"]+)"', resp.text).group(1)

    # Step 2: 表單登入
    s.post(f'{QC_BASE}/', data={
        '__VIEWSTATE': vs, '__VIEWSTATEGENERATOR': vsg, '__EVENTVALIDATION': ev,
        'Uid': USERNAME, 'Pwd': '123', 'Button1': '登入'
    }, timeout=20)

    # Step 3: 報表頁 ViewState
    resp3 = s.get(f'{QC_BASE}/ReportList.aspx?language=zh-TW', timeout=20)
    resp3.raise_for_status()
    vs2  = re.search(r'id="__VIEWSTATE"\s+value="([^"]+)"', resp3.text).group(1)
    vsg2 = re.search(r'id="__VIEWSTATEGENERATOR"\s+value="([^"]+)"', resp3.text).group(1)
    ev2  = re.search(r'id="__EVENTVALIDATION"\s+value="([^"]+)"', resp3.text).group(1)

    # Step 4: 點選「自訂日期範圍」radio
    resp_r = s.post(f'{QC_BASE}/ReportList.aspx?language=zh-TW', data={
        '__VIEWSTATE': vs2, '__VIEWSTATEGENERATOR': vsg2, '__EVENTVALIDATION': ev2,
        '__EVENTTARGET': 'ctl00$ContentPlaceHolder1$rdDateRange$2',
        '__EVENTARGUMENT': '',
        'ctl00$ContentPlaceHolder1$rdDateRange': '1',
    }, timeout=30)
    vs3  = re.search(r'id="__VIEWSTATE"\s+value="([^"]+)"', resp_r.text).group(1)
    ev3  = re.search(r'id="__EVENTVALIDATION"\s+value="([^"]+)"', resp_r.text).group(1)
    vsg3 = re.search(r'id="__VIEWSTATEGENERATOR"\s+value="([^"]+)"', resp_r.text).group(1)

    # Step 5: 匯出（最多 5 分鐘）
    print(f"  [QC] 匯出中 {sd_fmt}~{ed_fmt} ...")
    resp4 = s.post(f'{QC_BASE}/ReportList.aspx?language=zh-TW', data={
        '__VIEWSTATE': vs3, '__VIEWSTATEGENERATOR': vsg3, '__EVENTVALIDATION': ev3,
        '__EVENTTARGET': 'ctl00$ContentPlaceHolder1$lbtnExport',
        '__EVENTARGUMENT': '',
        'ctl00$ContentPlaceHolder1$rdDateRange': '1',
        'ctl00$ContentPlaceHolder1$txtInitDate': sd_fmt,
        'ctl00$ContentPlaceHolder1$txtEndDate':  ed_fmt,
        'ctl00$ContentPlaceHolder1$CheckBox1': 'on',
        'ctl00$ContentPlaceHolder1$hdPageIndex': '1',
    }, timeout=300, stream=True)

    ct = resp4.headers.get('Content-Type', '')
    cd4 = resp4.headers.get('Content-Disposition', '')
    # QC 伺服器回傳 text/html 但帶有 Content-Disposition filename（正常現象）
    if not cd4 and 'vnd.' not in ct.lower() and 'excel' not in ct.lower() and 'octet' not in ct.lower():
        raise Exception(f"QC 匯出失敗：Content-Type={ct}, 無 Content-Disposition")

    path, fname, kb = save_file(resp4, '驗布結論報表.xls')
    print(f"  [QC] ✓ {fname} ({kb} KB) → {path}")
    return path, fname

# ── 工廠系統下載 ──────────────────────────────────────────────────
def download_factory(sd_fmt, ed_fmt):
    """
    下載工廠後端驗布報表
    sd_fmt/ed_fmt: 'YYYY/MM/DD' 格式
    回傳: (save_path, filename) 或 raise Exception
    """
    print(f"  [Factory] 連線 {FAC_BASE} ...")
    s = requests.Session()
    s.verify = False
    s.auth = HttpNtlmAuth(USERNAME, PASSWORD)

    # Step 1: 登入
    resp = s.get(f'{FAC_BASE}/Login', timeout=20)
    resp.raise_for_status()
    vs  = re.search(r'id="__VIEWSTATE"\s+value="([^"]+)"', resp.text).group(1)
    vsg = re.search(r'id="__VIEWSTATEGENERATOR"\s+value="([^"]+)"', resp.text).group(1)
    ev_m = re.search(r'id="__EVENTVALIDATION"\s+value="([^"]+)"', resp.text)
    ev = ev_m.group(1) if ev_m else ''

    s.post(f'{FAC_BASE}/Login', data={
        '__VIEWSTATE': vs, '__VIEWSTATEGENERATOR': vsg, '__EVENTVALIDATION': ev,
        'txtAccount': USERNAME, 'txtPwd': PASSWORD, 'btnLogin': 'login'
    }, timeout=20)

    # Step 2: 載入主頁面（建立 session）
    s.get(f'{FAC_BASE}/WebPage/FabricResultAnalysis', timeout=20)

    # Step 3: 匯出（後端/FTY + 客戶分析）
    # dataSourceType=5@@FTY (後端), inspectionAnalysisType=客戶
    analysis = urllib.parse.quote('客戶', safe='')

    export_url = (
        f'{FAC_BASE}/WebPage/FabricResultAnalysis_Export.aspx'
        f'?sdate={urllib.parse.quote(sd_fmt)}'
        f'&edate={urllib.parse.quote(ed_fmt)}'
        f'&dataSourceType={urllib.parse.quote("5@@FTY", safe="")}'
        f'&country={urllib.parse.quote("-1@@ALL", safe="")}'
        f'&inspectionAnalysisType={analysis}'
        f'&yds=&c_Percentage='
        f'&inspectionPriority={urllib.parse.quote("-1@@ALL", safe="")}'
        f'&fty={urllib.parse.quote("-1@@ALL", safe="")}'
        f'&style=&c20=N&IsSalesMonthlyReport=N&SalesMonthlyReportVT='
    )

    print(f"  [Factory] 匯出中 {sd_fmt}~{ed_fmt} ...")
    resp4 = s.get(export_url,
                  headers={'Referer': f'{FAC_BASE}/WebPage/FabricResultAnalysis'},
                  timeout=300, stream=True)

    ct = resp4.headers.get('Content-Type', '')
    cd4 = resp4.headers.get('Content-Disposition', '')
    cl = int(resp4.headers.get('Content-Length', 0))
    if cl == 0 and 'text/html' in ct and not cd4:
        raise Exception(f"Factory 匯出失敗：Content-Length=0（可能無資料或參數錯誤）")

    path, fname, kb = save_file(resp4, '後端驗布報表.xls')
    print(f"  [Factory] ✓ {fname} ({kb} KB) → {path}")
    return path, fname

# ── 主程式 ────────────────────────────────────────────────────────
if __name__ == '__main__':
    """
    用法:
        python auto_download_http.py 260604 260610
        python auto_download_http.py 260601 260630
    """
    if len(sys.argv) < 3:
        print("用法: python auto_download_http.py <startDate_YYMMDD> <endDate_YYMMDD>")
        print("例如: python auto_download_http.py 260604 260610")
        sys.exit(1)

    try:
        sd_fmt = parse_date(sys.argv[1])
        ed_fmt = parse_date(sys.argv[2])
    except ValueError as e:
        print(f"ERROR: {e}")
        sys.exit(1)

    print(f"\n{'='*55}")
    print(f" 驗布報表自動下載")
    print(f" 日期範圍: {sd_fmt} ~ {ed_fmt}")
    print(f"{'='*55}")

    errors = []
    qc_path = fac_path = None

    # 平行下載 QC + Factory
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
        fut_qc  = executor.submit(download_qc,      sd_fmt, ed_fmt)
        fut_fac = executor.submit(download_factory, sd_fmt, ed_fmt)

        for fut, label in [(fut_qc, 'QC'), (fut_fac, 'Factory')]:
            try:
                path, fname = fut.result()
                if label == 'QC':
                    qc_path = path
                else:
                    fac_path = path
            except Exception as e:
                print(f"  [{label}] FAIL: {e}")
                errors.append(f"{label}: {e}")

    print(f"\n{'='*55}")
    if errors:
        print(f" ⚠ 部分失敗: {'; '.join(errors)}")
    else:
        print(f" ✓ 全部下載完成")
    print(f"{'='*55}\n")

    if errors:
        sys.exit(1)
    sys.exit(0)
