# -*- coding: utf-8 -*-
"""簽核紀錄檔：只記時間、帳號、單號、動作、結果；不記密碼。"""
import csv, os, sys
from datetime import datetime

COLUMNS = ["時間", "帳號", "單號", "標題", "動作", "意見", "結果", "訊息"]


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def log_path():
    return os.path.join(app_dir(), "簽核紀錄.csv")


def write(user, item, action, comment, ok, message=""):
    path = log_path()
    new = not os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if new:
            w.writerow(COLUMNS)
        w.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), user,
                    item.item_id, item.title, action, comment,
                    "成功" if ok else "失敗", message])
