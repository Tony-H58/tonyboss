# -*- coding: utf-8 -*-
"""簽核系統連線層。

登入後的 session 只存在記憶體；帳號密碼不寫入任何檔案。
list_pending() / sign() 需要簽核頁的實際 HTML 才能實作，目前會丟 NotImplementedError。
"""
from dataclasses import dataclass, field


class LoginError(Exception):
    """帳號或密碼錯誤、或無法連線。"""


class NotImplementedYet(NotImplementedError):
    """尚未取得簽核頁面結構，功能待接上。"""


@dataclass
class Item:
    """一張待簽單據。"""
    item_id: str
    title: str
    detail: str = ""
    extra: dict = field(default_factory=dict)  # 簽核送出時需要的隱藏欄位等


class ApprovalClient:
    def __init__(self, base_url, username, password, verify_ssl=True):
        self.base_url = base_url.rstrip("/")
        self.username = username
        self._password = password
        self.verify_ssl = verify_ssl
        self.session = None

    def login(self):
        import requests
        try:
            from requests_ntlm import HttpNtlmAuth
        except ImportError as e:
            raise LoginError("缺少 requests_ntlm 套件") from e
        s = requests.Session()
        s.verify = self.verify_ssl
        s.auth = HttpNtlmAuth(self.username, self._password)
        try:
            r = s.get(self.base_url, timeout=20)
        except requests.RequestException as e:
            raise LoginError(f"無法連線：{e}") from e
        if r.status_code in (401, 403):
            raise LoginError("帳號或密碼錯誤（只嘗試一次，避免帳號被鎖）")
        if r.status_code >= 400:
            raise LoginError(f"伺服器回應 {r.status_code}")
        self.session = s
        self._password = None  # 登入完成後不再保留密碼

    def list_pending(self):
        raise NotImplementedYet("待簽清單：需要簽核頁 HTML 才能解析")

    def sign(self, item, action, comment=""):
        """action: 'approve' 或 'reject'。成功回傳 None，失敗丟例外。"""
        raise NotImplementedYet("送出簽核：需要簽核頁 HTML 才能實作")


class DemoClient(ApprovalClient):
    """示範用，不連網路，用來測試介面與紀錄檔。"""
    def __init__(self, username="demo"):
        super().__init__("demo://", username, None)

    def login(self):
        self.session = object()

    def list_pending(self):
        return [Item(f"DEMO-{i:03d}", f"示範單據 {i}", f"C20 需出貨原因 範例 {i}")
                for i in range(1, 6)]

    def sign(self, item, action, comment=""):
        if item.item_id.endswith("003"):
            raise RuntimeError("示範：模擬簽核失敗")
