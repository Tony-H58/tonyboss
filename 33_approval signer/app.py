# -*- coding: utf-8 -*-
"""單機版簽核工具（視窗版）。

流程：輸入帳密登入 → 列出待簽 → 勾選 → 確認視窗 → 送出 → 寫紀錄檔。
帳密由使用者每次輸入，只存在記憶體。加 --demo 可用假資料測試介面。
"""
import configparser
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox

import audit_log
from client import ApprovalClient, DemoClient, LoginError, NotImplementedYet

DEMO = "--demo" in sys.argv


def load_config():
    cp = configparser.ConfigParser()
    cp.read(os.path.join(audit_log.app_dir(), "config.ini"), encoding="utf-8")
    sec = cp["server"] if cp.has_section("server") else {}
    return (sec.get("base_url", "").strip(),
            str(sec.get("verify_ssl", "true")).lower() != "false")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("簽核工具" + ("（示範模式）" if DEMO else ""))
        self.geometry("760x460")
        self.client = None
        self.items = {}
        self.base_url, self.verify_ssl = load_config()
        self.show_login()

    # ---------- 登入 ----------
    def show_login(self):
        for w in self.winfo_children():
            w.destroy()
        f = ttk.Frame(self, padding=24)
        f.pack(expand=True)
        self.v_url = tk.StringVar(value=self.base_url)
        self.v_user = tk.StringVar(value=os.environ.get("USERNAME", ""))
        self.v_pw = tk.StringVar()
        rows = [("系統網址", self.v_url, None), ("帳號", self.v_user, None),
                ("密碼", self.v_pw, "*")]
        for i, (label, var, show) in enumerate(rows):
            ttk.Label(f, text=label).grid(row=i, column=0, sticky="e", pady=4)
            e = ttk.Entry(f, textvariable=var, width=40, show=show or "")
            e.grid(row=i, column=1, pady=4, padx=6)
            if label == "密碼":
                e.focus_set()
                e.bind("<Return>", lambda _e: self.do_login())
        ttk.Button(f, text="登入", command=self.do_login).grid(
            row=3, column=1, sticky="e", pady=10)
        ttk.Label(f, foreground="#666",
                  text="密碼只留在記憶體，不會存檔；登入失敗只試一次。").grid(
            row=4, column=0, columnspan=2)

    def do_login(self):
        user, pw = self.v_user.get().strip(), self.v_pw.get()
        self.v_pw.set("")
        if not DEMO and not (self.v_url.get().strip() and user and pw):
            messagebox.showwarning("登入", "請輸入網址、帳號與密碼")
            return
        client = DemoClient(user or "demo") if DEMO else ApprovalClient(
            self.v_url.get().strip(), user, pw, self.verify_ssl)
        pw = None
        try:
            self.config(cursor="watch"); self.update_idletasks()
            client.login()
        except LoginError as e:
            messagebox.showerror("登入失敗", str(e))
            return
        finally:
            self.config(cursor="")
        self.client = client
        self.show_main()
        self.refresh()

    # ---------- 主畫面 ----------
    def show_main(self):
        for w in self.winfo_children():
            w.destroy()
        top = ttk.Frame(self, padding=8); top.pack(fill="x")
        ttk.Label(top, text=f"登入帳號：{self.client.username}").pack(side="left")
        ttk.Button(top, text="登出", command=self.logout).pack(side="right")
        ttk.Button(top, text="重新整理", command=self.refresh).pack(side="right", padx=6)

        self.tree = ttk.Treeview(self, columns=("id", "title", "detail"),
                                 show="headings", selectmode="extended")
        for c, t, w in (("id", "單號", 120), ("title", "標題", 220), ("detail", "內容", 380)):
            self.tree.heading(c, text=t); self.tree.column(c, width=w)
        self.tree.pack(fill="both", expand=True, padx=8)

        bot = ttk.Frame(self, padding=8); bot.pack(fill="x")
        ttk.Label(bot, text="意見").pack(side="left")
        self.v_comment = tk.StringVar()
        ttk.Entry(bot, textvariable=self.v_comment, width=40).pack(side="left", padx=6)
        ttk.Button(bot, text="退回", command=lambda: self.sign_selected("reject")).pack(side="right")
        ttk.Button(bot, text="核准", command=lambda: self.sign_selected("approve")).pack(side="right", padx=6)
        self.status = ttk.Label(self, text="", padding=(8, 0, 8, 6)); self.status.pack(fill="x")

    def logout(self):
        self.client = None
        self.show_login()

    def refresh(self):
        try:
            items = self.client.list_pending()
        except NotImplementedYet as e:
            messagebox.showinfo("尚未完成", str(e)); return
        except Exception as e:
            messagebox.showerror("讀取失敗", str(e)); return
        self.items = {i.item_id: i for i in items}
        self.tree.delete(*self.tree.get_children())
        for i in items:
            self.tree.insert("", "end", iid=i.item_id, values=(i.item_id, i.title, i.detail))
        self.status.config(text=f"待簽 {len(items)} 筆")

    # ---------- 簽核 ----------
    def sign_selected(self, action):
        ids = self.tree.selection()
        if not ids:
            messagebox.showinfo("簽核", "請先選取單據（可用 Ctrl／Shift 多選）"); return
        name = "核准" if action == "approve" else "退回"
        lines = "\n".join(f"{i}  {self.items[i].title}" for i in ids[:15])
        more = f"\n…共 {len(ids)} 筆" if len(ids) > 15 else ""
        if not messagebox.askyesno("確認送出", f"即將【{name}】以下 {len(ids)} 筆：\n\n{lines}{more}\n\n確定送出？"):
            return
        ok = fail = 0
        for i in ids:
            item = self.items[i]
            try:
                self.client.sign(item, action, self.v_comment.get())
                audit_log.write(self.client.username, item, name, self.v_comment.get(), True)
                ok += 1
            except NotImplementedYet as e:
                messagebox.showinfo("尚未完成", str(e)); return
            except Exception as e:
                audit_log.write(self.client.username, item, name, self.v_comment.get(), False, str(e))
                fail += 1
        messagebox.showinfo("結果", f"成功 {ok} 筆，失敗 {fail} 筆\n紀錄：{audit_log.log_path()}")
        self.refresh()


if __name__ == "__main__":
    App().mainloop()
