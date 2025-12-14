# -*- coding: utf-8 -*-
"""
Tooling Manager App (نسخه مدیر قالبسازی)
- ذخیره محلی در SQLite
- ارسال درخواست به Master با فایل JSON در پوشه Outbox
- پایش پوشه Inbox برای دریافت پاسخ از Master
- UI ساده با Tkinter (فارسی)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sqlite3
import os
import json
from datetime import datetime
import threading
import time
import traceback
import logging
import csv

# ----------------------------
# تنظیمات پیش‌فرض (قابل تغییر)
# ----------------------------
DB_FILE = "tooling_manager.db"
LOG_FILE = "tooling_manager_errors.log"
OUTBOX_FOLDER = os.path.abspath("outbox")   # پوشه پیش‌فرض خروجی به Master
INBOX_FOLDER = os.path.abspath("inbox")     # پوشه پیش‌فرض دریافت از Master
POLL_INTERVAL_SEC = 5                       # هر چند ثانیه inbox چک شود

# Logging
logging.basicConfig(filename=LOG_FILE, level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# ----------------------------
# دیتابیس ساده SQLite
# ----------------------------
def init_db():
    con = sqlite3.connect(DB_FILE)
    cur = con.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS requests (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT,
        type TEXT,
        priority TEXT,
        creator TEXT,
        created_at TEXT,
        status TEXT,
        attachment TEXT,
        notes TEXT,
        sent_to_master INTEGER DEFAULT 0,
        master_response TEXT
    )
    """)
    con.commit()
    con.close()

def db_insert_request(title, typ, priority, creator, notes, attachment):
    con = sqlite3.connect(DB_FILE)
    cur = con.cursor()
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute("""
        INSERT INTO requests (title,type,priority,creator,created_at,status,attachment,notes,sent_to_master)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, 0)
    """, (title, typ, priority, creator, created_at, 'draft', attachment, notes))
    con.commit()
    rid = cur.lastrowid
    con.close()
    return rid

def db_get_all_requests():
    con = sqlite3.connect(DB_FILE)
    cur = con.cursor()
    cur.execute("SELECT id,title,type,priority,creator,created_at,status,attachment,sent_to_master,master_response FROM requests ORDER BY id DESC")
    rows = cur.fetchall()
    con.close()
    return rows

def db_update_status(rid, status, master_response=None, sent_to_master=None):
    con = sqlite3.connect(DB_FILE)
    cur = con.cursor()
    params = []
    sql = "UPDATE requests SET status = ? "
    params.append(status)
    if master_response is not None:
        sql += ", master_response = ? "
        params.append(master_response)
    if sent_to_master is not None:
        sql += ", sent_to_master = ? "
        params.append(int(sent_to_master))
    sql += " WHERE id = ?"
    params.append(rid)
    cur.execute(sql, tuple(params))
    con.commit()
    con.close()

def db_get_request(rid):
    con = sqlite3.connect(DB_FILE)
    cur = con.cursor()
    cur.execute("SELECT id,title,type,priority,creator,created_at,status,attachment,notes,sent_to_master,master_response FROM requests WHERE id=?", (rid,))
    row = cur.fetchone()
    con.close()
    return row

# ----------------------------
# کمکی: اطمینان از پوشه‌ها
# ----------------------------
def ensure_folders():
    os.makedirs(OUTBOX_FOLDER, exist_ok=True)
    os.makedirs(INBOX_FOLDER, exist_ok=True)

# ----------------------------
# تولید فایل JSON برای ارسال به Master
# ----------------------------
def generate_request_json(rid):
    row = db_get_request(rid)
    if not row:
        return None
    keys = ["id","title","type","priority","creator","created_at","status","attachment","notes","sent_to_master","master_response"]
    data = dict(zip(keys, row))
    # update status locally to 'sent'
    data['sent_at'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return data

def write_outbox_file(data):
    ts = datetime.now().strftime("%Y%m%d%H%M%S")
    fname = f"request_{data['id']}_{ts}.json"
    path = os.path.join(OUTBOX_FOLDER, fname)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return path

# ----------------------------
# پردازش فایل‌های ورودی از Master
# ----------------------------
def process_inbox_file(path):
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        # انتظار: data شامل id و status و optional response message
        rid = data.get('id')
        status = data.get('status')
        response = data.get('response', '')
        if rid:
            # update DB
            db_update_status(rid, status, master_response=response)
            logging.info(f"Processed inbox file for request {rid}: status={status}")
        # بعد از پردازش، می‌توان فایل را به پوشه processed منتقل کرد
        processed_dir = os.path.join(INBOX_FOLDER, "processed")
        os.makedirs(processed_dir, exist_ok=True)
        os.rename(path, os.path.join(processed_dir, os.path.basename(path)))
    except Exception as e:
        logging.error(traceback.format_exc())

# ----------------------------
# Thread یا Polling برای بررسی inbox
# ----------------------------
class InboxWatcher(threading.Thread):
    def __init__(self, ui):
        super().__init__(daemon=True)
        self.ui = ui
        self._stop = threading.Event()

    def run(self):
        while not self._stop.is_set():
            try:
                files = [f for f in os.listdir(INBOX_FOLDER) if f.lower().endswith('.json')]
                for f in files:
                    p = os.path.join(INBOX_FOLDER, f)
                    try:
                        process_inbox_file(p)
                        # پس از پردازش، به UI اطلاع بده تا لیست بروز شود
                        self.ui.safe_refresh()
                    except Exception:
                        logging.error(traceback.format_exc())
                time.sleep(POLL_INTERVAL_SEC)
            except Exception:
                logging.error(traceback.format_exc())
                time.sleep(POLL_INTERVAL_SEC)

    def stop(self):
        self._stop.set()

# ----------------------------
# UI اصلی
# ----------------------------
class ToolingManagerUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Tooling Manager - مدیر قالبسازی")
        self.root.geometry("1100x650")
        ensure_folders()
        init_db()
        self.watcher = None

        # بالای UI: فرم ساخت درخواست
        frm = ttk.LabelFrame(root, text="ایجاد درخواست جدید")
        frm.pack(fill='x', padx=8, pady=6)

        ttk.Label(frm, text="عنوان:").grid(row=0, column=0, sticky='w', padx=4, pady=4)
        self.title_ent = ttk.Entry(frm, width=60)
        self.title_ent.grid(row=0, column=1, columnspan=3, sticky='w', padx=4, pady=4)

        ttk.Label(frm, text="نوع:").grid(row=1, column=0, sticky='w', padx=4, pady=4)
        self.type_cb = ttk.Combobox(frm, values=["تعمیر قالب","ساخت جدید","اصلاح طراحی","بازسازی"], state='readonly', width=30)
        self.type_cb.grid(row=1, column=1, sticky='w', padx=4, pady=4)
        self.type_cb.set("تعمیر قالب")

        ttk.Label(frm, text="اولویت:").grid(row=1, column=2, sticky='w', padx=4, pady=4)
        self.priority_cb = ttk.Combobox(frm, values=["فوری","عادی","کم‌اولویت"], state='readonly', width=20)
        self.priority_cb.grid(row=1, column=3, sticky='w', padx=4, pady=4)
        self.priority_cb.set("عادی")

        ttk.Label(frm, text="ایجادکننده:").grid(row=0, column=4, sticky='w', padx=4, pady=4)
        self.creator_ent = ttk.Entry(frm, width=25)
        self.creator_ent.grid(row=0, column=5, sticky='w', padx=4, pady=4)
        self.creator_ent.insert(0, "مدیر قالبسازی")

        ttk.Label(frm, text="ضمیمه:").grid(row=2, column=0, sticky='w', padx=4, pady=4)
        self.attach_ent = ttk.Entry(frm, width=60)
        self.attach_ent.grid(row=2, column=1, columnspan=3, sticky='w', padx=4, pady=4)
        ttk.Button(frm, text="انتخاب فایل", command=self.pick_attachment).grid(row=2, column=4, padx=4, pady=4)

        ttk.Label(frm, text="توضیحات:").grid(row=3, column=0, sticky='nw', padx=4, pady=4)
        self.notes_txt = tk.Text(frm, height=4, width=80)
        self.notes_txt.grid(row=3, column=1, columnspan=5, padx=4, pady=4)

        ttk.Button(frm, text="ذخیره محلی", command=self.save_local).grid(row=4, column=1, padx=6, pady=6, sticky='w')
        ttk.Button(frm, text="ارسال به Master (تولید فایل Outbox)", command=self.send_to_master).grid(row=4, column=2, padx=6, pady=6, sticky='w')
        ttk.Button(frm, text="پاکسازی فرم", command=self.clear_form).grid(row=4, column=3, padx=6, pady=6, sticky='w')

        for i in range(6):
            frm.grid_columnconfigure(i, weight=1)

        # میانه: لیست درخواست‌ها
        list_frm = ttk.LabelFrame(root, text="درخواست‌ها")
        list_frm.pack(fill='both', expand=True, padx=8, pady=6)

        cols = ("id","title","type","priority","creator","created_at","status","attachment","sent_to_master")
        self.tree = ttk.Treeview(list_frm, columns=cols, show='headings')
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=110, anchor='center')
        self.tree.column("title", width=220)
        self.tree.pack(side='left', fill='both', expand=True)

        vsb = ttk.Scrollbar(list_frm, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscroll=vsb.set)
        vsb.pack(side='right', fill='y')

        # پایین: عملیات روی رکورد انتخاب‌شده
        ops = ttk.Frame(root)
        ops.pack(fill='x', padx=8, pady=6)
        ttk.Button(ops, text="باز کردن ضمیمه", command=self.open_attachment).pack(side='left', padx=6)
        ttk.Button(ops, text="مشاهده جزئیات", command=self.view_details).pack(side='left', padx=6)
        ttk.Button(ops, text="ارسال مجدد به Master", command=self.resend_selected).pack(side='left', padx=6)
        ttk.Button(ops, text="صادرات CSV", command=self.export_csv).pack(side='right', padx=6)
        ttk.Button(ops, text="راه‌اندازی/توقف مانیتور Inbox", command=self.toggle_watcher).pack(side='right', padx=6)

        self.status_label = ttk.Label(root, text="آماده", relief='sunken', anchor='w')
        self.status_label.pack(fill='x', padx=8, pady=(0,6))

        self.safe_refresh()

    # ---------- UI helpers ----------
    def pick_attachment(self):
        p = filedialog.askopenfilename(title="انتخاب فایل پیوست")
        if p:
            self.attach_ent.delete(0, tk.END)
            self.attach_ent.insert(0, p)

    def clear_form(self):
        self.title_ent.delete(0, tk.END)
        self.type_cb.set("تعمیر قالب")
        self.priority_cb.set("عادی")
        self.attach_ent.delete(0, tk.END)
        self.notes_txt.delete("1.0", tk.END)

    def save_local(self):
        title = self.title_ent.get().strip()
        if not title:
            messagebox.showwarning("ورود اطلاعات", "عنوان وارد نشده است")
            return
        typ = self.type_cb.get().strip()
        priority = self.priority_cb.get().strip()
        creator = self.creator_ent.get().strip() or "مدیر قالبسازی"
        notes = self.notes_txt.get("1.0", tk.END).strip()
        attach = self.attach_ent.get().strip()
        try:
            rid = db_insert_request(title, typ, priority, creator, notes, attach)
            self.safe_refresh()
            self.clear_form()
            self.set_status(f"درخواست ذخیره شد (id={rid})")
        except Exception as e:
            logging.error(traceback.format_exc())
            messagebox.showerror("خطا", f"خطا در ذخیره: {e}")

    def send_to_master(self):
        sel = self.tree.selection()
        if sel:
            # اگر رکورد انتخاب شده -> ارسال همون رکورد
            item = self.tree.item(sel[0])['values']
            rid = int(item[0])
            self._send_request_by_id(rid)
            return
        # در غیر اینصورت، ارسال فرم فعلی اگر پر شده
        title = self.title_ent.get().strip()
        if not title:
            messagebox.showwarning("ارسال", "هیچ رکوردی انتخاب نشده و فرم خالی است")
            return
        # ابتدا ذخیره محلی
        typ = self.type_cb.get().strip()
        priority = self.priority_cb.get().strip()
        creator = self.creator_ent.get().strip() or "مدیر قالبسازی"
        notes = self.notes_txt.get("1.0", tk.END).strip()
        attach = self.attach_ent.get().strip()
        try:
            rid = db_insert_request(title, typ, priority, creator, notes, attach)
            self._send_request_by_id(rid)
            self.clear_form()
        except Exception as e:
            logging.error(traceback.format_exc())
            messagebox.showerror("خطا", f"خطا در ذخیره/ارسال: {e}")

    def _send_request_by_id(self, rid):
        try:
            data = generate_request_json(rid)
            if not data:
                messagebox.showerror("خطا", "رکورد پیدا نشد")
                return
            # update local status to 'sent'
            db_update_status(rid, "sent_to_master", sent_to_master=1)
            # write JSON to outbox
            outpath = write_outbox_file(data)
            self.safe_refresh()
            self.set_status(f"درخواست {rid} به Master ارسال شد -> {outpath}")
            messagebox.showinfo("ارسال", f"درخواست {rid} به Master ارسال شد.")
        except Exception as e:
            logging.error(traceback.format_exc())
            messagebox.showerror("خطا", f"خطا در ارسال: {e}")

    def resend_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("ارسال مجدد", "یک رکورد را انتخاب کنید")
            return
        item = self.tree.item(sel[0])['values']
        rid = int(item[0])
        self._send_request_by_id(rid)

    def open_attachment(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("باز کردن ضمیمه", "یک رکورد را انتخاب کنید")
            return
        item = self.tree.item(sel[0])['values']
        attach = item[7]
        if attach and os.path.exists(attach):
            try:
                if os.name == "nt":
                    os.startfile(attach)
                else:
                    os.system(f'xdg-open "{attach}"')
            except Exception as e:
                messagebox.showerror("خطا", f"باز کردن فایل ممکن نشد: {e}")
        else:
            messagebox.showwarning("ضمیمه", "مسیر ضمیمه معتبر نیست یا فایل حذف شده است")

    def view_details(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("جزئیات", "ابتدا یک رکورد را انتخاب کنید")
            return
        item = self.tree.item(sel[0])['values']
        rid = int(item[0])
        row = db_get_request(rid)
        if not row:
            messagebox.showerror("خطا", "رکورد پیدا نشد")
            return
        # نمایش جزئیات
        txt = (
            f"شناسه: {row[0]}\n"
            f"عنوان: {row[1]}\n"
            f"نوع: {row[2]}\n"
            f"اولویت: {row[3]}\n"
            f"ایجادکننده: {row[4]}\n"
            f"زمان: {row[5]}\n"
            f"وضعیت: {row[6]}\n"
            f"ضمیمه: {row[7]}\n\n"
            f"توضیحات:\n{row[8]}\n\n"
            f"پاسخ Master:\n{row[10] if row[10] else ''}"
        )
        dlg = tk.Toplevel(self.root)
        dlg.title(f"جزئیات درخواست {rid}")
        txtw = tk.Text(dlg, wrap='word', width=80, height=25)
        txtw.pack(fill='both', expand=True, padx=8, pady=8)
        txtw.insert('1.0', txt)
        txtw.config(state='disabled')
        ttk.Button(dlg, text="بستن", command=dlg.destroy).pack(pady=6)

    def export_csv(self):
        rows = db_get_all_requests()
        if not rows:
            messagebox.showwarning("صادرات", "هیچ رکوردی برای صادرات وجود ندارد")
            return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files","*.csv")])
        if not path:
            return
        try:
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(["id","title","type","priority","creator","created_at","status","attachment","sent_to_master","master_response"])
                for r in rows:
                    writer.writerow(r)
            messagebox.showinfo("صادرات", f"صادرات انجام شد: {path}")
        except Exception as e:
            logging.error(traceback.format_exc())
            messagebox.showerror("خطا", f"خطا در صادرات: {e}")

    def toggle_watcher(self):
        if self.watcher and self.watcher.is_alive():
            self.watcher.stop()
            self.watcher = None
            self.set_status("Inbox watcher stopped")
            return
        # start watcher
        self.watcher = InboxWatcher(self)
        self.watcher.start()
        self.set_status("Inbox watcher started (polling)")

    # ---------- UI refresh ----------
    def safe_refresh(self):
        # call from any thread to schedule refresh in main thread
        self.root.after(0, self.refresh_tree)

    def refresh_tree(self):
        try:
            rows = db_get_all_requests()
            self.tree.delete(*self.tree.get_children())
            for row in rows:
                self.tree.insert('', tk.END, values=row)
        except Exception:
            logging.error(traceback.format_exc())

    def set_status(self, text):
        self.status_label.config(text=text)

# ----------------------------
# Entrypoint
# اضافه کردن import های جدید
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ----------------------------
# Thread برای مانیتور real-time با watchdog
# ----------------------------
class InboxWatchdogHandler(FileSystemEventHandler):
    def __init__(self, ui):
        super().__init__()
        self.ui = ui

    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith(".json"):
            try:
                process_inbox_file(event.src_path)
                self.ui.safe_refresh()
            except Exception:
                logging.error(traceback.format_exc())

    def on_modified(self, event):
        if not event.is_directory and event.src_path.lower().endswith(".json"):
            try:
                process_inbox_file(event.src_path)
                self.ui.safe_refresh()
            except Exception:
                logging.error(traceback.format_exc())

class InboxWatcher(threading.Thread):
    """جایگزین polling با watchdog real-time"""
    def __init__(self, ui):
        super().__init__(daemon=True)
        self.ui = ui
        self._stop = threading.Event()
        self.observer = None

    def run(self):
        try:
            event_handler = InboxWatchdogHandler(self.ui)
            self.observer = Observer()
            self.observer.schedule(event_handler, INBOX_FOLDER, recursive=False)
            self.observer.start()
            logging.info(f"Watchdog started on {INBOX_FOLDER}")
            while not self._stop.is_set():
                time.sleep(1)
        except Exception:
            logging.error(traceback.format_exc())
        finally:
            if self.observer:
                self.observer.stop()
                self.observer.join()
                logging.info("Watchdog stopped")

    def stop(self):
        self._stop.set()
        if self.observer:
            self.observer.stop()
            self.observer.join()

# ----------------------------
def main():
    ensure_folders()
    init_db()
    root = tk.Tk()
    app = ToolingManagerUI(root)
    root.mainloop()
    # on exit, stop watcher if running
    if app.watcher and app.watcher.is_alive():
        app.watcher.stop()

if __name__ == '__main__':
    main()
