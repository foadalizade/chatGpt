# merged_report_app.py
# -*- coding: utf-8 -*-
"""
نسخه یکپارچه و تعمیرشده برنامه گزارش‌گیری قالبسازی
ویژگی‌ها:
- رابط کاربری اصلی با Tkinter
- امکان باز کردن پنجره‌های جداگانه برای: بارگذاری، فیلتر ساده، فیلتر پیشرفته، تحلیل، گزارش Power BI
- امکان راه‌اندازی پنجره PyQt (در صورت نصب PySide6) به عنوان جایگزین
- خواندن شیت‌ها و داده‌ها از اکسل، نرمالایز نوع تعمیر، فیلترها، گروه‌بندی، خروجی به Excel/CSV

نکات:
- برای اجرای اصلی کافی است: python merged_report_app.py
- وابستگی‌ها: pandas, openpyxl, tk (نصب شده با پایتون), optionally PySide6 for PyQt window
- تولید فایل Power BI: خروجی CSV تمیز برای ایمپورت در Power BI. خود فایل .pbix باید در Power BI Desktop ساخته شود.

توسعه‌دهنده: نسخه ادغام و بهینه‌سازی شده
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import sys
import traceback
import logging
from datetime import datetime

# Optional PyQt (PySide6) import guarded
try:
    from PySide6 import QtWidgets, QtCore
    PYQT_AVAILABLE = True
except Exception:
    PYQT_AVAILABLE = False

# Logging
logging.basicConfig(filename='merged_app_errors.log', level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# -----------------------------
# Utility functions
# -----------------------------

def safe_read_excel(path, sheet_name=None):
    """Read excel into pandas DataFrame robustly."""
    try:
        if sheet_name:
            df = pd.read_excel(path, sheet_name=sheet_name, engine='openpyxl')
        else:
            # read first sheet
            df = pd.read_excel(path, engine='openpyxl')
        # drop completely empty columns
        df = df.dropna(axis=1, how='all')
        return df
    except Exception as e:
        logging.error(f"Error reading excel {path} sheet {sheet_name}: {e}")
        raise


def find_column(columns, possible_names):
    for name in possible_names:
        for col in columns:
            try:
                if name.strip().lower() in str(col).strip().lower():
                    return col
            except Exception:
                continue
    return None


def normalize_repair_type(repair_type):
    try:
        if pd.isna(repair_type):
            return repair_type
        s = str(repair_type).strip()
        s_lower = s.lower()
        if 'قالب' in s_lower and 'تعمیر' in s_lower:
            return 'قالب تعمیری'
        if 'قطعه' in s_lower and 'تعمیر' in s_lower:
            return 'قطعه تعمیری'
        if 'دستگاه' in s_lower and 'تعمیر' in s_lower:
            return 'دستگاه تعمیری'
        if 'قالب' in s_lower:
            return 'قالب'
        if 'قطعه' in s_lower:
            return 'قطعه'
        if 'دستگاه' in s_lower:
            return 'دستگاه'
        if 'تعمیر' in s_lower:
            return 'تعمیری'
        return s
    except Exception:
        return str(repair_type)

# -----------------------------
# Core classes
# -----------------------------
class ExcelHandler:
    def __init__(self):
        self.file_path = None
        self.sheet_names = []
        self.df = None
        self.df_normalized = None
        self.cols = {}

    def load_file(self, path):
        if not path or not os.path.exists(path):
            raise FileNotFoundError('File not found')
        self.file_path = path
        try:
            # get sheet names via pandas
            xls = pd.ExcelFile(path, engine='openpyxl')
            self.sheet_names = xls.sheet_names
            return self.sheet_names
        except Exception:
            logging.error(traceback.format_exc())
            raise

    def load_sheet(self, sheet_name):
        if not self.file_path:
            raise RuntimeError('No file loaded')
        self.df = safe_read_excel(self.file_path, sheet_name)
        self._detect_columns()
        self._normalize()
        return self.df

    def _detect_columns(self):
        cols = list(self.df.columns)
        self.cols['repair'] = find_column(cols, ['نوع تعمیر', 'تعمیر', 'repair'])
        self.cols['part'] = find_column(cols, ['قالب', 'قطعه', 'دستگاه', 'part', 'device'])
        self.cols['date'] = find_column(cols, ['تاریخ', 'date'])
        self.cols['perf'] = find_column(cols, ['مقدار ساعت کار شده', 'ساعت', 'hour', 'time'])
        self.cols['req'] = find_column(cols, ['شماره نامه درخواست', 'شماره درخواست', 'request'])
        self.cols['code'] = find_column(cols, ['کد قالب', 'کد', 'code'])

    def _normalize(self):
        if self.df is None:
            return
        self.df_normalized = self.df.copy()
        if self.cols.get('repair') in self.df_normalized.columns:
            self.df_normalized[self.cols['repair']] = self.df_normalized[self.cols['repair']].apply(normalize_repair_type)

# -----------------------------
# GUI windows (Tkinter)
# -----------------------------
class MainAppTk:
    def __init__(self, root):
        self.root = root
        self.root.title('گزارش‌گیر قالبسازی - نسخه یکپارچه')
        self.root.geometry('1100x750')

        self.excel = ExcelHandler()

        # Menu
        menubar = tk.Menu(self.root)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label='باز کردن فایل اکسل', command=self.open_file_dialog)
        filemenu.add_command(label='خروج', command=self.root.quit)
        menubar.add_cascade(label='فایل', menu=filemenu)

        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label='باز کردن پنجره PyQt (در صورت نصب)', command=self.launch_pyqt_window)
        menubar.add_cascade(label='ابزارها', menu=tools_menu)

        self.root.config(menu=menubar)

        # Top frame: file & sheets
        top_frame = ttk.Frame(self.root, padding=8)
        top_frame.pack(fill='x')

        ttk.Label(top_frame, text='فایل:').grid(row=0, column=0, sticky='w')
        self.file_entry = ttk.Entry(top_frame, width=70)
        self.file_entry.grid(row=0, column=1, padx=5)
        ttk.Button(top_frame, text='انتخاب فایل', command=self.open_file_dialog).grid(row=0, column=2, padx=5)

        ttk.Label(top_frame, text='شیت:').grid(row=1, column=0, sticky='w')
        self.sheet_cb = ttk.Combobox(top_frame, state='readonly', width=40)
        self.sheet_cb.grid(row=1, column=1, sticky='w')
        ttk.Button(top_frame, text='بارگذاری شیت‌ها', command=self.load_sheets).grid(row=1, column=2, padx=5)

        # Buttons to open separate windows
        btn_frame = ttk.Frame(self.root, padding=8)
        btn_frame.pack(fill='x')
        ttk.Button(btn_frame, text='پنجره بارگذاری و نمایش', command=self.open_window_loader).pack(side='left', padx=6)
        ttk.Button(btn_frame, text='پنجره فیلتر ساده', command=self.open_window_simple_filter).pack(side='left', padx=6)
        ttk.Button(btn_frame, text='پنجره فیلتر پیشرفته', command=self.open_window_adv_filter).pack(side='left', padx=6)
        ttk.Button(btn_frame, text='پنجره تحلیل', command=self.open_window_analysis).pack(side='left', padx=6)
        ttk.Button(btn_frame, text='پنجره گزارش Power BI', command=self.open_window_powerbi).pack(side='left', padx=6)

        # Status bar
        self.status_var = tk.StringVar(value='آماده')
        status = ttk.Label(self.root, textvariable=self.status_var, relief='sunken', anchor='w')
        status.pack(side='bottom', fill='x')

        # Keep references to windows
        self.loader_win = None
        self.simple_filter_win = None
        self.adv_filter_win = None
        self.analysis_win = None
        self.powerbi_win = None

    # ---------------- Window launching ----------------
    def open_file_dialog(self):
        path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx *.xls')])
        if not path:
            return
        self.file_entry.delete(0, tk.END)
        self.file_entry.insert(0, path)
        self.status_var.set(f'فایل انتخاب شد: {os.path.basename(path)}')

    def load_sheets(self):
        path = self.file_entry.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror('خطا', 'لطفاً ابتدا یک فایل اکسل معتبر انتخاب کنید')
            return
        try:
            sheets = self.excel.load_file(path)
            self.sheet_cb['values'] = sheets
            if sheets:
                self.sheet_cb.set(sheets[0])
            self.status_var.set(f'{len(sheets)} شیت بارگذاری شد')
        except Exception as e:
            logging.error(traceback.format_exc())
            messagebox.showerror('خطا', f'خطا در بارگذاری شیت‌ها: {e}')

    def open_window_loader(self):
        if self.loader_win and tk.Toplevel.winfo_exists(self.loader_win):
            self.loader_win.lift()
            return
        self.loader_win = LoaderWindow(self)

    def open_window_simple_filter(self):
        if self.simple_filter_win and tk.Toplevel.winfo_exists(self.simple_filter_win):
            self.simple_filter_win.lift()
            return
        self.simple_filter_win = SimpleFilterWindow(self)

    def open_window_adv_filter(self):
        if self.adv_filter_win and tk.Toplevel.winfo_exists(self.adv_filter_win):
            self.adv_filter_win.lift()
            return
        self.adv_filter_win = AdvancedFilterWindow(self)

    def open_window_analysis(self):
        if self.analysis_win and tk.Toplevel.winfo_exists(self.analysis_win):
            self.analysis_win.lift()
            return
        self.analysis_win = AnalysisWindow(self)

    def open_window_powerbi(self):
        if self.powerbi_win and tk.Toplevel.winfo_exists(self.powerbi_win):
            self.powerbi_win.lift()
            return
        self.powerbi_win = PowerBIWindow(self)

    def launch_pyqt_window(self):
        if not PYQT_AVAILABLE:
            messagebox.showwarning('PyQt/PySide6 موجود نیست', 'برای باز شدن پنجره PyQt نیاز به نصب PySide6 دارید')
            return
        # Launch PyQt in a new process to avoid event loop conflicts
        try:
            import subprocess
            subprocess.Popen([sys.executable, __file__, '--pyqt'])
        except Exception as e:
            logging.error(traceback.format_exc())
            messagebox.showerror('خطا', f'خطا در راه‌اندازی PyQt: {e}')

# -----------------------------
# Loader Window
# -----------------------------
class LoaderWindow(tk.Toplevel):
    def __init__(self, app: MainAppTk):
        super().__init__(app.root)
        self.app = app
        self.title('بارگذاری و نمایش داده‌ها')
        self.geometry('900x600')

        frame = ttk.Frame(self, padding=8)
        frame.pack(fill='both', expand=True)

        # controls
        top = ttk.Frame(frame)
        top.pack(fill='x')
        ttk.Label(top, text='مسیر فایل:').grid(row=0, column=0, sticky='w')
        self.path_entry = ttk.Entry(top, width=70)
        self.path_entry.grid(row=0, column=1, padx=5)
        ttk.Button(top, text='انتخاب', command=self.browse).grid(row=0, column=2)

        ttk.Label(top, text='شیت:').grid(row=1, column=0, sticky='w')
        self.sheet_cb = ttk.Combobox(top, state='readonly', width=40)
        self.sheet_cb.grid(row=1, column=1, sticky='w')
        ttk.Button(top, text='بارگذاری شیت‌ها', command=self.load_sheets).grid(row=1, column=2)
        ttk.Button(top, text='بارگذاری داده‌ها', command=self.load_data).grid(row=1, column=3, padx=6)

        # treeview
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill='both', expand=True, pady=8)
        self.tree = ttk.Treeview(tree_frame, show='headings')
        vsb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        # export
        btns = ttk.Frame(frame)
        btns.pack(fill='x')
        ttk.Button(btns, text='صدور CSV', command=self.export_csv).pack(side='left', padx=4)
        ttk.Button(btns, text='صدور Excel', command=self.export_excel).pack(side='left', padx=4)

    def browse(self):
        p = filedialog.askopenfilename(filetypes=[('Excel', '*.xlsx *.xls')])
        if p:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, p)

    def load_sheets(self):
        p = self.path_entry.get().strip()
        if not p or not os.path.exists(p):
            messagebox.showerror('خطا', 'فایل معتبر انتخاب کنید')
            return
        try:
            sheets = self.app.excel.load_file(p)
            self.sheet_cb['values'] = sheets
            if sheets:
                self.sheet_cb.set(sheets[0])
            messagebox.showinfo('موفق', f'{len(sheets)} شیت بارگذاری شد')
        except Exception as e:
            logging.error(traceback.format_exc())
            messagebox.showerror('خطا', f'خطا: {e}')

    def load_data(self):
        p = self.path_entry.get().strip()
        sheet = self.sheet_cb.get().strip()
        if not p or not sheet:
            messagebox.showerror('خطا', 'فایل و شیت را مشخص کنید')
            return
        try:
            df = self.app.excel.load_sheet(sheet_name=sheet) if self.app.excel.file_path == p else None
            # if file differs load file first
            if df is None:
                self.app.excel.load_file(p)
                df = self.app.excel.load_sheet(sheet)
            self.populate_tree(df)
            self.app.status_var.set(f'{len(df)} رکورد بارگذاری شد')
        except Exception as e:
            logging.error(traceback.format_exc())
            messagebox.showerror('خطا', f'خطا در بارگذاری داده‌ها: {e}')

    def populate_tree(self, df: pd.DataFrame):
        for i in self.tree.get_children():
            self.tree.delete(i)
        cols = list(df.columns)
        self.tree['columns'] = cols
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120)
        for _, row in df.head(1000).iterrows():
            vals = [row[c] if pd.notna(row[c]) else '' for c in cols]
            self.tree.insert('', 'end', values=vals)

    def export_csv(self):
        try:
            df = self.app.excel.df
            if df is None:
                messagebox.showwarning('هشدار', 'داده‌ای بارگذاری نشده')
                return
            p = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV', '*.csv')])
            if p:
                df.to_csv(p, index=False, encoding='utf-8-sig')
                messagebox.showinfo('موفق', 'CSV ذخیره شد')
        except Exception as e:
            logging.error(traceback.format_exc())
            messagebox.showerror('خطا', f'خطا در ذخیره CSV: {e}')

    def export_excel(self):
        try:
            df = self.app.excel.df
            if df is None:
                messagebox.showwarning('هشدار', 'داده‌ای بارگذاری نشده')
                return
            p = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx')])
            if p:
                df.to_excel(p, index=False, engine='openpyxl')
                messagebox.showinfo('موفق', 'Excel ذخیره شد')
        except Exception as e:
            logging.error(traceback.format_exc())
            messagebox.showerror('خطا', f'خطا در ذخیره Excel: {e}')

# -----------------------------
# Simple Filter Window
# -----------------------------
class SimpleFilterWindow(tk.Toplevel):  # UPDATED grouping logic(tk.Toplevel):
    def __init__(self, app: MainAppTk):
        super().__init__(app.root)
        self.app = app
        self.title('فیلتر ساده')
        self.geometry('700x500')

        f = ttk.Frame(self, padding=8)
        f.pack(fill='both', expand=True)

        ttk.Label(f, text='نوع تعمیر:').grid(row=0, column=0, sticky='w')
        self.repair_cb = ttk.Combobox(f, state='readonly', width=40)
        self.repair_cb.grid(row=0, column=1, padx=6)

        ttk.Label(f, text='قالب/قطعه/دستگاه:').grid(row=1, column=0, sticky='w')
        self.part_cb = ttk.Combobox(f, state='readonly', width=40)
        self.part_cb.grid(row=1, column=1, padx=6)

        ttk.Button(f, text='بارگذاری گزینه‌ها', command=self.populate_options).grid(row=0, column=2, padx=6)
        ttk.Button(f, text='اعمال فیلتر', command=self.apply_filter).grid(row=2, column=1, pady=8)

        # tree
        self.tree = ttk.Treeview(f, show='headings')
        self.tree.grid(row=3, column=0, columnspan=3, sticky='nsew')
        f.rowconfigure(3, weight=1)
        f.columnconfigure(1, weight=1)

    def populate_options(self):
        dfn = self.app.excel.df_normalized
        df = self.app.excel.df
        if df is None:
            messagebox.showwarning('هشدار', 'ابتدا داده‌ها را در پنجره بارگذاری کنید')
            return
        repair_col = self.app.excel.cols.get('repair')
        part_col = self.app.excel.cols.get('part')
        if repair_col and repair_col in dfn.columns:
            vals = ['(همه)'] + sorted(dfn[repair_col].dropna().astype(str).unique().tolist())
            self.repair_cb['values'] = vals
            self.repair_cb.set(vals[0])
        if part_col and part_col in df.columns:
            vals2 = ['(همه)'] + sorted(df[part_col].dropna().astype(str).unique().tolist())
            self.part_cb['values'] = vals2
            self.part_cb.set(vals2[0])

    def apply_filter(self):
        df = self.app.excel.df
        dfn = self.app.excel.df_normalized
        if df is None:
            messagebox.showwarning('هشدار', 'ابتدا داده‌ها بارگذاری شود')
            return
        repair_sel = self.repair_cb.get()
        part_sel = self.part_cb.get()
        res = df.copy()
        if repair_sel and repair_sel != '(همه)':
            repair_col = self.app.excel.cols.get('repair')
            if repair_col and repair_col in dfn.columns:
                mask = dfn[repair_col].astype(str) == repair_sel
                res = res[mask]
        if part_sel and part_sel != '(همه)':
            part_col = self.app.excel.cols.get('part')
            if part_col and part_col in res.columns:
                res = res[res[part_col].astype(str) == part_sel]
        # --- NEW: group duplicates by key columns and sum hours
        rep_col = self.app.excel.cols.get('repair')
        part_col = self.app.excel.cols.get('part')
        code_col = self.app.excel.cols.get('code')
        req_col = self.app.excel.cols.get('req')
        perf_col = self.app.excel.cols.get('perf')

        grouping = []
        for c in [rep_col, part_col, code_col, req_col]:
            if c and c in res.columns:
                grouping.append(c)

        if perf_col and perf_col in res.columns and grouping:
            res[perf_col] = pd.to_numeric(res[perf_col], errors='coerce').fillna(0)
            res = res.groupby(grouping, as_index=False).agg({perf_col: 'sum'})

        self.populate_tree(res)
        self.app.status_var.set(f'فیلتر ساده اعمال شد - {len(res)} رکورد')

    def populate_tree(self, df):
        for i in self.tree.get_children():
            self.tree.delete(i)
        cols = list(df.columns)
        self.tree['columns'] = cols
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120)
        for _, row in df.head(1000).iterrows():
            vals = [row[c] if pd.notna(row[c]) else '' for c in cols]
            self.tree.insert('', 'end', values=vals)

# -----------------------------
# Advanced Filter Window
# -----------------------------
class AdvancedFilterWindow(tk.Toplevel):
    def __init__(self, app: MainAppTk):
        super().__init__(app.root)
        self.app = app
        self.title('فیلتر پیشرفته')
        self.geometry('800x600')

        f = ttk.Frame(self, padding=8)
        f.pack(fill='both', expand=True)

        ttk.Label(f, text='انواع تعمیر (چندتایی):').grid(row=0, column=0, sticky='w')
        self.lb = tk.Listbox(f, selectmode='multiple', height=6)
        self.lb.grid(row=0, column=1, sticky='ew')

        ttk.Label(f, text='بازه ساعت - از:').grid(row=1, column=0, sticky='w')
        self.hmin = ttk.Entry(f, width=10)
        self.hmin.grid(row=1, column=1, sticky='w')
        ttk.Label(f, text='تا:').grid(row=1, column=2, sticky='w')
        self.hmax = ttk.Entry(f, width=10)
        self.hmax.grid(row=1, column=3, sticky='w')

        ttk.Button(f, text='بارگذاری گزینه‌ها', command=self.populate_options).grid(row=0, column=2)
        ttk.Button(f, text='اعمال فیلتر پیشرفته', command=self.apply_adv_filter).grid(row=2, column=1, pady=6)
        ttk.Button(f, text='گروه‌بندی و جمع‌بندی', command=self.group_data).grid(row=2, column=2, pady=6)

        # tree
        self.tree = ttk.Treeview(f, show='headings')
        self.tree.grid(row=3, column=0, columnspan=4, sticky='nsew')
        f.rowconfigure(3, weight=1)
        f.columnconfigure(1, weight=1)

    def populate_options(self):
        dfn = self.app.excel.df_normalized
        if dfn is None:
            messagebox.showwarning('هشدار', 'ابتدا داده‌ها بارگذاری شود')
            return
        rep_col = self.app.excel.cols.get('repair')
        if rep_col and rep_col in dfn.columns:
            vals = sorted(dfn[rep_col].dropna().astype(str).unique().tolist())
            self.lb.delete(0, tk.END)
            for v in vals:
                self.lb.insert(tk.END, v)

    def apply_adv_filter(self):
        df = self.app.excel.df.copy()
        dfn = self.app.excel.df_normalized
        if df is None:
            messagebox.showwarning('هشدار', 'ابتدا داده‌ها بارگذاری شود')
            return
        sel = [self.lb.get(i) for i in self.lb.curselection()]
        if sel:
            rep_col = self.app.excel.cols.get('repair')
            mask = dfn[rep_col].astype(str).isin(sel)
            df = df[mask]
        # hours
        perf_col = self.app.excel.cols.get('perf')
        try:
            hmin = float(self.hmin.get()) if self.hmin.get().strip() else None
            hmax = float(self.hmax.get()) if self.hmax.get().strip() else None
        except ValueError:
            messagebox.showerror('خطا', 'مقادیر ساعت باید عددی باشند')
            return
        if perf_col in df.columns:
            df[perf_col] = pd.to_numeric(df[perf_col], errors='coerce')
            if hmin is not None:
                df = df[df[perf_col] >= hmin]
            if hmax is not None:
                df = df[df[perf_col] <= hmax]
        self.populate_tree(df)
        self.app.status_var.set(f'فیلتر پیشرفته اعمال شد - {len(df)} رکورد')

    def group_data(self):
        df = self.app.excel.df
        if df is None:
            messagebox.showwarning('هشدار', 'داده‌ای بارگذاری نشده')
            return
        grouping = []
        for key in ['part', 'code', 'req']:
            c = self.app.excel.cols.get(key)
            if c and c in df.columns:
                grouping.append(c)
        perf = self.app.excel.cols.get('perf')
        if not grouping or not perf:
            messagebox.showerror('خطا', 'ستون‌های لازم برای گروه‌بندی یافت نشد')
            return
        tmp = df.copy()
        tmp[perf] = pd.to_numeric(tmp[perf], errors='coerce').fillna(0)
        grouped = tmp.groupby(grouping, as_index=False).agg({perf: 'sum'}).sort_values(by=perf, ascending=False)
        self.populate_tree(grouped)
        self.app.status_var.set(f'گروه‌بندی انجام شد - {len(grouped)} گروه')

    def populate_tree(self, df):
        for i in self.tree.get_children():
            self.tree.delete(i)
        cols = list(df.columns)
        self.tree['columns'] = cols
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=140)
        for _, row in df.head(1000).iterrows():
            vals = [row[c] if pd.notna(row[c]) else '' for c in cols]
            self.tree.insert('', 'end', values=vals)

# -----------------------------
# Analysis Window
# -----------------------------
class AnalysisWindow(tk.Toplevel):
    def __init__(self, app: MainAppTk):
        super().__init__(app.root)
        self.app = app
        self.title('تجزیه و تحلیل')
        self.geometry('800x600')

        f = ttk.Frame(self, padding=8)
        f.pack(fill='both', expand=True)

        ttk.Button(f, text='آمار کلی', command=self.show_stats).pack(anchor='w')
        ttk.Button(f, text='Pivot ساده (گروه‌بندی)', command=self.show_pivot).pack(anchor='w', pady=4)

        self.txt = tk.Text(f)
        self.txt.pack(fill='both', expand=True)

    def show_stats(self):
        df = self.app.excel.df
        if df is None:
            messagebox.showwarning('هشدار', 'هیچ داده‌ای بارگذاری نشده')
            return
        info = []
        info.append(f"تعداد رکورد: {len(df)}")
        info.append(f"ستون‌ها: {', '.join(df.columns)}")
        perf = self.app.excel.cols.get('perf')
        if perf and perf in df.columns:
            s = pd.to_numeric(df[perf], errors='coerce')
            info.append(f"جمع ساعت: {s.sum():.2f}")
            info.append(f"میانگین ساعت: {s.mean():.2f}")
            info.append(f"مین و ماکس ساعت: {s.min():.2f} - {s.max():.2f}")
        self.txt.delete(1.0, tk.END)
        self.txt.insert(tk.END, "\n".join(info))

    def show_pivot(self):
        df = self.app.excel.df
        if df is None:
            messagebox.showwarning('هشدار', 'هیچ داده‌ای بارگذاری نشده')
            return
        grouping = []
        for key in ['part', 'code', 'req']:
            c = self.app.excel.cols.get(key)
            if c and c in df.columns:
                grouping.append(c)
        perf = self.app.excel.cols.get('perf')
        if not grouping or not perf:
            messagebox.showerror('خطا', 'ستون‌های لازم موجود نیست')
            return
        tmp = df.copy()
        tmp[perf] = pd.to_numeric(tmp[perf], errors='coerce').fillna(0)
        pv = tmp.groupby(grouping)[perf].sum().reset_index().sort_values(by=perf, ascending=False)
        self.txt.delete(1.0, tk.END)
        self.txt.insert(tk.END, pv.to_string(index=False))

# -----------------------------
# Power BI Window (prepares files)
# -----------------------------
class PowerBIWindow(tk.Toplevel):
    def __init__(self, app: MainAppTk):
        super().__init__(app.root)
        self.app = app
        self.title('گزارش‌گیری Power BI - آماده‌ساز فایل‌ها')
        self.geometry('600x400')

        f = ttk.Frame(self, padding=8)
        f.pack(fill='both', expand=True)

        ttk.Label(f, text='این پنجره فایل‌هایی را تولید می‌کند که به‌راحتی در Power BI ایمپورت می‌شوند').pack(anchor='w')
        ttk.Button(f, text='تولید CSV تمیز برای Power BI', command=self.prepare_powerbi_csv).pack(pady=6)
        ttk.Button(f, text='تولید Excel تمیز برای Power BI', command=self.prepare_powerbi_excel).pack(pady=6)

        ttk.Label(f, text='راهنما: فایل تولیدشده را در Power BI Desktop باز کنید و جداول را به مدل اضافه کنید.').pack(anchor='w', pady=8)

    def prepare_powerbi_csv(self):
        df = self.app.excel.df
        if df is None:
            messagebox.showwarning('هشدار', 'ابتدا داده‌ها را بارگذاری کنید')
            return
        # basic cleaning: remove fully empty rows/columns, normalize repair type column, convert dates
        tmp = df.dropna(how='all')
        # normalize repair
        rep = self.app.excel.cols.get('repair')
        if rep and rep in tmp.columns:
            tmp[rep] = tmp[rep].apply(normalize_repair_type)
        # try convert date col
        date_col = self.app.excel.cols.get('date')
        if date_col and date_col in tmp.columns:
            try:
                tmp[date_col] = pd.to_datetime(tmp[date_col], errors='coerce')
            except Exception:
                pass
        p = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV', '*.csv')])
        if p:
            tmp.to_csv(p, index=False, encoding='utf-8-sig')
            messagebox.showinfo('موفق', 'CSV آماده برای Power BI تولید شد')

    def prepare_powerbi_excel(self):
        df = self.app.excel.df
        if df is None:
            messagebox.showwarning('هشدار', 'ابتدا داده‌ها را بارگذاری کنید')
            return
        tmp = df.dropna(how='all')
        rep = self.app.excel.cols.get('repair')
        if rep and rep in tmp.columns:
            tmp[rep] = tmp[rep].apply(normalize_repair_type)
        p = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx')])
        if p:
            tmp.to_excel(p, index=False, engine='openpyxl')
            messagebox.showinfo('موفق', 'Excel آماده برای Power BI تولید شد')

# -----------------------------
# PyQt alternative simple viewer (minimal)
# -----------------------------
def launch_pyqt_app(file_path=None, sheet_name=None):
    if not PYQT_AVAILABLE:
        print('PySide6 not available')
        return
    app = QtWidgets.QApplication([])
    w = QtWidgets.QWidget()
    w.setWindowTitle('PyQt Alternative - Simple Viewer')
    w.resize(800, 600)
    layout = QtWidgets.QVBoxLayout(w)
    lbl = QtWidgets.QLabel('این پنجره با PySide6 ساخته شده است. اجرای حداقلی برای نمایش دیتافریم:')
    layout.addWidget(lbl)
    if file_path and os.path.exists(file_path) and sheet_name:
        try:
            df = safe_read_excel(file_path, sheet_name)
            txt = QtWidgets.QPlainTextEdit()
            txt.setPlainText(df.head(500).to_string())
            txt.setReadOnly(True)
            layout.addWidget(txt)
        except Exception as e:
            layout.addWidget(QtWidgets.QLabel(f'خطا در خواندن فایل: {e}'))
    w.show()
    app.exec()

# -----------------------------
# Entrypoint
# -----------------------------
if __name__ == '__main__':
    # special-case launching pyqt if requested
    if '--pyqt' in sys.argv:
        # try to read file and sheet from environment or arguments (not required)
        fp = None
        sh = None
        if len(sys.argv) >= 3:
            fp = sys.argv[2]
        if len(sys.argv) >= 4:
            sh = sys.argv[3]
        launch_pyqt_app(fp, sh)
        sys.exit(0)

    root = tk.Tk()
    app = MainAppTk(root)
    root.mainloop()
