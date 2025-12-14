# -*- coding: utf-8 -*-
"""
گزارش‌گیر قالبسازی حرفه‌ای (با پشتیبانی بهتر فارسی در نمودار)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib import font_manager
import warnings

# optional shaping for Arabic/Persian
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    HAS_ARABIC_TOOLS = True
except Exception:
    HAS_ARABIC_TOOLS = False

warnings.filterwarnings("ignore", category=UserWarning)

# ----------------------------
# Helpers for Persian support
# ----------------------------
PREFERRED_FONTS = [
    "Vazir", "IRANSans", "IRANSansX", "Tahoma", "DejaVu Sans", "Arial", "Noto Naskh Arabic", "Noto Sans Arabic"
]

def find_system_font_name(preferred_list=PREFERRED_FONTS):
    """Return the first available font family name from preferred_list, else None."""
    available = {f.name for f in font_manager.fontManager.ttflist}
    for name in preferred_list:
        if name in available:
            return name
    return None

def apply_matplotlib_font():
    """Set matplotlib rcParams to use a Persian-capable font if available."""
    font_name = find_system_font_name()
    if font_name:
        matplotlib.rcParams['font.family'] = font_name
        matplotlib.rcParams['axes.unicode_minus'] = False
    else:
        # fallback: keep default but disable unicode minus fix
        matplotlib.rcParams['axes.unicode_minus'] = False

def reshape_text_if_needed(text: str) -> str:
    """If arabic_reshaper + bidi available, reshape & bidi the text for correct Persian display."""
    if not text:
        return text
    if HAS_ARABIC_TOOLS:
        try:
            reshaped = arabic_reshaper.reshape(text)
            return get_display(reshaped)
        except Exception:
            return text
    return text

# Apply font at import time
apply_matplotlib_font()

# ======================
# Excel Handler (unchanged)
# ======================
class ExcelHandler:
    def __init__(self):
        self.file_path = None
        self.sheet_names = []
        self.df = None

    def load_file(self, path):
        if not os.path.exists(path):
            raise FileNotFoundError("File not found")
        self.file_path = path
        xls = pd.ExcelFile(path, engine='openpyxl')
        self.sheet_names = xls.sheet_names
        return self.sheet_names

    def load_sheet(self, sheet_name):
        self.df = pd.read_excel(self.file_path, sheet_name=sheet_name, engine='openpyxl')
        # normalize columns (strip)
        self.df.columns = [str(c).strip() for c in self.df.columns]
        return self.df

# ======================
# Main Application (unchanged structure)
# ======================
class MainAppTk:
    def __init__(self, root):
        self.root = root
        self.root.title('گزارش‌گیر قالبسازی حرفه‌ای')
        self.root.geometry('1000x700')
        self.excel = ExcelHandler()

        file_frame = ttk.Frame(root, padding=5)
        file_frame.pack(fill='x')
        ttk.Label(file_frame, text='فایل:').pack(side='left')
        self.file_entry = ttk.Entry(file_frame, width=70)
        self.file_entry.pack(side='left', padx=5)
        ttk.Button(file_frame, text='انتخاب فایل', command=self.open_file).pack(side='left')

        ttk.Label(file_frame, text='شیت:').pack(side='left', padx=(20,0))
        self.sheet_cb = ttk.Combobox(file_frame, state='readonly', width=30)
        self.sheet_cb.pack(side='left')
        ttk.Button(file_frame, text='بارگذاری شیت‌ها', command=self.load_sheets).pack(side='left', padx=5)

        btn_frame = ttk.Frame(root, padding=5)
        btn_frame.pack(fill='x')
        ttk.Button(btn_frame, text='نمایش دیتا', command=self.open_loader).pack(side='left', padx=5)
        ttk.Button(btn_frame, text='فیلتر ساده', command=self.open_simple_filter).pack(side='left', padx=5)
        ttk.Button(btn_frame, text='فیلتر پیشرفته', command=self.open_adv_filter).pack(side='left', padx=5)
        ttk.Button(btn_frame, text='تحلیل و نمودار', command=self.open_analysis).pack(side='left', padx=5)
        ttk.Button(btn_frame, text='خروج', command=self.on_closing).pack(side='right', padx=5)

        self.status_var = tk.StringVar(value='آماده')
        ttk.Label(root, textvariable=self.status_var, relief='sunken').pack(side='bottom', fill='x')

        self.loader_win = None
        self.simple_filter_win = None
        self.adv_filter_win = None
        self.analysis_win = None

    def open_file(self):
        path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx *.xls')])
        if path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, path)
            self.status_var.set(f'فایل انتخاب شد: {os.path.basename(path)}')

    def load_sheets(self):
        path = self.file_entry.get()
        if not path or not os.path.exists(path):
            messagebox.showerror('خطا', 'لطفاً یک فایل اکسل معتبر انتخاب کنید')
            return
        try:
            sheets = self.excel.load_file(path)
            self.sheet_cb['values'] = sheets
            if sheets:
                self.sheet_cb.set(sheets[0])
                self.excel.load_sheet(sheets[0])
                self.status_var.set(f'{len(sheets)} شیت بارگذاری شد. اولین شیت بارگذاری گردید.')
            else:
                self.status_var.set('هیچ شیتی در فایل پیدا نشد')
        except Exception as e:
            messagebox.showerror('خطا', f'خطا در بارگذاری فایل: {str(e)}')

    def update_all_windows(self):
        for window in [self.loader_win, self.simple_filter_win, self.adv_filter_win, self.analysis_win]:
            if window and hasattr(window, 'update_columns'):
                window.update_columns()

    def open_loader(self):
        if self.excel.df is None:
            messagebox.showerror('خطا', 'لطفاً ابتدا فایل و شیت را بارگذاری کنید')
            return
        if self.loader_win and tk.Toplevel.winfo_exists(self.loader_win):
            self.loader_win.lift()
            self.loader_win.load_data()
            return
        self.loader_win = LoaderWindow(self)

    def open_simple_filter(self):
        if self.excel.df is None:
            messagebox.showerror('خطا', 'لطفاً ابتدا فایل و شیت را بارگذاری کنید')
            return
        if self.simple_filter_win and tk.Toplevel.winfo_exists(self.simple_filter_win):
            self.simple_filter_win.lift()
            self.simple_filter_win.update_columns()
            return
        self.simple_filter_win = SimpleFilterWindow(self)

    def open_adv_filter(self):
        if self.excel.df is None:
            messagebox.showerror('خطا', 'لطفاً ابتدا فایل و شیت را بارگذاری کنید')
            return
        if self.adv_filter_win and tk.Toplevel.winfo_exists(self.adv_filter_win):
            self.adv_filter_win.lift()
            self.adv_filter_win.update_columns()
            return
        self.adv_filter_win = AdvancedFilterWindow(self)

    def open_analysis(self):
        if self.excel.df is None:
            messagebox.showerror('خطا', 'لطفاً ابتدا فایل و شیت را بارگذاری کنید')
            return
        if self.analysis_win and tk.Toplevel.winfo_exists(self.analysis_win):
            self.analysis_win.lift()
            self.analysis_win.update_columns()
            return
        self.analysis_win = AnalysisWindow(self)

    def on_closing(self):
        for window in [self.loader_win, self.simple_filter_win, self.adv_filter_win, self.analysis_win]:
            if window and tk.Toplevel.winfo_exists(window):
                window.destroy()
        self.root.quit()

# -------------------------
# Loader, SimpleFilter, AdvancedFilter (unchanged from your version)
# -------------------------
class LoaderWindow(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.title('نمایش دیتا')
        self.geometry('900x500')
        self.app = app
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill='x', pady=5)
        ttk.Button(btn_frame, text='بارگذاری دیتا', command=self.load_data).pack(side='left', padx=5)
        ttk.Button(btn_frame, text='ذخیره به Excel', command=self.save_excel).pack(side='left', padx=5)
        ttk.Button(btn_frame, text='ذخیره به CSV', command=self.save_csv).pack(side='left', padx=5)
        ttk.Label(self, text='دیتا پیش‌نمایش:').pack(anchor='w', padx=5, pady=5)
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill='both', expand=True)
        self.tree = ttk.Treeview(tree_frame, show='headings')
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        self.load_data()

    def load_data(self):
        try:
            if self.app.excel.df is None:
                messagebox.showerror('خطا', 'هیچ دیتایی برای نمایش وجود ندارد')
                return
            df = self.app.excel.df
            self.display_data(df)
            self.app.status_var.set(f'دیتا نمایش داده شد - {len(df)} رکورد')
        except Exception as e:
            messagebox.showerror('خطا', f'خطا در بارگذاری دیتا: {str(e)}')

    def display_data(self, df):
        self.tree.delete(*self.tree.get_children())
        self.tree['columns'] = list(df.columns)
        for col in df.columns:
            self.tree.heading(col, text=str(col))
            self.tree.column(col, width=100, anchor='center')
        for _, row in df.head(1000).iterrows():
            self.tree.insert('', tk.END, values=[str(x) if not pd.isna(x) else '' for x in row])

    def save_excel(self):
        if self.app.excel.df is None:
            messagebox.showerror('خطا', 'هیچ دیتایی برای ذخیره وجود ندارد')
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files","*.xlsx"),("All files","*.*")])
        if path:
            try:
                self.app.excel.df.to_excel(path, index=False, engine='openpyxl')
                messagebox.showinfo('موفق', 'فایل با موفقیت ذخیره شد')
            except Exception as e:
                messagebox.showerror('خطا', f'خطا در ذخیره فایل: {str(e)}')

    def save_csv(self):
        if self.app.excel.df is None:
            messagebox.showerror('خطا', 'هیچ دیتایی برای ذخیره وجود ندارد')
            return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files","*.csv"),("All files","*.*")])
        if path:
            try:
                self.app.excel.df.to_csv(path, index=False, encoding='utf-8-sig')
                messagebox.showinfo('موفق', 'فایل با موفقیت ذخیره شد')
            except Exception as e:
                messagebox.showerror('خطا', f'خطا در ذخیره فایل: {str(e)}')

class SimpleFilterWindow(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.title('فیلتر ساده')
        self.geometry('600x400')
        self.app = app
        form_frame = ttk.Frame(self)
        form_frame.pack(fill='x', pady=10)
        ttk.Label(form_frame, text='ستون برای فیلتر:').grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.col_cb = ttk.Combobox(form_frame, state='readonly', width=30)
        self.col_cb.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        ttk.Label(form_frame, text='مقدار مورد نظر:').grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.val_entry = ttk.Entry(form_frame, width=30)
        self.val_entry.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        ttk.Button(form_frame, text='اعمال فیلتر', command=self.apply_filter).grid(row=2, column=0, columnspan=2, pady=10)
        form_frame.grid_columnconfigure(1, weight=1)
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill='both', expand=True)
        self.tree = ttk.Treeview(tree_frame, show='headings')
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        self.update_columns()

    def update_columns(self):
        if self.app.excel.df is not None:
            self.col_cb['values'] = list(self.app.excel.df.columns)
            if self.col_cb['values']:
                self.col_cb.current(0)

    def apply_filter(self):
        try:
            df = self.app.excel.df
            if df is None or df.empty:
                messagebox.showerror('خطا', 'لطفاً ابتدا دیتا را بارگذاری کنید')
                return
            col = self.col_cb.get()
            val = self.val_entry.get().strip()
            if not col:
                messagebox.showerror('خطا', 'لطفاً ستون را انتخاب کنید')
                return
            if not val:
                messagebox.showerror('خطا', 'لطفاً مقدار فیلتر را وارد کنید')
                return
            try:
                filtered = df[df[col].astype(str).str.contains(val, na=False)]
            except Exception as e:
                messagebox.showerror('خطا', f'خطا در اعمال فیلتر: {str(e)}')
                return
            self.display_data(filtered)
            self.app.status_var.set(f'فیلتر اعمال شد - {len(filtered)} رکورد پیدا شد')
        except Exception as e:
            messagebox.showerror('خطا', f'خطای غیرمنتظره: {str(e)}')

    def display_data(self, df):
        self.tree.delete(*self.tree.get_children())
        if df.empty:
            messagebox.showwarning('هشدار', 'هیچ رکوردی با این مشخصات پیدا نشد')
            return
        self.tree['columns'] = list(df.columns)
        for col in df.columns:
            self.tree.heading(col, text=str(col))
            self.tree.column(col, width=100)
        for _, row in df.iterrows():
            self.tree.insert('', tk.END, values=[str(x) if not pd.isna(x) else '' for x in row])

class AdvancedFilterWindow(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.title('فیلتر پیشرفته')
        self.geometry('700x500')
        self.app = app
        form_frame = ttk.Frame(self)
        form_frame.pack(fill='x', pady=10)
        ttk.Label(form_frame, text='ستون اول:').grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.col1_cb = ttk.Combobox(form_frame, state='readonly', width=25)
        self.col1_cb.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        ttk.Label(form_frame, text='مقدار اول:').grid(row=0, column=2, padx=5, pady=5, sticky='w')
        self.val1_entry = ttk.Entry(form_frame, width=25)
        self.val1_entry.grid(row=0, column=3, padx=5, pady=5, sticky='w')
        ttk.Label(form_frame, text='ستون دوم (اختیاری):').grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.col2_cb = ttk.Combobox(form_frame, state='readonly', width=25)
        self.col2_cb.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        ttk.Label(form_frame, text='مقدار دوم:').grid(row=1, column=2, padx=5, pady=5, sticky='w')
        self.val2_entry = ttk.Entry(form_frame, width=25)
        self.val2_entry.grid(row=1, column=3, padx=5, pady=5, sticky='w')
        ttk.Button(form_frame, text='اعمال فیلتر', command=self.apply_filter).grid(row=2, column=0, columnspan=4, pady=10)
        for i in range(4):
            form_frame.grid_columnconfigure(i, weight=1)
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill='both', expand=True)
        self.tree = ttk.Treeview(tree_frame, show='headings')
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        self.update_columns()

    def update_columns(self):
        if self.app.excel.df is not None:
            columns = list(self.app.excel.df.columns)
            self.col1_cb['values'] = columns
            self.col2_cb['values'] = columns
            if columns:
                self.col1_cb.current(0)

    def apply_filter(self):
        try:
            df = self.app.excel.df
            if df is None or df.empty:
                messagebox.showerror('خطا', 'دیتا بارگذاری نشده است')
                return
            col1 = self.col1_cb.get()
            val1 = self.val1_entry.get().strip()
            if not col1:
                messagebox.showerror('خطا', 'لطفاً ستون اول را انتخاب کنید')
                return
            if not val1:
                messagebox.showerror('خطا', 'لطفاً مقدار اول را وارد کنید')
                return
            try:
                filtered = df[df[col1].astype(str).str.contains(val1, na=False)]
            except Exception as e:
                messagebox.showerror('خطا', f'خطا در اعمال فیلتر اول: {str(e)}')
                return
            col2 = self.col2_cb.get()
            val2 = self.val2_entry.get().strip()
            if col2 and val2:
                try:
                    filtered = filtered[filtered[col2].astype(str).str.contains(val2, na=False)]
                except Exception as e:
                    messagebox.showerror('خطا', f'خطا در اعمال فیلتر دوم: {str(e)}')
                    return
            self.display_data(filtered)
            self.app.status_var.set(f'فیلتر پیشرفته اعمال شد - {len(filtered)} رکورد پیدا شد')
        except Exception as e:
            messagebox.showerror('خطا', f'خطای غیرمنتظره: {str(e)}')

    def display_data(self, df):
        self.tree.delete(*self.tree.get_children())
        if df.empty:
            messagebox.showwarning('هشدار', 'هیچ رکوردی با این مشخصات پیدا نشد')
            return
        self.tree['columns'] = list(df.columns)
        for col in df.columns:
            self.tree.heading(col, text=str(col))
            self.tree.column(col, width=100)
        for _, row in df.iterrows():
            self.tree.insert('', tk.END, values=[str(x) if not pd.isna(x) else '' for x in row])

# ======================
# Analysis Window (with Persian fixes)
# ======================
class AnalysisWindow(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.title('تحلیل و نمودار')
        self.geometry('800x600')
        self.app = app

        form_frame = ttk.Frame(self)
        form_frame.pack(fill='x', pady=10)
        ttk.Label(form_frame, text='ستون تحلیل:').grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.col_cb = ttk.Combobox(form_frame, state='readonly', width=30)
        self.col_cb.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        ttk.Button(form_frame, text='نمایش نمودار', command=self.show_plot).grid(row=0, column=2, padx=5, pady=5)
        form_frame.grid_columnconfigure(1, weight=1)

        self.fig = plt.Figure(figsize=(6,4), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self)
        self.canvas.get_tk_widget().pack(fill='both', expand=True)

        self.update_columns()

    def update_columns(self):
        if self.app.excel.df is not None:
            self.col_cb['values'] = list(self.app.excel.df.columns)
            if self.col_cb['values']:
                self.col_cb.current(0)

    def _fix_xticklabels(self):
        """Apply Persian reshaping + bidi to xticklabels if needed."""
        labels = [t.get_text() for t in self.ax.get_xticklabels()]
        fixed = [reshape_text_if_needed(lbl) for lbl in labels]
        self.ax.set_xticklabels(fixed, rotation=45, ha='right')

    def show_plot(self):
        try:
            df = self.app.excel.df
            if df is None or df.empty:
                messagebox.showerror('خطا', 'دیتا بارگذاری نشده است')
                return
            col = self.col_cb.get()
            if not col:
                messagebox.showerror('خطا', 'لطفاً ستون را انتخاب کنید')
                return
            if col not in df.columns:
                messagebox.showerror('خطا', 'ستون انتخاب شده معتبر نیست')
                return
            self.ax.clear()
            if pd.api.types.is_numeric_dtype(df[col]):
                # numeric: histogram
                data = pd.to_numeric(df[col], errors='coerce').dropna()
                if data.empty:
                    messagebox.showwarning('هشدار', 'دادهٔ عددی برای نمایش وجود ندارد')
                    return
                self.ax.hist(data, bins=20, edgecolor='black')
                title = reshape_text_if_needed(f'توزیع {col}')
                self.ax.set_title(title)
                self.ax.set_ylabel(reshape_text_if_needed('تعداد'))
                self.ax.set_xlabel(reshape_text_if_needed(col))
            else:
                # categorical: bar chart
                value_counts = df[col].astype(str).value_counts().head(20)
                if value_counts.empty:
                    messagebox.showwarning('هشدار', 'داده‌ای برای نمایش وجود ندارد')
                    return
                # plot with labels fixed
                bars = self.ax.bar(range(len(value_counts)), value_counts.values, edgecolor='black')
                # set xticks with reshaped labels
                xt = [reshape_text_if_needed(str(s)) for s in value_counts.index.tolist()]
                self.ax.set_xticks(range(len(xt)))
                self.ax.set_xticklabels(xt, rotation=45, ha='right')
                title = reshape_text_if_needed(f'مقادیر {col}')
                self.ax.set_title(title)
                self.ax.set_ylabel(reshape_text_if_needed('تعداد'))
                self.ax.set_xlabel(reshape_text_if_needed(col))
            self.fig.tight_layout()
            self.canvas.draw()
            self.app.status_var.set(f'نمودار برای ستون {col} رسم شد')
        except Exception as e:
            messagebox.showerror('خطا', f'خطا در رسم نمودار: {str(e)}')

# ======================
# Entrypoint
# ======================
if __name__ == '__main__':
    root = tk.Tk()
    app = MainAppTk(root)
    root.mainloop()
