import pandas as pd
from tkinter import Tk, filedialog

# پنجره انتخاب فایل بدون پنجره اصلی Tkinter
root = Tk()
root.withdraw()
file_path = filedialog.askopenfilename(
    title="فایل قالبسازی.xlsx را انتخاب کنید",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)

if not file_path:
    print("❌ فایلی انتخاب نشد.")
    exit()

# انتخاب شیت (می‌توان ثابت یا پویا کرد)
sheet_name = 'فروردین'  # اگر خواستی از کاربر بگیریم، بعداً اضافه می‌کنم

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')

    print("\n✅ فایل با موفقیت خوانده شد!\n")
    print('📊 ستون‌ها:', df.columns.tolist())
    print(df.head(10))
    print('🧩 تعداد ردیف‌ها:', len(df))
    if 'نوع' in df.columns:
        print('🔹 مقادیر ستون «نوع»:', df['نوع'].dropna().unique()[:50])
    else:
        print("⚠️ ستون 'نوع' در فایل پیدا نشد.")

except Exception as e:
    print(f"\n❌ خطا در خواندن فایل یا شیت: {e}")

