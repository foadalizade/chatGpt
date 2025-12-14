from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

import arabic_reshaper
from bidi.algorithm import get_display

# ثبت فونت فارسی Vazirmatn-Black
pdfmetrics.registerFont(TTFont('VazirBlack', 'Vazirmatn-Black.ttf'))

# مسیر خروجی PDF
pdf_path = "portfolio_final_rtl.pdf"

# ایجاد سند PDF
doc = SimpleDocTemplate(pdf_path, pagesize=A4,
                        rightMargin=2*cm, leftMargin=2*cm,
                        topMargin=2*cm, bottomMargin=2*cm)
styles = getSampleStyleSheet()
story = []

# استایل‌ها با فونت Vazirmatn-Black
title_style = ParagraphStyle('Title', fontName='VazirBlack', fontSize=24, alignment=1, textColor=colors.darkblue)
subtitle_style = ParagraphStyle('Subtitle', fontName='VazirBlack', fontSize=16, alignment=1, textColor=colors.darkslategray)
section_style = ParagraphStyle('Section', fontName='VazirBlack', fontSize=18, textColor=colors.HexColor("#ffb703"))
normal_style = ParagraphStyle('Normal', fontName='VazirBlack', fontSize=12, textColor=colors.black, leading=16)

# تابع تبدیل متن فارسی برای RTL
def make_rtl(text):
    reshaped = arabic_reshaper.reshape(text)
    return get_display(reshaped)

# هدر
story.append(Paragraph(make_rtl("فؤاد علیزاده"), title_style))
story.append(Paragraph(make_rtl("برنامه‌نویس پایتون • تحلیل داده • ساخت اپلیکیشن دسکتاپ • گزارش‌گیری"), subtitle_style))
story.append(Spacer(1, 12))

# بخش مهارت‌ها
story.append(Paragraph(make_rtl("مهارت‌ها"), section_style))
skills_data = [[make_rtl("Python"), make_rtl("Pandas"), make_rtl("Tkinter")],
               [make_rtl("Openpyxl"), make_rtl("ReportLab"), make_rtl("Django")]]
table = Table(skills_data, colWidths=[5*cm]*3)
table.setStyle(TableStyle([
    ('BACKGROUND',(0,0),(-1,-1),colors.HexColor("#f0f8ff")),
    ('TEXTCOLOR',(0,0),(-1,-1),colors.HexColor("#0073ff")),
    ('ALIGN',(0,0),(-1,-1),'CENTER'),
    ('FONTNAME',(0,0),(-1,-1),'VazirBlack'),
    ('FONTSIZE',(0,0),(-1,-1),12),
    ('INNERGRID', (0,0), (-1,-1), 0.5, colors.gray),
    ('BOX', (0,0), (-1,-1), 1, colors.gray)
]))
story.append(table)
story.append(Spacer(1, 12))

# بخش پروژه‌ها
story.append(Paragraph(make_rtl("پروژه‌ها"), section_style))
projects = [
    (make_rtl("Report-Excel"), make_rtl("پایتون • دستورات • ایجاد گزارش • داده‌ها • یه سری تجربه کلی • نمایش فایل نمونه‌کار"), "https://picsum.photos/200/100"),
    
    (make_rtl("Desktop App"), make_rtl("طراحی برنامه دسکتاپ برای تحلیل داده‌ها و ساخت خروجی گرافیکی."), "https://picsum.photos/200/100?2"),
    
    (make_rtl("Data Analysis"), make_rtl("تحلیل و ارزیابی داده‌ها و ساخت داشبوردهای آماری."), "https://picsum.photos/200/100?3")
]

for title, desc, img_url in projects:
    story.append(Paragraph(title, ParagraphStyle('ProjTitle', fontName='VazirBlack', fontSize=14, textColor=colors.HexColor("#ff6f61"))))
    story.append(Paragraph(desc, normal_style))
    story.append(Spacer(1, 6))
    try:
        img = Image(img_url, width=10*cm, height=5*cm)
        story.append(img)
        story.append(Spacer(1, 12))
    except:
        pass

# بخش تماس
story.append(Spacer(1, 12))
story.append(Paragraph(make_rtl("تماس"), section_style))
story.append(Paragraph(make_rtl("ایمیل: f.alizadeh@example.com"), normal_style))

# ساخت PDF
doc.build(story)

print(f"PDF فارسی RTL حرفه‌ای ساخته شد: {pdf_path}")
