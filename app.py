import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Protection, PatternFill

# إعدادات الواجهة
st.set_page_config(page_title="نظام أتمتة المقرأة", layout="wide")
st.title("🗂️ نظام تجهيز جداول المعلمات")

# لوحة التحكم في الجانب
st.sidebar.header("⚙️ إعدادات الدورة")
days_input = st.sidebar.text_area("أيام الأسبوع (افصلي بينها بفاصلة)", "الاثنين 3/23, الثلاثاء 3/24, الأربعاء 3/25, الخميس 3/26, الجمعة 3/27, السبت 3/28, الأحد 3/29")
periods_input = st.sidebar.text_area("الفترات (افصلي بينها بفاصلة)", "فجرا 5.45-9.00, ضحى 9:15-12.30, ظهرا 12:45-4.15, عصرا 4.30-7.00, ليلا 7.15-9.30")

DAYS = [d.strip() for d in days_input.split(",")]
PERIODS = [p.strip() for p in periods_input.split(",")]
STATUS = ["أنهت المقرر", "لم تنه المقرر", "غائبة", "منسحبة"]

def shorten_name(name):
    special = ["سهيلا", "شام", "سنا"]
    parts = str(name).split()
    if not parts or parts[0] == 'nan': return "معلمة"
    if parts[0] in special: return parts[0]
    return f"{parts[0]}.{parts[-1][0]}" if len(parts) > 1 else parts[0]

uploaded_files = st.file_uploader("ارفعي ملفات المعلمات (Excel/CSV)", accept_multiple_files=True)

if st.button("🚀 ابدأ معالجة وتجهيز الملفات"):
    if uploaded_files:
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for uploaded_file in uploaded_files:
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(uploaded_file)
                
                teacher_name = df['المعلمة'].iloc[0] if 'المعلمة' in df.columns else "معلمة"
                short_name = shorten_name(teacher_name)
                
                wb = Workbook()
                ws = wb.active
                ws.sheet_view.rightToLeft = True
                
                headers = list(df.columns) + ['الحالة', 'يوم الاختبار', 'توقيت الاختبار', 'الفترة']
                ws.append(headers)
                for r in df.values.tolist():
                    ws.append(r + ["", "", "", ""])
                
                gray_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
                
                dv_status = DataValidation(type="list", formula1=f'"{",".join(STATUS)}"')
                dv_days = DataValidation(type="list", formula1=f'"{",".join(DAYS)}"')
                dv_periods = DataValidation(type="list", formula1=f'"{",".join(PERIODS)}"')
                
                ws.add_data_validation(dv_status)
                ws.add_data_validation(dv_days)
                ws.add_data_validation(dv_periods)
                
                max_r = ws.max_row + 20
                dv_status.add(f"I2:I{max_r}")
                dv_days.add(f"J2:J{max_r}")
                dv_periods.add(f"L2:L{max_r}")
                
                for row in range(1, max_r + 1):
                    for col in range(1, 14):
                        cell = ws.cell(row=row, column=col)
                        if col <= 8:
                            cell.protection = Protection(locked=True)
                            if row > 1: cell.fill = gray_fill
                        else:
                            cell.protection = Protection(locked=False)
                
                ws.protection.sheet = True
                ws.protection.enable()
                
                file_stream = BytesIO()
                wb.save(file_stream)
                zip_file.writestr(f"{short_name}.xlsx", file_stream.getvalue())
        
        st.success("✅ تم تجهيز جميع الملفات!")
        st.download_button(label="📥 تحميل الملفات الجاهزة (Zip)", data=zip_buffer.getvalue(), file_name="Ready_Files.zip", mime="application/zip")
    else:
        st.error("يرجى رفع ملف واحد على الأقل.")
