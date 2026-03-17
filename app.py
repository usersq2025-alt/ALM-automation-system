import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Protection

# إعدادات الواجهة
st.set_page_config(page_title="نظام أتمتة المقرأة", layout="wide")
st.title("🗂️ نظام تجهيز جداول المعلمات - النسخة المعتمدة النهائية")

# لوحة التحكم الجانبية
st.sidebar.header("⚙️ إعدادات الدورة")
days_input = st.sidebar.text_area("أيام الأسبوع", "الاثنين 3/23, الثلاثاء 3/24, الأربعاء 3/25, الخميس 3/26, الجمعة 3/27, السبت 3/28, الأحد 3/29")
periods_input = st.sidebar.text_area("الفترات", "فجرا من 5.45-9.00, ضحى 9:15-12.30, ظهرا 12:45-4.15, عصرا 4.30-7.00, ليلا 7.15-9.30")

# القائمة الحرفية كما طلبتِ تماماً
status_list = [
    "أنهت المقرر",
    "لم تنه المقرر",
    "ساكنة",
    "منسحبة",
    "أخرجتها الإدارة لأنها مخالفة",
    "لا يوجد واتس",
    "تم نقلها لغير مجموعة"
]
status_input = st.sidebar.text_area("خيارات الحالة", ", ".join(status_list))

DAYS = [d.strip() for d in days_input.split(",")]
PERIODS = [p.strip() for p in periods_input.split(",")]
STATUS = [s.strip() for s in status_input.split(",")]

def shorten_name(full_name):
    parts = str(full_name).split()
    if not parts or parts[0].lower() == 'nan': return "معلمة"
    if len(parts) == 1: return parts[0]
    # الاسم الأول + نقطة + أول حرف من الاسم الثاني (إيمان زياد -> إيمان.ز)
    return f"{parts[0]}.{parts[1][0]}"

uploaded_files = st.file_uploader("ارفعي ملفات الموقع الخام (Excel/CSV)", accept_multiple_files=True)

if st.button("🚀 ابدأ المعالجة الفورية"):
    if uploaded_files:
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for uploaded_file in uploaded_files:
                try:
                    # قراءة الملفات
                    if uploaded_file.name.endswith('.csv'):
                        df = pd.read_csv(uploaded_file)
                    else:
                        df = pd.read_excel(uploaded_file, engine='openpyxl')
                    
                    # استخراج اسم المعلمة
                    teacher_col = 'المعلمة' if 'المعلمة' in df.columns else df.columns[-1]
                    teacher_full = df[teacher_col].iloc[0]
                    short_name = shorten_name(teacher_full)
                    
                    # إنشاء ملف إكسل جديد (خالٍ من أخطاء التنسيق)
                    wb = Workbook()
                    ws = wb.active
                    ws.sheet_view.rightToLeft = True
                    
                    # الهيكل المعتمد (13 عموداً)
                    base_headers = ['الرقم', 'الاسم', 'رقم الواتس اب', 'المجموعة', 'البلد', 'المواليد', 'الإجازة', 'المعلمة']
                    extra_headers = ['الحالة', 'يوم الاختبار', 'توقيت الاختبار', 'الفترة', 'الملاحظات']
                    ws.append(base_headers + extra_headers)
                    
                    # تعبئة البيانات
                    for r in df[base_headers].values.tolist():
                        ws.append(r + [""] * len(extra_headers))
                    
                    # إعداد القوائم المنسدلة
                    dv_status = DataValidation(type="list", formula1=f'"{",".join(STATUS)}"')
                    dv_days = DataValidation(type="list", formula1=f'"{",".join(DAYS)}"')
                    dv_periods = DataValidation(type="list", formula1=f'"{",".join(PERIODS)}"')
                    
                    ws.add_data_validation(dv_status)
                    ws.add_data_validation(dv_days)
                    ws.add_data_validation(dv_periods)
                    
                    max_r = ws.max_row + 40 # مجال واسع للطالبات الجدد
                    dv_status.add(f"I2:I{max_r}")
                    dv_days.add(f"J2:J{max_r}")
                    dv_periods.add(f"L2:L{max_r}")
                    
                    # نظام الحماية (قفل A-H وفتح الباقي)
                    for row_idx in range(1, max_r + 1):
                        for col_idx in range(1, 14):
                            cell = ws.cell(row=row_idx, column=col_idx)
                            if col_idx <= 8:
                                cell.protection = Protection(locked=True)
                            else:
                                cell.protection = Protection(locked=False)
                    
                    ws.protection.sheet = True
                    ws.protection.enable()
                    
                    file_stream = BytesIO()
                    wb.save(file_stream)
                    zip_file.writestr(f"{short_name}.xlsx", file_stream.getvalue())
                    
                except Exception as e:
                    st.error(f"⚠️ خطأ في معالجة ملف {uploaded_file.name}: {e}")
        
        if zip_buffer.tell() > 0:
            st.success("✅ تم التجهيز بنجاح تام!")
            st.download_button("📥 تحميل المجلد الجاهز (Zip)", zip_buffer.getvalue(), "Quran_Files.zip", "application/zip")
    else:
        st.warning("يرجى رفع الملفات المصدريّة أولاً.")
