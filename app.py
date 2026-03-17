import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile

# إعدادات الواجهة
st.set_page_config(page_title="نظام أتمتة المقرأة", layout="wide")
st.title("🗂️ نظام أتمتة المقرأة - النسخة الاحترافية (محرك XlsxWriter)")

# لوحة التحكم الجانبية
st.sidebar.header("⚙️ إعدادات الدورة")
days_input = st.sidebar.text_area("أيام الأسبوع", "الاثنين 3/23, الثلاثاء 3/24, الأربعاء 3/25, الخميس 3/26, الجمعة 3/27, السبت 3/28, الأحد 3/29")
periods_input = st.sidebar.text_area("الفترات", "فجرا من 5.45-9.00, ضحى 9:15-12.30, ظهرا 12:45-4.15, عصرا 4.30-7.00, ليلا 7.15-9.30")

# القائمة الحرفية الدقيقة
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
    return f"{parts[0]}.{parts[1][0]}"

uploaded_files = st.file_uploader("ارفعي ملفات الموقع الخام (Excel/CSV)", accept_multiple_files=True)

if st.button("🚀 ابدأ المعالجة النهائية"):
    if uploaded_files:
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for uploaded_file in uploaded_files:
                try:
                    # قراءة الملف الأصلي
                    if uploaded_file.name.endswith('.csv'):
                        df_raw = pd.read_csv(uploaded_file)
                    else:
                        df_raw = pd.read_excel(uploaded_file)
                    
                    # استخراج اسم المعلمة للاختصار
                    teacher_col = 'المعلمة' if 'المعلمة' in df_raw.columns else df_raw.columns[-1]
                    teacher_full = df_raw[teacher_col].iloc[0]
                    short_name = shorten_name(teacher_full)
                    
                    # تجهيز البيانات (13 عموداً)
                    base_cols = ['الرقم', 'الاسم', 'رقم الواتس اب', 'المجموعة', 'البلد', 'المواليد', 'الإجازة', 'المعلمة']
                    df_final = df_raw[base_cols].copy()
                    for col in ['الحالة', 'يوم الاختبار', 'توقيت الاختبار', 'الفترة', 'الملاحظات']:
                        df_final[col] = ""
                    
                    # إنشاء ملف Excel باستخدام محرك XlsxWriter (الأكثر استقراراً)
                    output = BytesIO()
                    writer = pd.ExcelWriter(output, engine='xlsxwriter')
                    df_final.to_excel(writer, index=False, sheet_name='Sheet1')
                    
                    workbook = writer.book
                    worksheet = writer.sheets['Sheet1']
                    
                    # إعداد التنسيقات
                    worksheet.right_to_left()
                    
                    # تنسيق الرمادي للأعمدة المقفولة
                    locked_format = workbook.add_format({'bg_color': '#F2F2F2', 'locked': True, 'border': 1})
                    unlocked_format = workbook.add_format({'locked': False, 'border': 1})
                    header_format = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center'})
                    
                    # تطبيق التنسيق والحماية على الخلايا
                    num_rows = len(df_final) + 50 # ترك مجال للطالبات الجدد
                    
                    # قفل الأعمدة A-H وتلوينها
                    worksheet.conditional_format(1, 0, num_rows, 7, {
                        'type': 'no_blanks',
                        'format': locked_format
                    })
                    worksheet.conditional_format(1, 0, num_rows, 7, {
                        'type': 'blanks',
                        'format': locked_format
                    })
                    
                    # فتح الأعمدة I-M للكتابة
                    worksheet.set_column('I:M', 20, unlocked_format)
                    
                    # إضافة القوائم المنسدلة
                    worksheet.data_validation('I2:I100', {'validate': 'list', 'source': STATUS})
                    worksheet.data_validation('J2:J100', {'validate': 'list', 'source': DAYS})
                    worksheet.data_validation('L2:L100', {'validate': 'list', 'source': PERIODS})
                    
                    # حماية الورقة
                    worksheet.protect()
                    
                    writer.close()
                    zip_file.writestr(f"{short_name}.xlsx", output.getvalue())
                    
                except Exception as e:
                    st.error(f"⚠️ خطأ في معالجة ملف {uploaded_file.name}: {e}")
        
        if zip_buffer.tell() > 0:
            st.success("✅ تم التجهيز بنجاح تام وبدون أخطاء!")
            st.download_button("📥 تحميل المجلد الجاهز (Zip)", zip_buffer.getvalue(), "Quran_Files.zip", "application/zip")
    else:
        st.warning("يرجى رفع الملفات أولاً.")
