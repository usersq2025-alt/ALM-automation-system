import streamlit as st
import pandas as pd
import io
import zipfile
import xlsxwriter
import xml.etree.ElementTree as ET

# --- إعدادات الواجهة ---
st.set_page_config(page_title="نظام أتمتة المقرأة", page_icon="📖", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700;900&display=swap');
    html, body, [class*="css"] { font-family: 'Tajawal', sans-serif; direction: rtl; text-align: right; }
    .stApp { background-color: #f9fbf7; }
    [data-testid="stSidebar"] { background-color: #2d5016 !important; color: white !important; }
    </style>
    """, unsafe_allow_html=True)

# --- محرك القراءة الخام (الذي تجاوز الأخطاء التقنية) ---
def read_xlsx_raw(file_bytes):
    NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
        shared = []
        if "xl/sharedStrings.xml" in zf.namelist():
            tree = ET.parse(zf.open("xl/sharedStrings.xml"))
            for si in tree.getroot().iter("{" + NS + "}si"):
                texts = [t.text or "" for t in si.iter("{" + NS + "}t")]
                shared.append("".join(texts))

        wb_tree = ET.parse(zf.open("xl/workbook.xml"))
        rels_tree = ET.parse(zf.open("xl/_rels/workbook.xml.rels"))
        rels = {r.attrib["Id"]: r.attrib["Target"] for r in rels_tree.getroot()}
        sheets = wb_tree.getroot().find("{" + NS + "}sheets")
        r_id = sheets[0].attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        sheet_path = "xl/" + rels[r_id].lstrip("/").replace("xl/", "")
        
        sheet_tree = ET.parse(zf.open(sheet_path))
        rows_data = []
        for row_el in sheet_tree.getroot().iter("{" + NS + "}row"):
            row_vals = []
            for c in row_el.iter("{" + NS + "}c"):
                v_el = c.find("{" + NS + "}v")
                t = c.attrib.get("t", "")
                if v_el is None or v_el.text is None: row_vals.append("")
                elif t == "s": row_vals.append(shared[int(v_el.text)])
                else: row_vals.append(v_el.text)
            rows_data.append(row_vals)
    
    if not rows_data: return pd.DataFrame()
    # تنظيف العناوين (Headers)
    headers = [str(h).strip() for h in rows_data[0]]
    data = rows_data[1:]
    return pd.DataFrame(data, columns=headers)

# --- منطق الاسم المختصر (إيمان زياد -> إيمان.ز) ---
def get_short_name(full_name):
    if not full_name or str(full_name).lower() == 'nan': return "معلمة"
    parts = str(full_name).strip().split()
    if len(parts) >= 2:
        return f"{parts[0]}.{parts[1][0]}"
    return parts[0]

# --- بناء الملف بالهيكل المطلوب (13 عموداً دقيقاً) ---
def build_excel(raw_df, days, periods, statuses):
    output = io.BytesIO()
    
    # الأعمدة الـ 8 الأساسية من الملف الخام
    source_cols = ["الرقم", "الاسم", "رقم الواتس اب", "المجموعة", "البلد", "المواليد", "الإجازة", "المعلمة"]
    # الأعمدة الـ 13 المطلوبة في الملف النهائي
    target_headers = source_cols + ["الحالة", "يوم الاختبار", "توقيت الاختبار", "الفترة", "الملاحظات"]

    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    ws = workbook.add_worksheet("كشف الاختبار")
    ws.right_to_left()

    # التنسيقات (Formats)
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#2d5016', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    locked_fmt = workbook.add_format({'bg_color': '#F2F2F2', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'locked': True, 'font_name': 'Tajawal'})
    unlocked_fmt = workbook.add_format({'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'locked': False, 'font_name': 'Tajawal'})
    alert_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}) # لون أحمر لكلمة "شرطي"

    # ضبط عرض الأعمدة
    widths = [10, 25, 18, 15, 12, 10, 10, 15, 20, 15, 15, 15, 25]
    for i, w in enumerate(widths): ws.set_column(i, i, w)

    # كتابة العناوين
    for col_num, header in enumerate(target_headers):
        ws.write(0, col_num, header, header_fmt)

    # كتابة البيانات (تخطيط دقيق)
    num_rows = len(raw_df)
    for r_idx in range(num_rows):
        excel_row = r_idx + 1
        for c_idx in range(13):
            # الأعمدة من A إلى H (0 إلى 7): نأخذ البيانات من الملف الخام بالاسم
            if c_idx < 8:
                col_name = target_headers[c_idx]
                val = raw_df[col_name].iloc[r_idx] if col_name in raw_df.columns else ""
                ws.write(excel_row, c_idx, val, locked_fmt)
            # الأعمدة من I إلى M (8 إلى 12): نتركها فارغة للمعلمة
            else:
                ws.write(excel_row, c_idx, "", unlocked_fmt)

    # إضافة 50 صفاً فارغاً إضافياً للطالبات الجدد
    for r_idx in range(num_rows + 1, num_rows + 51):
        for c_idx in range(13):
            fmt = locked_fmt if c_idx < 8 else unlocked_fmt
            ws.write(r_idx, c_idx, "", fmt)

    total_rows = num_rows + 50
    # القوائم المنسدلة (Data Validation)
    ws.data_validation(1, 8, total_rows, 8, {'validate': 'list', 'source': statuses})
    ws.data_validation(1, 9, total_rows, 9, {'validate': 'list', 'source': days})
    ws.data_validation(1, 11, total_rows, 11, {'validate': 'list', 'source': periods})

    # تلوين خلية الملاحظات إذا كتبت المعلمة كلمة "شرطي"
    ws.conditional_format(1, 12, total_rows, 12, {
        'type': 'cell', 'criteria': 'containing', 'value': 'شرطي', 'format': alert_fmt
    })

    ws.protect() # تفعيل الحماية
    workbook.close()
    return output.getvalue()

# --- واجهة Streamlit ---
st.sidebar.markdown("### ⚙️ الإعدادات")
days = [l.strip() for l in st.sidebar.text_area("أيام الأسبوع", "الاثنين 3/23\nالثلاثاء 3/24\nالأربعاء 3/25\nالخميس 3/26\nالجمعة 3/27\nالسبت 3/28\nالأحد 3/29").splitlines() if l.strip()]
periods = [l.strip() for l in st.sidebar.text_area("الفترات", "فجراً من 5.45-9.00\nضحى 9:15-12.30\nظهراً 12:45-4.15\nعصراً 4.30-7.00\nليلاً 7.15-9.30").splitlines() if l.strip()]
statuses = [l.strip() for l in st.sidebar.text_area("قائمة الحالات", "أنهت المقرر\nلم تنه المقرر\nساكنة\nمنسحبة\nأخرجتها الإدارة لأنها مخالفة\nلا يوجد واتس\nتم نقلها لغير مجموعة").splitlines() if l.strip()]

st.title("📖 نظام أتمتة جداول المقرأة")
files = st.file_uploader("ارفعي ملفات الموقع الخام (Excel)", type=["xlsx"], accept_multiple_files=True)

if files and st.button("🚀 توليد الجداول النهائية", type="primary"):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for f in files:
            try:
                content = f.read()
                df = read_xlsx_raw(content)
                if not df.empty:
                    # استخراج اسم المعلمة للتسمية
                    teacher_name = df["المعلمة"].iloc[0] if "المعلمة" in df.columns else "معلمة"
                    short_n = get_short_name(teacher_name)
                    
                    # بناء الملف
                    xlsx_data = build_excel(df, days, periods, statuses)
                    zf.writestr(f"{short_n}.xlsx", xlsx_data)
            except Exception as e:
                st.error(f"خطأ في {f.name}: {e}")
    
    st.success("✅ تم التجهيز بنجاح!")
    st.download_button("📥 تحميل كافة الملفات (ZIP)", zip_buffer.getvalue(), "جداول_المعلمات.zip", "application/zip")
