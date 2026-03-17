import streamlit as st
import pandas as pd
import io
import zipfile
import xlsxwriter
import xml.etree.ElementTree as ET

# --- إعدادات الصفحة والواجهة ---
st.set_page_config(page_title="أداة أتمتة المقرأة", page_icon="📖", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700;900&display=swap');
    html, body, [class*="css"] { font-family: 'Tajawal', sans-serif; direction: rtl; text-align: right; }
    .stApp { background-color: #f9fbf7; }
    .main-title { color: #2d5016; font-weight: 900; font-size: 2.5rem; text-align: center; margin-bottom: 2rem; }
    [data-testid="stSidebar"] { background-color: #2d5016 !important; color: white !important; }
    </style>
    """, unsafe_allow_html=True)

# --- قارئ XML الخام لتجنب أخطاء التنسيق تماماً ---
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

    max_len = max(len(r) for r in rows_data)
    for r in rows_data:
        while len(r) < max_len: r.append("")
    return pd.DataFrame(rows_data[1:], columns=rows_data[0])

# --- مساعدات المنطق ---
def get_short_name(full_name):
    parts = str(full_name).strip().split()
    if len(parts) >= 2:
        return f"{parts[0]}.{parts[1][0]}"
    return parts[0] if parts else "معلمة"

def parse_list(text):
    return [line.strip() for line in text.strip().splitlines() if line.strip()]

# --- بناء ملف الإكسل النهائي (المحرك الرئيسي) ---
def build_excel(df, days, periods, statuses):
    output = io.BytesIO()
    columns_order = [
        "الرقم", "الاسم", "رقم الواتس اب", "المجموعة",
        "البلد", "المواليد", "الإجازة", "المعلمة",
        "الحالة", "يوم الاختبار", "توقيت الاختبار", "الفترة", "الملاحظات"
    ]

    for col in columns_order:
        if col not in df.columns: df[col] = ""
    
    df = df[columns_order]
    num_data_rows = len(df)
    extra_rows = 50 
    total_rows = num_data_rows + extra_rows

    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    ws = workbook.add_worksheet("كشف الاختبار")
    ws.right_to_left()

    # الأنماط والتنسيقات
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#2d5016', 'font_color': 'white', 'border': 1, 'align': 'center'})
    locked_fmt = workbook.add_format({'bg_color': '#f2f2f2', 'border': 1, 'align': 'center', 'locked': True})
    unlocked_fmt = workbook.add_format({'bg_color': '#ffffff', 'border': 1, 'align': 'center', 'locked': False})
    alert_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}) # لتلوين كلمة "شرطي"

    # ضبط عرض الأعمدة
    col_widths = [10, 25, 15, 15, 12, 10, 10, 15, 20, 15, 15, 15, 25]
    for i, w in enumerate(col_widths): ws.set_column(i, i, w)

    # كتابة العناوين
    for col_num, header in enumerate(columns_order):
        ws.write(0, col_num, header, header_fmt)

    # كتابة البيانات الأساسية (المقفلة) والأعمدة الجديدة (المفتوحة)
    for row_num, (_, row_data) in enumerate(df.iterrows(), start=1):
        for col_num, value in enumerate(row_data):
            fmt = locked_fmt if col_num < 8 else unlocked_fmt
            ws.write(row_num, col_num, value, fmt)

    # إضافة الصفوف الفارغة (للطالبات الجدد)
    for row_num in range(num_data_rows + 1, total_rows + 1):
        for col_num in range(13):
            fmt = locked_fmt if col_num < 8 else unlocked_fmt
            ws.write(row_num, col_num, "", fmt)

    # القوائم المنسدلة
    ws.data_validation(1, 8, total_rows, 8, {'validate': 'list', 'source': statuses})
    ws.data_validation(1, 9, total_rows, 9, {'validate': 'list', 'source': days})
    ws.data_validation(1, 11, total_rows, 11, {'validate': 'list', 'source': periods})

    # التنسيق الشرطي لخلية الملاحظات (تلوين إذا احتوت على "شرطي")
    ws.conditional_format(1, 12, total_rows, 12, {
        'type': 'cell', 'criteria': 'containing', 'value': 'شرطي', 'format': alert_fmt
    })

    # حماية الورقة
    ws.protect()
    workbook.close()
    return output.getvalue()

# --- واجهة المستخدم ---
st.sidebar.markdown("### ⚙️ إعدادات الجداول")
days_list = parse_list(st.sidebar.text_area("📅 أيام الأسبوع", "الاثنين 3/23\nالثلاثاء 3/24\nالأربعاء 3/25\nالخميس 3/26\nالجمعة 3/27\nالسبت 3/28\nالأحد 3/29"))
periods_list = parse_list(st.sidebar.text_area("⏰ الفترات", "فجراً من 5.45-9.00\nضحى 9:15-12.30\nظهراً 12:45-4.15\nعصراً 4.30-7.00\nليلاً 7.15-9.30"))
status_list = parse_list(st.sidebar.text_area("📋 قائمة الحالات", "أنهت المقرر\nلم تنه المقرر\nساكنة\nمنسحبة\nأخرجتها الإدارة لأنها مخالفة\nلا يوجد واتس\nتم نقلها لغير مجموعة"))

st.markdown('<h1 class="main-title">📖 نظام أتمتة جداول المقرأة</h1>', unsafe_allow_html=True)

uploaded_files = st.file_uploader("ارفعي ملفات الإكسل الخام هنا", type=["xlsx", "csv"], accept_multiple_files=True)

if uploaded_files and st.button("⚡ توليد الجداول النهائية", type="primary"):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for uf in uploaded_files:
            try:
                content = uf.read()
                df = pd.read_csv(io.BytesIO(content)) if uf.name.endswith('.csv') else read_xlsx_raw(content)
                
                # استخراج اسم المعلمة
                teacher_val = df['المعلمة'].iloc[0] if 'المعلمة' in df.columns else "معلمة"
                final_filename = f"{get_short_name(teacher_val)}.xlsx"
                
                # بناء الملف
                xlsx_data = build_excel(df, days_list, periods_list, status_list)
                zf.writestr(final_filename, xlsx_data)
            except Exception as e:
                st.error(f"خطأ في معالجة {uf.name}: {e}")
    
    st.success("✅ تمت المعالجة بنجاح!")
    st.download_button("📥 تحميل كافة الجداول (ZIP)", zip_buffer.getvalue(), "جداول_المعلمات.zip", "application/zip")
