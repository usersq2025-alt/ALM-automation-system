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

# --- محرك قراءة XML الخام (لتجاوز أخطاء التنسيق تماماً) ---
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
    max_len = max(len(r) for r in rows_data)
    for r in rows_data:
        while len(r) < max_len: r.append("")
    return pd.DataFrame(rows_data[1:], columns=rows_data[0])

# --- منطق اختصار الاسم (إيمان زياد -> إيمان.ز) ---
def get_short_name(full_name):
    parts = str(full_name).strip().split()
    if len(parts) >= 2:
        return f"{parts[0]}.{parts[1][0]}"
    return parts[0] if parts else "معلمة"

# --- بناء ملف الإكسل بالهيكل المطلوب (13 عموداً) ---
def build_excel(df, days, periods, statuses):
    output = io.BytesIO()
    # ترتيب الأعمدة الـ 13 المطلوبة حرفياً كما في ملف "معلمة.xlsx"
    target_columns = [
        "الرقم", "الاسم", "رقم الواتس اب", "المجموعة",
        "البلد", "المواليد", "الإجازة", "المعلمة",
        "الحالة", "يوم الاختبار", "توقيت الاختبار", "الفترة", "الملاحظات"
    ]

    # إعادة هيكلة البيانات: نأخذ أول 8 أعمدة من الملف الأصلي ونصفر الباقي
    final_df = pd.DataFrame(columns=target_columns)
    for i in range(8):
        col_name = target_columns[i]
        final_df[col_name] = df.iloc[:, i] if i < df.shape[1] else ""
    
    # الأعمدة الـ 5 الأخيرة يجب أن تكون فارغة تماماً لتملاها المعلمة
    for i in range(8, 13):
        final_df[target_columns[i]] = ""

    num_data_rows = len(final_df)
    total_rows = num_data_rows + 50 # إضافة صفوف للطالبات الجدد

    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    ws = workbook.add_worksheet("كشف الاختبار")
    ws.right_to_left()

    # التنسيقات
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#2d5016', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    locked_fmt = workbook.add_format({'bg_color': '#F2F2F2', 'border': 1, 'align': 'center', 'locked': True})
    unlocked_fmt = workbook.add_format({'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'locked': False})
    alert_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}) # لتلوين "شرطي"

    # عرض الأعمدة
    widths = [10, 30, 18, 15, 12, 10, 10, 15, 20, 18, 18, 18, 30]
    for i, w in enumerate(widths): ws.set_column(i, i, w)

    # كتابة العناوين
    for col_num, header in enumerate(target_columns):
        ws.write(0, col_num, header, header_fmt)

    # كتابة البيانات والحماية
    for r_idx in range(total_rows):
        excel_row = r_idx + 1
        for c_idx in range(13):
            val = final_df.iloc[r_idx, c_idx] if r_idx < num_data_rows else ""
            fmt = locked_fmt if c_idx < 8 else unlocked_fmt
            ws.write(excel_row, c_idx, val, fmt)

    # القوائم المنسدلة
    ws.data_validation(1, 8, total_rows, 8, {'validate': 'list', 'source': statuses})
    ws.data_validation(1, 9, total_rows, 9, {'validate': 'list', 'source': days})
    ws.data_validation(1, 11, total_rows, 11, {'validate': 'list', 'source': periods})

    # التنسيق الشرطي لخلية الملاحظات (تلوين إذا احتوت على "شرطي")
    ws.conditional_format(1, 12, total_rows, 12, {
        'type': 'cell', 'criteria': 'containing', 'value': 'شرطي', 'format': alert_fmt
    })

    ws.protect()
    workbook.close()
    return output.getvalue()

# --- واجهة التطبيق ---
st.sidebar.markdown("### ⚙️ إعدادات الجداول")
days = [line.strip() for line in st.sidebar.text_area("📅 الأيام", "الاثنين 3/23\nالثلاثاء 3/24\nالأربعاء 3/25\nالخميس 3/26\nالجمعة 3/27\nالسبت 3/28\nالأحد 3/29").splitlines() if line.strip()]
periods = [line.strip() for line in st.sidebar.text_area("⏰ الفترات", "فجراً من 5.45-9.00\nضحى 9:15-12.30\nظهراً 12:45-4.15\nعصراً 4.30-7.00\nليلاً 7.15-9.30").splitlines() if line.strip()]
statuses = [line.strip() for line in st.sidebar.text_area("📋 الحالات", "أنهت المقرر\nلم تنه المقرر\nساكنة\nمنسحبة\nأخرجتها الإدارة لأنها مخالفة\nلا يوجد واتس\nتم نقلها لغير مجموعة").splitlines() if line.strip()]

st.title("📖 نظام أتمتة جداول المقرأة")
files = st.file_uploader("ارفعي ملفات الموقع الخام", type=["xlsx", "csv"], accept_multiple_files=True)

if files and st.button("🚀 توليد الجداول النهائية", type="primary"):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for f in files:
            try:
                content = f.read()
                df = pd.read_csv(io.BytesIO(content)) if f.name.endswith('.csv') else read_xlsx_raw(content)
                teacher = df.iloc[0, 7] if df.shape[1] > 7 else "معلمة"
                short_name = get_short_name(teacher)
                xlsx_data = build_excel(df, days, periods, statuses)
                zf.writestr(f"{short_name}.xlsx", xlsx_data)
            except Exception as e:
                st.error(f"خطأ في {f.name}: {e}")
    
    st.success("✅ تمت المعالجة بنجاح!")
    st.download_button("📥 تحميل الجداول (ZIP)", zip_buffer.getvalue(), "جداول_المعلمات.zip", "application/zip")
