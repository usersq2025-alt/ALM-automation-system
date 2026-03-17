import streamlit as st
import pandas as pd
import io
import zipfile
import xlsxwriter
import xml.etree.ElementTree as ET

st.set_page_config(
    page_title="أداة مقرأة",
    page_icon="📖",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700;900&display=swap');
    html, body, [class*="css"] { font-family: 'Tajawal', sans-serif; direction: rtl; }
    .stApp { background: linear-gradient(135deg, #f8f4ef 0%, #ede8e0 100%); }
    h1, h2, h3 { font-family: 'Tajawal', sans-serif !important; }
    .hero-header {
        background: linear-gradient(135deg, #2d5016 0%, #4a7c28 50%, #6ba535 100%);
        border-radius: 16px; padding: 2rem 2.5rem; margin-bottom: 2rem;
        box-shadow: 0 8px 32px rgba(45,80,22,0.25); text-align: center; color: white;
    }
    .hero-header h1 { font-size: 2.4rem; font-weight: 900; margin: 0; text-shadow: 0 2px 4px rgba(0,0,0,0.2); }
    .hero-header p { font-size: 1.05rem; margin: 0.5rem 0 0; opacity: 0.88; font-weight: 300; }
    .stat-card {
        background: white; border-radius: 12px; padding: 1.2rem 1.5rem;
        box-shadow: 0 2px 12px rgba(0,0,0,0.07); border-right: 4px solid #4a7c28; margin-bottom: 1rem;
    }
    .stat-card .number { font-size: 2rem; font-weight: 900; color: #2d5016; line-height: 1; }
    .stat-card .label { font-size: 0.85rem; color: #777; margin-top: 4px; }
    .file-chip {
        display: inline-block; background: #e8f5e0; border: 1px solid #a8d878;
        color: #2d5016; border-radius: 20px; padding: 4px 14px;
        font-size: 0.82rem; margin: 3px; font-weight: 500;
    }
    .success-banner {
        background: linear-gradient(90deg, #e8f5e0, #d4edbe); border: 1px solid #a8d878;
        border-radius: 10px; padding: 1rem 1.5rem; color: #2d5016;
        font-weight: 600; font-size: 1.05rem; margin: 1rem 0;
    }
    .section-title {
        font-size: 1.1rem; font-weight: 700; color: #2d5016;
        border-bottom: 2px solid #a8d878; padding-bottom: 6px; margin: 1.5rem 0 1rem;
    }
    [data-testid="stSidebar"] { background: linear-gradient(180deg, #1a3a08 0%, #2d5016 100%) !important; }
    [data-testid="stSidebar"] * { color: #d8f0b8 !important; }
    [data-testid="stSidebar"] .stTextArea textarea {
        background: rgba(255,255,255,0.1) !important;
        border: 1px solid rgba(168,216,120,0.4) !important;
        color: #f0f8e8 !important;
        font-family: 'Tajawal', sans-serif !important;
        direction: rtl;
    }
    [data-testid="stSidebar"] label { font-weight: 600 !important; font-size: 0.9rem !important; }
    .stButton > button { font-family: 'Tajawal', sans-serif !important; font-weight: 700 !important; border-radius: 10px !important; }
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #2d5016, #4a7c28) !important;
        border: none !important; box-shadow: 0 4px 15px rgba(45,80,22,0.3) !important;
    }
    .stDownloadButton > button {
        background: linear-gradient(135deg, #1a5276, #2874a6) !important;
        color: white !important; font-family: 'Tajawal', sans-serif !important;
        font-weight: 700 !important; border: none !important; border-radius: 10px !important;
        padding: 0.6rem 2rem !important; font-size: 1rem !important;
        box-shadow: 0 4px 15px rgba(26,82,118,0.3) !important;
    }
    .upload-zone {
        background: white; border: 2px dashed #a8d878; border-radius: 16px;
        padding: 2rem; text-align: center; margin: 1rem 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def col_letter_to_index(col_str):
    """Convert Excel column letter(s) to 0-based index. e.g. A->0, B->1, Z->25, AA->26"""
    idx = 0
    for ch in col_str.upper():
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1


def read_xlsx_raw(file_bytes):
    """
    Reads .xlsx by parsing XML directly — handles both 'inlineStr' and shared-string cells.
    Completely bypasses openpyxl styles/Fill to avoid Fill errors.
    """
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:

        # 1. Shared strings table
        shared = []
        if "xl/sharedStrings.xml" in zf.namelist():
            tree = ET.parse(zf.open("xl/sharedStrings.xml"))
            for si in tree.getroot().iter("{" + NS + "}si"):
                texts = [t.text or "" for t in si.iter("{" + NS + "}t")]
                shared.append("".join(texts))

        # 2. Resolve first sheet path via relationships
        rels_tree = ET.parse(zf.open("xl/_rels/workbook.xml.rels"))
        rels = {r.attrib["Id"]: r.attrib["Target"] for r in rels_tree.getroot()}

        wb_tree = ET.parse(zf.open("xl/workbook.xml"))
        sheets_el = wb_tree.getroot().find("{" + NS + "}sheets")
        first_sheet = sheets_el[0]
        r_id = first_sheet.attrib.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        )
        target = rels[r_id]
        if target.startswith("/xl/"):
            sheet_path = target[1:]
        elif target.startswith("xl/"):
            sheet_path = target
        else:
            sheet_path = "xl/" + target

        # 3. Parse sheet rows — handle inlineStr, shared string, and numeric cells
        sheet_tree = ET.parse(zf.open(sheet_path))
        sheet_data = sheet_tree.getroot().find("{" + NS + "}sheetData")

        rows_dict = {}
        max_col = 0

        for row_el in sheet_data.iter("{" + NS + "}row"):
            row_num = int(row_el.attrib.get("r", 0)) - 1  # 0-based
            row_dict = {}

            for c in row_el.iter("{" + NS + "}c"):
                addr = c.attrib.get("r", "A1")
                # Extract column letters from address like "AB12"
                col_letters = "".join(ch for ch in addr if ch.isalpha())
                col_idx = col_letter_to_index(col_letters)
                max_col = max(max_col, col_idx)

                cell_type = c.attrib.get("t", "")

                if cell_type == "inlineStr":
                    # Text stored directly inside <is><t>...</t></is>
                    is_el = c.find("{" + NS + "}is")
                    if is_el is not None:
                        texts = [t.text or "" for t in is_el.iter("{" + NS + "}t")]
                        row_dict[col_idx] = "".join(texts)
                    else:
                        row_dict[col_idx] = ""

                elif cell_type == "s":
                    # Shared string index
                    v_el = c.find("{" + NS + "}v")
                    if v_el is not None and v_el.text is not None:
                        try:
                            row_dict[col_idx] = shared[int(v_el.text)]
                        except (IndexError, ValueError):
                            row_dict[col_idx] = v_el.text
                    else:
                        row_dict[col_idx] = ""

                elif cell_type == "b":
                    v_el = c.find("{" + NS + "}v")
                    row_dict[col_idx] = bool(int(v_el.text)) if v_el is not None else ""

                else:
                    # Numeric or formula result
                    v_el = c.find("{" + NS + "}v")
                    if v_el is not None and v_el.text is not None:
                        val = v_el.text
                        try:
                            val = int(val) if "." not in val else float(val)
                        except (ValueError, TypeError):
                            pass
                        row_dict[col_idx] = val
                    else:
                        row_dict[col_idx] = ""

            rows_dict[row_num] = row_dict

        if not rows_dict:
            raise ValueError("الملف فارغ")

        # Build full matrix
        max_row = max(rows_dict.keys())
        matrix = []
        for r in range(max_row + 1):
            row_data = rows_dict.get(r, {})
            matrix.append([row_data.get(c, "") for c in range(max_col + 1)])

        # First row as headers
        headers = [str(v).strip() if v != "" else "col_" + str(i)
                   for i, v in enumerate(matrix[0])]
        return pd.DataFrame(matrix[1:], columns=headers)


def get_first_name(full_name):
    """إيمان زياد الحموي → إيمان"""
    parts = str(full_name).strip().split()
    return parts[0] if parts else str(full_name)


def parse_list(text):
    return [line.strip() for line in text.strip().splitlines() if line.strip()]


def build_excel(df, days, periods, statuses):
    output = io.BytesIO()

    columns_order = [
        "الرقم", "الاسم", "رقم الواتس اب", "المجموعة",
        "البلد", "المواليد", "الإجازة", "المعلمة",
        "الحالة", "يوم الاختبار", "توقيت الاختبار", "الفترة", "الملاحظات",
    ]

    # Numeric columns (A, C, F) — store as numbers
    numeric_cols = {"الرقم", "رقم الواتس اب", "المواليد"}

    for col in columns_order:
        if col not in df.columns:
            df[col] = ""

    df = df[columns_order].copy()
    num_rows = len(df)
    extra_rows = 50

    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    ws = workbook.add_worksheet("الطالبات")
    ws.right_to_left()

    # ── Formats matching ابتسام_2 exactly ─────────────────────────────────────
    # Header: Calibri 11 Bold, white bg, thin border, center, locked
    header_fmt = workbook.add_format({
        "bold": True,
        "font_name": "Calibri", "font_size": 11,
        "align": "center", "valign": "vcenter",
        "border": 1,
        "locked": True,
    })

    # Locked text cells (A–H text): Calibri 11, thin border, center, locked
    locked_text_fmt = workbook.add_format({
        "font_name": "Calibri", "font_size": 11,
        "align": "center", "valign": "vcenter",
        "border": 1,
        "locked": True,
    })

    # Locked numeric cells (الرقم, رقم الواتس اب, المواليد): same + num format
    locked_num_fmt = workbook.add_format({
        "font_name": "Calibri", "font_size": 11,
        "align": "center", "valign": "vcenter",
        "border": 1,
        "locked": True,
        "num_format": "0",
    })

    # Unlocked cells I–L (Calibri): thin border, center, unlocked
    unlocked_fmt = workbook.add_format({
        "font_name": "Calibri", "font_size": 11,
        "align": "center", "valign": "vcenter",
        "border": 1,
        "locked": False,
    })

    # Unlocked cell L (الفترة) uses Arial per reference file
    unlocked_arial_fmt = workbook.add_format({
        "font_name": "Arial", "font_size": 11,
        "align": "center", "valign": "vcenter",
        "border": 1,
        "locked": False,
    })

    # Unlocked M (الملاحظات): wider, unlocked, border
    unlocked_notes_fmt = workbook.add_format({
        "font_name": "Calibri", "font_size": 11,
        "align": "center", "valign": "vcenter",
        "border": 1,
        "locked": False,
    })

    # ── Column widths matching ابتسام_2 exactly ───────────────────────────────
    # A=7, B=24, C=14.1, D=13.3, E=7, F=6, G=5.3, H=6.9
    # I=19.8, J=11.4, K=10.7, L=14, M=39.8
    col_widths = [7, 24, 14.1, 13.3, 7, 6, 5.3, 6.9, 19.8, 11.4, 10.7, 14, 39.8]
    for i, w in enumerate(col_widths):
        ws.set_column(i, i, w)

    # ── Header row ────────────────────────────────────────────────────────────
    for col_idx, col_name in enumerate(columns_order):
        ws.write(0, col_idx, col_name, header_fmt)

    # ── Data rows ─────────────────────────────────────────────────────────────
    for row_idx, row in df.iterrows():
        excel_row = row_idx + 1
        for col_idx, col_name in enumerate(columns_order):
            val = row[col_name]
            val = "" if pd.isna(val) else val

            if col_idx < 8:
                # Locked columns A–H
                if col_name in numeric_cols and val != "":
                    try:
                        val = int(str(val).replace(".0", ""))
                    except (ValueError, TypeError):
                        pass
                    ws.write(excel_row, col_idx, val, locked_num_fmt)
                else:
                    ws.write(excel_row, col_idx, str(val) if val != "" else "", locked_text_fmt)
            else:
                # Unlocked columns I–M
                if col_idx == 11:  # L = الفترة → Arial
                    ws.write(excel_row, col_idx, val, unlocked_arial_fmt)
                elif col_idx == 12:  # M = الملاحظات
                    ws.write(excel_row, col_idx, val, unlocked_notes_fmt)
                else:
                    ws.write(excel_row, col_idx, val, unlocked_fmt)

    # ── Extra blank rows (50) ─────────────────────────────────────────────────
    for extra in range(extra_rows):
        excel_row = num_rows + 1 + extra
        for col_idx in range(13):
            if col_idx < 8:
                ws.write(excel_row, col_idx, "", locked_text_fmt)
            elif col_idx == 11:
                ws.write(excel_row, col_idx, "", unlocked_arial_fmt)
            else:
                ws.write(excel_row, col_idx, "", unlocked_fmt)

    last_val_row = num_rows + extra_rows

    # ── Data validation ───────────────────────────────────────────────────────
    ws.data_validation(1, 8, last_val_row, 8, {
        "validate": "list", "source": statuses,
        "show_input": True, "show_error": True,
    })
    ws.data_validation(1, 9, last_val_row, 9, {
        "validate": "list", "source": days,
        "show_input": True, "show_error": True,
    })
    ws.data_validation(1, 11, last_val_row, 11, {
        "validate": "list", "source": periods,
        "show_input": True, "show_error": True,
    })

    # ── Sheet protection ──────────────────────────────────────────────────────
    ws.protect("", {
        "sheet": True,
        "objects": True,
        "scenarios": True,
        "insert_rows": True,
        "insert_columns": False,
        "delete_rows": False,
        "sort": False,
        "autofilter": False,
        "select_locked_cells": True,
        "select_unlocked_cells": True,
    })

    workbook.close()
    output.seek(0)
    return output.read()


def process_files(uploaded_files, days, periods, statuses):
    results = {}
    errors = []

    for uf in uploaded_files:
        try:
            file_bytes = uf.read()
            name_lower = uf.name.lower()

            if name_lower.endswith(".csv"):
                df = pd.read_csv(io.BytesIO(file_bytes))
            elif name_lower.endswith(".xls"):
                df = pd.read_excel(io.BytesIO(file_bytes), engine="xlrd")
            else:
                df = read_xlsx_raw(file_bytes)

            teacher_col = next((c for c in df.columns if "المعلمة" in str(c)), None)
            if teacher_col and not df[teacher_col].dropna().empty:
                raw_name = str(df[teacher_col].dropna().iloc[0]).strip()
                first_name = get_first_name(raw_name)
                # Write only first name in the المعلمة column
                df[teacher_col] = first_name
                short = first_name
            else:
                short = uf.name.rsplit(".", 1)[0]

            xlsx_bytes = build_excel(df.copy(), days, periods, statuses)
            out_name = short + ".xlsx"
            base = out_name
            counter = 1
            while out_name in results:
                out_name = base.replace(".xlsx", "_" + str(counter) + ".xlsx")
                counter += 1

            results[out_name] = xlsx_bytes

        except Exception as e:
            errors.append("❌ " + uf.name + ": " + str(e))

    return results, errors


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(
        """
        <div style='text-align:center; padding:1rem 0 0.5rem;'>
            <div style='font-size:2.5rem'>📖</div>
            <div style='font-size:1.2rem; font-weight:900; color:#d8f0b8;'>إعدادات الدورة</div>
            <div style='font-size:0.8rem; color:#a8d878; margin-top:4px;'>خصّصي القيم لكل دورة</div>
        </div>
        <hr style='border-color:rgba(168,216,120,0.3); margin:0.8rem 0;'>
        """,
        unsafe_allow_html=True,
    )

    days_text = st.text_area(
        "📅 أيام الأسبوع",
        value="الاثنين 3/23\nالثلاثاء 3/24\nالأربعاء 3/25\nالخميس 3/26\nالجمعة 3/27\nالسبت 3/28\nالأحد 3/29",
        height=160,
        help="كل يوم في سطر منفصل",
    )
    periods_text = st.text_area(
        "⏰ الفترات",
        value="فجراً\nضحى\nظهراً\nعصراً\nليلاً",
        height=140,
        help="كل فترة في سطر منفصل",
    )
    statuses_text = st.text_area(
        "📋 قائمة الحالات",
        value="أنهت المقرر\nلم تنه المقرر\nساكنة\nمنسحبة\nأخرجتها الإدارة لأنها مخالفة\nلا يوجد واتس\nتم نقلها لغير مجموعة",
        height=175,
        help="كل حالة في سطر منفصل",
    )

    days_list = parse_list(days_text)
    periods_list = parse_list(periods_text)
    statuses_list = parse_list(statuses_text)

    st.markdown(
        "<div style='margin-top:1rem; padding:0.8rem; background:rgba(255,255,255,0.08);"
        "border-radius:8px; font-size:0.82rem; color:#a8d878;'>"
        "✅ " + str(len(days_list)) + " أيام &nbsp;|&nbsp; ✅ "
        + str(len(periods_list)) + " فترات &nbsp;|&nbsp; ✅ "
        + str(len(statuses_list)) + " حالة</div>",
        unsafe_allow_html=True,
    )


# ── Main ──────────────────────────────────────────────────────────────────────
st.markdown(
    """
    <div class="hero-header">
        <h1>📖 أداة أتمتة جداول مقرأة</h1>
        <p>ارفعي ملفات Excel أو CSV الخام وستحصلين على جداول منسقة، محمية، وجاهزة للمعلمات</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="section-title">📂 رفع الملفات</div>', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "اسحبي الملفات هنا أو اضغطي للاختيار",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if uploaded_files:
    cols = st.columns(4)
    with cols[0]:
        st.markdown(
            '<div class="stat-card"><div class="number">' + str(len(uploaded_files))
            + '</div><div class="label">ملف مرفوع</div></div>',
            unsafe_allow_html=True,
        )
    with cols[1]:
        st.markdown(
            '<div class="stat-card"><div class="number">' + str(len(days_list))
            + '</div><div class="label">أيام الاختبار</div></div>',
            unsafe_allow_html=True,
        )
    with cols[2]:
        st.markdown(
            '<div class="stat-card"><div class="number">' + str(len(periods_list))
            + '</div><div class="label">فترة متاحة</div></div>',
            unsafe_allow_html=True,
        )
    with cols[3]:
        st.markdown(
            '<div class="stat-card"><div class="number">' + str(len(statuses_list))
            + '</div><div class="label">حالة في القائمة</div></div>',
            unsafe_allow_html=True,
        )

    chips = " ".join(
        ['<span class="file-chip">📄 ' + f.name + "</span>" for f in uploaded_files]
    )
    st.markdown(
        "<div style='margin:0.5rem 0 1.5rem'>" + chips + "</div>",
        unsafe_allow_html=True,
    )

    if st.button("⚡ معالجة الملفات وتوليد جداول المعلمات", type="primary", use_container_width=True):
        with st.spinner("جارٍ المعالجة..."):
            results, errors = process_files(uploaded_files, days_list, periods_list, statuses_list)

        for e in errors:
            st.error(e)

        if results:
            st.markdown(
                '<div class="success-banner">✅ تمت معالجة ' + str(len(results)) + ' ملف بنجاح!</div>',
                unsafe_allow_html=True,
            )

            st.markdown('<div class="section-title">👁️ معاينة الملفات الناتجة</div>', unsafe_allow_html=True)
            preview_data = [{"اسم الملف الناتج": fname, "الحالة": "✅ جاهز"} for fname in results]
            st.dataframe(pd.DataFrame(preview_data), use_container_width=True, hide_index=True)

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for fname, fbytes in results.items():
                    zf.writestr(fname, fbytes)
            zip_buffer.seek(0)

            st.markdown('<div class="section-title">📦 تحميل الملفات</div>', unsafe_allow_html=True)
            st.download_button(
                label="⬇️ تحميل جميع الملفات (" + str(len(results)) + " ملف) — ZIP",
                data=zip_buffer,
                file_name="جداول_المعلمات.zip",
                mime="application/zip",
                use_container_width=True,
            )

else:
    st.markdown(
        """
        <div class="upload-zone">
            <div style="font-size:3rem; margin-bottom:0.5rem">📂</div>
            <div style="font-size:1.1rem; font-weight:600; color:#4a7c28;">لم يتم رفع أي ملفات بعد</div>
            <div style="font-size:0.85rem; color:#888; margin-top:0.3rem;">يدعم الصيغ: xlsx, xls, csv</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown(
    """
    <hr style="margin:2rem 0 1rem; border-color:#d0e8c0;">
    <div style="text-align:center; color:#999; font-size:0.8rem; font-family:'Tajawal',sans-serif;">
        أداة مقرأة — مبنية بـ Python & Streamlit &nbsp;|&nbsp; 📖
    </div>
    """,
    unsafe_allow_html=True,
)
