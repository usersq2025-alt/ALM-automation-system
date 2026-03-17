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
    .stApp { background: linear-gradient(135deg, #f3f0f8 0%, #e8e0f0 100%); }
    h1, h2, h3 { font-family: 'Tajawal', sans-serif !important; }

    /* ── Hero ── */
    .hero-header {
        background: linear-gradient(135deg, #3d2060 0%, #6b3fa0 50%, #8b5cc8 100%);
        border-radius: 16px; padding: 2rem 2.5rem; margin-bottom: 2rem;
        box-shadow: 0 8px 32px rgba(61,32,96,0.3); text-align: center; color: white;
    }
    .hero-header h1 { font-size: 2.4rem; font-weight: 900; margin: 0; text-shadow: 0 2px 4px rgba(0,0,0,0.2); }
    .hero-header p { font-size: 1.05rem; margin: 0.5rem 0 0; opacity: 0.88; font-weight: 300; }

    /* ── Cards ── */
    .stat-card {
        background: white; border-radius: 12px; padding: 1.2rem 1.5rem;
        box-shadow: 0 2px 12px rgba(0,0,0,0.07); border-right: 4px solid #6b3fa0; margin-bottom: 1rem;
    }
    .stat-card .number { font-size: 2rem; font-weight: 900; color: #3d2060; line-height: 1; }
    .stat-card .label { font-size: 0.85rem; color: #777; margin-top: 4px; }

    .file-chip {
        display: inline-block; background: #f0e8fb; border: 1px solid #c4a0e8;
        color: #3d2060; border-radius: 20px; padding: 4px 14px;
        font-size: 0.82rem; margin: 3px; font-weight: 500;
    }
    .success-banner {
        background: linear-gradient(90deg, #f0e8fb, #e0d0f8); border: 1px solid #c4a0e8;
        border-radius: 10px; padding: 1rem 1.5rem; color: #3d2060;
        font-weight: 600; font-size: 1.05rem; margin: 1rem 0;
    }
    .section-title {
        font-size: 1.1rem; font-weight: 700; color: #3d2060;
        border-bottom: 2px solid #c4a0e8; padding-bottom: 6px; margin: 1.5rem 0 1rem;
    }
    .upload-zone {
        background: white; border: 2px dashed #c4a0e8; border-radius: 16px;
        padding: 2rem; text-align: center; margin: 1rem 0;
    }

    /* ── Sidebar background ── */
    [data-testid="stSidebar"] > div:first-child {
        background: linear-gradient(180deg, #2d1b4e 0%, #3d2060 100%) !important;
    }

    /* ── Sidebar — force ALL text to light purple ── */
    [data-testid="stSidebar"],
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] div,
    [data-testid="stSidebar"] label {
        color: #e8d5f8 !important;
        font-family: 'Tajawal', sans-serif !important;
    }
    [data-testid="stSidebar"] label {
        font-weight: 700 !important;
        font-size: 0.95rem !important;
    }

    /* ── Sidebar textarea — dark bg + white text ── */
    [data-testid="stSidebar"] textarea,
    [data-testid="stSidebar"] .stTextArea textarea {
        background-color: #1e1035 !important;
        border: 2px solid #9b6fd4 !important;
        border-radius: 8px !important;
        color: #f0e6ff !important;
        font-family: 'Tajawal', sans-serif !important;
        font-size: 0.92rem !important;
        direction: rtl !important;
        caret-color: #e8d5f8 !important;
    }
    [data-testid="stSidebar"] textarea:focus,
    [data-testid="stSidebar"] .stTextArea textarea:focus {
        border-color: #c4a0e8 !important;
        box-shadow: 0 0 0 2px rgba(196,160,232,0.3) !important;
    }
    [data-testid="stSidebar"] textarea::placeholder {
        color: #9b7dbf !important;
    }

    /* ── Sidebar tooltip/help text ── */
    [data-testid="stSidebar"] small,
    [data-testid="stSidebar"] .stMarkdown {
        color: #c4a0e8 !important;
    }

    /* ── Buttons ── */
    .stButton > button {
        font-family: 'Tajawal', sans-serif !important;
        font-weight: 700 !important;
        border-radius: 10px !important;
    }
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #3d2060, #6b3fa0) !important;
        border: none !important;
        color: white !important;
        box-shadow: 0 4px 15px rgba(45,27,78,0.3) !important;
    }
    .stDownloadButton > button {
        background: linear-gradient(135deg, #1a5276, #2874a6) !important;
        color: white !important; font-family: 'Tajawal', sans-serif !important;
        font-weight: 700 !important; border: none !important; border-radius: 10px !important;
        padding: 0.6rem 2rem !important; font-size: 1rem !important;
        box-shadow: 0 4px 15px rgba(26,82,118,0.3) !important;
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



def build_teacher_display_names(full_names_text):
    """
    تأخذ نصاً فيه اسم معلمة في كل سطر (الاسم الكامل)
    وترجع dict: {اسم_كامل: اسم_للعرض}
    القاعدة:
      - اسم فريد → الاسم الأول فقط:  ابتسام
      - اسم مكرر → أول حرف من الكنية: إيمان.ح / إيمان.ن
      - تعارض → يزداد عدد الأحرف: آلاء.شي / آلاء.شب / دعاء.سي / دعاء.سل
    """
    lines = [l.strip() for l in full_names_text.strip().splitlines() if l.strip()]
    if not lines:
        return {}

    def make_display(parts, n_chars=1):
        if len(parts) == 1:
            return parts[0]
        last = parts[-1]
        suffix = last[2:] if (last.startswith("ال") and len(last) > 2) else last
        return parts[0] + "." + suffix[:n_chars]

    first_count = {}
    for name in lines:
        f = name.split()[0] if name.split() else name
        first_count[f] = first_count.get(f, 0) + 1

    result = {}
    # الأسماء الفريدة أولاً
    for name in lines:
        parts = name.split()
        if not parts: continue
        if first_count[parts[0]] == 1:
            result[name] = parts[0]

    # الأسماء المكررة — زد الأحرف تدريجياً حتى لا يوجد تعارض
    duplicates = [n for n in lines if n not in result]
    for max_chars in range(1, 6):
        temp = {}
        for name in duplicates:
            if name in result: continue
            temp[name] = make_display(name.split(), max_chars)

        all_d = dict(result)
        all_d.update(temp)
        count = {}
        for d in all_d.values():
            count[d] = count.get(d, 0) + 1

        for name, display in temp.items():
            if count[display] == 1:
                result[name] = display

        if len(result) == len(lines):
            break

    # ما تبقى بدون حل → اسم كامل
    for name in lines:
        if name not in result:
            result[name] = name

    return result

def get_first_name(full_name):
    """ابتسام خالد سمونة → ابتسام"""
    parts = str(full_name).strip().split()
    return parts[0] if parts else str(full_name)


def get_teacher_name(full_name):
    """ابتسام خالد سمونة → ابتسام سمونة"""
    name_str = str(full_name).strip()
    parts = name_str.split()
    if len(parts) > 1:
        return parts[0] + " " + parts[-1]
    return parts[0] if parts else name_str


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

    # ── Extra blank rows (50) — كلها مفتوحة بالكامل للمعلمة ────────────────────
    for extra in range(extra_rows):
        excel_row = num_rows + 1 + extra
        for col_idx in range(13):
            if col_idx == 11:  # L = الفترة → Arial
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



def extract_teacher_names(uploaded_files):
    """
    يقرأ كل الملفات، يستخرج أسماء المعلمات، ويبني القاموس المقترح.
    يُعيد (preview_map, file_bytes_cache, errors)
    file_bytes_cache: {uf.name: bytes} لتجنب إعادة القراءة لاحقاً
    """
    errors = []
    raw_names_ordered = []
    file_bytes_cache  = {}

    for uf in uploaded_files:
        try:
            uf.seek(0)
            fb = uf.read()
            file_bytes_cache[uf.name] = fb          # احفظ الـ bytes
            name_lower = uf.name.lower()
            if name_lower.endswith(".csv"):
                df = pd.read_csv(io.BytesIO(fb))
            elif name_lower.endswith(".xls"):
                df = pd.read_excel(io.BytesIO(fb), engine="xlrd")
            else:
                df = read_xlsx_raw(fb)
            teacher_col = next((c for c in df.columns if "المعلمة" in str(c)), None)
            raw = ""
            if teacher_col and not df[teacher_col].dropna().empty:
                raw = str(df[teacher_col].dropna().iloc[0]).strip()
            if raw and raw not in raw_names_ordered:
                raw_names_ordered.append(raw)
        except Exception as e:
            errors.append("❌ " + uf.name + ": " + str(e))

    all_text = chr(10).join(raw_names_ordered)
    preview_map = build_teacher_display_names(all_text)
    return preview_map, file_bytes_cache, errors


def process_files_from_cache(file_bytes_cache, days, periods, statuses, teacher_map=None):
    """
    نفس process_files لكن تعمل على dict {filename: bytes} بدل uploaded_files
    لتجنب إعادة قراءة الملفات بعد انتهاء صلاحية الـ uploader
    """
    results = {}
    errors  = []

    file_data = []
    raw_names = []

    for fname, fb in file_bytes_cache.items():
        try:
            name_lower = fname.lower()
            if name_lower.endswith(".csv"):
                df = pd.read_csv(io.BytesIO(fb))
            elif name_lower.endswith(".xls"):
                df = pd.read_excel(io.BytesIO(fb), engine="xlrd")
            else:
                df = read_xlsx_raw(fb)
            teacher_col = next((c for c in df.columns if "المعلمة" in str(c)), None)
            raw = ""
            if teacher_col and not df[teacher_col].dropna().empty:
                raw = str(df[teacher_col].dropna().iloc[0]).strip()
            file_data.append((fname, df, teacher_col, raw))
            raw_names.append(raw)
        except Exception as e:
            errors.append("❌ " + fname + ": " + str(e))
            file_data.append((fname, None, None, ""))
            raw_names.append("")

    # استخدم teacher_map الممررة (المعتمدة من المشرفة) أو ابنِ آلياً
    if not teacher_map:
        all_names_text = chr(10).join(n for n in raw_names if n)
        teacher_map = build_teacher_display_names(all_names_text)

    for fname, df, teacher_col, raw_name in file_data:
        if df is None:
            continue
        try:
            if raw_name:
                col_h_name        = teacher_map.get(raw_name, get_first_name(raw_name))
                file_name_display = get_teacher_name(raw_name)
                if teacher_col:
                    df[teacher_col] = col_h_name
                short = file_name_display
            else:
                short = fname.rsplit(".", 1)[0]

            xlsx_bytes = build_excel(df.copy(), days, periods, statuses)
            out_name = short + ".xlsx"
            base = out_name
            counter = 1
            while out_name in results:
                out_name = base.replace(".xlsx", "_" + str(counter) + ".xlsx")
                counter += 1
            results[out_name] = xlsx_bytes
        except Exception as e:
            errors.append("❌ " + fname + ": " + str(e))

    return results, errors


def process_files(uploaded_files, days, periods, statuses, teacher_map=None):
    results = {}
    errors = []

    # ── المرور الأول: اقرأ كل الملفات واستخرج الأسماء الكاملة ───────────────
    file_data = []   # [(uf.name, file_bytes, df)]
    raw_names = []   # أسماء المعلمات الكاملة بنفس ترتيب الملفات

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
            raw = ""
            if teacher_col and not df[teacher_col].dropna().empty:
                raw = str(df[teacher_col].dropna().iloc[0]).strip()
            file_data.append((uf.name, file_bytes, df, teacher_col, raw))
            raw_names.append(raw)
        except Exception as e:
            errors.append("❌ " + uf.name + ": " + str(e))
            file_data.append((uf.name, None, None, None, ""))
            raw_names.append("")

    # ── بناء قاموس الأسماء المختصرة ─────────────────────────────────────────
    # إذا أُرسل teacher_map من الواجهة (بعد مراجعة المشرفة) → استخدمه
    # وإلا → ابنِ واحداً آلياً
    if not teacher_map:
        all_names_text = chr(10).join(n for n in raw_names if n)
        teacher_map = build_teacher_display_names(all_names_text)

    # ── المرور الثاني: اعالج كل ملف مع النسق الصحيح ─────────────────────────
    for fname, file_bytes, df, teacher_col, raw_name in file_data:
        if df is None:
            continue
        try:
            if raw_name:
                col_h_name        = (teacher_map or {}).get(raw_name, get_first_name(raw_name))
                file_name_display = get_teacher_name(raw_name)
                if teacher_col:
                    df[teacher_col] = col_h_name
                short = file_name_display
            else:
                short = fname.rsplit(".", 1)[0]

            xlsx_bytes = build_excel(df.copy(), days, periods, statuses)
            out_name = short + ".xlsx"
            base = out_name
            counter = 1
            while out_name in results:
                out_name = base.replace(".xlsx", "_" + str(counter) + ".xlsx")
                counter += 1
            results[out_name] = xlsx_bytes

        except Exception as e:
            errors.append("❌ " + fname + ": " + str(e))

    return results, errors


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(
        """
        <div style='text-align:center; padding:1rem 0 0.5rem;'>
            <div style='font-size:2.5rem'>📖</div>
            <div style='font-size:1.2rem; font-weight:900; color:#e8d5f8;'>إعدادات الدورة</div>
            <div style='font-size:0.8rem; color:#c4a0e8; margin-top:4px;'>خصّصي القيم لكل دورة</div>
        </div>
        <hr style='border-color:rgba(168,216,120,0.3); margin:0.8rem 0;'>
        """,
        unsafe_allow_html=True,
    )

    days_text = st.text_area(
        "📅 أيام الأسبوع",
        value="الإثنين\nالثلاثاء\nالأربعاء\nالخميس\nالجمعة\nالسبت\nالأحد",
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
        "border-radius:8px; font-size:0.82rem; color:#c4a0e8;'>"
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

    # ── خطوة 1: معاينة أسماء المعلمات قبل التنفيذ ──────────────────────────────
    if st.button("🔍 تحليل الملفات ومعاينة أسماء المعلمات", use_container_width=True):
        with st.spinner("جارٍ القراءة..."):
            preview_map, file_bytes_cache, preview_errors = extract_teacher_names(uploaded_files)
        st.session_state["preview_map"]        = preview_map
        st.session_state["file_bytes_cache"]   = file_bytes_cache
        st.session_state["preview_errors"]     = preview_errors
        st.session_state["stage1_results"]     = None
        st.session_state["stage1_errors"]      = None

    # ── عرض جدول المعاينة + حقول التعديل ────────────────────────────────────
    if st.session_state.get("preview_map"):
        preview_map = st.session_state["preview_map"]

        st.markdown('<div class="section-title">👁️ مراجعة أسماء المعلمات</div>', unsafe_allow_html=True)
        st.markdown(
            "<div style='font-size:0.88rem;color:#555;margin-bottom:0.8rem;direction:rtl;'>"
            "الكود اقترح هذه الأسماء تلقائياً — عدّلي أي اسم لا يناسبكِ ثم اضغطي تأكيد."
            "</div>",
            unsafe_allow_html=True,
        )

        edited_map = {}
        cols_h = st.columns([3, 2, 2])
        cols_h[0].markdown("**الاسم الكامل في الملف**")
        cols_h[1].markdown("**الاسم المقترح**")
        cols_h[2].markdown("**الاسم النهائي** (عدّلي هنا)")

        for i, (full_name, suggested) in enumerate(preview_map.items()):
            c1, c2, c3 = st.columns([3, 2, 2])
            c1.markdown(
                "<div style='direction:rtl;padding-top:8px;font-size:0.9rem;'>" + full_name + "</div>",
                unsafe_allow_html=True,
            )
            c2.markdown(
                "<div style='direction:rtl;padding-top:8px;font-size:0.9rem;"
                "font-weight:600;color:#6b3fa0;'>" + suggested + "</div>",
                unsafe_allow_html=True,
            )
            final = c3.text_input(
                "final_" + str(i),
                value=suggested,
                label_visibility="collapsed",
                key="teacher_edit_" + str(i),
            )
            edited_map[full_name] = final.strip() if final.strip() else suggested

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button("⚡ تأكيد ومعالجة الملفات", type="primary", use_container_width=True):
            file_bytes_cache = st.session_state.get("file_bytes_cache", {})
            with st.spinner("جارٍ المعالجة..."):
                results, errors = process_files_from_cache(
                    file_bytes_cache, days_list, periods_list, statuses_list, edited_map
                )
            st.session_state["stage1_results"] = results
            st.session_state["stage1_errors"]  = errors

    # ── عرض النتائج ──────────────────────────────────────────────────────────
    if st.session_state.get("stage1_results"):
        results = st.session_state["stage1_results"]
        errors  = st.session_state.get("stage1_errors", [])

        for e in errors:
            st.error(e)

        if results:
            st.markdown(
                '<div class="success-banner">✅ تمت معالجة ' + str(len(results)) + ' ملف بنجاح!</div>',
                unsafe_allow_html=True,
            )

            st.markdown('<div class="section-title">📦 تحميل الملفات</div>', unsafe_allow_html=True)
            preview_data = [{"اسم الملف الناتج": fname, "الحالة": "✅ جاهز"} for fname in results]
            st.dataframe(pd.DataFrame(preview_data), use_container_width=True, hide_index=True)

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for fname, fbytes in results.items():
                    zf.writestr(fname, fbytes)
            zip_buffer.seek(0)

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
            <div style="font-size:1.1rem; font-weight:600; color:#6b3fa0;">لم يتم رفع أي ملفات بعد</div>
            <div style="font-size:0.85rem; color:#888; margin-top:0.3rem;">يدعم الصيغ: xlsx, xls, csv</div>
        </div>
        """,
        unsafe_allow_html=True,
    )



# ═══════════════════════════════════════════════════════════════════════════════
# المرحلة الثانية — معالجة ملفات المعلمات المُعادة
# ═══════════════════════════════════════════════════════════════════════════════

def parse_time_to_minutes(time_str):
    """تحويل نص الوقت (8:30 أو 8.30) إلى دقائق منذ منتصف الليل"""
    s = str(time_str).strip().replace(".", ":").replace("٫", ":")
    parts = s.split(":")
    try:
        h = int(parts[0])
        m = int(parts[1]) if len(parts) > 1 else 0
        return h * 60 + m
    except Exception:
        return None


def get_period_from_time(minutes, period_schedule):
    """
    period_schedule: list of (period_name, start_min, end_min)
    يرجع اسم الفترة المناسبة للوقت، أو None
    """
    if minutes is None:
        return None
    for name, start, end in period_schedule:
        if start <= minutes <= end:
            return name
    return None


def format_time(time_str):
    """تنسيق الوقت: 8.30 / 8:30 / 830 → 8:30"""
    s = str(time_str).strip()
    if not s:
        return ""
    s = s.replace(".", ":").replace("٫", ":")
    if ":" not in s:
        if len(s) <= 2:
            return s + ":00"
        s = s[:-2] + ":" + s[-2:]
    parts = s.split(":")
    try:
        h = int(parts[0])
        m = int(parts[1]) if len(parts) > 1 else 0
        return f"{h}:{m:02d}"
    except Exception:
        return str(time_str)


# جدول الفترات الثابت (من الصورة المرفقة)
PERIOD_SCHEDULE = [
    ("فجراً",  5*60+45,  9*60+0),
    ("ضحى",    9*60+15, 12*60+30),
    ("ظهراً", 12*60+45, 16*60+15),
    ("عصراً", 16*60+30, 19*60+0),
    ("ليلاً", 19*60+15, 21*60+30),
]

KEYWORD_RED   = "شرطي"
COLOR_RED     = "#FF9999"   # أحمر فاتح — شرطي
COLOR_ORANGE  = "#FFD580"   # برتقالي — ملاحظات جوهرية
COLOR_YELLOW  = "#FFFF99"   # أصفر — تعارض فترة/وقت
COLOR_HEADER  = "#D9D9D9"   # رمادي فاتح للهيدر

NOTES_KEYWORDS = ["تغيير رقم", "تعديل مواليد", "تعديل اسم", "تغيير اسم",
                  "تصحيح رقم", "تصحيح اسم", "تصحيح مواليد"]


def analyze_day_distribution(students_df, days_list, day_col, status_col):
    """
    يحلل توزيع الطالبات على الأيام ويُنتج تقريراً يوضح:
    - عدد الطالبات الكلي اللواتي أنهين المقرر
    - الحصة المثالية لكل يوم
    - الأيام التي فيها ضغط (أكثر من الحصة) أو فراغ (أقل من الحصة)
    - الطالبات غير الموزعات (بدون يوم)
    يُعيد: dict يحتوي على ملخص التقرير
    """
    finished_mask = students_df[status_col].astype(str).str.strip() == "أنهت المقرر"
    finished_df   = students_df[finished_mask]
    total         = len(finished_df)

    if total == 0 or not days_list:
        return {"total": 0, "days": {}, "unassigned": 0, "has_issue": False, "ideal": 0}

    d           = len(days_list)
    base, extra = divmod(total, d)

    # الحصة المثالية لكل يوم
    ideal = {}
    for i, day in enumerate(days_list):
        ideal[day] = base + (1 if i < extra else 0)

    # العدد الفعلي لكل يوم
    actual = {day: 0 for day in days_list}
    unassigned = 0
    for _, row in finished_df.iterrows():
        day_val = str(row[day_col]).strip()
        if day_val in ("", "nan"):
            unassigned += 1
        else:
            matched = False
            for day in days_list:
                if day_val in day or day in day_val:
                    actual[day] = actual.get(day, 0) + 1
                    matched = True
                    break
            if not matched:
                unassigned += 1

    # بناء تقرير لكل يوم
    days_report = {}
    has_issue   = False
    for day in days_list:
        a = actual.get(day, 0)
        i = ideal.get(day, 0)
        status = "✅ مناسب"
        if a > i:
            status = "🔴 ضغط — يجب تحويل " + str(a - i) + " طالبة"
            has_issue = True
        elif a < i and unassigned > 0:
            status = "🟢 متاح — يستوعب " + str(i - a) + " طالبة إضافية"
        days_report[day] = {"actual": a, "ideal": i, "status": status}

    return {
        "total":      total,
        "days":       days_report,
        "unassigned": unassigned,
        "has_issue":  has_issue,
        "ideal_base": base,
        "ideal_extra": extra,
    }


def excel_serial_to_time_str(val):
    """تحويل Excel time serial (0.375) أو نص وقت (9:00 / 9.30) إلى HH:MM"""
    if val is None or str(val).strip() in ("", "nan"):
        return ""
    s = str(val).strip()
    # إذا كانت قيمة عشرية (Excel serial)
    try:
        f = float(s)
        if 0 < f < 1:
            total_min = round(f * 24 * 60)
            h = total_min // 60
            m = total_min % 60
            return f"{h}:{m:02d}"
    except ValueError:
        pass
    # نص وقت عادي
    return format_time(s)




def build_distribution_report(day_reports):
    """
    يبني ملف Excel واحد يحتوي:
    - ورقة "ملخص" تجمع كل المعلمات في جدول واحد
    - ورقة منفصلة لكل معلمة تفصيلية
    """
    output   = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})

    # ── صيغ مشتركة ───────────────────────────────────────────────────────────
    def fmt(bold=False, bg=None, align="center"):
        d = {"font_name": "Calibri", "font_size": 11,
             "align": align, "valign": "vcenter", "border": 1}
        if bold: d["bold"] = True
        if bg:   d["bg_color"] = bg
        return workbook.add_format(d)

    hdr_fmt  = fmt(bold=True, bg="#3d2060", align="center")
    hdr_fmt.set_font_color("white")
    ok_fmt   = fmt()
    red_fmt  = fmt(bg="#FF9999")
    grn_fmt  = fmt(bg="#C6EFCE")
    bold_fmt = fmt(bold=True, align="right")
    title_fmt = workbook.add_format({"bold": True, "font_name": "Calibri",
                                     "font_size": 13, "align": "center",
                                     "valign": "vcenter", "bg_color": "#EDE8F5"})

    # ── ورقة الملخص ──────────────────────────────────────────────────────────
    ws_sum = workbook.add_worksheet("ملخص")
    ws_sum.right_to_left()
    ws_sum.set_column(0, 0, 25)   # المعلمة
    ws_sum.set_column(1, 1, 10)   # إجمالي
    ws_sum.set_column(2, 2, 10)   # بدون يوم
    ws_sum.set_column(3, 3, 12)   # فيها ضغط
    ws_sum.set_column(4, 4, 15)   # الحالة العامة

    ws_sum.merge_range(0, 0, 0, 4, "ملخص توزيع الأيام — جميع المعلمات", title_fmt)
    ws_sum.set_row(0, 25)

    for ci, h in enumerate(["المعلمة", "أنهين المقرر", "بدون يوم", "أيام مكتظة", "الحالة"]):
        ws_sum.write(1, ci, h, hdr_fmt)

    for ri, (fname, report) in enumerate(day_reports.items()):
        r          = ri + 2
        teacher    = fname.replace(".xlsx", "")
        total      = report.get("total", 0)
        unassigned = report.get("unassigned", 0)
        over_days  = sum(1 for d in report.get("days", {}).values() if "🔴" in d["status"])
        has_issue  = report.get("has_issue", False) or unassigned > 0

        row_fmt = red_fmt if has_issue else ok_fmt
        status_txt = "⚠️ يحتاج تدخل" if has_issue else "✅ موزّع بشكل جيد"
        ws_sum.write(r, 0, teacher,    row_fmt)
        ws_sum.write(r, 1, total,      row_fmt)
        ws_sum.write(r, 2, unassigned, red_fmt if unassigned else ok_fmt)
        ws_sum.write(r, 3, over_days,  red_fmt if over_days  else ok_fmt)
        ws_sum.write(r, 4, status_txt, row_fmt)

    # ── ورقة لكل معلمة ───────────────────────────────────────────────────────
    for fname, report in day_reports.items():
        teacher  = fname.replace(".xlsx", "")
        # اسم الورقة: أول 31 حرف (حد Excel)
        sh_name  = teacher[:31]
        ws       = workbook.add_worksheet(sh_name)
        ws.right_to_left()
        ws.set_column(0, 0, 22)
        ws.set_column(1, 1, 10)
        ws.set_column(2, 2, 10)
        ws.set_column(3, 3, 35)

        # عنوان
        ws.merge_range(0, 0, 0, 3, "تقرير توزيع الأيام — " + teacher, title_fmt)
        ws.set_row(0, 25)

        # معلومات عامة
        total      = report.get("total", 0)
        unassigned = report.get("unassigned", 0)
        base       = report.get("ideal_base", 0)
        xtra       = report.get("ideal_extra", 0)
        ideal_txt  = str(base) + (" (+1 لأول " + str(xtra) + " أيام)" if xtra else " لكل يوم")

        ws.write(1, 0, "إجمالي اللواتي أنهين المقرر:", bold_fmt)
        ws.write(1, 1, total, ok_fmt)
        ws.write(2, 0, "التوزيع المثالي:", bold_fmt)
        ws.write(2, 1, ideal_txt, ok_fmt)
        if unassigned:
            ws.write(3, 0, "⚠️ بدون يوم محدد:", bold_fmt)
            ws.write(3, 1, unassigned, red_fmt)

        # هيدر جدول الأيام
        for ci, h in enumerate(["اليوم", "الفعلي", "المثالي", "الحالة"]):
            ws.write(5, ci, h, hdr_fmt)

        for ri, (day, info) in enumerate(report.get("days", {}).items()):
            r       = ri + 6
            is_over = "🔴" in info["status"]
            is_avail= "🟢" in info["status"]
            df      = red_fmt if is_over else (grn_fmt if is_avail else ok_fmt)
            ws.write(r, 0, day,            df)
            ws.write(r, 1, info["actual"], df)
            ws.write(r, 2, info["ideal"],  df)
            ws.write(r, 3, info["status"], df)

    workbook.close()
    output.seek(0)
    return output.read()

def process_stage2_file(file_bytes, days_list, statuses_list, periods_list):
    """معالجة ملف واحد مُعاد من المعلمة"""
    import xlsxwriter

    df = read_xlsx_raw(file_bytes)

    col_map = {
        "status": next((c for c in df.columns if "الحالة"         in str(c)), None),
        "day":    next((c for c in df.columns if "يوم الاختبار"   in str(c)), None),
        "time":   next((c for c in df.columns if "توقيت الاختبار" in str(c)), None),
        "period": next((c for c in df.columns if "الفترة"         in str(c)), None),
        "notes":  next((c for c in df.columns if "الملاحظات"      in str(c)), None),
    }

    columns_order = [
        "الرقم", "الاسم", "رقم الواتس اب", "المجموعة",
        "البلد", "المواليد", "الإجازة", "المعلمة",
        "الحالة", "يوم الاختبار", "توقيت الاختبار", "الفترة", "الملاحظات",
    ]
    for col in columns_order:
        if col not in df.columns:
            df[col] = ""

    # ── تحليل توزيع الأيام (تقرير فقط) ──────────────────────────────────────
    day_report = {}
    if col_map["status"] and col_map["day"] and days_list:
        day_report = analyze_day_distribution(df, days_list, col_map["day"], col_map["status"])

    # ── فحص كل صف ────────────────────────────────────────────────────────────
    red_rows          = []   # شرطي
    orange_rows       = []   # ملاحظات جوهرية
    empty_status_rows = []   # حالة فارغة
    wrong_data_rows   = []   # بيانات خاطئة (أنهت بلا يوم / لم تنه بها بيانات)

    STATUS_FINISHED = "أنهت المقرر"

    for idx, row in df.iterrows():
        status = str(row.get(col_map["status"] or "الحالة", "")).strip()
        day    = str(row.get(col_map["day"]    or "يوم الاختبار",   "")).strip()
        note   = str(row.get(col_map["notes"]  or "الملاحظات",  "")).strip()
        time_raw = row.get(col_map["time"]   or "توقيت الاختبار", "")
        period   = str(row.get(col_map["period"] or "الفترة", "")).strip()

        # تحويل الوقت من serial إلى HH:MM وتخزينه
        time_str = excel_serial_to_time_str(time_raw)
        if col_map["time"] and time_str:
            df.at[idx, col_map["time"]] = time_str

        # 1. حالة فارغة
        if not status or status == "nan":
            empty_status_rows.append(idx)
            continue

        # 2. أنهت المقرر → يجب أن يكون لها يوم
        if status == STATUS_FINISHED:
            if not day or day == "nan":
                wrong_data_rows.append(idx)



        # 3. غير أنهت → يجب أن تكون خلايا اليوم/الوقت/الفترة فارغة
        else:
            has_extra = (
                (day    and day    != "nan") or
                (time_str and time_str != "") or
                (period and period != "nan")
            )
            if has_extra:
                # ميّز الصف فقط بدون مسح
                wrong_data_rows.append(idx)

        # 4. كلمة شرطي
        if KEYWORD_RED in note:
            red_rows.append(idx)
        # 5. ملاحظات جوهرية
        elif any(kw in note for kw in NOTES_KEYWORDS):
            orange_rows.append(idx)

    # ── بناء Excel ───────────────────────────────────────────────────────────
    output   = io.BytesIO()
    df_out   = df[columns_order].copy()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    ws       = workbook.add_worksheet("الطالبات")
    ws.right_to_left()

    def fmt(extra=None):
        base = {"font_name": "Calibri", "font_size": 11,
                "align": "center", "valign": "vcenter",
                "border": 1, "locked": False}
        if extra:
            base.update(extra)
        return workbook.add_format(base)

    header_fmt     = workbook.add_format({"bold": True, "font_name": "Calibri",
                                          "font_size": 11, "align": "center",
                                          "valign": "vcenter", "border": 1,
                                          "bg_color": COLOR_HEADER, "locked": False})
    normal_fmt     = fmt()
    num_fmt        = fmt({"num_format": "0"})
    red_fmt        = fmt({"bg_color": COLOR_RED})
    red_num_fmt    = fmt({"bg_color": COLOR_RED,    "num_format": "0"})
    orange_fmt     = fmt({"bg_color": COLOR_ORANGE})
    orange_num_fmt = fmt({"bg_color": COLOR_ORANGE, "num_format": "0"})
    yellow_fmt     = fmt({"bg_color": COLOR_YELLOW})
    arial_fmt      = fmt({"font_name": "Arial"})
    red_arial      = fmt({"bg_color": COLOR_RED,    "font_name": "Arial"})
    orange_arial   = fmt({"bg_color": COLOR_ORANGE, "font_name": "Arial"})
    yellow_arial   = fmt({"bg_color": COLOR_YELLOW, "font_name": "Arial"})

    col_widths = [7, 24, 14.1, 13.3, 7, 6, 5.3, 6.9, 19.8, 11.4, 10.7, 14, 39.8]
    for i, w in enumerate(col_widths):
        ws.set_column(i, i, w)

    for ci, cn in enumerate(columns_order):
        ws.write(0, ci, cn, header_fmt)

    numeric_cols = {"الرقم", "رقم الواتس اب", "المواليد"}

    for row_idx, row in df_out.iterrows():
        er = row_idx + 1
        is_red    = row_idx in red_rows
        is_orange = row_idx in orange_rows
        is_warn   = row_idx in empty_status_rows or row_idx in wrong_data_rows

        for ci, cn in enumerate(columns_order):
            val = row[cn]
            val = "" if pd.isna(val) else val

            # اختر لون الصف (الأولوية: أحمر > برتقالي > أصفر > عادي)
            # أصفر: تعارض وقت/فترة + بيانات خاطئة (حالة فارغة أو غير أنهت مع بيانات)
            if is_red:
                row_color = "red"
            elif is_orange:
                row_color = "orange"
            elif is_warn:
                row_color = "yellow"
            else:
                row_color = "normal"

            if cn == "الفترة":
                f = {"red": red_arial, "orange": orange_arial,
                     "yellow": yellow_arial}.get(row_color, arial_fmt)
            elif cn in numeric_cols and val != "":
                try:
                    val = int(str(val).replace(".0", ""))
                except Exception:
                    pass
                f = {"red": red_num_fmt, "orange": orange_num_fmt,
                     "yellow": yellow_fmt}.get(row_color, num_fmt)
            else:
                f = {"red": red_fmt, "orange": orange_fmt,
                     "yellow": yellow_fmt}.get(row_color, normal_fmt)

            ws.write(er, ci, str(val) if isinstance(val, str) else val, f)

    # ── القوائم المنسدلة (نفس المرحلة الأولى) ───────────────────────────────
    num_rows    = len(df_out)
    last_dv_row = num_rows + 50

    if days_list:
        ws.data_validation(1, 9, last_dv_row, 9, {
            "validate": "list", "source": days_list,
            "show_input": True, "show_error": True,
        })
    if statuses_list:
        ws.data_validation(1, 8, last_dv_row, 8, {
            "validate": "list", "source": statuses_list,
            "show_input": True, "show_error": True,
        })
    if periods_list:
        ws.data_validation(1, 11, last_dv_row, 11, {
            "validate": "list", "source": periods_list,
            "show_input": True, "show_error": True,
        })

    workbook.close()
    output.seek(0)
    return (output.read(), len(red_rows), len(orange_rows),
            len(empty_status_rows) + len(wrong_data_rows), day_report)


# ── واجهة المرحلة الثانية ───────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    """
    <div style="background:linear-gradient(135deg,#1a3a5c,#2874a6);border-radius:16px;
    padding:1.8rem 2.5rem;margin-bottom:1.5rem;box-shadow:0 8px 32px rgba(26,58,92,0.3);
    text-align:center;color:white;">
        <div style="font-size:1.8rem;font-weight:900;margin:0;">🔄 المرحلة الثانية — مراجعة ملفات المعلمات</div>
        <div style="font-size:0.95rem;margin:0.4rem 0 0;opacity:0.88;">ارفعي الملفات المُعادة من المعلمات لمعالجتها وتدقيقها</div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="section-title">📂 رفع ملفات المعلمات المُعادة</div>', unsafe_allow_html=True)

uploaded_stage2 = st.file_uploader(
    "ارفعي ملفات المعلمات بعد تعبئتها",
    type=["xlsx"],
    accept_multiple_files=True,
    label_visibility="collapsed",
    key="stage2_uploader",
)

if uploaded_stage2:
    chips2 = " ".join(
        ['<span class="file-chip">📄 ' + f.name + "</span>" for f in uploaded_stage2]
    )
    st.markdown("<div style='margin:0.5rem 0 1rem'>" + chips2 + "</div>", unsafe_allow_html=True)

    # لون الأسطر
    st.markdown(
        """
        <div style="background:white;border-radius:10px;padding:0.8rem 1.2rem;
        margin-bottom:1rem;border:1px solid #e0d0f8;font-size:0.88rem;direction:rtl;">
            <b>دليل الألوان:</b> &nbsp;
            <span style="background:#FF9999;padding:2px 10px;border-radius:4px;">🔴 شرطي</span> &nbsp;
            <span style="background:#FFD580;padding:2px 10px;border-radius:4px;">🟠 ملاحظة جوهرية</span> &nbsp;
            <span style="background:#FFFF99;padding:2px 10px;border-radius:4px;">🟡 حالة فارغة · بيانات خاطئة تحتاج مراجعة</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.button("🔄 معالجة ومراجعة الملفات", type="primary", use_container_width=True, key="stage2_btn"):
        with st.spinner("جارٍ التدقيق والمعالجة..."):
            stage2_results = {}
            stage2_errors  = []
            day_reports    = {}
            total_red = total_orange = total_issues = 0

            for uf in uploaded_stage2:
                try:
                    fb = uf.read()
                    out_bytes, n_red, n_orange, n_issues, d_report = process_stage2_file(fb, days_list, statuses_list, periods_list)
                    out_name = uf.name
                    stage2_results[out_name] = out_bytes
                    total_red    += n_red
                    total_orange += n_orange
                    total_issues += n_issues
                    if d_report.get("has_issue") or d_report.get("unassigned", 0) > 0:
                        day_reports[uf.name] = d_report
                except Exception as e:
                    stage2_errors.append("❌ " + uf.name + ": " + str(e))

        for e in stage2_errors:
            st.error(e)

        if stage2_results:
            cols2 = st.columns(5)
            with cols2[0]:
                st.markdown('<div class="stat-card"><div class="number">' + str(len(stage2_results)) + '</div><div class="label">ملف معالج</div></div>', unsafe_allow_html=True)
            with cols2[1]:
                st.markdown('<div class="stat-card"><div class="number" style="color:#c0392b">' + str(total_red) + '</div><div class="label">شرطي 🔴</div></div>', unsafe_allow_html=True)
            with cols2[2]:
                st.markdown('<div class="stat-card"><div class="number" style="color:#e67e22">' + str(total_orange) + '</div><div class="label">ملاحظة جوهرية 🟠</div></div>', unsafe_allow_html=True)
            with cols2[3]:
                st.markdown('<div class="stat-card"><div class="number" style="color:#b7950b">' + str(total_issues) + '</div><div class="label">يحتاج مراجعة 🟡</div></div>', unsafe_allow_html=True)
            with cols2[4]:
                st.markdown('<div class="stat-card"><div class="number" style="color:#555">' + str(sum(len(r["days"]) for r in day_reports.values())) + '</div><div class="label">أيام مكتظة 📊</div></div>', unsafe_allow_html=True)

            # ── تقرير توزيع الأيام — ملف Excel واحد ──────────────────────────
            if day_reports:
                report_bytes = build_distribution_report(day_reports)
                has_any_issue = any(
                    r.get("has_issue") or r.get("unassigned", 0) > 0
                    for r in day_reports.values()
                )
                st.markdown('<div class="section-title">📊 تقرير توزيع الأيام</div>', unsafe_allow_html=True)
                if has_any_issue:
                    st.warning("⚠️ يوجد أيام مكتظة أو طالبات بدون يوم — راجعي التقرير")
                else:
                    st.success("✅ التوزيع متوازن لدى جميع المعلمات")
                st.download_button(
                    label="⬇️ تحميل تقرير التوزيع — Excel",
                    data=report_bytes,
                    file_name="تقرير_توزيع_الأيام.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="report_download",
                )

            zip2 = io.BytesIO()
            with zipfile.ZipFile(zip2, "w", zipfile.ZIP_DEFLATED) as zf:
                for fname, fbytes in stage2_results.items():
                    zf.writestr(fname, fbytes)
            zip2.seek(0)

            st.markdown('<div class="section-title">📦 تحميل الملفات المُدققة</div>', unsafe_allow_html=True)
            st.download_button(
                label="⬇️ تحميل جميع الملفات المُعالجة — ZIP",
                data=zip2,
                file_name="مراجعة_المعلمات.zip",
                mime="application/zip",
                use_container_width=True,
                key="stage2_download",
            )

st.markdown(
    """
    <hr style="margin:2rem 0 1rem; border-color:#d8c8f0;">
    <div style="text-align:center; color:#999; font-size:0.8rem; font-family:'Tajawal',sans-serif;">
        أداة مقرأة — مبنية بـ Python & Streamlit &nbsp;|&nbsp; 📖
    </div>
    """,
    unsafe_allow_html=True,
)
