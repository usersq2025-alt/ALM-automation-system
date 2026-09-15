import streamlit as st
import pandas as pd
import io
import zipfile
import xlsxwriter
import xml.etree.ElementTree as ET
import json
import base64
import datetime
import streamlit.components.v1 as components

try:
    import extra_streamlit_components as stx
    HAS_COOKIE_MGR = True
except ImportError:
    stx = None
    HAS_COOKIE_MGR = False

# ── اللوغو — ضعي ملف logo.png في نفس مجلد app.py ─────────────────────────
import os, base64

def load_logo_b64(path="logo.png"):
    """يحوّل الصورة لـ base64 لاستخدامها في HTML"""
    if os.path.exists(path):
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None

LOGO_B64 = load_logo_b64()
LOGO_SRC  = f"data:image/png;base64,{LOGO_B64}" if LOGO_B64 else None

st.set_page_config(
    page_title="أداة مقرأة",
    page_icon="logo.png" if os.path.exists("logo.png") else "📖",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700;900&display=swap');

    html, body, [class*="css"] { font-family: 'Tajawal', sans-serif; direction: rtl; }

    /* ── خلفية بيضاء ناعمة ── */
    .stApp { background: #F8F7FC !important; }

    /* ── Hero — بنفسجي داكن فاخر ── */
    .hero-header {
        background: linear-gradient(135deg, #1a0533 0%, #3b0f72 60%, #5a1fa8 100%);
        border-radius: 20px; padding: 2rem 2.5rem; margin-bottom: 1.5rem;
        box-shadow: 0 12px 40px rgba(26,5,51,0.35); text-align: center; color: white;
        border: 1px solid rgba(255,255,255,0.08);
    }
    .hero-header h1 { font-size: 2.2rem; font-weight: 900; margin: 0; letter-spacing: -0.5px; }
    .hero-header p  { font-size: 0.95rem; margin: 0.4rem 0 0; opacity: 0.75; font-weight: 300; }

    /* ── بطاقات إحصاء ── */
    .stat-card {
        background: white; border-radius: 14px; padding: 1.1rem 1.3rem;
        box-shadow: 0 2px 14px rgba(0,0,0,0.07); margin-bottom: 1rem;
        border-top: 3px solid #7c3aed; text-align: center;
    }
    .stat-card .number { font-size: 1.9rem; font-weight: 900; color: #3b0f72; line-height: 1; }
    .stat-card .label  { font-size: 0.8rem; color: #888; margin-top: 5px; font-weight: 500; }

    /* ── File chips ── */
    .file-chip {
        display: inline-block; background: #f3f0fc; border: 1px solid #c4b5fd;
        color: #4c1d95; border-radius: 20px; padding: 3px 13px;
        font-size: 0.8rem; margin: 3px; font-weight: 600;
    }

    /* ── بانر نجاح ── */
    .success-banner {
        background: linear-gradient(90deg, #f0fdf4, #dcfce7);
        border: 1.5px solid #86efac; border-radius: 12px;
        padding: 0.9rem 1.5rem; color: #166534;
        font-weight: 700; font-size: 1rem; margin: 1rem 0;
    }

    /* ── عناوين الأقسام ── */
    .section-title {
        font-size: 0.95rem; font-weight: 800; color: #3b0f72;
        padding: 0.45rem 1rem; margin: 1.2rem 0 0.8rem;
        border-right: 4px solid #7c3aed;
        background: linear-gradient(90deg, rgba(124,58,237,0.08), transparent);
        border-radius: 0 10px 10px 0; display: inline-block; width: 100%;
    }

    /* ── منطقة الرفع ── */
    .upload-zone {
        background: white; border: 2px dashed #c4b5fd;
        border-radius: 16px; padding: 2rem; text-align: center; margin: 1rem 0;
    }

    /* ── Sidebar ── */
    [data-testid="stSidebar"] > div:first-child {
        background: linear-gradient(180deg, #0f0520 0%, #1e0a40 50%, #2d1060 100%) !important;
    }
    [data-testid="stSidebar"],
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] div,
    [data-testid="stSidebar"] label {
        color: #e2d9f3 !important; font-family: 'Tajawal', sans-serif !important;
    }
    [data-testid="stSidebar"] label { font-weight: 700 !important; font-size: 0.88rem !important; }
    [data-testid="stSidebar"] textarea {
        background-color: #0a0318 !important; border: 1.5px solid #6d28d9 !important;
        border-radius: 8px !important; color: #ede9fe !important;
        font-family: 'Tajawal', sans-serif !important; font-size: 0.88rem !important;
        direction: rtl !important;
    }
    [data-testid="stSidebar"] textarea:focus {
        border-color: #a78bfa !important; box-shadow: 0 0 0 2px rgba(167,139,250,0.2) !important;
    }
    [data-testid="stSidebar"] textarea::placeholder { color: #6d28d9 !important; }
    [data-testid="stSidebar"] small,
    [data-testid="stSidebar"] .stMarkdown { color: #a78bfa !important; }

    /* ── Sidebar مطوي ── */
    [data-testid="collapsedControl"] { display: none !important; }
    section[data-testid="stSidebar"][aria-expanded="false"] { width: 0 !important; min-width: 0 !important; }
    section[data-testid="stSidebar"][aria-expanded="false"] * { visibility: hidden !important; opacity: 0 !important; }

    /* ── أزرار ── */
    .stButton > button {
        font-family: 'Tajawal', sans-serif !important;
        font-weight: 700 !important; border-radius: 12px !important;
        transition: all 0.15s ease !important;
    }
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #3b0f72, #7c3aed) !important;
        border: none !important; color: white !important;
        box-shadow: 0 4px 18px rgba(59,15,114,0.4) !important;
    }
    .stButton > button[kind="primary"]:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 22px rgba(59,15,114,0.5) !important;
    }
    .stButton > button[kind="secondary"] {
        background: white !important; border: 2px solid #c4b5fd !important;
        color: #4c1d95 !important;
    }
    .stDownloadButton > button {
        background: linear-gradient(135deg, #1e3a5f, #1d4ed8) !important;
        color: white !important; font-family: 'Tajawal', sans-serif !important;
        font-weight: 700 !important; border: none !important;
        border-radius: 12px !important; font-size: 0.95rem !important;
        box-shadow: 0 4px 16px rgba(29,78,216,0.35) !important;
    }

    /* ── Tabs ── */
    [data-testid="stTabs"] [data-baseweb="tab-list"] {
        background: white; border-radius: 16px; padding: 6px;
        box-shadow: 0 2px 12px rgba(0,0,0,0.08); gap: 4px; margin-bottom: 1.5rem;
    }
    [data-testid="stTabs"] [data-baseweb="tab"] {
        border-radius: 11px !important; font-family: 'Tajawal', sans-serif !important;
        font-weight: 700 !important; font-size: 0.87rem !important;
        color: #6d28d9 !important; padding: 0.5rem 1rem !important;
    }
    [data-testid="stTabs"] [aria-selected="true"] {
        background: linear-gradient(135deg, #3b0f72, #7c3aed) !important;
        color: white !important; box-shadow: 0 3px 10px rgba(59,15,114,0.35) !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ── اللوغو
import zipfile
import xlsxwriter
import xml.etree.ElementTree as ET

# ── اللوغو — ضعي ملف logo.png في نفس مجلد app.py ─────────────────────────
import os, base64

def load_logo_b64(path="logo.png"):
    """يحوّل الصورة لـ base64 لاستخدامها في HTML"""
    if os.path.exists(path):
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None

LOGO_B64 = load_logo_b64()
LOGO_SRC  = f"data:image/png;base64,{LOGO_B64}" if LOGO_B64 else None

st.set_page_config(
    page_title="أداة مقرأة",
    page_icon="logo.png" if os.path.exists("logo.png") else "📖",
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
        background: linear-gradient(180deg, #16082b 0%, #2a1248 45%, #3b1a66 100%) !important;
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
        font-size: 0.9rem !important;
    }

    /* بطاقات الإعدادات داخل الشريط الجانبي */
    [data-testid="stSidebar"] [data-testid="stTextArea"] {
        background: rgba(255,255,255,0.07) !important;
        border: 1px solid rgba(167,139,250,0.4) !important;
        border-radius: 14px !important;
        padding: 10px 12px 6px !important;
        margin-bottom: 0.75rem !important;
        box-shadow: 0 6px 18px rgba(0,0,0,0.18) !important;
    }
    [data-testid="stSidebar"] .sidebar-card-title {
        display: block;
        font-size: 0.78rem;
        font-weight: 800;
        color: #c4b5fd !important;
        margin: 0.15rem 0 0.35rem;
        letter-spacing: 0.2px;
    }
    [data-testid="stSidebar"] .sidebar-summary {
        background: rgba(124,58,237,0.22);
        border: 1px solid rgba(167,139,250,0.45);
        border-radius: 12px;
        padding: 0.75rem 0.9rem;
        margin-top: 0.4rem;
        font-size: 0.8rem;
        color: #ddd6fe !important;
        line-height: 1.7;
    }
    [data-testid="stSidebar"] .sidebar-summary b {
        color: #f5f3ff !important;
    }

    /* ── Sidebar textarea — dark bg + white text ── */
    [data-testid="stSidebar"] textarea,
    [data-testid="stSidebar"] .stTextArea textarea {
        background-color: #12071f !important;
        border: 1.5px solid #7c3aed !important;
        border-radius: 10px !important;
        color: #f5f3ff !important;
        font-family: 'Tajawal', sans-serif !important;
        font-size: 0.9rem !important;
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

    /* ── إخفاء محتوى الـ Sidebar عند طيّه ── */
    [data-testid="collapsedControl"] {
        display: none !important;
    }
    section[data-testid="stSidebar"][aria-expanded="false"] {
        width: 0 !important;
        min-width: 0 !important;
    }
    section[data-testid="stSidebar"][aria-expanded="false"] > div {
        overflow: hidden !important;
        width: 0 !important;
    }
    /* إخفاء النصوص في الـ collapsed sidebar */
    section[data-testid="stSidebar"][aria-expanded="false"] * {
        visibility: hidden !important;
        opacity: 0 !important;
    }
    /* زر الفتح/الإغلاق */
    [data-testid="collapsedControl"] svg {
        fill: #6b3fa0 !important;
    }
    button[kind="header"] {
        background: white !important;
        border-radius: 0 8px 8px 0 !important;
        box-shadow: 2px 2px 8px rgba(0,0,0,0.1) !important;
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
                col_h_name = teacher_map.get(raw_name, get_first_name(raw_name))
                if teacher_col:
                    df[teacher_col] = col_h_name
                short = col_h_name  # اسم الملف = نفس اسم العمود H
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
                col_h_name = (teacher_map or {}).get(raw_name, get_first_name(raw_name))
                if teacher_col:
                    df[teacher_col] = col_h_name
                short = col_h_name  # اسم الملف = نفس اسم العمود H
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


# ── إعدادات الدورة الافتراضية + حفظ محلي ─────────────────────────────────────
DEFAULT_DAYS = "الإثنين\nالثلاثاء\nالأربعاء\nالخميس\nالجمعة\nالسبت\nالأحد"
DEFAULT_PERIODS = "فجراً\nضحى\nظهراً\nعصراً\nليلاً"
DEFAULT_STATUSES = (
    "أنهت المقرر\nلم تنه المقرر\nساكنة\nمنسحبة\nأخرجتها الإدارة لأنها مخالفة\n"
    "لا يوجد واتس\nتم نقلها لغير مجموعة"
)
DEFAULT_PERIOD_SCHEDULE = (
    "فجراً: 4:00-8:45\nضحى: 9:00-11:45\nظهراً: 12:00-15:45\n"
    "عصراً: 16:00-18:45\nليلاً: 19:00-21:30"
)
SETTINGS_COOKIE_KEY = "alm_course_settings_v1"


def _get_cookie_manager():
    if not HAS_COOKIE_MGR:
        return None
    return stx.CookieManager(key="alm_cookie_mgr")


def load_persisted_settings(cookie_manager):
    """تحميل الإعدادات المحفوظة محلياً (كوكي المتصفح) مرة واحدة."""
    # ضمان وجود قيم للودجات حتى قبل وصول الكوكي
    st.session_state.setdefault("cfg_days", DEFAULT_DAYS)
    st.session_state.setdefault("cfg_periods", DEFAULT_PERIODS)
    st.session_state.setdefault("cfg_statuses", DEFAULT_STATUSES)
    st.session_state.setdefault("cfg_period_schedule", DEFAULT_PERIOD_SCHEDULE)

    if st.session_state.get("_settings_from_cookie"):
        return

    if cookie_manager is None:
        st.session_state["_settings_from_cookie"] = True
        return

    try:
        raw = cookie_manager.get(SETTINGS_COOKIE_KEY)
    except Exception:
        raw = None

    if raw:
        try:
            saved = json.loads(raw) if isinstance(raw, str) else raw
            if isinstance(saved, dict):
                st.session_state["cfg_days"] = saved.get("days", DEFAULT_DAYS)
                st.session_state["cfg_periods"] = saved.get("periods", DEFAULT_PERIODS)
                st.session_state["cfg_statuses"] = saved.get("statuses", DEFAULT_STATUSES)
                st.session_state["cfg_period_schedule"] = saved.get(
                    "period_schedule", DEFAULT_PERIOD_SCHEDULE
                )
        except Exception:
            pass
        st.session_state["_settings_from_cookie"] = True
        return

    # أول تحميل قد يعود None قبل جاهزية مكوّن الكوكي → إعادة محاولة لمرة واحدة
    if not st.session_state.get("_cookie_retry_done"):
        st.session_state["_cookie_retry_done"] = True
        return

    st.session_state["_settings_from_cookie"] = True


def persist_settings_to_browser(cookie_manager):
    """حفظ إعدادات الشريط الجانبي في كوكي المتصفح."""
    payload = {
        "days": st.session_state.get("cfg_days", DEFAULT_DAYS),
        "periods": st.session_state.get("cfg_periods", DEFAULT_PERIODS),
        "statuses": st.session_state.get("cfg_statuses", DEFAULT_STATUSES),
        "period_schedule": st.session_state.get("cfg_period_schedule", DEFAULT_PERIOD_SCHEDULE),
    }
    # نسخة احتياطية داخل الجلسة
    st.session_state["_last_saved_settings"] = payload
    if cookie_manager is None:
        return
    try:
        cookie_manager.set(
            SETTINGS_COOKIE_KEY,
            json.dumps(payload, ensure_ascii=False),
            expires_at=datetime.datetime.now() + datetime.timedelta(days=400),
            key="alm_set_settings_cookie",
        )
    except Exception:
        pass


def render_batch_separate_download(files_dict, widget_key="batch_sep"):
    """تنزيل كل الملفات في مجلد واحد: اختيار المجلد مرة واحدة ثم الحفظ تلقائياً."""
    if not files_dict:
        return
    items = []
    total_b64 = 0
    for name, data in files_dict.items():
        b64 = base64.b64encode(data).decode("ascii")
        total_b64 += len(b64)
        items.append({"name": name, "b64": b64})
    if total_b64 >= 12_000_000:
        st.caption("الملفات كبيرة للتنزيل الجماعي المنفصل — استخدمي ZIP أو التنزيل الفردي.")
        return
    files_json = json.dumps(items, ensure_ascii=False)
    # نفتح نافذة مستقلة (وليس داخل iframe ستريملت) لأن showDirectoryPicker
    # غالباً يُحجب داخل الإطار، بينما يعمل في نافذة عادية: اختيار مجلد مرة واحدة ثم حفظ الكل.
    components.html(
        f"""
        <div style="font-family:Tajawal,Tahoma,sans-serif;direction:rtl;">
          <button id="dl_all_{widget_key}" style="width:100%;padding:12px 14px;border:none;border-radius:12px;
            background:linear-gradient(135deg,#0f766e,#0d9488);color:#fff;font-weight:800;
            cursor:pointer;font-size:0.95rem;box-shadow:0 4px 14px rgba(13,148,136,0.3);">
            ⬇️ تنزيل الجميع منفصلاً — مجلد واحد ({len(items)})
          </button>
          <div id="dl_msg_{widget_key}" style="margin-top:6px;font-size:0.75rem;color:#64748b;text-align:center;min-height:1.1em;">
            ستظهر نافذة: اختاري المجلد مرة واحدة ثم تُحفظ كل الملفات فيه تلقائياً
          </div>
        </div>
        <script>
          const files_{widget_key} = {files_json};
          const btn_{widget_key} = document.getElementById('dl_all_{widget_key}');
          const msg_{widget_key} = document.getElementById('dl_msg_{widget_key}');

          btn_{widget_key}.addEventListener('click', () => {{
            const payload = JSON.stringify(files_{widget_key});
            const html = `<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8" />
  <title>حفظ ملفات المعلمات</title>
  <style>
    body {{ font-family: Tajawal, Tahoma, sans-serif; background:#F8F7FC; color:#3b0f72;
      display:flex; align-items:center; justify-content:center; min-height:100vh; margin:0; }}
    .card {{ background:#fff; border:1px solid #e0d0f8; border-radius:18px; padding:2rem;
      max-width:420px; width:90%; text-align:center; box-shadow:0 8px 28px rgba(59,15,114,0.08); }}
    h1 {{ font-size:1.25rem; margin:0 0 0.5rem; }}
    p {{ color:#6b7280; font-size:0.92rem; line-height:1.7; }}
    button {{ margin-top:1rem; width:100%; padding:14px; border:none; border-radius:12px;
      background:linear-gradient(135deg,#3b0f72,#7c3aed); color:#fff; font-weight:800;
      font-size:1rem; cursor:pointer; font-family:inherit; }}
    button:disabled {{ opacity:0.7; cursor:wait; }}
    #msg {{ margin-top:0.9rem; font-size:0.85rem; color:#0f766e; min-height:1.3em; }}
  </style>
</head>
<body>
  <div class="card">
    <h1>حفظ الملفات في مجلد واحد</h1>
    <p>اضغطي الزر، اختاري المجلد مرة واحدة فقط، ثم تُحفظ كل الملفات تلقائياً في نفس المكان دون سؤال متكرر.</p>
    <button id="go">اختيار المجلد وحفظ الكل (${{files_{widget_key}.length}} ملف)</button>
    <div id="msg"></div>
  </div>
  <script>
    const files = ${{payload}};
    function b64ToBytes(b64) {{
      const bin = atob(b64);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      return bytes;
    }}
    const go = document.getElementById('go');
    const msg = document.getElementById('msg');
    go.addEventListener('click', async () => {{
      if (!window.showDirectoryPicker) {{
        msg.textContent = 'هذا المتصفح لا يدعم اختيار مجلد. استخدمي Chrome أو Edge.';
        return;
      }}
      go.disabled = true;
      try {{
        const dirHandle = await window.showDirectoryPicker({{ mode: 'readwrite' }});
        let done = 0;
        for (const f of files) {{
          msg.textContent = 'جارٍ الحفظ: ' + f.name + ' (' + (done + 1) + '/' + files.length + ')';
          const fileHandle = await dirHandle.getFileHandle(f.name, {{ create: true }});
          const writable = await fileHandle.createWritable();
          await writable.write(b64ToBytes(f.b64));
          await writable.close();
          done += 1;
        }}
        msg.textContent = 'تم حفظ ' + done + ' ملف بنجاح. يمكن إغلاق هذه النافذة.';
        go.textContent = 'تم الحفظ';
      }} catch (err) {{
        go.disabled = false;
        if (err && err.name === 'AbortError') msg.textContent = 'تم إلغاء اختيار المجلد';
        else msg.textContent = 'تعذّر الحفظ. تأكدي من منح صلاحية الكتابة للمجلد.';
      }}
    }});
  <\\/script>
</body>
</html>`;
            const blob = new Blob([html], {{ type: 'text/html;charset=utf-8' }});
            const url = URL.createObjectURL(blob);
            const w = window.open(url, '_blank');
            if (!w) {{
              msg_{widget_key}.textContent = 'اسمحي بالنوافذ المنبثقة ثم أعيدي المحاولة';
            }} else {{
              msg_{widget_key}.textContent = 'فُتحت نافذة الحفظ — اختاري المجلد مرة واحدة فقط';
            }}
            setTimeout(() => URL.revokeObjectURL(url), 60000);
          }});
        </script>
        """,
        height=78,
    )


def render_stage2_results_ui(stage2_results, day_reports, stage2_errors,
                             time_fmt_warnings, period_warnings,
                             total_red, total_issues):
    """واجهة نتائج المرحلة 2 — مختصرة ومركّزة على التحميل."""
    for e in stage2_errors:
        st.error(e)

    if not stage2_results:
        return

    n_files = len(stage2_results)
    has_balance_issue = any(
        r.get("has_issue") or r.get("unassigned", 0) > 0
        for r in (day_reports or {}).values()
    )
    n_alerts = len(time_fmt_warnings or {}) + len(period_warnings or {})

    # بطاقة الملخص + أسماء المعلمات (كل اسم = تنزيل)
    st.markdown(
        f"""
        <div style="background:white;border:1px solid #e0d0f8;border-radius:16px;
        padding:1.1rem 1.3rem;margin:0.6rem 0 1rem;box-shadow:0 2px 12px rgba(59,15,114,0.06);">
            <div style="display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;">
                <div style="font-size:1.05rem;font-weight:900;color:#3b0f72;">
                    ✅ تمت معالجة {n_files} ملف
                </div>
                <div style="font-size:0.82rem;color:#6b7280;">
                    اضغطي اسم المعلمة لتنزيل ملفها
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    names = list(stage2_results.keys())
    # صفوف أزرار بأسماء المعلمات
    per_row = 3 if n_files > 3 else max(n_files, 1)
    for start in range(0, n_files, per_row):
        chunk = names[start:start + per_row]
        cols = st.columns(len(chunk))
        for i, fname in enumerate(chunk):
            teacher_label = fname[:-5] if fname.lower().endswith(".xlsx") else fname
            with cols[i]:
                st.download_button(
                    label="📄 " + teacher_label,
                    data=stage2_results[fname],
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"s2_teacher_{start + i}",
                )

    st.markdown("<div style='height:0.4rem'></div>", unsafe_allow_html=True)

    # تنزيل الجميع
    c_all1, c_all2 = st.columns(2)
    with c_all1:
        render_batch_separate_download(stage2_results, widget_key="stage2_batch")
    with c_all2:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for fname, fbytes in stage2_results.items():
                zf.writestr(fname, fbytes)
        zip_buffer.seek(0)
        st.download_button(
            label=f"📦 تنزيل الجميع — ZIP ({n_files})",
            data=zip_buffer,
            file_name="مراجعة_المعلمات.zip",
            mime="application/zip",
            use_container_width=True,
            key="stage2_all_zip",
        )

    # تقرير التوزيع
    if day_reports:
        report_bytes = build_distribution_report(day_reports)
        st.download_button(
            label="📊 تنزيل تقرير توزيع الأيام",
            data=report_bytes,
            file_name="تقرير_توزيع_الأيام.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="report_download",
        )
        if has_balance_issue:
            st.caption("⚠️ التقرير يشير إلى أيام مكتظة أو ناقصة — راجعيه بعد التنزيل.")
        else:
            st.caption("✅ التوزيع متوازن حسب التقرير.")

    # تنبيهات مطوية — لا تزاحم أزرار التحميل
    if n_alerts or total_issues or total_red:
        with st.expander("🔎 تفاصيل التدقيق والتنبيهات", expanded=False):
            m1, m2 = st.columns(2)
            with m1:
                st.metric("صفوف ملوّنة بملاحظات", total_red)
            with m2:
                st.metric("صفوف تحتاج مراجعة بيانات", total_issues)

            if time_fmt_warnings:
                st.markdown("**تنسيق التوقيت**")
                for fname, errors in time_fmt_warnings.items():
                    rows_str = "، ".join(f"صف {idx+2}" for idx, raw in errors)
                    st.warning(f"{fname}: {rows_str}")

            if period_warnings:
                st.markdown("**تعارض الفترة مع الوقت**")
                for fname, mismatches in period_warnings.items():
                    lines = [
                        f"صف {idx+2}: {actual} ← {correct}"
                        for idx, actual, correct in mismatches
                    ]
                    st.info(f"**{fname}** — " + " | ".join(lines))


def render_separate_downloads(files_dict, key_prefix, zip_name=None):
    """تنزيل ZIP + تنزيل منفصل لكل ملف + زر لتنزيل الكل كملفات منفصلة."""
    if not files_dict:
        return

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, fbytes in files_dict.items():
            zf.writestr(fname, fbytes)
    zip_buffer.seek(0)

    st.download_button(
        label="⬇️ تحميل الكل — ZIP (" + str(len(files_dict)) + " ملف)",
        data=zip_buffer,
        file_name=zip_name or (key_prefix + ".zip"),
        mime="application/zip",
        use_container_width=True,
        key=key_prefix + "_zip",
    )

    st.markdown("##### 📂 تنزيل منفصل")
    render_batch_separate_download(files_dict, widget_key=key_prefix + "_batch")

    cols = st.columns(2)
    for i, (fname, fbytes) in enumerate(files_dict.items()):
        with cols[i % 2]:
            st.download_button(
                label="📄 " + fname,
                data=fbytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"{key_prefix}_file_{i}",
            )


# ── Sidebar ───────────────────────────────────────────────────────────────────
cookie_manager = _get_cookie_manager()
load_persisted_settings(cookie_manager)

with st.sidebar:
    # ── لوغو السايدبار ────────────────────────────────────────────────────────
    if LOGO_SRC:
        st.markdown(
            f"""
            <div style='text-align:center; padding:0.6rem 0 0.2rem;'>
                <img src="{LOGO_SRC}" style='width:82px; border-radius:14px;
                box-shadow:0 6px 20px rgba(0,0,0,0.35); border:2px solid rgba(167,139,250,0.35);'>
                <div style='font-size:1.05rem; font-weight:900; color:#f5f3ff; margin-top:0.7rem;'>إعدادات الدورة</div>
                <div style='font-size:0.75rem; color:#c4b5fd; margin-top:4px;'>تُحفظ تلقائياً على هذا الجهاز</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            """
            <div style='text-align:center; padding:0.8rem 0 0.3rem;'>
                <div style='font-size:2.2rem'>📖</div>
                <div style='font-size:1.1rem; font-weight:900; color:#f5f3ff;'>إعدادات الدورة</div>
                <div style='font-size:0.75rem; color:#c4b5fd; margin-top:4px;'>تُحفظ تلقائياً على هذا الجهاز</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<div class='sidebar-card-title'>📅 أيام الأسبوع</div>", unsafe_allow_html=True)
    days_text = st.text_area(
        "أيام الأسبوع",
        height=150,
        help="كل يوم في سطر منفصل",
        key="cfg_days",
        label_visibility="collapsed",
    )

    st.markdown("<div class='sidebar-card-title'>⏰ الفترات</div>", unsafe_allow_html=True)
    periods_text = st.text_area(
        "الفترات",
        height=130,
        help="كل فترة في سطر منفصل",
        key="cfg_periods",
        label_visibility="collapsed",
    )

    st.markdown("<div class='sidebar-card-title'>📋 قائمة الحالات</div>", unsafe_allow_html=True)
    statuses_text = st.text_area(
        "قائمة الحالات",
        height=160,
        help="كل حالة في سطر منفصل",
        key="cfg_statuses",
        label_visibility="collapsed",
    )

    st.markdown("<div class='sidebar-card-title'>🕐 أوقات الفترات (للمطابقة)</div>", unsafe_allow_html=True)
    periods_schedule_text = st.text_area(
        "أوقات الفترات",
        height=140,
        help="النسق: اسم الفترة: HH:MM-HH:MM",
        key="cfg_period_schedule",
        label_visibility="collapsed",
    )

    days_list = parse_list(days_text)
    periods_list = parse_list(periods_text)
    statuses_list = parse_list(statuses_text)

    def parse_period_schedule(text):
        """فجراً: 4:00-8:45 → [("فجراً", 240, 525)]"""
        schedule = []
        for line in text.strip().splitlines():
            line = line.strip()
            if ":" not in line:
                continue
            parts = line.split(":", 1)
            if len(parts) < 2:
                continue
            name = parts[0].strip()
            times = parts[1].strip()
            if "-" not in times:
                continue
            t_parts = times.split("-")
            try:
                def to_min(t):
                    t = t.strip()
                    h, m = (t.split(":") + ["0"])[:2]
                    return int(h) * 60 + int(m)
                schedule.append((name, to_min(t_parts[0]), to_min(t_parts[1])))
            except Exception:
                continue
        return schedule

    period_schedule = parse_period_schedule(periods_schedule_text)

    if st.button("💾 حفظ الإعدادات", use_container_width=True, key="btn_save_settings"):
        persist_settings_to_browser(cookie_manager)
        st.toast("تم حفظ الإعدادات على هذا الجهاز")

    # الحفظ التلقائي بعد الزر لتفادي تعارض مفاتيح الكوكي في نفس الدورة
    _current_payload = {
        "days": st.session_state.get("cfg_days", days_text),
        "periods": st.session_state.get("cfg_periods", periods_text),
        "statuses": st.session_state.get("cfg_statuses", statuses_text),
        "period_schedule": st.session_state.get("cfg_period_schedule", periods_schedule_text),
    }
    if st.session_state.get("_last_saved_settings") != _current_payload:
        # تجنّب set مكرر إذا ضغطت «حفظ» للتو في نفس الدورة
        if not st.session_state.get("btn_save_settings"):
            persist_settings_to_browser(cookie_manager)

    st.markdown(
        "<div class='sidebar-summary'>"
        "<b>الملخص:</b><br>"
        "✅ " + str(len(days_list)) + " أيام &nbsp;·&nbsp; ✅ "
        + str(len(periods_list)) + " فترات<br>✅ "
        + str(len(statuses_list)) + " حالة &nbsp;·&nbsp; ✅ "
        + str(len(period_schedule)) + " أوقات مطابقة"
        + ("<br><span style='opacity:.85'>💾 الحفظ المحلي مفعّل</span>" if HAS_COOKIE_MGR else
           "<br><span style='opacity:.85'>⚠️ ثبّتي extra-streamlit-components للحفظ المحلي</span>")
        + "</div>",
        unsafe_allow_html=True,
    )


# ── Main ──────────────────────────────────────────────────────────────────────
# ── Hero مع أو بدون لوغو ──────────────────────────────────────────────────
if LOGO_SRC:
    st.markdown(
        f"""
        <div style="margin-bottom:1.5rem;">
            <!-- اللوغو فوق الهيرو -->
            <div style="display:flex; justify-content:center; margin-bottom:-55px; position:relative; z-index:10;">
                <div style="
                    background:white;
                    border-radius:50%;
                    padding:12px;
                    box-shadow:0 8px 32px rgba(0,0,0,0.25), 0 0 0 5px rgba(124,58,237,0.2);
                    width:120px; height:120px;
                    display:flex; align-items:center; justify-content:center;
                ">
                    <img src="{LOGO_SRC}" style="width:96px;height:96px;border-radius:50%;object-fit:contain;">
                </div>
            </div>
            <!-- الهيرو -->
            <div class="hero-header" style="padding-top:4rem;">
                <h1 style="margin:0;font-size:2rem;">أداة أتمتة جداول مقرأة</h1>
                <p style="margin:0.4rem 0 0;opacity:0.8;font-size:0.95rem;">
                    ارفعي ملفات Excel أو CSV الخام وستحصلين على جداول منسقة، محمية، وجاهزة للمعلمات
                </p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        """
        <div class="hero-header">
            <h1>📖 أداة أتمتة جداول مقرأة</h1>
            <p>ارفعي ملفات Excel أو CSV الخام وستحصلين على جداول منسقة، محمية، وجاهزة للمعلمات</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    """
    <div style="background:linear-gradient(135deg,#064e3b,#065f46,#047857);border-radius:16px;
    padding:1.8rem 2.5rem;margin-bottom:1.5rem;box-shadow:0 8px 32px rgba(6,78,59,0.35);
    text-align:center;color:white;">
        <div style="font-size:1.8rem;font-weight:900;margin:0;">📋 المرحلة الأولى — توليد جداول المعلمات</div>
        <div style="font-size:0.95rem;margin:0.4rem 0 0;opacity:0.88;">ارفعي ملفات Excel أو CSV الخام لتوليد جداول منسقة وجاهزة للمعلمات</div>
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
            render_separate_downloads(results, key_prefix="stage1_dl", zip_name="جداول_المعلمات.zip")





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

KEYWORD_RED    = "شرطي"
KEYWORD_CAMERA = "كاميرا"
COLOR_RED      = "#FF9999"   # كاميرا — صف كامل
COLOR_ORANGE   = "#FFB347"   # ملاحظة جوهرية — خلية الاسم
COLOR_YELLOW   = "#FFFF99"   # شرطي فقط — خلية الحالة
COLOR_PURPLE   = "#CE93D8"   # حالة فارغة — خلية الحالة (بنفسجي واضح)
COLOR_BLUE     = "#AED6F1"   # نقص يوم/فترة إلزامي — الخلية الناقصة
COLOR_TEAL     = "#76D7C4"   # عدم تطابق الفترة مع الوقت — خلية الفترة
COLOR_PINK     = "#F5B7B1"   # يوم/وقت/فترة معبأة رغم أن الحالة ليست «أنهت المقرر»
COLOR_HEADER   = "#D9D9D9"

ISSUE_EMPTY_STATUS    = "empty_status"
ISSUE_MISSING         = "missing"
ISSUE_PERIOD_MISMATCH = "period_mismatch"
ISSUE_UNEXPECTED      = "unexpected"

ISSUE_COLORS = {
    ISSUE_EMPTY_STATUS:    COLOR_PURPLE,
    ISSUE_MISSING:         COLOR_BLUE,
    ISSUE_PERIOD_MISMATCH: COLOR_TEAL,
    ISSUE_UNEXPECTED:      COLOR_PINK,
}

NOTES_KEYWORDS = ["تغيير رقم", "تعديل مواليد", "تعديل اسم", "تغيير اسم",
                  "تصحيح رقم", "تصحيح اسم", "تصحيح مواليد"]

VALID_MINUTES  = [0, 15, 30, 45]

def resolve_period_for_time(fixed_time, period_schedule, hinted_period=None):
    """
    يطابق وقتاً بصيغة H:MM مع جدول الفترات اليدوي.
    يدعم أوقات 12 ساعة بدون AM/PM بتجربة h و h+12.
    يُرجع: (اسم_الفترة_الصحيحة أو None, هل_تطابق_الفترة_المكتوبة)
    """
    if not fixed_time or not period_schedule:
        return None, True
    try:
        parts = str(fixed_time).strip().split(":")
        h = int(parts[0])
        m = int(parts[1]) if len(parts) > 1 else 0
    except Exception:
        return None, True

    candidates = [h * 60 + m]
    if 1 <= h <= 11:
        candidates.append((h + 12) * 60 + m)

    def period_at(mins):
        for name, start, end in period_schedule:
            if start <= mins <= end:
                return name
        return None

    # إن وُجدت فترة مكتوبة: فضّل التفسير الذي يقع داخلها
    if hinted_period:
        for mins in candidates:
            for name, start, end in period_schedule:
                if name == hinted_period and start <= mins <= end:
                    return hinted_period, True

    matches = []
    for mins in candidates:
        found = period_at(mins)
        if found:
            matches.append(found)

    if not matches:
        return None, False if hinted_period else True

    # عند الغموض (صباح/مساء): فضّل التفسير المسائي (h+12) لأنه الأقرب لسياق الاختبارات
    chosen = matches[-1]
    return chosen, (chosen == hinted_period) if hinted_period else True


def _match_day_name(day_val, days_list):
    """يرجع اسم اليوم المطابق من القائمة، أو None."""
    s = str(day_val).strip() if day_val is not None else ""
    if not s or s.lower() == "nan":
        return None
    for day in days_list:
        if s == day or s in day or day in s:
            return day
    return None


def analyze_day_distribution(students_df, days_list, day_col, status_col, notes_col=None):
    """
    يحلل توزيع الأيام:
    - المثالي والحالة مبنيان على اللواتي «أنهت المقرر»
    - عمود شرطي: طالبات لسن «أنهت المقرر» ولهن يوم اختبار (مثل لم تنه المقرر + موعد)
    """
    if not days_list:
        return {
            "total": 0, "total_shurty": 0, "days": {},
            "unassigned": 0, "has_issue": False, "ideal_base": 0, "ideal_extra": 0,
        }

    status_s = students_df[status_col].astype(str).str.strip()
    finished_mask = status_s == "أنهت المقرر"
    finished_df = students_df[finished_mask]
    total = len(finished_df)

    d = len(days_list)
    base, extra = (0, 0) if total == 0 else divmod(total, d)
    ideal = {
        day: (base + (1 if i < extra else 0)) if total else 0
        for i, day in enumerate(days_list)
    }

    finished_actual = {day: 0 for day in days_list}
    shurty_actual = {day: 0 for day in days_list}
    unassigned = 0
    total_shurty = 0

    for _, row in students_df.iterrows():
        name_val = str(row.get("الاسم", "")).strip()
        if not name_val or name_val.lower() == "nan":
            continue

        st_val = str(row.get(status_col, "")).strip()
        note_val = ""
        if notes_col and notes_col in students_df.columns:
            note_val = str(row.get(notes_col, "")).strip()
            if note_val.lower() == "nan":
                note_val = ""

        matched = _match_day_name(row.get(day_col, ""), days_list)

        if st_val == "أنهت المقرر":
            if matched is None:
                unassigned += 1
            else:
                finished_actual[matched] += 1
            continue

        # شرطي على جدول الأيام: لها يوم وليست «أنهت المقرر»
        # (يشمل لم تنه المقرر + موعد، أو ملاحظة شرطي مع يوم)
        is_shurty_case = (
            matched is not None
            and (
                KEYWORD_RED in note_val
                or st_val == "لم تنه المقرر"
            )
        )
        if is_shurty_case:
            shurty_actual[matched] += 1
            total_shurty += 1

    days_report = {}
    has_issue = False
    for day in days_list:
        fin = finished_actual.get(day, 0)
        shu = shurty_actual.get(day, 0)
        ide = ideal.get(day, 0)
        if total == 0 and total_shurty == 0:
            status = "✅ لا يوجد توزيع"
        elif fin > ide:
            status = "🔴 ضغط — يجب تحويل " + str(fin - ide) + " طالبة"
            has_issue = True
        elif fin < ide:
            status = "🟡 نقص — يحتاج " + str(ide - fin) + " طالبة"
            has_issue = True
        else:
            status = "✅ مناسب"
        if shu > 0:
            status = status + " — مع " + str(shu) + " شرطي"
        days_report[day] = {
            "actual": fin,
            "shurty": shu,
            "total": fin + shu,
            "ideal": ide,
            "status": status,
        }

    if unassigned > 0:
        has_issue = True

    return {
        "total": total,
        "total_shurty": total_shurty,
        "days": days_report,
        "unassigned": unassigned,
        "has_issue": has_issue,
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




def _safe_sheet_name(base_name, used_names):
    """اسم ورقة Excel صالح وفريد (حد 31 حرفاً، بدون أحرف محظورة)."""
    cleaned = str(base_name or "معلمة").strip()
    for ch in ('\\', '/', '?', '*', '[', ']', ':'):
        cleaned = cleaned.replace(ch, " ")
    cleaned = " ".join(cleaned.split()) or "معلمة"
    candidate = cleaned[:31]
    if candidate not in used_names:
        used_names.add(candidate)
        return candidate
    i = 2
    while True:
        suffix = f" ({i})"
        candidate = cleaned[: 31 - len(suffix)] + suffix
        if candidate not in used_names:
            used_names.add(candidate)
            return candidate
        i += 1


def build_distribution_report(day_reports):
    """
    يبني ملف Excel واحد يحتوي:
    - ورقة "ملخص" تجمع كل المعلمات في جدول واحد
    - ورقة منفصلة لكل معلمة تفصيلية (أنهين + شرطي)
    """
    output   = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})

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
    yel_fmt  = fmt(bg="#FFE699")
    soft_yel = fmt(bg="#FFFF99")
    bold_fmt = fmt(bold=True, align="right")
    title_fmt = workbook.add_format({"bold": True, "font_name": "Calibri",
                                     "font_size": 13, "align": "center",
                                     "valign": "vcenter", "bg_color": "#EDE8F5"})
    note_fmt = workbook.add_format({
        "font_name": "Calibri", "font_size": 10, "align": "right",
        "valign": "vcenter", "font_color": "#555555",
    })

    # ── ورقة الملخص ──────────────────────────────────────────────────────────
    ws_sum = workbook.add_worksheet("ملخص")
    ws_sum.right_to_left()
    ws_sum.set_column(0, 0, 25)
    ws_sum.set_column(1, 1, 12)
    ws_sum.set_column(2, 2, 10)
    ws_sum.set_column(3, 3, 12)
    ws_sum.set_column(4, 4, 12)
    ws_sum.set_column(5, 5, 16)

    ws_sum.merge_range(0, 0, 0, 5, "ملخص توزيع الأيام — جميع المعلمات", title_fmt)
    ws_sum.set_row(0, 25)

    for ci, h in enumerate(["المعلمة", "أنهين المقرر", "شرطي", "أيام مكتظة", "أيام ناقصة", "الحالة"]):
        ws_sum.write(1, ci, h, hdr_fmt)

    for ri, (teacher, report) in enumerate(day_reports.items()):
        r          = ri + 2
        total      = report.get("total", 0)
        shurty_n   = report.get("total_shurty", 0)
        unassigned = report.get("unassigned", 0)
        over_days  = sum(1 for d in report.get("days", {}).values() if "🔴" in d["status"])
        under_days = sum(1 for d in report.get("days", {}).values() if "🟡" in d["status"])
        has_issue  = report.get("has_issue", False) or unassigned > 0

        row_fmt = red_fmt if has_issue else ok_fmt
        status_txt = "⚠️ يحتاج تدخل" if has_issue else "✅ موزّع بشكل جيد"
        ws_sum.write(r, 0, teacher,    row_fmt)
        ws_sum.write(r, 1, total,      row_fmt)
        ws_sum.write(r, 2, shurty_n,   soft_yel if shurty_n else ok_fmt)
        ws_sum.write(r, 3, over_days,  red_fmt if over_days  else ok_fmt)
        ws_sum.write(r, 4, under_days, yel_fmt if under_days else ok_fmt)
        ws_sum.write(r, 5, status_txt, row_fmt)

    # ── ورقة لكل معلمة ───────────────────────────────────────────────────────
    used_sheet_names = {"ملخص"}
    for teacher, report in day_reports.items():
        sh_name  = _safe_sheet_name(teacher, used_sheet_names)
        ws       = workbook.add_worksheet(sh_name)
        ws.right_to_left()
        for c, w in enumerate([18, 10, 10, 10, 10, 38]):
            ws.set_column(c, c, w)

        ws.merge_range(0, 0, 0, 5, "تقرير توزيع الأيام — " + teacher, title_fmt)
        ws.set_row(0, 25)

        total      = report.get("total", 0)
        shurty_n   = report.get("total_shurty", 0)
        unassigned = report.get("unassigned", 0)
        base       = report.get("ideal_base", 0)
        xtra       = report.get("ideal_extra", 0)
        ideal_txt  = str(base) + (" (+1 لأول " + str(xtra) + " أيام)" if xtra else " لكل يوم")

        ws.write(1, 0, "إجمالي اللواتي أنهين المقرر:", bold_fmt)
        ws.write(1, 1, total, ok_fmt)
        ws.write(2, 0, "عدد الشرطي (بمواعيد):", bold_fmt)
        ws.write(2, 1, shurty_n, soft_yel if shurty_n else ok_fmt)
        ws.write(3, 0, "التوزيع المثالي (حسب من أنهين):", bold_fmt)
        ws.write(3, 1, ideal_txt, ok_fmt)
        row_cursor = 4
        if unassigned:
            ws.write(row_cursor, 0, "⚠️ أنهين المقرر بدون يوم:", bold_fmt)
            ws.write(row_cursor, 1, unassigned, red_fmt)
            row_cursor += 1
        ws.merge_range(
            row_cursor, 0, row_cursor, 5,
            "المثالي والحالة لـ«أنهين المقرر» فقط — عمود شرطي إضافي للمواعيد المحجوزة",
            note_fmt,
        )
        header_row = row_cursor + 2

        for ci, h in enumerate(["اليوم", "أنهين", "شرطي", "الإجمالي", "المثالي", "الحالة"]):
            ws.write(header_row, ci, h, hdr_fmt)

        for ri, (day, info) in enumerate(report.get("days", {}).items()):
            r       = header_row + 1 + ri
            is_over = "🔴" in info["status"]
            is_under= "🟡" in info["status"]
            row_f   = red_fmt if is_over else (yel_fmt if is_under else ok_fmt)
            shu     = info.get("shurty", 0)
            ws.write(r, 0, day,                     row_f)
            ws.write(r, 1, info.get("actual", 0),   row_f)
            ws.write(r, 2, shu, soft_yel if shu else row_f)
            ws.write(r, 3, info.get("total", info.get("actual", 0)), row_f)
            ws.write(r, 4, info.get("ideal", 0),    row_f)
            ws.write(r, 5, info.get("status", ""),  row_f)

    workbook.close()
    output.seek(0)
    return output.read()


def fix_time_minutes(time_raw):
    """
    يُصحح الوقت — الدقائق المقبولة: 00، 15، 30، 45 فقط
    رقم واحد (1-5) بعد النقطتين → ×10  (12:3 → 12:30)
    رقم واحد (6-9)              → يُقرَّب (9:7  → 9:00)
    Excel serial                → يُحوَّل ويُقرَّب
    """
    if time_raw is None or str(time_raw).strip() in ("", "nan"):
        return ""
    s = str(time_raw).strip()
    # Excel serial
    try:
        f = float(s)
        if 0 < f < 1:
            total_min = round(f * 24 * 60)
            h = total_min // 60
            m = total_min % 60
            closest = min(VALID_MINUTES, key=lambda x: abs(x - m))
            return f"{h}:{closest:02d}"
    except ValueError:
        pass
    # نص وقت
    s = s.replace(".", ":").replace("٫", ":")
    if ":" not in s:
        s = (s[:-2] + ":" + s[-2:]) if len(s) > 2 else s + ":00"
    parts = s.split(":")
    try:
        h = int(parts[0])
        raw_m = parts[1].strip() if len(parts) > 1 else "0"
        m = int(raw_m)
        if len(raw_m) == 1 and m <= 5:
            m = m * 10
        closest = min(VALID_MINUTES, key=lambda x: abs(x - m))
        return f"{h}:{closest:02d}"
    except Exception:
        return str(time_raw)


def process_stage2_file(file_bytes, days_list, statuses_list, periods_list, period_schedule=None):
    """معالجة ملف واحد مُعاد من المعلمة"""

    df = read_xlsx_raw(file_bytes)

    col_map = {
        "status":  next((c for c in df.columns if "الحالة"         in str(c)), None),
        "day":     next((c for c in df.columns if "يوم الاختبار"   in str(c)), None),
        "time":    next((c for c in df.columns if "توقيت الاختبار" in str(c)), None),
        "period":  next((c for c in df.columns if "الفترة"         in str(c)), None),
        "notes":   next((c for c in df.columns if "الملاحظات"      in str(c)), None),
        "teacher": next(
            (c for c in df.columns if str(c).strip() in ("اسم المعلمة", "المعلمة")),
            next((c for c in df.columns if "المعلمة" in str(c)), None),
        ),
    }

    # اسم المعلمة الحقيقي من العمود (الأكثر تكراراً)، وليس من اسم الملف
    teacher_col_val = ""
    if col_map["teacher"]:
        vals = df[col_map["teacher"]].dropna().astype(str).str.strip()
        vals = vals[(vals != "") & (vals.str.lower() != "nan")]
        if not vals.empty:
            teacher_col_val = vals.mode().iloc[0] if not vals.mode().empty else vals.iloc[0]

    columns_order = [
        "الرقم", "الاسم", "رقم الواتس اب", "المجموعة",
        "البلد", "المواليد", "الإجازة", "المعلمة",
        "الحالة", "يوم الاختبار", "توقيت الاختبار", "الفترة", "الملاحظات",
    ]
    for col in columns_order:
        if col not in df.columns:
            df[col] = ""

    # ── فحص كل صف ────────────────────────────────────────────────────────────
    camera_rows         = []   # كاميرا  → صف كامل أحمر
    shurty_rows         = []   # شرطي فقط → خلية الحالة أصفر
    note_rows           = []   # ملاحظة جوهرية فقط → خلية الاسم برتقالي
    both_rows           = []   # شرطي + جوهرية → حالة أصفر + اسم برتقالي
    # خلايا المشاكل: {row_idx: {col_name: issue_type}}
    issue_cells         = {}
    time_format_errors  = []
    period_mismatch_rows = []

    STATUS_FINISHED = "أنهت المقرر"

    def mark_issue(row_i, col_name, issue_type):
        issue_cells.setdefault(row_i, {})[col_name] = issue_type

    def is_blank(val):
        if val is None:
            return True
        try:
            if pd.isna(val):
                return True
        except Exception:
            pass
        s = str(val).strip()
        return (not s) or s.lower() in ("nan", "none", "nat", "<na>")

    for idx, row in df.iterrows():
        status_raw = row.get(col_map["status"] or "الحالة", "")
        day_raw    = row.get(col_map["day"]    or "يوم الاختبار", "")
        note_raw   = row.get(col_map["notes"]  or "الملاحظات", "")
        time_raw   = row.get(col_map["time"]   or "توقيت الاختبار", "")
        period_raw = row.get(col_map["period"] or "الفترة", "")

        status = "" if is_blank(status_raw) else str(status_raw).strip()
        day    = "" if is_blank(day_raw)    else str(day_raw).strip()
        note   = "" if is_blank(note_raw)   else str(note_raw).strip()
        period = "" if is_blank(period_raw) else str(period_raw).strip()

        # تحقق من تنسيق التوقيت — قيمة عددية >= 1 تعني تاريخ وليس ساعة H:MM
        try:
            raw_s = str(time_raw).strip()
            if raw_s and ":" not in raw_s.replace(".", ":"):
                fval = float(raw_s)
                if fval >= 1:
                    time_format_errors.append((idx, str(time_raw)))
        except (ValueError, TypeError):
            pass

        # تصحيح الوقت وتخزينه كنص H:MM (مثل 1:45 / 8:30)
        fixed_time = fix_time_minutes(time_raw)
        has_time = bool(fixed_time) and not is_blank(fixed_time)
        if col_map["time"] and has_time:
            df.at[idx, col_map["time"]] = fixed_time

        # تخطي الصفوف الفارغة كلياً (بعد نهاية البيانات)
        name_val = str(row.get("الاسم", "")).strip()
        if not name_val or name_val == "nan":
            continue

        # الحالة إلزامية لكل الطالبات
        if is_blank(status):
            mark_issue(idx, "الحالة", ISSUE_EMPTY_STATUS)
        elif status == STATUS_FINISHED:
            # أنهت المقرر: يوم + فترة إلزاميان، التوقيت اختياري
            if is_blank(day):
                mark_issue(idx, "يوم الاختبار", ISSUE_MISSING)
            if is_blank(period):
                mark_issue(idx, "الفترة", ISSUE_MISSING)
            elif has_time and period_schedule:
                correct_period, matched = resolve_period_for_time(
                    fixed_time, period_schedule, hinted_period=period
                )
                if not matched:
                    mark_issue(idx, "الفترة", ISSUE_PERIOD_MISMATCH)
                    period_mismatch_rows.append(
                        (idx, period, correct_period or "—")
                    )
        else:
            # ليست «أنهت المقرر»
            has_exam_fields = (
                not is_blank(day) or has_time or not is_blank(period)
            )
            if status == "لم تنه المقرر" and has_exam_fields:
                # نمط شائع: لم تنه المقرر + موعد اختبار → شرطي تلقائياً + تمييز أصفر
                if KEYWORD_RED not in note:
                    note = f"{note} {KEYWORD_RED}".strip() if note else KEYWORD_RED
                    notes_col = col_map["notes"] or "الملاحظات"
                    df.at[idx, notes_col] = note
            elif has_exam_fields:
                # حالات أخرى مع حقول اختبار غير متوقعة → تلوين وردي
                if not is_blank(day):
                    mark_issue(idx, "يوم الاختبار", ISSUE_UNEXPECTED)
                if has_time:
                    mark_issue(idx, "توقيت الاختبار", ISSUE_UNEXPECTED)
                if not is_blank(period):
                    mark_issue(idx, "الفترة", ISSUE_UNEXPECTED)

        # منطق ألوان الملاحظات — كاميرا لها أولوية قصوى
        has_camera = KEYWORD_CAMERA in note
        has_shurty = KEYWORD_RED    in note
        has_note   = any(kw in note for kw in NOTES_KEYWORDS)

        if has_camera:
            camera_rows.append(idx)
        elif has_shurty and has_note:
            both_rows.append(idx)
        elif has_shurty:
            shurty_rows.append(idx)
        elif has_note:
            note_rows.append(idx)

    # ── تحليل توزيع الأيام بعد إضافة شرطي التلقائي ────────────────────────────
    day_report = {}
    if col_map["status"] and col_map["day"] and days_list:
        day_report = analyze_day_distribution(
            df, days_list, col_map["day"], col_map["status"],
            notes_col=col_map["notes"] or "الملاحظات",
        )

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

    header_fmt  = workbook.add_format({
        "bold": True, "font_name": "Calibri", "font_size": 11,
        "align": "center", "valign": "vcenter",
        "border": 1, "bg_color": COLOR_HEADER, "locked": False,
    })
    normal_fmt  = fmt()
    num_fmt     = fmt({"num_format": "0"})
    phone_fmt   = fmt({"num_format": "0"})
    time_fmt    = fmt({"num_format": "h:mm"})
    arial_fmt   = fmt({"font_name": "Arial"})
    # كاميرا — صف كامل أحمر
    cam_fmt       = fmt({"bg_color": COLOR_RED})
    cam_phone     = fmt({"bg_color": COLOR_RED, "num_format": "0"})
    cam_time      = fmt({"bg_color": COLOR_RED, "num_format": "h:mm"})
    # شرطي — أصفر فقط
    yellow_cell   = fmt({"bg_color": COLOR_YELLOW})
    # ملاحظة جوهرية — برتقالي (مميز عن الكاميرا)
    note_name_fmt = fmt({"bg_color": COLOR_ORANGE})
    # صيغ المشاكل حسب النوع
    issue_fmts = {
        itype: {
            "cell": fmt({"bg_color": color}),
            "time": fmt({"bg_color": color, "num_format": "h:mm"}),
        }
        for itype, color in ISSUE_COLORS.items()
    }

    col_widths = [7, 24, 14.1, 13.3, 7, 6, 5.3, 6.9, 19.8, 11.4, 10.7, 14, 39.8]
    for i, w in enumerate(col_widths):
        ws.set_column(i, i, w)

    for ci, cn in enumerate(columns_order):
        ws.write(0, ci, cn, header_fmt)

    for row_idx, row in df_out.iterrows():
        er        = row_idx + 1
        is_camera = row_idx in camera_rows
        is_shurty = row_idx in shurty_rows
        is_note   = row_idx in note_rows
        is_both   = row_idx in both_rows
        row_issues = issue_cells.get(row_idx, {})

        for ci, cn in enumerate(columns_order):
            val = row[cn]
            val = "" if pd.isna(val) else val

            def write_cell(f, phone_f=None, time_f=None):
                if cn == "رقم الواتس اب" and val != "":
                    use_f = phone_f if phone_f else f
                    try:
                        ws.write_number(er, ci, int(float(str(val).replace(".0",""))), use_f)
                    except Exception:
                        ws.write(er, ci, str(val), use_f)
                elif cn == "توقيت الاختبار" and val != "":
                    use_f = time_f if time_f else f
                    ws.write_string(er, ci, str(val), use_f)
                elif cn in {"الرقم", "المواليد"} and val != "":
                    try:
                        ws.write_number(er, ci, int(str(val).replace(".0", "")), f)
                    except Exception:
                        ws.write(er, ci, str(val) if isinstance(val, str) else val, f)
                else:
                    ws.write(er, ci, str(val) if isinstance(val, str) else val, f)

            def normal_f():
                if cn == "رقم الواتس اب" and val != "":
                    return phone_fmt
                if cn == "توقيت الاختبار" and val != "":
                    return time_fmt
                if cn in {"الرقم", "المواليد"} and val != "":
                    return num_fmt
                if cn == "الفترة":
                    return arial_fmt
                return normal_fmt

            def write_issue_or_normal(prefer_shurty_status=False, prefer_note_name=False):
                # ألوان المشاكل أولاً حتى لا يغطي الأصفر (شرطي) الحالة الفارغة
                if cn in row_issues:
                    itype = row_issues[cn]
                    pair = issue_fmts[itype]
                    write_cell(pair["cell"], time_f=pair["time"])
                    return
                if prefer_shurty_status and cn == "الحالة":
                    write_cell(yellow_cell)
                    return
                if prefer_note_name and cn == "الاسم":
                    write_cell(note_name_fmt)
                    return
                write_cell(normal_f(), phone_f=phone_fmt, time_f=time_fmt)

            if is_camera:
                write_cell(cam_fmt, phone_f=cam_phone, time_f=cam_time)
            elif is_both:
                write_issue_or_normal(prefer_shurty_status=True, prefer_note_name=True)
            elif is_shurty:
                write_issue_or_normal(prefer_shurty_status=True)
            elif is_note:
                write_issue_or_normal(prefer_note_name=True)
            else:
                write_issue_or_normal()

    # ── القوائم المنسدلة ─────────────────────────────────────────────────────
    last_dv_row = len(df_out) + 50

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
    n_colored = len(camera_rows) + len(shurty_rows) + len(note_rows) + len(both_rows)
    n_issues  = len(issue_cells)
    return (output.read(), n_colored, 0,
            n_issues, day_report,
            time_format_errors, period_mismatch_rows, teacher_col_val)


# ── واجهة المرحلة الثانية ───────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    """
    <div style="background:linear-gradient(135deg,#1e1b4b,#3730a3,#4338ca);border-radius:16px;
    padding:1.8rem 2.5rem;margin-bottom:1.5rem;box-shadow:0 8px 32px rgba(30,27,75,0.35);
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
        margin-bottom:1rem;border:1px solid #e0d0f8;font-size:0.82rem;direction:rtl;line-height:2;">
            <b>دليل الألوان:</b><br>
            <span style="background:#FFFF99;padding:2px 10px;border-radius:4px;">🟡 شرطي — خلية الحالة</span> &nbsp;
            <span style="background:#FF9999;padding:2px 10px;border-radius:4px;">🔴 كاميرا — صف كامل</span> &nbsp;
            <span style="background:#FFB347;padding:2px 10px;border-radius:4px;">🟠 ملاحظة جوهرية — خلية الاسم</span><br>
            <span style="background:#CE93D8;padding:2px 10px;border-radius:4px;">🟣 حالة فارغة — يجب تعبئة الحالة</span> &nbsp;
            <span style="background:#AED6F1;padding:2px 10px;border-radius:4px;">🔵 نقص يوم/فترة لمن أنهت المقرر</span><br>
            <span style="background:#76D7C4;padding:2px 10px;border-radius:4px;">🟢 عدم تطابق الفترة مع الوقت</span> &nbsp;
            <span style="background:#F5B7B1;padding:2px 10px;border-radius:4px;">🩷 يوم/وقت/فترة معبأة وحالتها ليست «أنهت المقرر»</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.button("🔄 معالجة ومراجعة الملفات", type="primary", use_container_width=True, key="stage2_btn"):
        with st.spinner("جارٍ التدقيق والمعالجة..."):
            stage2_results    = {}
            stage2_errors     = []
            day_reports       = {}
            time_fmt_warnings = {}
            period_warnings   = {}
            total_red = total_issues = 0

            for uf in uploaded_stage2:
                try:
                    fb = uf.read()
                    out_bytes, n_colored, _, n_issues, d_report, t_errors, p_mismatches, teacher_name = process_stage2_file(
                        fb, days_list, statuses_list, periods_list, period_schedule
                    )
                    display_name = teacher_name.strip() if teacher_name else uf.name.rsplit(".", 1)[0]
                    unique_name = display_name
                    n = 2
                    while (unique_name + ".xlsx") in stage2_results or unique_name in day_reports:
                        unique_name = f"{display_name} ({n})"
                        n += 1
                    out_name = unique_name + ".xlsx"
                    stage2_results[out_name] = out_bytes
                    total_red    += n_colored
                    total_issues += n_issues
                    if d_report:
                        day_reports[unique_name] = d_report
                    if t_errors:
                        time_fmt_warnings[unique_name] = t_errors
                    if p_mismatches:
                        period_warnings[unique_name] = p_mismatches
                except Exception as e:
                    stage2_errors.append("❌ " + uf.name + ": " + str(e))

            st.session_state["stage2_results"] = stage2_results
            st.session_state["stage2_errors"] = stage2_errors
            st.session_state["stage2_day_reports"] = day_reports
            st.session_state["stage2_time_fmt_warnings"] = time_fmt_warnings
            st.session_state["stage2_period_warnings"] = period_warnings
            st.session_state["stage2_total_red"] = total_red
            st.session_state["stage2_total_issues"] = total_issues

    if st.session_state.get("stage2_results") is not None:
        render_stage2_results_ui(
            stage2_results=st.session_state["stage2_results"],
            day_reports=st.session_state.get("stage2_day_reports", {}),
            stage2_errors=st.session_state.get("stage2_errors", []),
            time_fmt_warnings=st.session_state.get("stage2_time_fmt_warnings", {}),
            period_warnings=st.session_state.get("stage2_period_warnings", {}),
            total_red=st.session_state.get("stage2_total_red", 0),
            total_issues=st.session_state.get("stage2_total_issues", 0),
        )

# ═══════════════════════════════════════════════════════════════════════════════
# المرحلة الثالثة — تجميع ملفات المعلمات في ملف لجان واحد
# ═══════════════════════════════════════════════════════════════════════════════

DAYS_ORDER = ["الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت", "الأحد"]

def day_sort_key(day_val, days_list):
    day_str = str(day_val).strip()
    for i, d in enumerate(days_list):
        if day_str == d or day_str in d or d in day_str:
            return i
    for i, d in enumerate(DAYS_ORDER):
        if day_str == d or day_str in d or d in day_str:
            return i + 100
    return 999


def build_stage3_file(files_dict, days_list, existing_bytes=None):
    all_rows = []
    for fname, fb in files_dict.items():
        try:
            df = read_xlsx_raw(fb)
            df.columns = [str(c).strip() for c in df.columns]
            df = df.dropna(how="all")
            if "الاسم" in df.columns:
                df = df[df["الاسم"].astype(str).str.strip().replace("nan", "") != ""]
            all_rows.append(df)
        except Exception:
            pass

    if not all_rows:
        raise ValueError("لم يتم قراءة أي بيانات")

    combined = pd.concat(all_rows, ignore_index=True)

    # ── إذا وُجد ملف أم: استخرج بياناته وادمجها ─────────────────────────────
    if existing_bytes:
        try:
            existing_sheets = read_existing_stage3(existing_bytes)
            existing_rows = []
            for sheet_name, df_ex in existing_sheets.items():
                if df_ex.empty: continue
                df_ex.columns = [str(c).strip() for c in df_ex.columns]
                df_ex = df_ex.dropna(how="all")
                if "الاسم" in df_ex.columns:
                    df_ex = df_ex[df_ex["الاسم"].astype(str).str.strip().replace("nan","") != ""]
                if not df_ex.empty:
                    existing_rows.append(df_ex)
            if existing_rows:
                combined = pd.concat([pd.concat(existing_rows, ignore_index=True), combined],
                                     ignore_index=True)
        except Exception as ex:
            pass  # إذا فشلت القراءة نكمل بالملفات الجديدة فقط

    cols = [
        "الرقم", "الاسم", "رقم الواتس اب", "المجموعة", "البلد",
        "المواليد", "الإجازة", "المعلمة", "الحالة",
        "يوم الاختبار", "توقيت الاختبار", "الفترة", "الملاحظات",
    ]
    for c in cols:
        if c not in combined.columns:
            combined[c] = ""
    combined = combined[cols].copy()

    # تنظيف شامل — يزيل المسافات الخفية وكل أشكال الفراغ
    for c in cols:
        combined[c] = (combined[c]
                       .fillna("")
                       .astype(str)
                       .str.strip()
                       .str.replace("\u00a0", "", regex=False)
                       .replace("nan", ""))

    # حذف صفوف اسمها فارغ
    combined = combined[combined["الاسم"] != ""].reset_index(drop=True)

    # ── تقسيم الأوراق ────────────────────────────────────────────────────────
    # الترتيب: اختبار مبكر → (يوم+موعد أو شرطي أو أنهت المقرر) → المتقدمات
    #           وإلا → غير متقدمات
    def _cell_filled(val):
        s = str(val or "").strip()
        return bool(s) and s.lower() not in ("nan", "none", "nat", "<na>")

    def _has_shurty(notes_val, status_val=""):
        note = str(notes_val or "")
        st   = str(status_val or "")
        return (KEYWORD_RED in note) or (KEYWORD_RED in st)

    mask_early = combined["الملاحظات"].astype(str).str.contains(
        "قدمت الاختبار", na=False
    )
    mask_shurty = (~mask_early) & combined.apply(
        lambda r: _has_shurty(r["الملاحظات"], r["الحالة"]), axis=1
    )
    # موعد مكتمل: يوم الاختبار + توقيت الاختبار معاً → متقدمة بغض النظر عن «لم تنه المقرر»
    mask_has_slot = (~mask_early) & combined.apply(
        lambda r: _cell_filled(r["يوم الاختبار"]) and _cell_filled(r["توقيت الاختبار"]),
        axis=1,
    )
    mask_finished_course = combined["الحالة"].astype(str).str.strip() == "أنهت المقرر"
    mask_applicants = (~mask_early) & (
        mask_finished_course | mask_shurty | mask_has_slot
    )
    mask_others = (~mask_applicants) & (~mask_early)

    df_finished = combined[mask_applicants].copy()
    df_others   = combined[mask_others].copy()
    df_early    = combined[mask_early].copy()
    # أصفر على الحالة: شرطي صراحة، أو موعد مع حالة ليست «أنهت المقرر» (تمييز شرطي)
    df_finished["_is_shurty"] = (
        mask_shurty.loc[df_finished.index]
        | (
            mask_has_slot.loc[df_finished.index]
            & ~mask_finished_course.loc[df_finished.index]
        )
    ).values
    # إحصاء الشرطي للعرض = كل من يُبرز بالأصفر في المتقدمات
    mask_shurty_highlight = mask_shurty | (
        mask_has_slot & ~mask_finished_course & ~mask_early
    )

    # ── الترتيب ──────────────────────────────────────────────────────────────
    df_finished["_day"]  = df_finished["يوم الاختبار"].apply(lambda x: day_sort_key(x, days_list))
    df_finished["_time"] = pd.to_numeric(df_finished["توقيت الاختبار"], errors="coerce").fillna(999)
    df_finished = (
        df_finished
        .sort_values(["المعلمة", "_day", "_time"])
        .drop(columns=["_day", "_time"])
        .reset_index(drop=True)
    )

    df_others = df_others.sort_values(["المعلمة", "الاسم"]).reset_index(drop=True)
    df_early  = df_early.sort_values(["المعلمة", "الاسم"]).reset_index(drop=True)

    # ── بناء Excel ───────────────────────────────────────────────────────────
    output   = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})

    header_fmt = workbook.add_format({"bold": True, "font_name": "Calibri", "font_size": 11,
        "align": "center", "valign": "vcenter", "border": 1})
    cell_fmt  = workbook.add_format({"font_name": "Calibri", "font_size": 11,
        "align": "center", "valign": "vcenter", "border": 1})
    num_fmt   = workbook.add_format({"font_name": "Calibri", "font_size": 11,
        "align": "center", "valign": "vcenter", "border": 1, "num_format": "0"})
    time_fmt  = workbook.add_format({"font_name": "Calibri", "font_size": 11,
        "align": "center", "valign": "vcenter", "border": 1, "num_format": "h:mm"})
    arial_fmt = workbook.add_format({"font_name": "Arial", "font_size": 11,
        "align": "center", "valign": "vcenter", "border": 1})
    # شرطي — خلية الحالة أصفر (نفس لون المرحلة الثانية)
    yellow_status_fmt = workbook.add_format({
        "font_name": "Calibri", "font_size": 11,
        "align": "center", "valign": "vcenter", "border": 1,
        "bg_color": COLOR_YELLOW,
    })

    numeric_set = {"الرقم", "رقم الواتس اب", "المواليد"}
    col_widths  = {
        "الرقم": 7, "الاسم": 24, "رقم الواتس اب": 14,
        "المجموعة": 13, "البلد": 7, "المواليد": 6,
        "الإجازة": 5.3, "المعلمة": 7, "الحالة": 20,
        "يوم الاختبار": 11, "توقيت الاختبار": 11,
        "الفترة": 14, "الملاحظات": 40,
    }

    def write_sheet(name, df_sheet, highlight_shurty=False):
        ws = workbook.add_worksheet(name)
        ws.right_to_left()
        for ci, cn in enumerate(cols):
            ws.set_column(ci, ci, col_widths.get(cn, 12))
        for ci, cn in enumerate(cols):
            ws.write(0, ci, cn, header_fmt)
        for ri, row in df_sheet.iterrows():
            er = ri + 1
            is_shurty = bool(highlight_shurty and row.get("_is_shurty", False))
            for ci, cn in enumerate(cols):
                val = row[cn]
                if cn == "الحالة" and is_shurty:
                    ws.write(er, ci, val, yellow_status_fmt)
                elif cn in numeric_set and val not in ("", "nan"):
                    try:
                        ws.write_number(er, ci, int(str(val).replace(".0", "")), num_fmt)
                    except Exception:
                        ws.write(er, ci, val, cell_fmt)
                elif cn == "توقيت الاختبار" and val not in ("", "nan"):
                    try:
                        fval = float(val)
                        if 0 < fval < 1:
                            ws.write_number(er, ci, fval, time_fmt)
                        else:
                            ws.write(er, ci, val, cell_fmt)
                    except Exception:
                        ws.write(er, ci, val, cell_fmt)
                elif cn == "الفترة":
                    ws.write(er, ci, val, arial_fmt)
                else:
                    ws.write(er, ci, val, cell_fmt)

    write_sheet("المتقدمات للاختبار", df_finished, highlight_shurty=True)
    write_sheet("غير متقدمات",        df_others)
    write_sheet("اختبار مبكر",        df_early)

    workbook.close()
    output.seek(0)
    n_shurty = int(mask_shurty_highlight.sum())
    return output.read(), len(df_finished), len(df_others), len(df_early), n_shurty



def read_existing_stage3(file_bytes):
    """
    يقرأ ملف اللجان الأم الموجود (3 أوراق) ويُعيد dict:
    {"المتقدمات للاختبار": df, "غير متقدمات": df, "اختبار مبكر": df}
    """
    import zipfile as zf_mod
    NS2 = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    def col_idx(col_str):
        idx = 0
        for ch in col_str.upper():
            idx = idx * 26 + (ord(ch) - ord('A') + 1)
        return idx - 1

    result = {}
    with zf_mod.ZipFile(io.BytesIO(file_bytes)) as zf:
        shared = []
        if "xl/sharedStrings.xml" in zf.namelist():
            tree = ET.parse(zf.open("xl/sharedStrings.xml"))
            for si in tree.getroot().iter("{" + NS2 + "}si"):
                shared.append("".join(t.text or "" for t in si.iter("{" + NS2 + "}t")))

        rels = {r.attrib["Id"]: r.attrib["Target"]
                for r in ET.parse(zf.open("xl/_rels/workbook.xml.rels")).getroot()}
        sheets_el = ET.parse(zf.open("xl/workbook.xml")).getroot().find("{" + NS2 + "}sheets")

        for sh in sheets_el:
            sh_name = sh.attrib.get("name", "")
            r_id = sh.attrib.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            )
            target = rels.get(r_id, "")
            sheet_path = (target[1:] if target.startswith("/xl/")
                         else ("xl/" + target if not target.startswith("xl/") else target))

            rows_dict = {}; max_col = 0
            sheet_data = ET.parse(zf.open(sheet_path)).getroot().find("{" + NS2 + "}sheetData")
            all_rows = list(sheet_data.iter("{" + NS2 + "}row"))
            if not all_rows:
                result[sh_name] = pd.DataFrame()
                continue

            for row_el in all_rows:
                row_num = int(row_el.attrib.get("r", 0)) - 1
                row_dict = {}
                for c in row_el.iter("{" + NS2 + "}c"):
                    addr = c.attrib.get("r", "A1")
                    ci = col_idx("".join(ch for ch in addr if ch.isalpha()))
                    max_col = max(max_col, ci)
                    t = c.attrib.get("t", "")
                    if t == "inlineStr":
                        is_el = c.find("{" + NS2 + "}is")
                        row_dict[ci] = ("".join(x.text or "" for x in is_el.iter("{" + NS2 + "}t"))
                                        if is_el is not None else "")
                    elif t == "s":
                        v_el = c.find("{" + NS2 + "}v")
                        row_dict[ci] = shared[int(v_el.text)] if v_el is not None and v_el.text else ""
                    else:
                        v_el = c.find("{" + NS2 + "}v")
                        if v_el is not None and v_el.text:
                            val = v_el.text
                            try: val = int(val) if "." not in val else float(val)
                            except: pass
                            row_dict[ci] = val
                        else: row_dict[ci] = ""
                rows_dict[row_num] = row_dict

            max_row = max(rows_dict.keys())
            matrix = [[rows_dict.get(r, {}).get(c, "") for c in range(max_col + 1)]
                      for r in range(max_row + 1)]
            headers = [str(v).strip() if v != "" else "col_" + str(i)
                       for i, v in enumerate(matrix[0])]
            df = pd.DataFrame(matrix[1:], columns=headers)
            result[sh_name] = df

    return result

# ── واجهة المرحلة الثالثة ────────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    """
    <div style="background:linear-gradient(135deg,#7c2d12,#9a3412,#c2410c);border-radius:16px;
    padding:1.8rem 2.5rem;margin-bottom:1.5rem;box-shadow:0 8px 32px rgba(124,45,18,0.35);
    text-align:center;color:white;">
        <div style="font-size:1.8rem;font-weight:900;margin:0;">📊 المرحلة الثالثة — تجميع اللجان</div>
        <div style="font-size:0.95rem;margin:0.4rem 0 0;opacity:0.88;">
            ارفعي ملفات المعلمات المُراجعة لتجميعها في ملف لجان واحد — وطالبات «شرطي» تُدرج مع المتقدمات بلون أصفر
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="section-title">📂 رفع ملفات المعلمات الجديدة</div>', unsafe_allow_html=True)

uploaded_stage3 = st.file_uploader(
    "ارفعي ملفات المعلمات",
    type=["xlsx"],
    accept_multiple_files=True,
    label_visibility="collapsed",
    key="stage3_uploader",
)

if uploaded_stage3:
    chips3 = " ".join(
        ['<span class="file-chip">📄 ' + f.name + "</span>" for f in uploaded_stage3]
    )
    st.markdown("<div style='margin:0.5rem 0 1rem'>" + chips3 + "</div>", unsafe_allow_html=True)

    # ── هل يوجد ملف أم سابق؟ ─────────────────────────────────────────────────
    st.markdown('<div class="section-title">📁 هل لديكِ ملف لجان سابق؟</div>', unsafe_allow_html=True)
    has_existing = st.radio(
        "اختاري:",
        ["لا — أنشئي ملفاً جديداً من الملفات المرفوعة",
         "نعم — أضيفي الملفات المرفوعة للملف الأم الموجود"],
        key="stage3_mode",
        label_visibility="collapsed",
    )

    existing_file = None
    if "نعم" in has_existing:
        st.markdown(
            "<div style='font-size:0.88rem;color:#555;direction:rtl;margin-bottom:0.5rem;'>"
            "ارفعي ملف اللجان الأم (الذي يحتوي 3 أوراق: متقدمات / غير متقدمات / اختبار مبكر)"
            "</div>", unsafe_allow_html=True,
        )
        existing_file = st.file_uploader(
            "ملف اللجان الأم",
            type=["xlsx"],
            key="stage3_existing",
            label_visibility="collapsed",
        )
        if existing_file:
            st.success("✅ تم رفع الملف الأم: " + existing_file.name)

    output_name = st.text_input(
        "اسم الملف الناتج",
        value="اللجان_المجمعة.xlsx",
        key="stage3_output_name",
    )

    btn_label = "📊 إضافة للملف الأم وتحميل" if "نعم" in has_existing else "📊 تجميع وبناء اللجان"

    if st.button(btn_label, type="primary", use_container_width=True, key="stage3_btn"):
        with st.spinner("جارٍ التجميع..."):
            files_dict = {}
            read_errors = []
            for uf in uploaded_stage3:
                try:
                    uf.seek(0)
                    files_dict[uf.name] = uf.read()
                except Exception as e:
                    read_errors.append("❌ " + uf.name + ": " + str(e))

            for e in read_errors:
                st.error(e)

            if files_dict:
                try:
                    # ── إذا يوجد ملف أم: اقرأ بياناته وادمجه مع الجديد ─────
                    existing_bytes = None
                    if "نعم" in has_existing and existing_file:
                        existing_file.seek(0)
                        existing_bytes = existing_file.read()

                    result_bytes, n_fin, n_oth, n_ear, n_shurty = build_stage3_file(
                        files_dict, days_list, existing_bytes=existing_bytes
                    )

                    mode_label = "إضافة للملف الأم" if existing_bytes else "ملف جديد"
                    st.success("✅ " + mode_label + " — تم بنجاح")

                    cols3 = st.columns(5)
                    with cols3[0]:
                        st.markdown('<div class="stat-card"><div class="number">' + str(len(files_dict)) + '</div><div class="label">ملف مُضاف</div></div>', unsafe_allow_html=True)
                    with cols3[1]:
                        st.markdown('<div class="stat-card"><div class="number" style="color:#1a4e1a;">' + str(n_fin) + '</div><div class="label">متقدمة ✅</div></div>', unsafe_allow_html=True)
                    with cols3[2]:
                        st.markdown('<div class="stat-card"><div class="number" style="color:#b7950b;">' + str(n_shurty) + '</div><div class="label">شرطي (أصفر) 🟨</div></div>', unsafe_allow_html=True)
                    with cols3[3]:
                        st.markdown('<div class="stat-card"><div class="number" style="color:#b7950b;">' + str(n_oth) + '</div><div class="label">غير متقدمة 🕐</div></div>', unsafe_allow_html=True)
                    with cols3[4]:
                        st.markdown('<div class="stat-card"><div class="number" style="color:#1a5276;">' + str(n_ear) + '</div><div class="label">اختبار مبكر 🎓</div></div>', unsafe_allow_html=True)

                    fname_out = output_name if output_name.endswith(".xlsx") else output_name + ".xlsx"
                    st.download_button(
                        label="⬇️ تحميل ملف اللجان المجمعة",
                        data=result_bytes,
                        file_name=fname_out,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="stage3_download",
                    )

                except Exception as e:
                    st.error("❌ خطأ في التجميع: " + str(e))
                    
                    
st.markdown(
    """
    <hr style="margin:2rem 0 1rem; border-color:#d8c8f0;">
    <div style="text-align:center; color:#999; font-size:0.8rem; font-family:'Tajawal',sans-serif;">
        أداة مقرأة — مبنية بـ Python & Streamlit &nbsp;|&nbsp; 📖
    </div>
    """,
    unsafe_allow_html=True,
)
