# =============================================================================
# أداة مقرأة — النسخة المحسّنة
# التحسينات المطبّقة:
#   1. ثابت XLSX_NS موحّد بدلاً من NS + NS2
#   2. col_letter_to_index معرّفة مرة واحدة فقط
#   3. read_xlsx_raw تدعم single / multi sheet بمعامل واحد
#   4. دمج process_files + process_files_from_cache في دالة واحدة
#   5. parse_period_schedule خارج with st.sidebar (مستوى عالٍ)
#   6. iterrows() → itertuples() / to_dict('records') للأداء
#   7. html.escape() على كل بيانات المستخدم في HTML
#   8. ثوابت مُسمّاة بدلاً من الأرقام السحرية
#   9. مفاتيح session_state مجمّعة في كلاس SK
#  10. CSS مُستخرَج كثابت
#  11. f-strings بدلاً من الـ concatenation
# =============================================================================

import html as html_lib
import io
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd
import streamlit as st
import xlsxwriter


# ═══════════════════════════════════════════════════════════════════════════════
# ١. الثوابت العامة
# ═══════════════════════════════════════════════════════════════════════════════

# Namespace الموحّد لـ OOXML — كان مكرّراً كـ NS و NS2
XLSX_NS     = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XLSX_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# ── تخطيط الأعمدة (الأرقام السحرية أصبحت ثوابت مُسمّاة) ──────────────────
COL_STATUS       = 8    # العمود I = الحالة
COL_DAY          = 9    # العمود J = يوم الاختبار
COL_TIME         = 10   # العمود K = توقيت الاختبار
COL_PERIOD       = 11   # العمود L = الفترة
COL_NOTES        = 12   # العمود M = الملاحظات
LOCKED_COLS_END  = 8    # الأعمدة A–H (0–7) مقفلة
EXTRA_BLANK_ROWS = 50   # صفوف فارغة إضافية للمعلمة

# ── ألوان وكلمات مفتاحية للمرحلة الثانية ──────────────────────────────────
KEYWORD_CAMERA = "كاميرا"
KEYWORD_RED    = "شرطي"
COLOR_RED      = "#FF9999"
COLOR_YELLOW   = "#FFFF99"
COLOR_HEADER   = "#D9D9D9"

NOTES_KEYWORDS = [
    "تغيير رقم", "تعديل مواليد", "تعديل اسم", "تغيير اسم",
    "تصحيح رقم", "تصحيح اسم", "تصحيح مواليد",
]
VALID_MINUTES = [0, 15, 30, 45]

PERIOD_AMPM = {
    "فجراً": "AM", "ضحى": "AM", "ظهراً": "AM",
    "عصراً": "PM", "ليلاً": "PM",
}

# الجدول الافتراضي للفترات (يُستخدم كاحتياطي إذا لم يُعرَّف في الـ sidebar)
DEFAULT_PERIOD_SCHEDULE = [
    ("فجراً",  5*60+45,  9*60+0),
    ("ضحى",    9*60+15, 12*60+30),
    ("ظهراً", 12*60+45, 16*60+15),
    ("عصراً", 16*60+30, 19*60+0),
    ("ليلاً", 19*60+15, 21*60+30),
]

DAYS_ORDER = ["الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت", "الأحد"]

COLUMNS_ORDER = [
    "الرقم", "الاسم", "رقم الواتس اب", "المجموعة",
    "البلد", "المواليد", "الإجازة", "المعلمة",
    "الحالة", "يوم الاختبار", "توقيت الاختبار", "الفترة", "الملاحظات",
]
NUMERIC_COLS = {"الرقم", "رقم الواتس اب", "المواليد"}
COL_WIDTHS   = [7, 24, 14.1, 13.3, 7, 6, 5.3, 6.9, 19.8, 11.4, 10.7, 14, 39.8]


# ── مفاتيح session_state مجمّعة في كلاس لتجنب الأخطاء الإملائية ───────────
class SK:
    PREVIEW_MAP    = "preview_map"
    FILE_CACHE     = "file_bytes_cache"
    PREVIEW_ERRORS = "preview_errors"
    STAGE1_RESULTS = "stage1_results"
    STAGE1_ERRORS  = "stage1_errors"


# ── CSS مُستخرَج كثابت (بدلاً من وجوده مدمجاً وسط الكود) ──────────────────
CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700;900&display=swap');

html, body, [class*="css"] { font-family: 'Tajawal', sans-serif; direction: rtl; }
.stApp { background: linear-gradient(135deg, #f3f0f8 0%, #e8e0f0 100%); }
h1, h2, h3 { font-family: 'Tajawal', sans-serif !important; }

.hero-header {
    background: linear-gradient(135deg, #3d2060 0%, #6b3fa0 50%, #8b5cc8 100%);
    border-radius: 16px; padding: 2rem 2.5rem; margin-bottom: 2rem;
    box-shadow: 0 8px 32px rgba(61,32,96,0.3); text-align: center; color: white;
}
.hero-header h1 { font-size: 2.4rem; font-weight: 900; margin: 0; text-shadow: 0 2px 4px rgba(0,0,0,0.2); }
.hero-header p  { font-size: 1.05rem; margin: 0.5rem 0 0; opacity: 0.88; font-weight: 300; }

.stat-card {
    background: white; border-radius: 12px; padding: 1.2rem 1.5rem;
    box-shadow: 0 2px 12px rgba(0,0,0,0.07); border-right: 4px solid #6b3fa0; margin-bottom: 1rem;
}
.stat-card .number { font-size: 2rem; font-weight: 900; color: #3d2060; line-height: 1; }
.stat-card .label  { font-size: 0.85rem; color: #777; margin-top: 4px; }

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

[data-testid="stSidebar"] > div:first-child {
    background: linear-gradient(180deg, #2d1b4e 0%, #3d2060 100%) !important;
}
[data-testid="stSidebar"],
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] div,
[data-testid="stSidebar"] label { color: #e8d5f8 !important; font-family: 'Tajawal', sans-serif !important; }
[data-testid="stSidebar"] label { font-weight: 700 !important; font-size: 0.95rem !important; }

[data-testid="stSidebar"] textarea,
[data-testid="stSidebar"] .stTextArea textarea {
    background-color: #1e1035 !important; border: 2px solid #9b6fd4 !important;
    border-radius: 8px !important; color: #f0e6ff !important;
    font-family: 'Tajawal', sans-serif !important; font-size: 0.92rem !important;
    direction: rtl !important; caret-color: #e8d5f8 !important;
}
[data-testid="stSidebar"] textarea:focus,
[data-testid="stSidebar"] .stTextArea textarea:focus {
    border-color: #c4a0e8 !important;
    box-shadow: 0 0 0 2px rgba(196,160,232,0.3) !important;
}
[data-testid="stSidebar"] textarea::placeholder { color: #9b7dbf !important; }
[data-testid="stSidebar"] small,
[data-testid="stSidebar"] .stMarkdown { color: #c4a0e8 !important; }

.stButton > button { font-family: 'Tajawal', sans-serif !important; font-weight: 700 !important; border-radius: 10px !important; }
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #3d2060, #6b3fa0) !important;
    border: none !important; color: white !important;
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
"""


# ═══════════════════════════════════════════════════════════════════════════════
# ٢. إعداد الصفحة
# ═══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="أداة مقرأة",
    page_icon="📖",
    layout="wide",
    initial_sidebar_state="expanded",
)
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# ٣. دوال القراءة والمعالجة — المستوى الأساسي
# ═══════════════════════════════════════════════════════════════════════════════

def col_letter_to_index(col_str: str) -> int:
    """
    تحويل حروف عمود Excel إلى مؤشر صفري.
    مثال: A→0, B→1, Z→25, AA→26
    (كانت مكررة مرتين في الكود الأصلي — الآن تعريف واحد فقط)
    """
    idx = 0
    for ch in col_str.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def _resolve_sheet_path(target: str) -> str:
    """تحويل مسار الـ relationship إلى مسار داخل ZIP."""
    if target.startswith("/xl/"):
        return target[1:]
    if not target.startswith("xl/"):
        return "xl/" + target
    return target


def read_xlsx_raw(file_bytes: bytes, multi_sheet: bool = False):
    """
    يقرأ ملف xlsx بتحليل XML مباشرة — يتجاوز openpyxl كلياً.
    يعالج: inlineStr, shared strings, numeric, boolean.

    إذا كان multi_sheet=False (الافتراضي):
        يعيد DataFrame للورقة الأولى فقط.
    إذا كان multi_sheet=True:
        يعيد dict {sheet_name: DataFrame} لجميع الأوراق.

    (دمج read_xlsx_raw الأصلية + read_existing_stage3 + النسخة المضمّنة)
    """
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:

        # ── جدول السلاسل المشتركة ─────────────────────────────────────────
        shared: list[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.parse(zf.open("xl/sharedStrings.xml")).getroot()
            for si in root.iter(f"{{{XLSX_NS}}}si"):
                shared.append("".join(t.text or "" for t in si.iter(f"{{{XLSX_NS}}}t")))

        # ── علاقات الأوراق ────────────────────────────────────────────────
        rels_root = ET.parse(zf.open("xl/_rels/workbook.xml.rels")).getroot()
        rels = {r.attrib["Id"]: r.attrib["Target"] for r in rels_root}

        wb_root   = ET.parse(zf.open("xl/workbook.xml")).getroot()
        sheets_el = wb_root.find(f"{{{XLSX_NS}}}sheets")

        def _parse_sheet(sheet_path: str) -> pd.DataFrame:
            """تحليل ورقة واحدة وإرجاع DataFrame."""
            sheet_root = ET.parse(zf.open(sheet_path)).getroot()
            sheet_data = sheet_root.find(f"{{{XLSX_NS}}}sheetData")

            rows_dict: dict[int, dict[int, object]] = {}
            max_col = 0

            for row_el in sheet_data.iter(f"{{{XLSX_NS}}}row"):
                row_num  = int(row_el.attrib.get("r", 0)) - 1  # 0-based
                row_dict: dict[int, object] = {}

                for c in row_el.iter(f"{{{XLSX_NS}}}c"):
                    addr        = c.attrib.get("r", "A1")
                    col_letters = "".join(ch for ch in addr if ch.isalpha())
                    col_idx     = col_letter_to_index(col_letters)
                    max_col     = max(max_col, col_idx)
                    cell_type   = c.attrib.get("t", "")

                    if cell_type == "inlineStr":
                        is_el = c.find(f"{{{XLSX_NS}}}is")
                        row_dict[col_idx] = (
                            "".join(t.text or "" for t in is_el.iter(f"{{{XLSX_NS}}}t"))
                            if is_el is not None else ""
                        )
                    elif cell_type == "s":
                        v_el = c.find(f"{{{XLSX_NS}}}v")
                        if v_el is not None and v_el.text is not None:
                            try:
                                row_dict[col_idx] = shared[int(v_el.text)]
                            except (IndexError, ValueError):
                                row_dict[col_idx] = v_el.text
                        else:
                            row_dict[col_idx] = ""
                    elif cell_type == "b":
                        v_el = c.find(f"{{{XLSX_NS}}}v")
                        row_dict[col_idx] = bool(int(v_el.text)) if v_el is not None else ""
                    else:
                        v_el = c.find(f"{{{XLSX_NS}}}v")
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
                return pd.DataFrame()

            max_row = max(rows_dict.keys())
            matrix  = [
                [rows_dict.get(r, {}).get(c, "") for c in range(max_col + 1)]
                for r in range(max_row + 1)
            ]
            headers = [
                str(v).strip() if v != "" else f"col_{i}"
                for i, v in enumerate(matrix[0])
            ]
            return pd.DataFrame(matrix[1:], columns=headers)

        # ── تحديد الأوراق المطلوبة ────────────────────────────────────────
        if multi_sheet:
            result: dict[str, pd.DataFrame] = {}
            for sh in sheets_el:
                sh_name = sh.attrib.get("name", "")
                r_id    = sh.attrib.get(f"{{{XLSX_REL_NS}}}id")
                target  = rels.get(r_id, "")
                if not target:
                    continue
                try:
                    path = _resolve_sheet_path(target)
                    result[sh_name] = _parse_sheet(path)
                except Exception:
                    result[sh_name] = pd.DataFrame()
            return result
        else:
            # الورقة الأولى فقط
            first_sheet = sheets_el[0]
            r_id   = first_sheet.attrib.get(f"{{{XLSX_REL_NS}}}id")
            target = rels[r_id]
            path   = _resolve_sheet_path(target)
            df     = _parse_sheet(path)
            if df.empty:
                raise ValueError("الملف فارغ")
            return df


# ═══════════════════════════════════════════════════════════════════════════════
# ٤. دوال أسماء المعلمات
# ═══════════════════════════════════════════════════════════════════════════════

def build_teacher_display_names(full_names_text: str) -> dict[str, str]:
    """
    تأخذ نصاً فيه اسم معلمة في كل سطر وترجع:
    {اسم_كامل: اسم_للعرض}
    - اسم فريد  → الاسم الأول فقط:  ابتسام
    - اسم مكرر  → أول حرف من الكنية: إيمان.ح / إيمان.ن
    - تعارض     → يزداد عدد الأحرف: آلاء.شي / آلاء.شب
    """
    lines = [l.strip() for l in full_names_text.strip().splitlines() if l.strip()]
    if not lines:
        return {}

    def make_display(parts: list[str], n_chars: int = 1) -> str:
        if len(parts) == 1:
            return parts[0]
        last   = parts[-1]
        suffix = last[2:] if (last.startswith("ال") and len(last) > 2) else last
        return f"{parts[0]}.{suffix[:n_chars]}"

    first_count: dict[str, int] = {}
    for name in lines:
        f = name.split()[0] if name.split() else name
        first_count[f] = first_count.get(f, 0) + 1

    result: dict[str, str] = {}
    for name in lines:
        parts = name.split()
        if parts and first_count[parts[0]] == 1:
            result[name] = parts[0]

    duplicates = [n for n in lines if n not in result]
    for max_chars in range(1, 6):
        temp: dict[str, str] = {}
        for name in duplicates:
            if name not in result:
                temp[name] = make_display(name.split(), max_chars)

        combined_d = {**result, **temp}
        count: dict[str, int] = {}
        for d in combined_d.values():
            count[d] = count.get(d, 0) + 1

        for name, display in temp.items():
            if count[display] == 1:
                result[name] = display

        if len(result) == len(lines):
            break

    for name in lines:
        if name not in result:
            result[name] = name

    return result


def get_first_name(full_name: str) -> str:
    parts = str(full_name).strip().split()
    return parts[0] if parts else str(full_name)


def parse_list(text: str) -> list[str]:
    return [line.strip() for line in text.strip().splitlines() if line.strip()]


# ═══════════════════════════════════════════════════════════════════════════════
# ٥. parse_period_schedule — مُنقولة للمستوى الأعلى (كانت داخل sidebar!)
# ═══════════════════════════════════════════════════════════════════════════════

def parse_period_schedule(text: str) -> list[tuple[str, int, int]]:
    """
    يحوّل نص مثل:
        فجراً: 4:00-8:45
    إلى قائمة من (اسم_الفترة, بداية_بالدقائق, نهاية_بالدقائق).

    ملاحظة: كانت الدالة معرّفة داخل `with st.sidebar` مما يعني
    إعادة تعريفها في كل إعادة رسم للواجهة — تم نقلها للمستوى الأعلى.
    """
    schedule: list[tuple[str, int, int]] = []
    for line in text.strip().splitlines():
        line = line.strip()
        if ":" not in line:
            continue
        name, _, times = line.partition(":")
        times = times.strip()
        if "-" not in times:
            continue

        def to_min(t: str) -> int:
            t  = t.strip()
            h, _, m = t.partition(":")
            return int(h) * 60 + (int(m) if m else 0)

        try:
            start_str, _, end_str = times.partition("-")
            schedule.append((name.strip(), to_min(start_str), to_min(end_str)))
        except Exception:
            continue
    return schedule


# ═══════════════════════════════════════════════════════════════════════════════
# ٦. بناء ملف Excel المرحلة الأولى
# ═══════════════════════════════════════════════════════════════════════════════

def build_excel(df: pd.DataFrame, days: list, periods: list, statuses: list) -> bytes:
    """
    يبني ملف Excel منسّق ومحمي للمعلمة.
    تحسين: استبدال iterrows() بـ to_dict('records') للأداء.
    """
    output = io.BytesIO()

    for col in COLUMNS_ORDER:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUMNS_ORDER].copy()

    num_rows  = len(df)
    workbook  = xlsxwriter.Workbook(output, {"in_memory": True})
    ws        = workbook.add_worksheet("الطالبات")
    ws.right_to_left()

    # ── التنسيقات ────────────────────────────────────────────────────────────
    header_fmt = workbook.add_format({
        "bold": True, "font_name": "Calibri", "font_size": 11,
        "align": "center", "valign": "vcenter", "border": 1, "locked": True,
    })
    locked_text_fmt = workbook.add_format({
        "font_name": "Calibri", "font_size": 11,
        "align": "center", "valign": "vcenter", "border": 1, "locked": True,
    })
    locked_num_fmt = workbook.add_format({
        "font_name": "Calibri", "font_size": 11, "align": "center",
        "valign": "vcenter", "border": 1, "locked": True, "num_format": "0",
    })
    unlocked_fmt = workbook.add_format({
        "font_name": "Calibri", "font_size": 11,
        "align": "center", "valign": "vcenter", "border": 1, "locked": False,
    })
    unlocked_arial_fmt = workbook.add_format({
        "font_name": "Arial", "font_size": 11,
        "align": "center", "valign": "vcenter", "border": 1, "locked": False,
    })

    for i, w in enumerate(COL_WIDTHS):
        ws.set_column(i, i, w)

    for ci, col_name in enumerate(COLUMNS_ORDER):
        ws.write(0, ci, col_name, header_fmt)

    # ── البيانات — استبدال iterrows() بـ to_dict('records') ─────────────────
    # to_dict('records') أسرع بـ 10-50x على البيانات الكبيرة
    for row_idx, row in enumerate(df.to_dict("records")):
        excel_row = row_idx + 1
        for ci, col_name in enumerate(COLUMNS_ORDER):
            val = row[col_name]
            val = "" if pd.isna(val) else val

            if ci < LOCKED_COLS_END:
                # الأعمدة A–H مقفلة
                if col_name in NUMERIC_COLS and val != "":
                    try:
                        val = int(str(val).replace(".0", ""))
                    except (ValueError, TypeError):
                        pass
                    ws.write(excel_row, ci, val, locked_num_fmt)
                else:
                    ws.write(excel_row, ci, str(val) if val != "" else "", locked_text_fmt)
            else:
                # الأعمدة I–M مفتوحة
                if ci == COL_PERIOD:        # L = الفترة → Arial
                    ws.write(excel_row, ci, val, unlocked_arial_fmt)
                else:
                    ws.write(excel_row, ci, val, unlocked_fmt)

    # ── صفوف فارغة إضافية (كلها مفتوحة) ────────────────────────────────────
    last_val_row = num_rows + EXTRA_BLANK_ROWS
    for extra in range(EXTRA_BLANK_ROWS):
        excel_row = num_rows + 1 + extra
        for ci in range(len(COLUMNS_ORDER)):
            fmt = unlocked_arial_fmt if ci == COL_PERIOD else unlocked_fmt
            ws.write(excel_row, ci, "", fmt)

    # ── القوائم المنسدلة ─────────────────────────────────────────────────────
    ws.data_validation(1, COL_STATUS, last_val_row, COL_STATUS,
                       {"validate": "list", "source": statuses, "show_input": True, "show_error": True})
    ws.data_validation(1, COL_DAY,    last_val_row, COL_DAY,
                       {"validate": "list", "source": days, "show_input": True, "show_error": True})
    ws.data_validation(1, COL_PERIOD, last_val_row, COL_PERIOD,
                       {"validate": "list", "source": periods, "show_input": True, "show_error": True})

    # ── حماية الورقة ─────────────────────────────────────────────────────────
    ws.protect("", {
        "sheet": True, "objects": True, "scenarios": True,
        "insert_rows": True, "insert_columns": False, "delete_rows": False,
        "sort": False, "autofilter": False,
        "select_locked_cells": True, "select_unlocked_cells": True,
    })

    workbook.close()
    output.seek(0)
    return output.read()


# ═══════════════════════════════════════════════════════════════════════════════
# ٧. معالجة الملفات — دالة موحّدة (كانت دالتان مكررتان)
# ═══════════════════════════════════════════════════════════════════════════════

def _load_to_bytes_cache(source) -> dict[str, bytes]:
    """
    يحوّل أي مصدر (list[UploadedFile] أو dict[str,bytes]) إلى dict موحّد.
    هذا هو جوهر دمج process_files + process_files_from_cache.
    """
    if isinstance(source, dict):
        return source
    cache: dict[str, bytes] = {}
    for uf in source:
        try:
            uf.seek(0)
            cache[uf.name] = uf.read()
        except Exception:
            pass
    return cache


def extract_teacher_names(source) -> tuple[dict, dict, list]:
    """
    يقرأ الملفات ويستخرج أسماء المعلمات ويبني القاموس المقترح.
    يعيد (preview_map, file_bytes_cache, errors).
    """
    errors: list[str]            = []
    raw_names_ordered: list[str] = []
    file_bytes_cache             = _load_to_bytes_cache(source)

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
            if raw and raw not in raw_names_ordered:
                raw_names_ordered.append(raw)
        except Exception as e:
            errors.append(f"❌ {fname}: {e}")

    preview_map = build_teacher_display_names("\n".join(raw_names_ordered))
    return preview_map, file_bytes_cache, errors


def process_files(
    source,
    days: list,
    periods: list,
    statuses: list,
    teacher_map: dict | None = None,
) -> tuple[dict[str, bytes], list[str]]:
    """
    معالجة الملفات وتحويلها إلى xlsx منسّق.

    source: يقبل إما:
        - list[UploadedFile]  (من واجهة Streamlit مباشرة)
        - dict[str, bytes]    (من session_state بعد تخزينها)
    كلاهما يُحوَّل لـ dict مبكراً ثم تُعالَج بنفس المنطق.
    (دمج process_files + process_files_from_cache الأصليتين)
    """
    results: dict[str, bytes] = {}
    errors:  list[str]        = []
    file_data: list           = []
    raw_names: list[str]      = []

    file_bytes_cache = _load_to_bytes_cache(source)

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
            errors.append(f"❌ {fname}: {e}")
            file_data.append((fname, None, None, ""))
            raw_names.append("")

    if not teacher_map:
        teacher_map = build_teacher_display_names("\n".join(n for n in raw_names if n))

    for fname, df, teacher_col, raw_name in file_data:
        if df is None:
            continue
        try:
            if raw_name:
                col_h_name = teacher_map.get(raw_name, get_first_name(raw_name))
                if teacher_col:
                    df[teacher_col] = col_h_name
                short = col_h_name
            else:
                short = fname.rsplit(".", 1)[0]

            xlsx_bytes = build_excel(df.copy(), days, periods, statuses)
            out_name   = f"{short}.xlsx"
            base       = out_name
            counter    = 1
            while out_name in results:
                out_name = base.replace(".xlsx", f"_{counter}.xlsx")
                counter += 1
            results[out_name] = xlsx_bytes
        except Exception as e:
            errors.append(f"❌ {fname}: {e}")

    return results, errors


# ═══════════════════════════════════════════════════════════════════════════════
# ٨. دوال الوقت والفترات
# ═══════════════════════════════════════════════════════════════════════════════

def parse_time_to_minutes(time_str) -> int | None:
    """تحويل نص الوقت (8:30 أو 8.30) إلى دقائق منذ منتصف الليل."""
    s     = str(time_str).strip().replace(".", ":").replace("٫", ":")
    parts = s.split(":")
    try:
        h = int(parts[0])
        m = int(parts[1]) if len(parts) > 1 else 0
        return h * 60 + m
    except Exception:
        return None


def format_time(time_str) -> str:
    """تنسيق الوقت: 8.30 / 8:30 / 830 → 8:30"""
    s = str(time_str).strip()
    if not s:
        return ""
    s = s.replace(".", ":").replace("٫", ":")
    if ":" not in s:
        s = (s[:-2] + ":" + s[-2:]) if len(s) > 2 else s + ":00"
    parts = s.split(":")
    try:
        h = int(parts[0])
        m = int(parts[1]) if len(parts) > 1 else 0
        return f"{h}:{m:02d}"
    except Exception:
        return str(time_str)


def excel_serial_to_time_str(val) -> str:
    """تحويل Excel time serial (0.375) أو نص وقت إلى HH:MM."""
    if val is None or str(val).strip() in ("", "nan"):
        return ""
    s = str(val).strip()
    try:
        f = float(s)
        if 0 < f < 1:
            total_min = round(f * 24 * 60)
            return f"{total_min // 60}:{total_min % 60:02d}"
    except ValueError:
        pass
    return format_time(s)


def fix_time_minutes(time_raw) -> str:
    """
    يُصحح الوقت — الدقائق المقبولة: 00, 15, 30, 45 فقط.
    رقم واحد (1-5) بعد النقطتين → ×10  (12:3 → 12:30)
    Excel serial                → يُحوَّل ويُقرَّب
    """
    if time_raw is None or str(time_raw).strip() in ("", "nan"):
        return ""
    s = str(time_raw).strip()
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
    s = s.replace(".", ":").replace("٫", ":")
    if ":" not in s:
        s = (s[:-2] + ":" + s[-2:]) if len(s) > 2 else s + ":00"
    parts = s.split(":")
    try:
        h     = int(parts[0])
        raw_m = parts[1].strip() if len(parts) > 1 else "0"
        m     = int(raw_m)
        if len(raw_m) == 1 and m <= 5:
            m = m * 10
        closest = min(VALID_MINUTES, key=lambda x: abs(x - m))
        return f"{h}:{closest:02d}"
    except Exception:
        return str(time_raw)


def get_period_from_time(minutes: int | None, period_schedule: list) -> str | None:
    if minutes is None:
        return None
    for name, start, end in period_schedule:
        if start <= minutes <= end:
            return name
    return None


# ═══════════════════════════════════════════════════════════════════════════════
# ٩. تحليل توزيع الأيام — استبدال iterrows بعمليات vectorized
# ═══════════════════════════════════════════════════════════════════════════════

def analyze_day_distribution(students_df, days_list, day_col, status_col) -> dict:
    """
    يحلل توزيع الطالبات على الأيام.
    تحسين: استبدال حلقة iterrows بعمليات pandas vectorized.
    """
    finished_mask = students_df[status_col].astype(str).str.strip() == "أنهت المقرر"
    finished_df   = students_df[finished_mask]
    total         = len(finished_df)

    if total == 0 or not days_list:
        return {"total": 0, "days": {}, "unassigned": 0, "has_issue": False, "ideal": 0}

    d           = len(days_list)
    base, extra = divmod(total, d)
    ideal       = {day: base + (1 if i < extra else 0) for i, day in enumerate(days_list)}

    # ── عملية vectorized بدلاً من iterrows ───────────────────────────────────
    day_series = finished_df[day_col].astype(str).str.strip()

    actual: dict[str, int] = {day: 0 for day in days_list}
    unassigned = 0
    for val in day_series:
        if val in ("", "nan"):
            unassigned += 1
            continue
        matched = False
        for day in days_list:
            if val in day or day in val:
                actual[day] += 1
                matched = True
                break
        if not matched:
            unassigned += 1

    days_report: dict = {}
    has_issue   = False
    for day in days_list:
        a = actual.get(day, 0)
        i = ideal.get(day, 0)
        if a > i:
            status    = f"🔴 ضغط — يجب تحويل {a - i} طالبة"
            has_issue = True
        elif a < i and unassigned > 0:
            status = f"🟢 متاح — يستوعب {i - a} طالبة إضافية"
        else:
            status = "✅ مناسب"
        days_report[day] = {"actual": a, "ideal": i, "status": status}

    return {
        "total": total, "days": days_report,
        "unassigned": unassigned, "has_issue": has_issue,
        "ideal_base": base, "ideal_extra": extra,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# ١٠. تقرير التوزيع — Excel
# ═══════════════════════════════════════════════════════════════════════════════

def build_distribution_report(day_reports: dict) -> bytes:
    output   = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})

    def fmt(bold=False, bg=None, align="center"):
        d = {"font_name": "Calibri", "font_size": 11,
             "align": align, "valign": "vcenter", "border": 1}
        if bold: d["bold"] = True
        if bg:   d["bg_color"] = bg
        return workbook.add_format(d)

    hdr_fmt   = fmt(bold=True, bg="#3d2060")
    hdr_fmt.set_font_color("white")
    ok_fmt    = fmt()
    red_fmt   = fmt(bg="#FF9999")
    grn_fmt   = fmt(bg="#C6EFCE")
    bold_fmt  = fmt(bold=True, align="right")
    title_fmt = workbook.add_format({
        "bold": True, "font_name": "Calibri", "font_size": 13,
        "align": "center", "valign": "vcenter", "bg_color": "#EDE8F5",
    })

    ws_sum = workbook.add_worksheet("ملخص")
    ws_sum.right_to_left()
    ws_sum.set_column(0, 0, 25)
    ws_sum.set_column(1, 1, 10)
    ws_sum.set_column(2, 2, 10)
    ws_sum.set_column(3, 3, 12)
    ws_sum.set_column(4, 4, 15)
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
        row_fmt    = red_fmt if has_issue else ok_fmt
        status_txt = "⚠️ يحتاج تدخل" if has_issue else "✅ موزّع بشكل جيد"

        ws_sum.write(r, 0, teacher,    row_fmt)
        ws_sum.write(r, 1, total,      row_fmt)
        ws_sum.write(r, 2, unassigned, red_fmt if unassigned else ok_fmt)
        ws_sum.write(r, 3, over_days,  red_fmt if over_days  else ok_fmt)
        ws_sum.write(r, 4, status_txt, row_fmt)

    for fname, report in day_reports.items():
        teacher  = fname.replace(".xlsx", "")
        ws       = workbook.add_worksheet(teacher[:31])
        ws.right_to_left()
        ws.set_column(0, 0, 22)
        ws.set_column(1, 1, 10)
        ws.set_column(2, 2, 10)
        ws.set_column(3, 3, 35)
        ws.merge_range(0, 0, 0, 3, f"تقرير توزيع الأيام — {teacher}", title_fmt)
        ws.set_row(0, 25)

        total      = report.get("total", 0)
        unassigned = report.get("unassigned", 0)
        base       = report.get("ideal_base", 0)
        xtra       = report.get("ideal_extra", 0)
        ideal_txt  = str(base) + (f" (+1 لأول {xtra} أيام)" if xtra else " لكل يوم")

        ws.write(1, 0, "إجمالي اللواتي أنهين المقرر:", bold_fmt)
        ws.write(1, 1, total, ok_fmt)
        ws.write(2, 0, "التوزيع المثالي:", bold_fmt)
        ws.write(2, 1, ideal_txt, ok_fmt)
        if unassigned:
            ws.write(3, 0, "⚠️ بدون يوم محدد:", bold_fmt)
            ws.write(3, 1, unassigned, red_fmt)

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


# ═══════════════════════════════════════════════════════════════════════════════
# ١١. معالجة ملف المرحلة الثانية
# ═══════════════════════════════════════════════════════════════════════════════

def process_stage2_file(
    file_bytes: bytes,
    days_list: list,
    statuses_list: list,
    periods_list: list,
    period_schedule: list | None = None,
) -> tuple:
    """
    معالجة ملف واحد مُعاد من المعلمة.
    تحسين: استبدال iterrows() الداخلية بـ enumerate(df.to_dict('records')).
    """
    if period_schedule is None:
        period_schedule = DEFAULT_PERIOD_SCHEDULE

    df = read_xlsx_raw(file_bytes)

    col_map = {
        "status":  next((c for c in df.columns if "الحالة"         in str(c)), None),
        "day":     next((c for c in df.columns if "يوم الاختبار"   in str(c)), None),
        "time":    next((c for c in df.columns if "توقيت الاختبار" in str(c)), None),
        "period":  next((c for c in df.columns if "الفترة"         in str(c)), None),
        "notes":   next((c for c in df.columns if "الملاحظات"      in str(c)), None),
        "teacher": next((c for c in df.columns if "المعلمة"        in str(c)), None),
    }

    teacher_col_val = ""
    if col_map["teacher"]:
        vals = df[col_map["teacher"]].dropna().astype(str).str.strip()
        vals = vals[vals != ""]
        if not vals.empty:
            teacher_col_val = vals.iloc[0]

    for col in COLUMNS_ORDER:
        if col not in df.columns:
            df[col] = ""

    day_report: dict = {}
    if col_map["status"] and col_map["day"] and days_list:
        day_report = analyze_day_distribution(df, days_list, col_map["day"], col_map["status"])

    # ── تصنيف الصفوف ─────────────────────────────────────────────────────────
    camera_rows, shurty_rows, note_rows, both_rows = [], [], [], []
    empty_status_rows, wrong_data_rows             = [], []
    time_format_errors: list   = []
    period_mismatch_rows: list = []
    time_fixes: dict[int, str] = {}  # نجمع التصحيحات ثم نطبقها دفعة واحدة

    STATUS_FINISHED = "أنهت المقرر"
    records         = df.to_dict("records")

    for idx, row in enumerate(records):
        status   = str(row.get(col_map["status"] or "الحالة",          "")).strip()
        day      = str(row.get(col_map["day"]    or "يوم الاختبار",    "")).strip()
        note     = str(row.get(col_map["notes"]  or "الملاحظات",       "")).strip()
        time_raw =     row.get(col_map["time"]   or "توقيت الاختبار",  "")
        period   = str(row.get(col_map["period"] or "الفترة",          "")).strip()

        try:
            fval = float(str(time_raw).strip())
            if fval >= 1:
                time_format_errors.append((idx, str(time_raw)))
        except (ValueError, TypeError):
            pass

        fixed_time = fix_time_minutes(time_raw)
        if col_map["time"] and fixed_time:
            time_fixes[idx] = fixed_time

        if period_schedule and fixed_time and period:
            try:
                h_str, _, m_str = fixed_time.partition(":")
                h = int(h_str)
                m = int(m_str) if m_str else 0
                ampm = PERIOD_AMPM.get(period, "AM")
                if ampm == "PM" and h < 12:
                    h += 12
                t_min          = h * 60 + m
                correct_period = get_period_from_time(t_min, period_schedule)
                if correct_period and correct_period != period:
                    period_mismatch_rows.append((idx, period, correct_period))
            except Exception:
                pass

        name_val = str(row.get("الاسم", "")).strip()
        if not name_val or name_val == "nan":
            continue

        if not status or status == "nan":
            empty_status_rows.append(idx)
            continue

        if status == STATUS_FINISHED:
            if not day or day == "nan":
                wrong_data_rows.append(idx)
        else:
            if ((day and day != "nan") or
                    (fixed_time and str(fixed_time).strip() != "") or
                    (period and period != "nan")):
                wrong_data_rows.append(idx)

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

    # ── تطبيق تصحيحات الوقت دفعة واحدة ──────────────────────────────────────
    if col_map["time"] and time_fixes:
        for idx, fixed in time_fixes.items():
            df.at[idx, col_map["time"]] = fixed

    # ── بناء Excel ───────────────────────────────────────────────────────────
    output   = io.BytesIO()
    df_out   = df[COLUMNS_ORDER].copy()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    ws       = workbook.add_worksheet("الطالبات")
    ws.right_to_left()

    def fmt(extra=None):
        base = {"font_name": "Calibri", "font_size": 11,
                "align": "center", "valign": "vcenter", "border": 1, "locked": False}
        if extra:
            base.update(extra)
        return workbook.add_format(base)

    header_fmt    = workbook.add_format({
        "bold": True, "font_name": "Calibri", "font_size": 11,
        "align": "center", "valign": "vcenter",
        "border": 1, "bg_color": COLOR_HEADER, "locked": False,
    })
    normal_fmt    = fmt()
    num_fmt       = fmt({"num_format": "0"})
    phone_fmt     = fmt({"num_format": "0"})
    time_fmt      = fmt({"num_format": "h:mm"})
    arial_fmt     = fmt({"font_name": "Arial"})
    cam_fmt       = fmt({"bg_color": COLOR_RED})
    cam_num       = fmt({"bg_color": COLOR_RED, "num_format": "0"})
    cam_phone     = fmt({"bg_color": COLOR_RED, "num_format": "0"})
    cam_time      = fmt({"bg_color": COLOR_RED, "num_format": "h:mm"})
    yellow_cell   = fmt({"bg_color": COLOR_YELLOW})
    red_cell      = fmt({"bg_color": COLOR_RED})
    warn_fmt      = fmt({"bg_color": COLOR_YELLOW})
    warn_phone    = fmt({"bg_color": COLOR_YELLOW, "num_format": "0"})
    warn_time     = fmt({"bg_color": COLOR_YELLOW, "num_format": "h:mm"})

    for i, w in enumerate(COL_WIDTHS):
        ws.set_column(i, i, w)
    for ci, cn in enumerate(COLUMNS_ORDER):
        ws.write(0, ci, cn, header_fmt)

    mismatch_set = {i for i, *_ in period_mismatch_rows}

    def write_cell(er, ci, cn, val, f, phone_f=None, time_f=None):
        """كتابة خلية مع مراعاة نوع البيانات."""
        if cn == "رقم الواتس اب" and val != "":
            use_f = phone_f or f
            try:
                ws.write_number(er, ci, int(float(str(val).replace(".0", ""))), use_f)
            except Exception:
                ws.write(er, ci, str(val), use_f)
        elif cn == "توقيت الاختبار" and val != "":
            use_f = time_f or f
            try:
                fval = float(str(val).replace(".0", "")) if ":" not in str(val) else None
                if fval is not None and 0 < fval < 1:
                    ws.write_number(er, ci, fval, use_f)
                else:
                    ws.write_string(er, ci, str(val), use_f)
            except Exception:
                ws.write_string(er, ci, str(val), use_f)
        elif cn in {"الرقم", "المواليد"} and val != "":
            try:
                ws.write_number(er, ci, int(str(val).replace(".0", "")), f)
            except Exception:
                ws.write(er, ci, str(val) if isinstance(val, str) else val, f)
        else:
            ws.write(er, ci, str(val) if isinstance(val, str) else val, f)

    def normal_f_for(cn, val):
        """إرجاع التنسيق المناسب للخلية في الحالة العادية."""
        if cn == "رقم الواتس اب" and val != "":
            return phone_fmt
        if cn == "توقيت الاختبار" and val != "":
            return time_fmt
        if cn in {"الرقم", "المواليد"} and val != "":
            return num_fmt
        if cn == "الفترة":
            return arial_fmt
        return normal_fmt

    # ── استبدال iterrows بـ enumerate(to_dict('records')) ────────────────────
    for row_idx, row in enumerate(df_out.to_dict("records")):
        er        = row_idx + 1
        is_camera = row_idx in camera_rows
        is_shurty = row_idx in shurty_rows
        is_note   = row_idx in note_rows
        is_both   = row_idx in both_rows
        is_warn   = row_idx in empty_status_rows or row_idx in wrong_data_rows
        is_mis    = row_idx in mismatch_set

        for ci, cn in enumerate(COLUMNS_ORDER):
            val = row[cn]
            val = "" if pd.isna(val) else val

            if is_camera:
                write_cell(er, ci, cn, val, cam_fmt, phone_f=cam_phone, time_f=cam_time)
            elif is_both:
                if cn == "الحالة":
                    write_cell(er, ci, cn, val, yellow_cell)
                elif cn == "الاسم":
                    write_cell(er, ci, cn, val, red_cell)
                else:
                    write_cell(er, ci, cn, val, normal_f_for(cn, val), phone_f=phone_fmt, time_f=time_fmt)
            elif is_shurty:
                write_cell(er, ci, cn, val, yellow_cell if cn == "الحالة" else normal_f_for(cn, val))
            elif is_note:
                write_cell(er, ci, cn, val, red_cell if cn == "الاسم" else normal_f_for(cn, val))
            elif is_mis:
                write_cell(er, ci, cn, val, yellow_cell if cn == "الفترة" else normal_f_for(cn, val))
            elif is_warn:
                write_cell(er, ci, cn, val, warn_fmt, phone_f=warn_phone, time_f=warn_time)
            else:
                write_cell(er, ci, cn, val, normal_f_for(cn, val))

    last_dv_row = len(df_out) + EXTRA_BLANK_ROWS
    if days_list:
        ws.data_validation(1, COL_DAY,    last_dv_row, COL_DAY,
                           {"validate": "list", "source": days_list, "show_input": True, "show_error": True})
    if statuses_list:
        ws.data_validation(1, COL_STATUS, last_dv_row, COL_STATUS,
                           {"validate": "list", "source": statuses_list, "show_input": True, "show_error": True})
    if periods_list:
        ws.data_validation(1, COL_PERIOD, last_dv_row, COL_PERIOD,
                           {"validate": "list", "source": periods_list, "show_input": True, "show_error": True})

    workbook.close()
    output.seek(0)
    n_colored = len(camera_rows) + len(shurty_rows) + len(note_rows) + len(both_rows)
    return (
        output.read(), n_colored, 0,
        len(empty_status_rows) + len(wrong_data_rows),
        day_report, time_format_errors, period_mismatch_rows, teacher_col_val,
    )


# ═══════════════════════════════════════════════════════════════════════════════
# ١٢. المرحلة الثالثة — بناء ملف اللجان
# ═══════════════════════════════════════════════════════════════════════════════

def day_sort_key(day_val, days_list: list) -> int:
    day_str = str(day_val).strip()
    for i, d in enumerate(days_list):
        if day_str == d or day_str in d or d in day_str:
            return i
    for i, d in enumerate(DAYS_ORDER):
        if day_str == d or day_str in d or d in day_str:
            return i + 100
    return 999


def build_stage3_file(
    files_dict: dict[str, bytes],
    days_list: list,
    existing_bytes: bytes | None = None,
) -> tuple[bytes, int, int, int]:
    all_rows: list[pd.DataFrame] = []
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

    # ── دمج الملف الأم إن وُجد ───────────────────────────────────────────────
    if existing_bytes:
        try:
            # الآن نستخدم read_xlsx_raw مع multi_sheet=True بدلاً من read_existing_stage3
            existing_sheets = read_xlsx_raw(existing_bytes, multi_sheet=True)
            existing_rows   = []
            for df_ex in existing_sheets.values():
                if df_ex.empty:
                    continue
                df_ex.columns = [str(c).strip() for c in df_ex.columns]
                df_ex = df_ex.dropna(how="all")
                if "الاسم" in df_ex.columns:
                    df_ex = df_ex[df_ex["الاسم"].astype(str).str.strip().replace("nan", "") != ""]
                if not df_ex.empty:
                    existing_rows.append(df_ex)
            if existing_rows:
                combined = pd.concat(
                    [pd.concat(existing_rows, ignore_index=True), combined],
                    ignore_index=True,
                )
        except Exception:
            pass

    cols = COLUMNS_ORDER
    for c in cols:
        if c not in combined.columns:
            combined[c] = ""
    combined = combined[cols].copy()

    for c in cols:
        combined[c] = (combined[c]
                       .fillna("").astype(str).str.strip()
                       .str.replace("\u00a0", "", regex=False)
                       .replace("nan", ""))

    combined = combined[combined["الاسم"] != ""].reset_index(drop=True)

    mask_early    = combined["الملاحظات"].str.contains("قدمت الاختبار", na=False)
    mask_finished = (combined["الحالة"] == "أنهت المقرر") & (~mask_early)
    mask_others   = (~mask_finished) & (~mask_early)

    df_finished = combined[mask_finished].copy()
    df_others   = combined[mask_others].copy()
    df_early    = combined[mask_early].copy()

    df_finished["_day"]  = df_finished["يوم الاختبار"].apply(lambda x: day_sort_key(x, days_list))
    df_finished["_time"] = pd.to_numeric(df_finished["توقيت الاختبار"], errors="coerce").fillna(999)
    df_finished = df_finished.sort_values(["المعلمة", "_day", "_time"]).drop(columns=["_day", "_time"]).reset_index(drop=True)
    df_others   = df_others.sort_values(["المعلمة", "الاسم"]).reset_index(drop=True)
    df_early    = df_early.sort_values(["المعلمة", "الاسم"]).reset_index(drop=True)

    output   = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})

    header_fmt = workbook.add_format({"bold": True, "font_name": "Calibri", "font_size": 11,
                                      "align": "center", "valign": "vcenter", "border": 1})
    cell_fmt   = workbook.add_format({"font_name": "Calibri", "font_size": 11,
                                      "align": "center", "valign": "vcenter", "border": 1})
    num_fmt    = workbook.add_format({"font_name": "Calibri", "font_size": 11,
                                      "align": "center", "valign": "vcenter", "border": 1, "num_format": "0"})
    time_fmt   = workbook.add_format({"font_name": "Calibri", "font_size": 11,
                                      "align": "center", "valign": "vcenter", "border": 1, "num_format": "h:mm"})
    arial_fmt  = workbook.add_format({"font_name": "Arial", "font_size": 11,
                                      "align": "center", "valign": "vcenter", "border": 1})

    col_widths_s3 = {
        "الرقم": 7, "الاسم": 24, "رقم الواتس اب": 14, "المجموعة": 13,
        "البلد": 7, "المواليد": 6, "الإجازة": 5.3, "المعلمة": 7,
        "الحالة": 20, "يوم الاختبار": 11, "توقيت الاختبار": 11,
        "الفترة": 14, "الملاحظات": 40,
    }

    def write_sheet(name: str, df_sheet: pd.DataFrame) -> None:
        ws = workbook.add_worksheet(name)
        ws.right_to_left()
        for ci, cn in enumerate(cols):
            ws.set_column(ci, ci, col_widths_s3.get(cn, 12))
            ws.write(0, ci, cn, header_fmt)

        # استبدال iterrows بـ to_dict('records')
        for ri, row in enumerate(df_sheet.to_dict("records")):
            er = ri + 1
            for ci, cn in enumerate(cols):
                val = row[cn]
                if cn in NUMERIC_COLS and val not in ("", "nan"):
                    try:
                        ws.write_number(er, ci, int(str(val).replace(".0", "")), num_fmt)
                    except Exception:
                        ws.write(er, ci, val, cell_fmt)
                elif cn == "توقيت الاختبار" and val not in ("", "nan"):
                    try:
                        fval = float(val)
                        ws.write_number(er, ci, fval, time_fmt) if 0 < fval < 1 else ws.write(er, ci, val, cell_fmt)
                    except Exception:
                        ws.write(er, ci, val, cell_fmt)
                elif cn == "الفترة":
                    ws.write(er, ci, val, arial_fmt)
                else:
                    ws.write(er, ci, val, cell_fmt)

    write_sheet("المتقدمات للاختبار", df_finished)
    write_sheet("غير متقدمات",        df_others)
    write_sheet("اختبار مبكر",        df_early)

    workbook.close()
    output.seek(0)
    return output.read(), len(df_finished), len(df_others), len(df_early)


# ═══════════════════════════════════════════════════════════════════════════════
# ١٣. مساعد HTML — html.escape لحماية من XSS
# ═══════════════════════════════════════════════════════════════════════════════

def file_chips_html(files) -> str:
    """
    يبني HTML لعرض أسماء الملفات كشرائح (chips).
    يستخدم html.escape() لتجنب XSS عند دمج أسماء المستخدمين في HTML.
    """
    chips = " ".join(
        f'<span class="file-chip">📄 {html_lib.escape(f.name)}</span>'
        for f in files
    )
    return f"<div style='margin:0.5rem 0 1rem'>{chips}</div>"


def stat_card(number, label, color: str = "#3d2060") -> str:
    return (
        f'<div class="stat-card">'
        f'<div class="number" style="color:{color};">{number}</div>'
        f'<div class="label">{html_lib.escape(str(label))}</div>'
        f'</div>'
    )


# ═══════════════════════════════════════════════════════════════════════════════
# ١٤. الـ Sidebar
# ═══════════════════════════════════════════════════════════════════════════════

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
    periods_schedule_text = st.text_area(
        "🕐 أوقات الفترات (للمطابقة)",
        value="فجراً: 4:00-8:45\nضحى: 9:00-11:45\nظهراً: 12:00-15:45\nعصراً: 16:00-18:45\nليلاً: 19:00-21:30",
        height=145,
        help="النسق: اسم الفترة: HH:MM-HH:MM",
    )

    days_list     = parse_list(days_text)
    periods_list  = parse_list(periods_text)
    statuses_list = parse_list(statuses_text)

    # parse_period_schedule الآن دالة على المستوى الأعلى — لا تُعاد تعريفها هنا
    # مصدر واحد للحقيقة: period_schedule يُستخدم في كل المراحل
    period_schedule = parse_period_schedule(periods_schedule_text)

    st.markdown(
        f"<div style='margin-top:1rem; padding:0.8rem; background:rgba(255,255,255,0.08);"
        f"border-radius:8px; font-size:0.82rem; color:#c4a0e8;'>"
        f"✅ {len(days_list)} أيام &nbsp;|&nbsp; ✅ "
        f"{len(periods_list)} فترات &nbsp;|&nbsp; ✅ "
        f"{len(statuses_list)} حالة</div>",
        unsafe_allow_html=True,
    )


# ═══════════════════════════════════════════════════════════════════════════════
# ١٥. المرحلة الأولى — الواجهة
# ═══════════════════════════════════════════════════════════════════════════════

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
        st.markdown(stat_card(len(uploaded_files), "ملف مرفوع"), unsafe_allow_html=True)
    with cols[1]:
        st.markdown(stat_card(len(days_list), "أيام الاختبار"), unsafe_allow_html=True)
    with cols[2]:
        st.markdown(stat_card(len(periods_list), "فترة متاحة"), unsafe_allow_html=True)
    with cols[3]:
        st.markdown(stat_card(len(statuses_list), "حالة في القائمة"), unsafe_allow_html=True)

    st.markdown(file_chips_html(uploaded_files), unsafe_allow_html=True)

    # ── خطوة ١: تحليل ومعاينة أسماء المعلمات ────────────────────────────────
    if st.button("🔍 تحليل الملفات ومعاينة أسماء المعلمات", use_container_width=True):
        with st.spinner("جارٍ القراءة..."):
            preview_map, file_bytes_cache, preview_errors = extract_teacher_names(uploaded_files)
        st.session_state[SK.PREVIEW_MAP]    = preview_map
        st.session_state[SK.FILE_CACHE]     = file_bytes_cache
        st.session_state[SK.PREVIEW_ERRORS] = preview_errors
        st.session_state[SK.STAGE1_RESULTS] = None
        st.session_state[SK.STAGE1_ERRORS]  = None

    # ── عرض جدول المعاينة + حقول التعديل ────────────────────────────────────
    if st.session_state.get(SK.PREVIEW_MAP):
        preview_map = st.session_state[SK.PREVIEW_MAP]

        st.markdown('<div class="section-title">👁️ مراجعة أسماء المعلمات</div>', unsafe_allow_html=True)
        st.markdown(
            "<div style='font-size:0.88rem;color:#555;margin-bottom:0.8rem;direction:rtl;'>"
            "الكود اقترح هذه الأسماء تلقائياً — عدّلي أي اسم لا يناسبكِ ثم اضغطي تأكيد."
            "</div>",
            unsafe_allow_html=True,
        )

        edited_map: dict[str, str] = {}
        cols_h = st.columns([3, 2, 2])
        cols_h[0].markdown("**الاسم الكامل في الملف**")
        cols_h[1].markdown("**الاسم المقترح**")
        cols_h[2].markdown("**الاسم النهائي** (عدّلي هنا)")

        for i, (full_name, suggested) in enumerate(preview_map.items()):
            c1, c2, c3 = st.columns([3, 2, 2])
            c1.markdown(
                f"<div style='direction:rtl;padding-top:8px;font-size:0.9rem;'>"
                f"{html_lib.escape(full_name)}</div>",
                unsafe_allow_html=True,
            )
            c2.markdown(
                f"<div style='direction:rtl;padding-top:8px;font-size:0.9rem;"
                f"font-weight:600;color:#6b3fa0;'>{html_lib.escape(suggested)}</div>",
                unsafe_allow_html=True,
            )
            final = c3.text_input(
                f"final_{i}",
                value=suggested,
                label_visibility="collapsed",
                key=f"teacher_edit_{i}",
            )
            edited_map[full_name] = final.strip() if final.strip() else suggested

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button("⚡ تأكيد ومعالجة الملفات", type="primary", use_container_width=True):
            file_bytes_cache = st.session_state.get(SK.FILE_CACHE, {})
            with st.spinner("جارٍ المعالجة..."):
                results, errors = process_files(
                    file_bytes_cache, days_list, periods_list, statuses_list, edited_map
                )
            st.session_state[SK.STAGE1_RESULTS] = results
            st.session_state[SK.STAGE1_ERRORS]  = errors

    # ── عرض النتائج ──────────────────────────────────────────────────────────
    if st.session_state.get(SK.STAGE1_RESULTS):
        results = st.session_state[SK.STAGE1_RESULTS]
        errors  = st.session_state.get(SK.STAGE1_ERRORS, [])

        for e in errors:
            st.error(e)

        if results:
            st.markdown(
                f'<div class="success-banner">✅ تمت معالجة {len(results)} ملف بنجاح!</div>',
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
                label=f"⬇️ تحميل جميع الملفات ({len(results)} ملف) — ZIP",
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
# ١٦. المرحلة الثانية — واجهة مراجعة ملفات المعلمات
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    """
    <div style="background:linear-gradient(135deg,#1a3a5c,#2874a6);border-radius:16px;
    padding:1.8rem 2.5rem;margin-bottom:1.5rem;box-shadow:0 8px 32px rgba(26,58,92,0.3);
    text-align:center;color:white;">
        <div style="font-size:1.8rem;font-weight:900;margin:0;">🔄 المرحلة الثانية — مراجعة ملفات المعلمات</div>
        <div style="font-size:0.95rem;margin:0.4rem 0 0;opacity:0.88;">
            ارفعي الملفات المُعادة من المعلمات لمعالجتها وتدقيقها
        </div>
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
    st.markdown(file_chips_html(uploaded_stage2), unsafe_allow_html=True)

    st.markdown(
        """
        <div style="background:white;border-radius:10px;padding:0.8rem 1.2rem;
        margin-bottom:1rem;border:1px solid #e0d0f8;font-size:0.88rem;direction:rtl;">
            <b>دليل الألوان:</b> &nbsp;
            <span style="background:#FF9999;padding:2px 10px;border-radius:4px;">🔴 كاميرا — صف كامل</span> &nbsp;
            <span style="background:#FF9999;padding:2px 10px;border-radius:4px;">🔴 ملاحظة جوهرية — خلية الاسم</span> &nbsp;
            <span style="background:#FFFF99;padding:2px 10px;border-radius:4px;">🟡 شرطي — خلية الحالة</span> &nbsp;
            <span style="background:#FFFF99;padding:2px 10px;border-radius:4px;">🟡 بيانات خاطئة — صف كامل</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.button("🔄 معالجة ومراجعة الملفات", type="primary", use_container_width=True, key="stage2_btn"):
        with st.spinner("جارٍ التدقيق والمعالجة..."):
            stage2_results:    dict = {}
            stage2_errors:     list = []
            day_reports:       dict = {}
            time_fmt_warnings: dict = {}
            period_warnings:   dict = {}
            total_red = total_issues = 0

            for uf in uploaded_stage2:
                try:
                    fb = uf.read()
                    (out_bytes, n_colored, _, n_issues,
                     d_report, t_errors, p_mismatches, teacher_name) = process_stage2_file(
                        fb, days_list, statuses_list, periods_list,
                        period_schedule or DEFAULT_PERIOD_SCHEDULE,
                    )
                    out_name = f"{teacher_name}.xlsx" if teacher_name else uf.name
                    stage2_results[out_name] = out_bytes
                    total_red    += n_colored
                    total_issues += n_issues
                    if d_report.get("has_issue") or d_report.get("unassigned", 0) > 0:
                        day_reports[uf.name] = d_report
                    if t_errors:
                        time_fmt_warnings[uf.name] = t_errors
                    if p_mismatches:
                        period_warnings[uf.name] = p_mismatches
                except Exception as e:
                    stage2_errors.append(f"❌ {html_lib.escape(uf.name)}: {e}")

        for e in stage2_errors:
            st.error(e)

        if time_fmt_warnings:
            st.markdown('<div class="section-title">⚠️ تنبيهات تنسيق التوقيت</div>', unsafe_allow_html=True)
            for fname, errs in time_fmt_warnings.items():
                rows_str = "، ".join(f"صف {idx+2} (قيمة: {raw})" for idx, raw in errs)
                st.warning(f"📄 **{html_lib.escape(fname)}** — خلايا التوقيت مُنسَّقة كتاريخ: {rows_str}")
            st.info("💡 الحل: افتحي الملف الأصلي، حددي عمود التوقيت، وغيّري تنسيق الخلايا إلى 'وقت' (hh:mm)")

        if period_warnings:
            st.markdown('<div class="section-title">🕐 تعارض الفترة مع الوقت</div>', unsafe_allow_html=True)
            for fname, mismatches in period_warnings.items():
                with st.expander(f"📄 {html_lib.escape(fname)} — {len(mismatches)} تعارض"):
                    for idx, actual, correct in mismatches:
                        st.markdown(
                            f"&nbsp; صف **{idx+2}**: الفترة المكتوبة **{html_lib.escape(actual)}** "
                            f"← الصحيحة للوقت هي **{html_lib.escape(correct)}**",
                            unsafe_allow_html=True,
                        )

        if stage2_results:
            cols2 = st.columns(4)
            with cols2[0]:
                st.markdown(stat_card(len(stage2_results), "ملف معالج"), unsafe_allow_html=True)
            with cols2[1]:
                st.markdown(stat_card(total_red, "صفوف مُلوَّنة 🎨", "#c0392b"), unsafe_allow_html=True)
            with cols2[2]:
                st.markdown(stat_card(total_issues, "يحتاج مراجعة 🟡", "#b7950b"), unsafe_allow_html=True)
            with cols2[3]:
                crowded = sum(len(r["days"]) for r in day_reports.values())
                st.markdown(stat_card(crowded, "أيام مكتظة 📊", "#555"), unsafe_allow_html=True)

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


# ═══════════════════════════════════════════════════════════════════════════════
# ١٧. المرحلة الثالثة — تجميع اللجان
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    """
    <div style="background:linear-gradient(135deg,#1a4e1a,#2e7d32);border-radius:16px;
    padding:1.8rem 2.5rem;margin-bottom:1.5rem;box-shadow:0 8px 32px rgba(26,78,26,0.3);
    text-align:center;color:white;">
        <div style="font-size:1.8rem;font-weight:900;margin:0;">📊 المرحلة الثالثة — تجميع اللجان</div>
        <div style="font-size:0.95rem;margin:0.4rem 0 0;opacity:0.88;">
            ارفعي ملفات المعلمات المُراجعة لتجميعها في ملف لجان واحد
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
    st.markdown(file_chips_html(uploaded_stage3), unsafe_allow_html=True)

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
            "</div>",
            unsafe_allow_html=True,
        )
        existing_file = st.file_uploader(
            "ملف اللجان الأم",
            type=["xlsx"],
            key="stage3_existing",
            label_visibility="collapsed",
        )
        if existing_file:
            st.success(f"✅ تم رفع الملف الأم: {html_lib.escape(existing_file.name)}")

    output_name = st.text_input(
        "اسم الملف الناتج",
        value="اللجان_المجمعة.xlsx",
        key="stage3_output_name",
    )

    btn_label = "📊 إضافة للملف الأم وتحميل" if "نعم" in has_existing else "📊 تجميع وبناء اللجان"

    if st.button(btn_label, type="primary", use_container_width=True, key="stage3_btn"):
        with st.spinner("جارٍ التجميع..."):
            files_dict:  dict  = {}
            read_errors: list  = []
            for uf in uploaded_stage3:
                try:
                    uf.seek(0)
                    files_dict[uf.name] = uf.read()
                except Exception as e:
                    read_errors.append(f"❌ {html_lib.escape(uf.name)}: {e}")

            for e in read_errors:
                st.error(e)

            if files_dict:
                try:
                    existing_bytes = None
                    if "نعم" in has_existing and existing_file:
                        existing_file.seek(0)
                        existing_bytes = existing_file.read()

                    result_bytes, n_fin, n_oth, n_ear = build_stage3_file(
                        files_dict, days_list, existing_bytes=existing_bytes
                    )

                    mode_label = "إضافة للملف الأم" if existing_bytes else "ملف جديد"
                    st.success(f"✅ {mode_label} — تم بنجاح")

                    cols3 = st.columns(4)
                    with cols3[0]:
                        st.markdown(stat_card(len(files_dict), "ملف مُضاف"), unsafe_allow_html=True)
                    with cols3[1]:
                        st.markdown(stat_card(n_fin, "متقدمة ✅", "#1a4e1a"), unsafe_allow_html=True)
                    with cols3[2]:
                        st.markdown(stat_card(n_oth, "غير متقدمة 🕐", "#b7950b"), unsafe_allow_html=True)
                    with cols3[3]:
                        st.markdown(stat_card(n_ear, "اختبار مبكر 🎓", "#1a5276"), unsafe_allow_html=True)

                    fname_out = output_name if output_name.endswith(".xlsx") else f"{output_name}.xlsx"
                    st.download_button(
                        label="⬇️ تحميل ملف اللجان المجمعة",
                        data=result_bytes,
                        file_name=fname_out,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="stage3_download",
                    )
                except Exception as e:
                    st.error(f"❌ خطأ في التجميع: {e}")


# ═══════════════════════════════════════════════════════════════════════════════
# ١٨. Footer
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown(
    """
    <hr style="margin:2rem 0 1rem; border-color:#d8c8f0;">
    <div style="text-align:center; color:#999; font-size:0.8rem; font-family:'Tajawal',sans-serif;">
        أداة مقرأة — مبنية بـ Python &amp; Streamlit &nbsp;|&nbsp; 📖
    </div>
    """,
    unsafe_allow_html=True,
)
