import io
import re
from datetime import datetime
from typing import Any, Dict, List, Optional

from PIL import Image

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor, Mm
from uuid import uuid4


# =========================
# Small, stable helpers
# =========================
def get_bool(row, base_key):
    return row.get(base_key, row.get(f"{base_key}_label"))

def clear_paragraph(p) -> None:
    p.text = ""
    for r in p.runs:
        r.text = ""

def add_paragraphs(doc: Document, text: Any, align=WD_ALIGN_PARAGRAPH.JUSTIFY) -> None:
    """
    Convert text with blank lines into multiple real Word paragraphs.
    Prevents huge gaps caused by JUSTIFY + line breaks inside one paragraph.
    """
    t = s(text)
    if not t:
        return

    # Normalize line endings
    t = t.replace("\r\n", "\n").replace("\r", "\n")

    # Split into paragraphs by blank lines
    parts = [p.strip() for p in re.split(r"\n\s*\n+", t) if p.strip()]

    for part in parts:
        # Inside a paragraph, remove single line breaks and extra spaces
        part = re.sub(r"\s*\n\s*", " ", part)  # turn internal newlines to spaces
        part = re.sub(r"[ \t]{2,}", " ", part)  # collapse multiple spaces

        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(8)  # ✅ فاصله استاندارد بین پاراگراف‌ها
        p.paragraph_format.line_spacing = 1.15  # ✅ خوانا و مشابه ریپورت

        r = p.add_run(part)
        set_run(r, "Times New Roman", 12, False)

CHECKED = "☒"
UNCHECKED = "☐"


def s(value: Any) -> str:
    return "" if value is None else str(value).strip()


# =========================
# Page setup (A4)
# =========================
def set_doc_a4(doc: Document) -> None:
    """Force A4 page size for ALL sections."""
    for sec in doc.sections:
        sec.page_width = Mm(210)
        sec.page_height = Mm(297)


def _set_row_height_exact(row, height_twips: int) -> None:
    """
    height_twips:
      1 point = 20 twips
      مثال:
        260 = ~13pt
        280 = ~14pt
    """
    tr = row._tr
    trPr = tr.get_or_add_trPr()

    trHeight = OxmlElement("w:trHeight")
    trHeight.set(qn("w:val"), str(int(height_twips)))
    trHeight.set(qn("w:hRule"), "exact")
    trPr.append(trHeight)


def get_usable_width(doc: Document):
    sec = doc.sections[0]
    return sec.page_width - sec.left_margin - sec.right_margin


# ✅ EMU -> TWIPS
def _emu_to_twips(emu: int) -> int:
    # 1 inch = 914400 EMU = 1440 twips => twips = emu / 635
    return int(round(int(emu) / 635.0))


def _set_table_fixed_layout(table) -> None:
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblLayout = tblPr.find(qn("w:tblLayout"))
    if tblLayout is None:
        tblLayout = OxmlElement("w:tblLayout")
        tblPr.append(tblLayout)
    tblLayout.set(qn("w:type"), "fixed")


def _set_table_width_and_indent(table, width_emu: int) -> None:
    """
    width_emu: table width in EMU (int)
    Sets:
      - tblW in twips
      - indent = 0   (prevents right overflow and border mismatch)
    """
    width_emu = int(width_emu)

    tbl = table._tbl
    tblPr = tbl.tblPr

    # tblW
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(_emu_to_twips(width_emu)))

    # tblInd = 0
    tblInd = tblPr.find(qn("w:tblInd"))
    if tblInd is None:
        tblInd = OxmlElement("w:tblInd")
        tblPr.append(tblInd)
    tblInd.set(qn("w:type"), "dxa")
    tblInd.set(qn("w:w"), "0")

def _set_table_inside_borders_only(table, color_hex: str = "A6A6A6", size: str = "8") -> None:
    """
    Keep only insideH/insideV borders.
    Remove outside borders (top/left/bottom/right).
    """
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)

    def set_edge(edge: str, val: str, sz: str = "8", color: str = "A6A6A6"):
        el = borders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            borders.append(el)
        el.set(qn("w:val"), val)
        if val != "nil":
            el.set(qn("w:sz"), sz)
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), color)

    # outside OFF
    for edge in ("top", "left", "bottom", "right"):
        set_edge(edge, "nil")

    # inside ON
    set_edge("insideH", "single", size, color_hex)
    set_edge("insideV", "single", size, color_hex)


def _remove_all_table_borders(table) -> None:
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = tblPr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tblPr.append(borders)

    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            borders.append(el)
        el.set(qn("w:val"), "nil")


def _remove_all_cell_borders(cell) -> None:
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)

    for edge in ("top", "left", "bottom", "right"):
        el = tcBorders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            tcBorders.append(el)
        el.set(qn("w:val"), "nil")


def _set_cell_margins_zero(cell) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = tcPr.find(qn("w:tcMar"))
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)

    for side in ("top", "left", "bottom", "right"):
        node = tcMar.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            tcMar.append(node)
        node.set(qn("w:w"), "0")
        node.set(qn("w:type"), "dxa")


# =========================
# ✅ NEW: Inner table borders connect-to-outer + inside ON
# =========================
def _set_table_borders_connect_to_outer(table, color_hex="A6A6A6", size="8"):
    """
    Inner table borders:
      - left/right ON  -> connect visually to outer table borders
      - top/bottom OFF
      - insideH/insideV ON
    """
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)

    def edge(name, val):
        el = borders.find(qn(f"w:{name}"))
        if el is None:
            el = OxmlElement(f"w:{name}")
            borders.append(el)
        el.set(qn("w:val"), val)
        if val != "nil":
            el.set(qn("w:sz"), size)
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), color_hex)

    # ❌ no top/bottom border
    edge("top", "nil")
    edge("bottom", "nil")

    # ✅ connect to outer borders
    edge("left", "single")
    edge("right", "single")

    # ✅ inside borders
    edge("insideH", "single")
    edge("insideV", "single")


def set_run(run, name: str, size: int, bold: bool = False, color: Optional[RGBColor] = None) -> None:
    run.font.name = name
    try:
        run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    except Exception:
        pass
    run.font.size = Pt(size)
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color


def set_style(doc: Document, style_name: str, name: str, size: int, bold: bool = False) -> None:
    if style_name not in doc.styles:
        return
    st = doc.styles[style_name]
    st.font.name = name
    st.font.size = Pt(size)
    st.font.bold = bold
    try:
        st.element.rPr.rFonts.set(qn("w:eastAsia"), name)
    except Exception:
        pass


def apply_heading_rules(doc: Document) -> None:
    set_style(doc, "Heading 1", "Calibri (Headings)", 18, True)
    set_style(doc, "Heading 2", "Calibri (Headings)", 16, True)
    set_style(doc, "Heading 3", "Calibri (Headings)", 14, True)
    set_style(doc, "Normal", "Times New Roman", 12, False)


def shade(cell, fill_hex: str = "D9E2F3") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tc_pr.append(shd)


def set_table_borders(table, color_hex: str = "A6A6A6") -> None:
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)

    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            borders.append(el)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "8")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color_hex)


def underline_heading(paragraph, color_hex: str = "F4A261") -> None:
    ppr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), color_hex)
    pbdr.append(bottom)
    ppr.append(pbdr)


def h1(doc: Document, text: str, underline: bool = True) -> None:
    p = doc.add_paragraph(text, style="Heading 1")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if underline:
        underline_heading(p)


def h2(doc: Document, text: str) -> None:
    p = doc.add_paragraph(text, style="Heading 2")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def h3(doc: Document, text: str) -> None:
    p = doc.add_paragraph(text, style="Heading 3")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def body(doc: Document, text: Any) -> None:
    t = s(text)
    if not t:
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(t)
    set_run(r, "Times New Roman", 12, False)


# =========================
# Image normalization
# =========================
def _bytes_look_like_html(data: bytes) -> bool:
    if not data:
        return True
    head = data[:400].lower()
    return head.startswith(b"<!doctype html") or b"<html" in head or b"<head" in head


def normalize_image_bytes_for_docx(img_bytes: bytes) -> bytes:
    """
    Convert any valid image bytes to clean PNG bytes that python-docx accepts.
    Raises ValueError if bytes are not an image (e.g., HTML/login page).
    """
    if _bytes_look_like_html(img_bytes):
        raise ValueError("Not an image (HTML/login page).")
    img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    out = io.BytesIO()
    img.save(out, format="PNG", optimize=True)
    return out.getvalue()


# =========================
# Data cleaning helpers
# =========================
def _parse_bool(value: Any) -> Optional[bool]:
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        if value == 1:
            return True
        if value == 0:
            return False
    sv = str(value).strip().lower()
    if sv in {"yes", "y", "true", "t", "1", "checked", "✔", "✅"}:
        return True
    if sv in {"no", "n", "false", "f", "0", "unchecked", "✘", "❌"}:
        return False
    return None


def yes_no_line(value: Any, label_yes: str = "Yes", label_no: str = "No") -> str:
    b = _parse_bool(value)
    if b is True:
        return f"{CHECKED} {label_yes}    {UNCHECKED} {label_no}"
    if b is False:
        return f"{UNCHECKED} {label_yes}    {CHECKED} {label_no}"
    return f"{UNCHECKED} {label_yes}    {UNCHECKED} {label_no}"


def two_option_checkbox(value: Any, opt1: str, opt2: str) -> str:
    """
    Supports:
      - strings: 'male'/'female', 'm'/'f'
      - numbers: 1=Male, 2=Female (your dataset)
    """
    # --- numeric mapping (dataset codes) ---
    if isinstance(value, (int, float)):
        if int(value) == 1:
            return f"{CHECKED} {opt1}    {UNCHECKED} {opt2}"
        if int(value) == 2:
            return f"{UNCHECKED} {opt1}    {CHECKED} {opt2}"

    v = s(value).strip().lower()
    o1 = opt1.strip().lower()
    o2 = opt2.strip().lower()

    # also handle numeric strings "1"/"2"
    if v == "1":
        return f"{CHECKED} {opt1}    {UNCHECKED} {opt2}"
    if v == "2":
        return f"{UNCHECKED} {opt1}    {CHECKED} {opt2}"

    if v == o1:
        return f"{CHECKED} {opt1}    {UNCHECKED} {opt2}"
    if v == o2:
        return f"{UNCHECKED} {opt1}    {CHECKED} {opt2}"

    if v in {"m", "male", "man"} and o1 == "male":
        return f"{CHECKED} {opt1}    {UNCHECKED} {opt2}"
    if v in {"f", "female", "woman"} and o2 == "female":
        return f"{UNCHECKED} {opt1}    {CHECKED} {opt2}"

    return f"{UNCHECKED} {opt1}    {UNCHECKED} {opt2}"



def three_option_checkbox(value: Any, a: str, b: str, c: str) -> str:
    v = s(value).strip().lower()
    aa, bb, cc = a.lower(), b.lower(), c.lower()

    def box(txt, is_on):
        return f"{CHECKED if is_on else UNCHECKED} {txt}"

    return "    ".join([
        box(a, v == aa),
        box(b, v == bb),
        box(c, v == cc),
    ])


def remove_tool_prefix(text: Any) -> str:
    t = s(text)
    return re.sub(r"^Tool\s*\d+\s*", "", t, flags=re.IGNORECASE).strip()


def extract_date_only(value: Any) -> str:
    sv = s(value)
    if not sv:
        return ""
    return sv.split(" ")[0].strip() if " " in sv else sv


def format_date_dd_mon_yyyy(value: Any) -> str:
    sv = s(value)
    if not sv:
        return ""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(sv.split(".")[0], fmt)
            return dt.strftime("%d/%b/%Y")
        except Exception:
            pass
    return extract_date_only(sv)


def format_af_phone(value: Any) -> str:
    sv = re.sub(r"\D+", "", s(value))
    if not sv:
        return ""
    if sv.startswith("0"):
        sv = sv[1:]
    if sv.startswith("93"):
        return f"+{sv}"
    return f"+93{sv}"


def na_if_empty(value: Any) -> str:
    sv = s(value)
    return sv if sv else "N/A"


def extract_village_name(village_raw: Any) -> str:
    """
    'Com13731-Saracha village' -> 'Saracha village'
    'Com20832 - Taghaye Khwajaulya' -> 'Taghaye Khwajaulya'
    """
    t = s(village_raw)
    if not t:
        return ""
    if "-" in t:
        _, right = t.split("-", 1)
        right = right.strip()
        return right if right else t.strip()
    return t.strip()


def extract_cdc_code_from_village(village_raw: Any) -> str:
    """'Com20832 - Taghaye Khwajaulya' -> '20832'"""
    t = s(village_raw)
    if not t:
        return ""
    m = re.search(r"(\d{3,})", t)
    return m.group(1) if m else ""


def normalize_email_or_na_strict(value: Any) -> str:
    """
    - If empty -> N/A
    - If contains 'example' or looks invalid -> N/A
    """
    email = s(value).strip()
    if not email:
        return "N/A"
    low = email.lower()
    if "example" in low or "test" in low:
        return "N/A"
    if "@" not in email or "." not in email.split("@")[-1]:
        return "N/A"
    return email


def donor_upper_and_pipe(value: Any) -> str:
    """
    - Uppercase
    - If list or multi values, join by |
    """
    if value is None:
        return ""
    if isinstance(value, list):
        vals = [s(x).upper() for x in value if s(x)]
        return " | ".join(vals)

    t = s(value)
    if not t:
        return ""
    parts = re.split(r"[;,/|]+", t)
    parts = [p.strip().upper() for p in parts if p.strip()]
    return " | ".join(parts)


def _set_cell_paragraph_tight(p, align=WD_ALIGN_PARAGRAPH.LEFT):
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1


# =========================
# ✅ FINAL: Available documents inner table (CONNECT to outer)
# =========================
def add_available_documents_inner_table(details_cell, row: Dict[str, Any], details_width_emu: int):
    """
    Nested table whose left/right borders CONNECT to outer table borders.
    ❗️IMPORTANT: width must match Details cell width exactly.
    """
    details_cell.text = ""
    _set_cell_margins_zero(details_cell)

    inner = details_cell.add_table(rows=4, cols=4)
    inner.autofit = False
    inner.alignment = WD_TABLE_ALIGNMENT.LEFT
    _set_table_fixed_layout(inner)

    # ⭐️ Key: exact width equals Details cell width
    _set_table_width_and_indent(inner, details_width_emu)

    # ⭐️ Border logic: left/right connect, top/bottom off, inside on
    _set_table_borders_connect_to_outer(inner)

    # row height
    for r in inner.rows:
        _set_row_height_exact(r, 360)

    left_items = [
        ("Contract", row.get("B3_Contract")),
        ("Journal", row.get("B4_Journal")),
        ("BOQ", row.get("B5_BOQ")),
        ("Drawing", row.get("B1_Design_drawings")),
    ]

    right_items = [
        ("Site engineer", row.get("B6_Site_engineer")),
        ("Geophysical / hydrological tests", row.get("B7_Geophysical")),
        ("Water quality tests", row.get("D5_1_water_quality_compliance")),
        ("Pump test result", row.get("D4_pump_test_results_available")),
    ]

    # column widths MUST sum to details_width
    inner.columns[0].width = int(details_width_emu * 0.18)
    inner.columns[1].width = int(details_width_emu * 0.17)
    inner.columns[2].width = int(details_width_emu * 0.33)
    inner.columns[3].width = int(details_width_emu * 0.32)

    for i in range(4):
        inner.cell(i, 0).text = left_items[i][0]
        inner.cell(i, 2).text = right_items[i][0]

        inner.cell(i, 1).text = yes_no_line(left_items[i][1])
        inner.cell(i, 3).text = yes_no_line(right_items[i][1])

        for j in range(4):
            c = inner.cell(i, j)
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            _set_cell_margins_zero(c)

            for p in c.paragraphs:
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j in (1, 3) else WD_ALIGN_PARAGRAPH.LEFT
                for rrun in p.runs:
                    set_run(rrun, "Times New Roman", 11, False)


# =========================
# TOC field
# =========================
def add_toc_field(paragraph) -> None:
    run = paragraph.add_run()

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = r'TOC \o "1-3" \h \z \u'

    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")

    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")

    run._r.append(begin)
    run._r.append(instr)
    run._r.append(sep)

    ph = paragraph.add_run("Right-click and choose 'Update Field' to generate the table of contents.")
    set_run(ph, "Times New Roman", 11, False)

    run._r.append(end)


# =========================
# Cover page
# =========================
def add_cover_page(
    doc: Document,
    row: Dict[str, Any],
    cover_image_bytes: bytes,
    unicef_logo_path: Optional[str],
    act_logo_path: Optional[str],
    ppc_logo_path: Optional[str],
) -> None:
    title_blue = RGBColor(0, 112, 192)
    cover_font = "Cambria"

    section = doc.sections[0]
    usable_width = section.page_width - section.left_margin - section.right_margin
    usable_width_emu = int(usable_width)

    if unicef_logo_path:
        hp = section.header.paragraphs[0] if section.header.paragraphs else section.header.add_paragraph()
        clear_paragraph(hp)
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hp.paragraph_format.space_before = Pt(0)
        hp.paragraph_format.space_after = Pt(0)
        hp.add_run().add_picture(unicef_logo_path, width=Inches(1.4))

    if act_logo_path or ppc_logo_path:
        ft = section.footer.add_table(rows=1, cols=2, width=usable_width)
        ft.autofit = False
        _set_table_fixed_layout(ft)
        _set_table_width_and_indent(ft, usable_width_emu)

        logo_h_act = Inches(1.0)
        logo_h_ppc = Inches(0.35)

        if act_logo_path:
            c0 = ft.cell(0, 0)
            c0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p0 = c0.paragraphs[0]
            clear_paragraph(p0)
            p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p0.paragraph_format.space_before = Pt(0)
            p0.paragraph_format.space_after = Pt(0)
            p0.add_run().add_picture(act_logo_path, height=logo_h_act)

        if ppc_logo_path:
            c1 = ft.cell(0, 1)
            c1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p1 = c1.paragraphs[0]
            clear_paragraph(p1)
            p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p1.paragraph_format.space_before = Pt(0)
            p1.paragraph_format.space_after = Pt(0)
            p1.add_run().add_picture(ppc_logo_path, height=logo_h_ppc)

    for line in ["Third-party Monitoring (TPM) — WASH PROGRAMME", "FIELD Visit REPORT"]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(line)
        set_run(r, cover_font, 20, True, title_blue)

    doc.add_paragraph("")

    try:
        cover_clean = normalize_image_bytes_for_docx(cover_image_bytes)
        pimg = doc.add_paragraph()
        pimg.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pimg.add_run().add_picture(io.BytesIO(cover_clean), width=usable_width)
    except Exception:
        doc.add_paragraph("Cover photo could not be embedded.")

    doc.add_paragraph("")
    province = s(row.get("A01_Province"))
    district = s(row.get("A02_District"))
    village = extract_village_name(row.get("Village"))
    location = ", ".join([x for x in [province, district, village] if x])

    tool_value = row.get("Tool_Name", row.get("Tool", ""))

    cover_rows = [
        ("Project Title:", s(row.get("Activity_Name"))),
        ("Visit No.:", s(row.get("A26_Visit_number"))),
        ("Type of Intervention:", remove_tool_prefix(tool_value)),
        ("Province / District / Village:", location),
        ("Date of Visit:", format_date_dd_mon_yyyy(row.get("starttime"))),
        ("Implementing Partner (IP):", s(row.get("Primary_Partner_Name"))),
        ("Prepared by:", "Premium Performance Consulting (PPC) & Act for Performance"),
        ("Prepared for:", "UNICEF"),
    ]

    table = doc.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    set_table_borders(table)

    _set_table_fixed_layout(table)
    _set_table_width_and_indent(table, usable_width_emu)

    left_w = usable_width // 2
    right_w = usable_width - left_w

    for label, value in cover_rows:
        row_cells = table.add_row().cells
        c0, c1 = row_cells[0], row_cells[1]
        c0.width = left_w
        c1.width = right_w

        shade(c0, "D9E2F3")

        r0 = c0.paragraphs[0].add_run(label)
        set_run(r0, cover_font, 12, True)

        r1 = c1.paragraphs[0].add_run(s(value))
        set_run(r1, cover_font, 11, False)


# =========================
# PAGE 2: TOC
# =========================
def add_toc_page(doc: Document) -> None:
    doc.add_page_break()
    h1(doc, "TABLE OF CONTENTS", underline=True)
    doc.add_paragraph("")
    for i in range(1, 4):
        set_style(doc, f"TOC {i}", "Times New Roman", 11, False)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    add_toc_field(p)


# =========================
# PAGE 3: General Project Information (✅ FIXED CALL + WIDTH PASS)
# =========================
def add_general_project_information(doc: Document, row: Dict[str, Any], overrides: Optional[Dict[str, Any]] = None) -> None:
    overrides = overrides or {}
    doc.add_page_break()
    h1(doc, "1.        General Project Information:", underline=True)
    doc.add_paragraph("")

    sec0 = doc.sections[0]
    usable_width = sec0.page_width - sec0.left_margin - sec0.right_margin
    usable_width_emu = int(usable_width)

    province = s(overrides.get("Province", row.get("A01_Province")))
    district = s(overrides.get("District", row.get("A02_District")))

    village_raw = overrides.get("Village / Community", row.get("Village"))
    village = extract_village_name(village_raw)

    gps_lat = s(row.get("GPS_1-Latitude"))
    gps_lon = s(row.get("GPS_1-Longitude"))

    project_name = s(overrides.get("Project Name", row.get("Activity_Name")))
    visit_date = s(overrides.get("Date of Visit", format_date_dd_mon_yyyy(row.get("starttime"))))
    ip_name = s(overrides.get("Name of the IP, Organization / NGO", row.get("Primary_Partner_Name")))

    monitor_name = s(overrides.get("Name of the monitor Engineer", row.get("A07_Monitor_name")))
    monitor_email = s(overrides.get("Email of the monitor engineer", row.get("A12_Monitor_email")))

    respondent_name = s(overrides.get("Name of the respondent (Participant / UNICEF / IPs)", row.get("A08_Respondent_name")))
    respondent_sex_val = overrides.get("Sex of Respondent", row.get("A09_Respondent_sex"))
    respondent_sex = two_option_checkbox(respondent_sex_val, "Male", "Female")

    respondent_phone = s(overrides.get("Contact Number of the Respondent", format_af_phone(row.get("A10_Respondent_phone"))))
    respondent_email = s(overrides.get("Email Address of the Respondent", normalize_email_or_na_strict(row.get("A11_Respondent_email"))))

    cost_label = s(row.get("A14_Project_cost_label")).lower()
    estimated_amount = s(row.get("A14_Estimated_cost_amount") or row.get("Estimated_Project_Cost_amount") or row.get("A14_Estimated_cost_amount_label"))
    contracted_amount = s(row.get("A14_Contracted_cost_amount") or row.get("Contracted_Project_Cost_amount") or row.get("A14_Contracted_cost_amount_label"))
    estimated_cost = estimated_amount if "estimated" in cost_label else ""
    contracted_cost = contracted_amount if "contract" in cost_label else ""

    project_status_val = overrides.get("Project Status", row.get("A15_Project_status_label"))
    project_status = three_option_checkbox(project_status_val, "Ongoing", "Completed", "Suspended")

    reason_delay = s(overrides.get("Reason for delay", na_if_empty(row.get("B8_Reasons_for_delay"))))

    progress_val = overrides.get("Project progress", row.get("Project_progress") or row.get("A22_Project_progress_label"))
    project_progress = three_option_checkbox(progress_val, "Ahead of Schedule", "On Schedule", "Running behind")

    contract_start = s(overrides.get("Contract Start Date", format_date_dd_mon_yyyy(row.get("A16_Start_date"))))
    contract_end = s(overrides.get("Contract End Date", format_date_dd_mon_yyyy(row.get("A17_End_date"))))

    prev_phys = s(overrides.get("Previous Physical Progress (%)", row.get("A18_Previous_progress")))
    curr_phys = s(overrides.get("Current Physical Progress (%)", row.get("A19_Current_progress")))

    cdc_code = s(overrides.get("CDC Code", extract_cdc_code_from_village(row.get("Village"))))
    donor_name = donor_upper_and_pipe(overrides.get("Donor Name", row.get("A24_Donor_name")))

    monitoring_report_no = s(overrides.get("Monitoring Report Number", row.get("A25_Monitoring_report_number")))
    current_report_date = s(overrides.get("Date of Current Report", format_date_dd_mon_yyyy(row.get("A20_Current_report_date"))))
    last_report_date = s(overrides.get("Date of Last Monitoring Report", format_date_dd_mon_yyyy(row.get("A21_Previous_report_date"))))
    sites_visited = s(overrides.get("Number of Sites Visited", row.get("A26_Visit_number")))

    community_agreement = overrides.get(
        "community agreement - Is the community/user group agreed on the well site?",
        row.get("community_agreement") or row.get("community_agreement_label") or row.get("Community_agreement")
    )
    work_safety = overrides.get("Is work_safety_considered -", row.get("A13_Is_work_safety_considered_label"))
    env_risk = overrides.get("environmental risk -", row.get("C2_environmental_risk_label"))

    # --- Table layout ---
    field_w = Inches(4.5)
    field_w_emu = int(field_w)

    # ✅ This is the REAL Details cell width (MUST be passed to inner table)
    details_w_emu = int(usable_width_emu) - int(field_w_emu)

    c1_w = int(details_w_emu * 0.25)
    c2_w = int(details_w_emu * 0.25)
    c3_w = int(details_w_emu * 0.25)
    c4_w = int(details_w_emu - (c1_w + c2_w + c3_w))

    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    set_table_borders(table)

    _set_table_fixed_layout(table)
    _set_table_width_and_indent(table, usable_width_emu)

    table.columns[0].width = field_w_emu
    table.columns[1].width = c1_w
    table.columns[2].width = c2_w
    table.columns[3].width = c3_w
    table.columns[4].width = c4_w

    # Header
    hdr = table.rows[0].cells
    hdr[0].text = "Field"
    hdr[1].text = "Details"
    hdr[1].merge(hdr[2]).merge(hdr[3]).merge(hdr[4])
    for cell in (hdr[0], hdr[1]):
        shade(cell, "D9E2F3")
        for p in cell.paragraphs:
            for r in p.runs:
                set_run(r, "Times New Roman", 12, True)

    def style_field(cell) -> None:
        shade(cell, "D9E2F3")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for r in p.runs:
                set_run(r, "Times New Roman", 12, False)

    def add_row(field: str, value: str) -> None:
        cells = table.add_row().cells
        cells[0].text = field
        style_field(cells[0])

        cells[1].text = s(value)
        cells[1].merge(cells[2]).merge(cells[3]).merge(cells[4])
        for p in cells[1].paragraphs:
            for r in p.runs:
                set_run(r, "Times New Roman", 12, False)

    def add_row_with_hint(field: str, value: str, hint: str) -> None:
        cells = table.add_row().cells
        cells[0].text = field
        style_field(cells[0])

        d = cells[1].merge(cells[2]).merge(cells[3]).merge(cells[4])
        d.text = ""
        p = d.paragraphs[0]
        _set_cell_paragraph_tight(p, align=WD_ALIGN_PARAGRAPH.LEFT)

        r1 = p.add_run(s(value))
        set_run(r1, "Times New Roman", 12, False)

        r2 = p.add_run(f"  ({hint})")
        set_run(r2, "Times New Roman", 10, False)

    def add_row_custom(field: str, renderer_fn) -> None:
        cells = table.add_row().cells
        cells[0].text = field
        style_field(cells[0])

        details = cells[1].merge(cells[2]).merge(cells[3]).merge(cells[4])
        details.text = ""
        renderer_fn(details)

    # Rows
    add_row("Province", province)
    add_row("District", district)
    add_row("Village / Community", village)

    add_row("GPS points", f"{gps_lat}, {gps_lon}")
    add_row("Project Name", project_name)
    add_row("Date of Visit", visit_date)
    add_row("Name of the IP, Organization / NGO", ip_name)
    add_row("Name of the monitor Engineer", monitor_name)
    add_row("Email of the monitor engineer", monitor_email)

    add_row("Name of the respondent (Participant / UNICEF / IPs)", respondent_name)
    add_row("Sex of Respondent", respondent_sex)
    add_row("Contact Number of the Respondent", respondent_phone)
    add_row("Email Address of the Respondent", respondent_email)

    add_row("Estimated Project Cost", na_if_empty(estimated_cost))
    add_row("Contracted Project Cost", na_if_empty(contracted_cost))

    add_row("Project Status", project_status)
    add_row("Reason for delay", reason_delay)

    add_row("Project progress", project_progress)
    add_row("Contract Start Date", na_if_empty(contract_start))
    add_row("Contract End Date", na_if_empty(contract_end))
    add_row("Previous Physical Progress (%)", na_if_empty(prev_phys))
    add_row("Current Physical Progress (%)", na_if_empty(curr_phys))

    add_row("CDC Code", cdc_code)
    add_row("Donor Name", donor_name)

    add_row_with_hint("Monitoring Report Number", monitoring_report_no, "Please verify in report")
    add_row("Date of Current Report", current_report_date)
    add_row("Date of Last Monitoring Report", na_if_empty(last_report_date))
    add_row("Number of Sites Visited", sites_visited)

    # Available documents (inner 4x4 table like screenshot)
    add_row_custom(
        "Available documents in the site",
        lambda cell: add_available_documents_inner_table(cell, row, details_w_emu),
    )

    add_row("community agreement - Is the community/user group agreed on the well site?", yes_no_line(community_agreement))
    add_row("Is work_safety_considered -", yes_no_line(work_safety))
    add_row("environmental risk -", yes_no_line(env_risk))


# =========================
# PAGE 4: Executive Summary
# =========================
def build_executive_summary_auto(row: Dict[str, Any], component_observations: Optional[List[Dict[str, Any]]] = None, max_chars: int = 2500) -> str:
    if row is None or row is ... or not isinstance(row, dict):
        row = {}

    component_observations = component_observations or []

    province = s(row.get("A01_Province"))
    district = s(row.get("A02_District"))
    village = extract_village_name(row.get("Village"))
    location = ", ".join([x for x in [village, district, province] if x])

    project_name = s(row.get("Activity_Name"))
    intervention = remove_tool_prefix(s(row.get("Tool_Name") or row.get("Tool") or "WASH intervention"))
    ip = s(row.get("Primary_Partner_Name"))
    visit_date = format_date_dd_mon_yyyy(row.get("starttime"))
    cdc = extract_cdc_code_from_village(row.get("Village"))

    # Project status and progress (if available)
    status = s(row.get("A15_Project_status_label") or row.get("Project_Status") or "")
    prev_prog = s(row.get("A18_Previous_progress") or "")
    curr_prog = s(row.get("A19_Current_progress") or "")

    # Extract quick signals from component_observations
    comp_titles = []
    major_findings_texts = []
    for comp in component_observations:
        t = s(comp.get("title"))
        if t:
            comp_titles.append(t)

        for sub in (comp.get("subsections") or []):
            mt = sub.get("major_table")
            if isinstance(mt, list):
                for rr in mt:
                    ftxt = s(rr.get("Findings") or rr.get("Finding") or "")
                    compv = s(rr.get("Compliance") or "")
                    if ftxt:
                        major_findings_texts.append((ftxt, compv))

    # Limit for readability
    major_findings_texts = major_findings_texts[:8]

    # Paragraph 1: purpose + scope
    p1 = (
        f"This Third-Party Monitoring (TPM) field visit was conducted to assess the technical implementation, "
        f"functionality, and compliance of the {intervention} project ({project_name}) in {location}. "
        f"The visit was carried out on {visit_date} to verify system functionality, adherence to approved designs "
        f"and BOQ, and to identify any technical or operational issues that may affect long-term performance."
    )

    # Paragraph 2: IP + CDC + status
    bits = []
    if ip:
        bits.append(f"The implementing partner is {ip}.")
    if cdc:
        bits.append(f"CDC code: {cdc}.")
    if status:
        bits.append(f"Project status reported as: {status}.")
    if prev_prog or curr_prog:
        prog = []
        if prev_prog:
            prog.append(f"previous physical progress {prev_prog}%")
        if curr_prog:
            prog.append(f"current physical progress {curr_prog}%")
        if prog:
            bits.append("Reported " + ", ".join(prog) + ".")
    p2 = " ".join(bits).strip()
    if not p2:
        p2 = "The monitoring included site observation, document review, and stakeholder consultations."

    # Paragraph 3: what was confirmed / functional (use components if provided)
    if comp_titles:
        p3 = (
            "The assessment confirmed that key infrastructure components were observed on site and reviewed for "
            "construction quality and operational readiness. The review covered major system elements such as bore wells, "
            "solar-powered pumping systems, storage reservoirs, distribution network, and associated civil structures "
            "including protective and auxiliary works, in line with the project scope."
        )
    else:
        p3 = (
            "The assessment reviewed the main water supply infrastructure, including source works, solar pumping system, "
            "reservoirs, distribution pipelines, and associated civil works, to confirm operational performance and compliance."
        )

    # Paragraph 4: gaps / issues (summarize major findings table)
    if major_findings_texts:
        issue_lines = []
        for ftxt, compv in major_findings_texts:
            # highlight non-compliance more
            if s(compv).strip().lower() in {"no", "n", "non", "not compliant"}:
                issue_lines.append(ftxt)
        if not issue_lines:
            issue_lines = [x[0] for x in major_findings_texts]

        p4 = (
            "Some technical and operational gaps were identified during the monitoring. Key issues observed include: "
            + "; ".join(issue_lines[:6]).rstrip("; ")
            + ". These gaps require corrective actions to prevent reduced efficiency, durability concerns, and operational risks."
        )
    else:
        p4 = (
            "Some technical and operational gaps were identified during the monitoring, primarily related to construction quality, "
            "system operation arrangements, and compliance with technical standards. These require follow-up actions to ensure sustainability."
        )

    # Paragraph 5: conclusion
    p5 = (
        "Overall, the project is functional and providing services to the beneficiary community. Addressing the identified issues through "
        "the recommended corrective actions will enhance system reliability, improve safety, and support long-term sustainability of the service."
    )

    text = "\n\n".join([p1, p2, p3, p4, p5]).strip()

    if len(text) <= max_chars:
        return text
    return text[: max_chars - 3] + "..."



def add_executive_summary(
    doc: Document,
    row: Dict[str, Any],
    executive_summary_text: str = "",
    component_observations: Optional[List[Dict[str, Any]]] = None
) -> None:
    doc.add_page_break()
    h1(doc, "2.        Executive Summary:", underline=True)
    doc.add_paragraph("")

    component_observations = component_observations or []

    if s(executive_summary_text):
        add_paragraphs(doc, executive_summary_text, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    else:
        auto_text = build_executive_summary_auto(
            row=row,
            component_observations=component_observations,
            max_chars=2500
        )
        add_paragraphs(doc, auto_text, align=WD_ALIGN_PARAGRAPH.JUSTIFY)


def build_summary_and_conclusion_auto(
    row: Dict[str, Any],
    component_observations: Optional[List[Dict[str, Any]]] = None,
    work_progress_findings: Optional[List[Dict[str, Any]]] = None,
) -> str:
    """
    Auto-generate a detailed and standard Summary & Conclusion section
    based on project status, progress, findings, and component observations.
    """
    component_observations = component_observations or []
    work_progress_findings = work_progress_findings or []

    # Location
    province = s(row.get("A01_Province"))
    district = s(row.get("A02_District"))
    village = extract_village_name(row.get("Village"))
    location = ", ".join([x for x in [village, district, province] if x])

    project_name = s(row.get("Activity_Name"))
    intervention = remove_tool_prefix(s(row.get("Tool_Name") or row.get("Tool") or "WASH intervention"))

    # Status & progress
    status = s(row.get("A15_Project_status_label") or "").lower()
    prev_prog = s(row.get("A18_Previous_progress") or "")
    curr_prog = s(row.get("A19_Current_progress") or "")

    # Determine completion wording
    if "completed" in status:
        status_sentence = "has been completed"
    elif "ongoing" in status:
        status_sentence = "is currently under implementation"
    elif "suspend" in status:
        status_sentence = "is currently suspended"
    else:
        status_sentence = "is under implementation"

    # Count components and findings
    component_count = 0
    major_findings = []
    non_compliance_count = 0

    for comp in component_observations:
        component_count += 1
        for sub in (comp.get("subsections") or []):
            mt = sub.get("major_table")
            if isinstance(mt, list):
                for rr in mt:
                    ftxt = s(rr.get("Findings") or rr.get("Finding") or "")
                    compv = s(rr.get("Compliance") or "").lower()
                    if ftxt:
                        major_findings.append(ftxt)
                        if compv in {"no", "n", "not compliant", "non"}:
                            non_compliance_count += 1

    summary_findings_count = len([
        x for x in work_progress_findings
        if isinstance(x, dict) and any(s(v) for v in x.values())
    ])

    # ---- Paragraph 1: Overall project status ----
    p1 = (
        f"The {intervention} project"
        + (f" in {location}" if location else "")
        + f" {status_sentence}, and the main infrastructure components have been installed and reviewed during the monitoring visit. "
        "The system was observed to be generally functional and capable of delivering the intended services to the target community."
    )

    # ---- Paragraph 2: Progress & implementation quality ----
    prog_bits = []
    if prev_prog:
        prog_bits.append(f"previous physical progress of {prev_prog}%")
    if curr_prog:
        prog_bits.append(f"current physical progress of {curr_prog}%")

    if prog_bits:
        p2 = (
            "Based on the monitoring records and site observations, the project implementation reflects "
            + ", ".join(prog_bits)
            + ". Construction quality and workmanship were assessed across key components, with overall progress broadly aligned with the approved scope of work."
        )
    else:
        p2 = (
            "Construction quality and workmanship were assessed across key components, with overall implementation broadly aligned with the approved scope of work."
        )

    # ---- Paragraph 3: Issues and gaps ----
    if major_findings:
        p3 = (
            "Despite the overall functionality of the system, several technical and operational gaps were identified during the assessment. "
            "Key issues observed include "
            + "; ".join(major_findings[:6]).rstrip("; ")
            + ". These gaps mainly relate to construction detailing, system protection, operational controls, and operation and maintenance arrangements."
        )
    else:
        p3 = (
            "No critical technical deficiencies were observed during the monitoring; however, minor improvements related to operation, maintenance, and system protection may further enhance performance."
        )

    # ---- Paragraph 4: Compliance & risk ----
    if non_compliance_count > 0 or summary_findings_count > 0:
        p4 = (
            f"The identified findings (total recorded issues: {summary_findings_count or len(major_findings)}, "
            f"including {non_compliance_count} non-compliance items) were assessed in terms of severity and potential risk to system performance. "
            "These findings require timely corrective actions to mitigate operational risks and prevent deterioration of the infrastructure."
        )
    else:
        p4 = (
            "The system generally complies with the applicable technical standards and project requirements, with no significant risks identified that would hinder service delivery."
        )

    # ---- Paragraph 5: Conclusion & sustainability ----
    p5 = (
        "In conclusion, the project is providing water services to the beneficiary community. "
        "Addressing the identified issues through the recommended corrective actions, strengthening operation and maintenance practices, "
        "and enhancing community capacity will contribute to improved system reliability, safe operation, and long-term sustainability of the water supply service."
    )

    return "\n\n".join([p1, p2, p3, p4, p5]).strip()


# =========================
# Section 3 + 4 (minimal)
# =========================
def build_data_collection_methods_auto(
    row: Dict[str, Any],
    component_observations: Optional[List[Dict[str, Any]]] = None,
    work_progress_findings: Optional[List[Dict[str, Any]]] = None,
    max_chars: int = 1800,
) -> str:
    if row is None or row is ... or not isinstance(row, dict):
        row = {}

    component_observations = component_observations or []
    work_progress_findings = work_progress_findings or []

    intervention = remove_tool_prefix(s(row.get("Tool_Name") or row.get("Tool") or "WASH intervention"))
    project_name = s(row.get("Activity_Name"))
    visit_date = format_date_dd_mon_yyyy(row.get("starttime"))

    # Simple signals
    ip = s(row.get("Primary_Partner_Name"))
    province = s(row.get("A01_Province"))
    district = s(row.get("A02_District"))
    village = extract_village_name(row.get("Village"))
    location = ", ".join([x for x in [village, district, province] if x])

    # Count components checked
    comp_count = len([c for c in component_observations if isinstance(c, dict)])
    finding_count = len([f for f in (work_progress_findings or []) if isinstance(f, dict) and any(s(v) for v in f.values())])

    # Detect doc availability from section 1 fields
    docs_yes = 0
    for k in [
        "B3_Contract",
        "B4_Journal",
        "B5_BOQ",
        "B1_Design_drawings",
        "B6_Site_engineer",
        "B7_Geophysical",
        "D5_1_water_quality_compliance",
        "D4_pump_test_results_available",
    ]:
        if _parse_bool(row.get(k)) is True:
            docs_yes += 1

        if _parse_bool(row.get(k)) is True:
            docs_yes += 1

    bits = []
    bits.append(
        f"The Third-Party Monitoring (TPM) assessment for the {intervention}"
        + (f" project ({project_name})" if project_name else " project")
        + (f" in {location}" if location else "")
        + (f" was conducted on {visit_date}." if visit_date else ".")
    )
    if ip:
        bits.append(f"The implementing partner was {ip}.")
    bits.append(
        "A mixed-methods approach was applied, combining direct on-site technical observation, "
        "structured compliance checklists, document review (approved drawings, BOQ, and contract documents), "
        "and semi-structured interviews with implementing partner staff and community representatives."
    )

    if docs_yes > 0:
        bits.append(
            f"During the visit, available project documents and records were checked (documents confirmed: {docs_yes})."
        )

    if comp_count > 0:
        bits.append(
            f"Physical verification covered key project components (components reviewed: {comp_count}), "
            "focusing on construction quality, workmanship, system protection measures, and functionality."
        )
    else:
        bits.append(
            "Physical evidence was inspected across visible project components, focusing on construction quality, "
            "functionality, and compliance with the approved scope of work."
        )

    if finding_count > 0:
        bits.append(
            f"Findings were recorded (total issues captured: {finding_count}), categorized by risk/severity, "
            "and linked to practical corrective actions to support sustainability in line with WASH standards."
        )
    else:
        bits.append(
            "Observations were documented and assessed against technical requirements, and practical corrective actions were proposed where needed."
        )

    text = " ".join(bits).strip()
    return text if len(text) <= max_chars else (text[: max_chars - 3] + "...")

def add_section_3_and_4_auto(
    doc: Document,
    row: Dict[str, Any],
    component_observations: Optional[List[Dict[str, Any]]] = None,
    work_progress_findings: Optional[List[Dict[str, Any]]] = None,
) -> None:
    # --- Section 3
    doc.add_page_break()
    h1(doc, "3.        Data Collection Methods", underline=True)
    doc.add_paragraph("")

    items = [
        "Direct observation of work progress on-site.",
        "Technical Compliance Checklists",
        "Review of the documents (BOQ, Drawings, Contract, Test results)",
        "Interviews with technical staff of the contracted company and CDC members",
        "Review of physical evidence related to construction activities",
    ]
    for it in items:
        p = doc.add_paragraph(it, style="List Number")
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        for r in p.runs:
            set_run(r, "Times New Roman", 12, False)

    # ✅ Auto narrative paragraph (richer + dynamic)
    doc.add_paragraph("")
    auto_text = build_data_collection_methods_auto(
        row=row,
        component_observations=component_observations or [],
        work_progress_findings=work_progress_findings or [],
    )
    add_paragraphs(doc, auto_text, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # --- Section 4
    doc.add_page_break()
    h1(doc, "4.        Work Progress Summary during the Visit.", underline=True)
    doc.add_paragraph("")
    body(doc, "Work progress table can be added here (already handled in your UI pipeline).")

# =========================
# Section 6 + 7 + 8 (kept from your code)
# =========================
def _set_cell_borders(cell, size=8, color="000000") -> None:
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)

    for edge in ("top", "left", "bottom", "right"):
        tag = qn(f"w:{edge}")
        element = tcBorders.find(tag)
        if element is None:
            element = OxmlElement(f"w:{edge}")
            tcBorders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), str(int(size)))
        element.set(qn("w:color"), color)


def _normalize_severity(val: Any) -> str:
    v = s(val).strip().lower()
    if not v:
        return ""
    m = {
        "low": "Low",
        "medium": "Medium",
        "med": "Medium",
        "high": "High",
        "critical": "High",
        "very high": "High",
    }
    return m.get(v, s(val).strip())


def add_summary_of_findings_table(doc: Document, findings_rows: List[Dict[str, Any]]) -> None:
    doc.add_page_break()
    h1(doc, "6.        Summary of the findings:", underline=True)
    doc.add_paragraph("")

    findings_rows = findings_rows or []

    sec0 = doc.sections[0]
    usable_width = sec0.page_width - sec0.left_margin - sec0.right_margin
    usable_width_emu = int(usable_width)

    table = doc.add_table(rows=1, cols=4)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    set_table_borders(table)

    _set_table_fixed_layout(table)
    _set_table_width_and_indent(table, usable_width_emu)

    w_no = Inches(0.55)
    w_find = Inches(2.55)
    w_sev = Inches(1.05)
    w_rec = usable_width - w_no - w_find - w_sev

    widths = [w_no, w_find, w_sev, w_rec]
    headers = ["No.", "Finding", "Severity", "Recommendation / Corrective Action"]

    header_row = table.rows[0]
    _set_row_height_exact(header_row, 260)

    dark_blue = "1F4E79"
    for i, cell in enumerate(header_row.cells):
        cell.width = widths[i]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        shade(cell, dark_blue)
        _set_cell_borders(cell, size=10, color="000000")

        cell.text = ""
        p = cell.paragraphs[0]
        _set_cell_paragraph_tight(p, align=WD_ALIGN_PARAGRAPH.CENTER)

        r = p.add_run(headers[i])
        set_run(r, "Times New Roman", 10, True, RGBColor(255, 255, 255))

    for idx, rr in enumerate(findings_rows, start=1):
        finding = s(rr.get("Finding") or rr.get("finding") or rr.get("Issue") or rr.get("Observation"))
        severity = _normalize_severity(rr.get("Severity") or rr.get("severity") or rr.get("Priority"))
        reco = s(rr.get("Recommendation") or rr.get("recommendation") or rr.get("Corrective Action") or rr.get("Action"))

        if not (finding or reco or severity):
            continue

        row_cells = table.add_row().cells
        for i, cell in enumerate(row_cells):
            cell.width = widths[i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            _set_cell_borders(cell, size=8, color="000000")

        row_cells[0].text = ""
        p0 = row_cells[0].paragraphs[0]
        _set_cell_paragraph_tight(p0, align=WD_ALIGN_PARAGRAPH.LEFT)
        set_run(p0.add_run(str(idx)), "Times New Roman", 10, False)

        row_cells[1].text = ""
        p1 = row_cells[1].paragraphs[0]
        _set_cell_paragraph_tight(p1, align=WD_ALIGN_PARAGRAPH.LEFT)
        set_run(p1.add_run(finding), "Times New Roman", 10, False)

        row_cells[2].text = ""
        p2 = row_cells[2].paragraphs[0]
        _set_cell_paragraph_tight(p2, align=WD_ALIGN_PARAGRAPH.LEFT)
        set_run(p2.add_run(severity), "Times New Roman", 10, False)

        row_cells[3].text = ""
        p3 = row_cells[3].paragraphs[0]
        _set_cell_paragraph_tight(p3, align=WD_ALIGN_PARAGRAPH.LEFT)
        set_run(p3.add_run(reco), "Times New Roman", 10, False)

    doc.add_paragraph("")


def add_summary_and_conclusion(
    doc: Document,
    row: Dict[str, Any],
    conclusion_text: str = "",
    component_observations: Optional[List[Dict[str, Any]]] = None,
    work_progress_findings: Optional[List[Dict[str, Any]]] = None,
) -> None:
    doc.add_page_break()
    h1(doc, "7.        Summary and Conclusion", underline=True)
    doc.add_paragraph("")

    if s(conclusion_text):
        add_paragraphs(doc, conclusion_text, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    else:
        auto_text = build_summary_and_conclusion_auto(
            row=row,
            component_observations=component_observations or [],
            work_progress_findings=work_progress_findings or [],
        )
        add_paragraphs(doc, auto_text, align=WD_ALIGN_PARAGRAPH.JUSTIFY)



def add_photo_documentation(
    doc: Document,
    photo_selections: Dict[str, str],
    photo_bytes: Dict[str, bytes],
    photo_field_map: Optional[Dict[str, str]] = None,
) -> None:
    photo_field_map = photo_field_map or {}
    photo_selections = photo_selections or {}
    photo_bytes = photo_bytes or {}

    findings = [u for u, p in photo_selections.items() if p == "Findings" and u in photo_bytes]
    obs = [u for u, p in photo_selections.items() if p == "Observations" and u in photo_bytes]

    if not findings and not obs:
        return

    doc.add_page_break()
    h1(doc, "8.        Photo Documentation", underline=True)
    doc.add_paragraph("")

    def add_group(title: str, urls: List[str]) -> None:
        if not urls:
            return
        h2(doc, title)

        sec0 = doc.sections[0]
        usable_width = sec0.page_width - sec0.left_margin - sec0.right_margin
        img_w = int(usable_width * 0.48)

        table = doc.add_table(rows=0, cols=2)
        table.autofit = False

        for i in range(0, len(urls), 2):
            row_cells = table.add_row().cells
            for j in range(2):
                idx = i + j
                if idx >= len(urls):
                    continue

                u = urls[idx]
                label = photo_field_map.get(u, "Photo")

                p = row_cells[j].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

                try:
                    clean = normalize_image_bytes_for_docx(photo_bytes[u])
                    p.add_run().add_picture(io.BytesIO(clean), width=img_w)
                except Exception:
                    row_cells[j].add_paragraph("Photo could not be embedded.")

                cap = row_cells[j].add_paragraph(label)
                cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for rr in cap.runs:
                    set_run(rr, "Times New Roman", 10, False)

        doc.add_paragraph("")

    add_group("8.1 Findings Photos", findings)
    add_group("8.2 Observation Photos", obs)


# ============================================================
# ✅✅ SECTION 5 RESTORED (from your Streamlit Tool_6 structure)
# ============================================================

def _compliance_checkbox(value: Any) -> str:
    """
    Streamlit major_table uses "Yes"/"No" for compliance.
    We'll render it as checkbox line.
    """
    v = s(value).strip().lower()
    if v in {"yes", "y", "true", "1"}:
        return f"{CHECKED} Yes    {UNCHECKED} No"
    if v in {"no", "n", "false", "0"}:
        return f"{UNCHECKED} Yes    {CHECKED} No"
    return f"{UNCHECKED} Yes    {UNCHECKED} No"


def add_major_findings_table_tool6(
    doc: Document,
    major_rows: List[Dict[str, Any]],
    photo_bytes: Dict[str, bytes],
    photo_field_map: Optional[Dict[str, str]] = None,
) -> None:
    """
    Table columns: NO | Findings | Compliance (Yes/No) | Photos
    Photos can embed image if URL exists in photo_bytes.
    """
    photo_field_map = photo_field_map or {}
    major_rows = major_rows or []

    sec0 = doc.sections[0]
    usable_width = sec0.page_width - sec0.left_margin - sec0.right_margin
    usable_width_emu = int(usable_width)

    table = doc.add_table(rows=1, cols=4)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    set_table_borders(table)

    _set_table_fixed_layout(table)
    _set_table_width_and_indent(table, usable_width_emu)

    # width plan
    w_no = Inches(0.55)
    w_find = Inches(3.65)
    w_comp = Inches(1.35)
    w_photo = usable_width - w_no - w_find - w_comp

    widths = [w_no, w_find, w_comp, w_photo]
    headers = ["NO", "Findings", "Compliance\n", "Photos"]

    header_row = table.rows[0]
    _set_row_height_exact(header_row, 260)

    dark_blue = "1F4E79"
    for i, cell in enumerate(header_row.cells):
        cell.width = widths[i]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        shade(cell, dark_blue)
        _set_cell_borders(cell, size=10, color="000000")

        cell.text = ""
        p = cell.paragraphs[0]
        _set_cell_paragraph_tight(p, align=WD_ALIGN_PARAGRAPH.CENTER)
        r = p.add_run(headers[i])
        set_run(r, "Times New Roman", 10, True, RGBColor(255, 255, 255))

    # body
    idx = 0
    for rr in major_rows:
        finding = s(rr.get("Findings") or rr.get("Finding") or rr.get("finding"))
        compliance = rr.get("Compliance") or rr.get("compliance")
        photo_url = s(rr.get("Photo") or rr.get("photo"))

        if not (finding or compliance or photo_url):
            continue

        idx += 1
        row_cells = table.add_row().cells
        for i, cell in enumerate(row_cells):
            cell.width = widths[i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            _set_cell_borders(cell, size=8, color="000000")

        # NO
        row_cells[0].text = ""
        p0 = row_cells[0].paragraphs[0]
        _set_cell_paragraph_tight(p0, align=WD_ALIGN_PARAGRAPH.LEFT)
        set_run(p0.add_run(str(idx)), "Times New Roman", 10, False)

        # Findings
        row_cells[1].text = ""
        p1 = row_cells[1].paragraphs[0]
        _set_cell_paragraph_tight(p1, align=WD_ALIGN_PARAGRAPH.LEFT)
        set_run(p1.add_run(finding), "Times New Roman", 10, False)

        # Compliance checkbox
        row_cells[2].text = ""
        p2 = row_cells[2].paragraphs[0]
        _set_cell_paragraph_tight(p2, align=WD_ALIGN_PARAGRAPH.LEFT)
        set_run(p2.add_run(_compliance_checkbox(compliance)), "Times New Roman", 10, False)

        # Photos: embed if possible
        row_cells[3].text = ""
        p3 = row_cells[3].paragraphs[0]
        _set_cell_paragraph_tight(p3, align=WD_ALIGN_PARAGRAPH.CENTER)

        if photo_url and photo_url in (photo_bytes or {}):
            try:
                clean = normalize_image_bytes_for_docx(photo_bytes[photo_url])
                p3.add_run().add_picture(io.BytesIO(clean), width=Inches(1.35))
                cap = row_cells[3].add_paragraph(photo_field_map.get(photo_url, "Photo"))
                cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for rr2 in cap.runs:
                    set_run(rr2, "Times New Roman", 9, False)
            except Exception:
                # fallback to label only
                set_run(p3.add_run(photo_field_map.get(photo_url, "Photo")), "Times New Roman", 9, False)
        else:
            # keep empty if no bytes
            pass


def add_component_wise_key_observations_tool6(
    doc: Document,
    component_observations: List[Dict[str, Any]],
    photo_bytes: Dict[str, bytes],
    photo_field_map: Optional[Dict[str, str]] = None,
) -> None:
    """
    Expects the SAME structure produced in Tool_6.py:
      [
        {
          "comp_id": "5.1",
          "title": "5.1. ...:",
          "paragraphs": [...],
          "subsections": [
              {"title": "Major findings:", "major_table": [...]},
              {"title": "Recommendations ...:", "paragraphs": [...]}
          ],
          "photos": [url1,url2,url3]
        }, ...
      ]
    """
    photo_field_map = photo_field_map or {}
    component_observations = component_observations or []
    if not component_observations:
        return

    doc.add_page_break()
    h1(doc, "5.        Project Component-Wise Key Observations", underline=True)
    doc.add_paragraph("")

    for comp in component_observations:
        title = s(comp.get("title"))
        if title:
            h2(doc, title)
            doc.add_paragraph("")

        # Observation paragraphs
        for para in (comp.get("paragraphs") or []):
            body(doc, para)

        doc.add_paragraph("")

        # Subsections: Major findings table + recommendations
        comp_id = s(comp.get("comp_id") or "").strip()

        for sub_idx, sub in enumerate((comp.get("subsections") or []), start=1):
            stitle = s(sub.get("title")).strip()
            if stitle:
                # ✅ 5.1.1 / 5.1.2 ...
                if comp_id:
                    numbered_title = f"{comp_id}.{sub_idx} {stitle}"
                else:
                    numbered_title = f"{sub_idx}. {stitle}"

                h3(doc, numbered_title)
                doc.add_paragraph("")

            if isinstance(sub.get("major_table"), list):
                add_major_findings_table_tool6(
                    doc,
                    major_rows=sub.get("major_table") or [],
                    photo_bytes=photo_bytes or {},
                    photo_field_map=photo_field_map or {},
                )
                doc.add_paragraph("")
            else:
                for para in (sub.get("paragraphs") or []):
                    body(doc, para)
                doc.add_paragraph("")



# =========================
# MASTER BUILDER
# =========================
def build_tool6_full_report_docx(
    row: Dict[str, Any],
    cover_image_bytes: bytes,
    unicef_logo_path: Optional[str],
    act_logo_path: Optional[str],
    ppc_logo_path: Optional[str],
    general_info_overrides: Optional[Dict[str, Any]] = None,
    executive_summary_text: str = "",
    data_collection_text: str = "",
    work_progress_text: str = "",
    work_progress_findings: Optional[List[Dict[str, Any]]] = None,
    component_observations: Optional[List[Dict[str, Any]]] = None,
    findings_text: str = "",
    conclusion_text: str = "",
    photo_selections: Optional[Dict[str, str]] = None,
    photo_bytes: Optional[Dict[str, bytes]] = None,
    photo_field_map: Optional[Dict[str, str]] = None,
) -> bytes:
    general_info_overrides = general_info_overrides or {}
    component_observations = component_observations or []
    photo_selections = photo_selections or {}
    photo_bytes = photo_bytes or {}
    photo_field_map = photo_field_map or {}

    work_progress_findings = work_progress_findings or []

    doc = Document()

    set_doc_a4(doc)

    sec0 = doc.sections[0]
    sec0.top_margin = Inches(0.25)
    sec0.bottom_margin = Inches(0.25)
    sec0.left_margin = Inches(0.5)
    sec0.right_margin = Inches(0.5)
    try:
        sec0.header_distance = Inches(0.2)
        sec0.footer_distance = Inches(0.2)
    except Exception:
        pass

    apply_heading_rules(doc)

    add_cover_page(doc, row, cover_image_bytes, unicef_logo_path, act_logo_path, ppc_logo_path)
    add_toc_page(doc)
    add_general_project_information(doc, row, overrides=general_info_overrides)
    add_executive_summary(
        doc,
        row=row,
        executive_summary_text=executive_summary_text,
        component_observations=component_observations,
    )

    add_section_3_and_4_auto(
        doc,
        row=row,
        component_observations=component_observations,
        work_progress_findings=work_progress_findings,
    )

    # ✅ Section 5 restored (Component-wise Observations + Major findings tables)
    add_component_wise_key_observations_tool6(
        doc,
        component_observations=component_observations,
        photo_bytes=photo_bytes,
        photo_field_map=photo_field_map,
    )

    # Section 6 uses findings rows directly
    add_summary_of_findings_table(doc, work_progress_findings)

    # Section 7
    add_summary_and_conclusion(
        doc,
        row=row,
        conclusion_text=conclusion_text,
        component_observations=component_observations,
        work_progress_findings=work_progress_findings,
    )

    # Section 8
    add_photo_documentation(doc, photo_selections, photo_bytes, photo_field_map=photo_field_map)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue()
