# src/report_sections/work_progress_summary.py
from __future__ import annotations

from typing import List

from docx.document import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH

from src.report_sections._word_common import (
    s,
    # title (standard)
    add_section_title_h1,

    # paragraph
    tight_paragraph,

    # table
    set_table_fixed_layout,
    set_table_width_from_section,
    set_table_borders,
    shade_cell,
    set_cell_margins,
    set_row_cant_split,

    # helpers
    strip_heading_numbering,
    write_cell_text,
)

# -----------------
# Style (match report rule)
# -----------------
TITLE_FONT = "Cambria"
TITLE_SIZE = 16
TITLE_BLUE = None  # let add_section_title_h1 use provided color if you want
ORANGE_HEX = "ED7D31"

BODY_FONT = "Times New Roman"
BODY_SIZE = 11

HEADER_FILL_HEX = "D9E2F3"
BORDER_HEX = "A6A6A6"


def add_work_progress_summary_during_visit(
    doc: Document,
    *,
    activity_titles_from_section5: List[str],
    title_text: str = "4.    Work Progress Summary during the Visit.",
) -> None:
    """
    Section 4 table.

    IMPORTANT:
      - NO page break here (so it stays on same page as section 3).
      - Activity list comes from section 5 headings (titles only).
      - Planned/Achieved/Progress/Remarks left blank for manual entry.
    """

    # ✅ Title = Heading 1 (TOC-safe) + size 16 + orange underline
    add_section_title_h1(
        doc,
        s(title_text),
        font=TITLE_FONT,
        size=TITLE_SIZE,
        color=None,              # keep default (or set RGBColor(0,112,192) if you want blue)
        orange_hex=ORANGE_HEX,
        after_pt=6,
    )

    # Small spacing after title
    psp = doc.add_paragraph("")
    tight_paragraph(psp, align=WD_ALIGN_PARAGRAPH.LEFT, before_pt=0, after_pt=6, line_spacing=1)

    # ---- Table
    table = doc.add_table(rows=1, cols=6)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.style = "Table Grid"

    set_table_fixed_layout(table)
    set_table_borders(table, color_hex=BORDER_HEX)

    # ✅ Make table match actual usable width of current section
    # (prevents overflow when margins differ across sections)
    set_table_width_from_section(table, doc, section_index=len(doc.sections) - 1)

    headers = ["No.", "Activities", "Planned", "Achieved", "Progress", "Remarks"]

    # Column widths as percentages via relative sizing:
    # We'll set widths by proportion using the section width.
    # (write_cell_text + fixed layout will keep things stable)
    # proportions sum to 1.0
    proportions = [0.07, 0.33, 0.15, 0.15, 0.12, 0.18]

    # Apply column widths in DXA-ish through python-docx cell widths:
    # We rely on set_table_width_from_section + fixed layout for stability,
    # so we just keep columns consistent by setting widths on the first row.
    # (No Inches hardcoding => safer across margin changes.)
    hdr = table.rows[0]
    set_row_cant_split(hdr, cant_split=True)

    for i, cell in enumerate(hdr.cells):
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        shade_cell(cell, HEADER_FILL_HEX)
        set_cell_margins(cell, top_dxa=80, bottom_dxa=80, left_dxa=120, right_dxa=120)

        write_cell_text(
            cell,
            headers[i],
            font=BODY_FONT,
            size=BODY_SIZE,
            bold=True,
            align=WD_ALIGN_PARAGRAPH.CENTER,
        )

    # ---- Body rows
    clean_acts = [
        strip_heading_numbering(s(t)).strip()
        for t in (activity_titles_from_section5 or [])
        if s(t).strip()
    ]

    # If empty, still create a few blank rows for manual entry
    if not clean_acts:
        clean_acts = ["", "", ""]

    for idx, act in enumerate(clean_acts, start=1):
        r = table.add_row()

        # ✅ Allow split to avoid awkward whitespace if a row grows
        set_row_cant_split(r, cant_split=False)

        cells = r.cells
        for c in cells:
            c.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            set_cell_margins(c, top_dxa=80, bottom_dxa=80, left_dxa=120, right_dxa=120)

        write_cell_text(cells[0], str(idx), font=BODY_FONT, size=BODY_SIZE, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT)
        write_cell_text(cells[1], act, font=BODY_FONT, size=BODY_SIZE, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT)

        # Leave blanks for manual fill
        for j in [2, 3, 4, 5]:
            write_cell_text(cells[j], "", font=BODY_FONT, size=BODY_SIZE, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT)
