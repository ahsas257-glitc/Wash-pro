# src/report_sections/toc_page.py
from __future__ import annotations

from docx.document import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import RGBColor

from src.report_sections._word_common import (
    title_with_orange_line,   # ✅ NOT a heading => TOC page title will NOT appear in TOC
    tight_paragraph,
    set_run,
    add_toc_field,
)

TITLE_TEXT = "Table of Contents"
TITLE_FONT = "Cambria"
TITLE_SIZE = 16
TITLE_BLUE = RGBColor(0, 112, 192)

BODY_FONT = "Times New Roman"
BODY_SIZE = 11


def add_toc_page(
    doc: Document,
    *,
    toc_levels: str = "1-3",
    include_hyperlinks: bool = True,
    hide_page_numbers_in_web_layout: bool = False,
) -> None:
    """
    Adds a TOC page using a Word field code.
    NOTE: This page title is NOT a Heading, so it won't appear in TOC.
    """
    doc.add_page_break()

    title_with_orange_line(
        doc,
        text=TITLE_TEXT,
        font=TITLE_FONT,
        size=TITLE_SIZE,
        color=TITLE_BLUE,
        orange_hex="ED7D31",
        after_pt=10,
    )

    p = doc.add_paragraph()
    tight_paragraph(p, align=WD_ALIGN_PARAGRAPH.LEFT, before_pt=0, after_pt=8, line_spacing=1)
    set_run(
        p.add_run("Right-click the table and select “Update Field” in Word to refresh page numbers."),
        BODY_FONT,
        BODY_SIZE,
        bold=False,
    )

    add_toc_field(
        doc,
        levels=toc_levels,
        hyperlinks=include_hyperlinks,
        hide_page_numbers_in_web_layout=hide_page_numbers_in_web_layout,
    )
