# src/report_sections/general_project_information.py
from __future__ import annotations

from typing import Any, Dict, Optional

from docx.document import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.shared import Inches, RGBColor, Mm

from src.report_sections._word_common import (
    s,

    # ✅ standard H1 title (size 16 + TOC + orange line)
    add_section_title_h1,

    # table/layout
    set_table_fixed_layout,
    set_table_width_exact,
    set_table_borders,
    shade_cell,
    set_cell_margins,
    set_row_cant_split,
    write_cell_text,
    set_table_columns_exact,
    set_repeat_table_header,

    # helpers
    na_if_empty,
    format_af_phone,
    normalize_email_or_na_strict,
    donor_upper_and_pipe,
    format_date_dd_mon_yyyy,
    three_option_checkbox_line,
    write_yes_no_checkboxes,
    write_two_option_checkboxes,
    add_available_documents_inner_table,
)

# -------------------------
# Section-specific style only
# -------------------------
TITLE_TEXT = "1.  General Project Information:"
TITLE_FONT = "Cambria"
TITLE_SIZE = 16  # ✅ MUST be 16
TITLE_BLUE = RGBColor(0, 112, 192)

BODY_FONT = "Times New Roman"
BODY_SIZE = 11

FIELD_FILL_HEX = "D9E2F3"
HEADER_FILL_HEX = "E7EEF9"
BORDER_HEX = "A6A6A6"

# ✅ ONLY change this number to move the middle divider line
FIELD_COL_WIDTH_IN = 2.30

# cell margins (DXA)
M_TOP = 80
M_BOTTOM = 80
M_LEFT = 120
M_RIGHT = 120

DATA_KEYS = {
    "PROVINCE": "A01_Province",
    "DISTRICT": "A02_District",
    "VILLAGE": "Village",
    "GPS_LAT": "GPS_1-Latitude",
    "GPS_LON": "GPS_1-Longitude",
    "STARTTIME": "starttime",
    "ACTIVITY_NAME": "Activity_Name",
    "PRIMARY_PARTNER": "Primary_Partner_Name",
    "MONITOR_NAME": "A07_Monitor_name",
    "MONITOR_EMAIL": "A12_Monitor_email",
    "RESPONDENT_NAME": "A08_Respondent_name",
    "RESPONDENT_PHONE": "A10_Respondent_phone",
    "RESPONDENT_EMAIL": "A11_Respondent_email",
    "RESPONDENT_SEX": "A09_Respondent_sex",
    "PROJECT_COST_LABEL": "A14_Estimated_cost_amount_label",
    "EST_COST_AMOUNT": "Estimated_Project_Cost_amount",
    "CONTRACT_COST_AMOUNT": "Contracted_Project_Cost_amount",
    "PROJECT_STATUS_LABEL": "Project_Status",
    "PROJECT_PROGRESS_LABEL": "Project_progress",
    "START_DATE": "A15_Contract_start_date",
    "END_DATE": "A16_Contract_end_date",
    "PREV_PROGRESS": "A17_Previous_physical_progress",
    "CURR_PROGRESS": "A18_Current_physical_progress",
    "DONOR_NAME": "A24_Donor_name",
    "MONITORING_REPORT_NO": "A25_Monitoring_report_number",
    "CURRENT_REPORT_DATE": "A20_Current_report_date",
    "PREV_REPORT_DATE": "A21_Last_report_date",
    "VISIT_NO": "A26_Visit_number",
    "DOC_CONTRACT": "D1_contract_available",
    "DOC_JOURNAL": "D1_journal_available",
    "DOC_BOQ": "D2_boq_available",
    "DOC_DRAWINGS": "D2_drawings_available",
    "DOC_SITE_ENGINEER": "D3_site_engineer_available",
    "DOC_GEOPHYSICAL": "D3_geophysical_tests_available",
    "DOC_WQ_TEST": "D4_water_quality_tests_available",
    "DOC_PUMP_TEST": "D4_pump_test_results_available",
    "COMMUNITY_AGREEMENT": "community_agreement",
    "WORK_SAFETY": "work_safety_considered",
    "ENV_RISK": "environmental_risk",
}


def _set_a4_narrow(section) -> None:
    """A4 + Word 'Narrow' margins (0.5 inch = 12.7mm)."""
    section.page_width = Mm(210)
    section.page_height = Mm(297)

    section.top_margin = Mm(12.7)
    section.bottom_margin = Mm(12.7)
    section.left_margin = Mm(12.7)
    section.right_margin = Mm(12.7)

    section.header_distance = Mm(5)
    section.footer_distance = Mm(5)


def _usable_width_inches(section) -> float:
    usable_emu = section.page_width.emu - section.left_margin.emu - section.right_margin.emu
    return float(usable_emu) / 914400.0


def add_general_project_information(
    doc: Document,
    row: Dict[str, Any],
    overrides: Optional[Dict[str, Any]] = None,
    respondent_sex_val: Any = None,
) -> None:
    overrides = overrides or {}
    row = row or {}

    # ✅ section break so Narrow margins apply only to this section (table can span pages)
    doc.add_section(WD_SECTION.NEW_PAGE)
    section = doc.sections[-1]
    _set_a4_narrow(section)

    # ✅ Heading 1, size 16, TOC friendly + orange line
    add_section_title_h1(
        doc,
        TITLE_TEXT,
        font=TITLE_FONT,
        size=TITLE_SIZE,
        color=TITLE_BLUE,
        orange_hex="ED7D31",
        after_pt=6,
    )

    # ---------- Extract values ----------
    province = s(overrides.get("Province", row.get(DATA_KEYS["PROVINCE"])))
    district = s(overrides.get("District", row.get(DATA_KEYS["DISTRICT"])))
    village = s(overrides.get("Village / Community", row.get(DATA_KEYS["VILLAGE"])))

    gps_lat = s(row.get(DATA_KEYS["GPS_LAT"]))
    gps_lon = s(row.get(DATA_KEYS["GPS_LON"]))

    project_name = s(overrides.get("Project Name", row.get(DATA_KEYS["ACTIVITY_NAME"])))
    visit_date = s(overrides.get("Date of Visit", format_date_dd_mon_yyyy(row.get(DATA_KEYS["STARTTIME"]))))
    ip_name = s(overrides.get("Name of the IP, Organization / NGO", row.get(DATA_KEYS["PRIMARY_PARTNER"])))

    monitor_name = s(overrides.get("Name of the monitor Engineer", row.get(DATA_KEYS["MONITOR_NAME"])))
    monitor_email = s(overrides.get("Email of the monitor engineer", row.get(DATA_KEYS["MONITOR_EMAIL"])))

    respondent_name = s(overrides.get("Name of the respondent (Participant / UNICEF / IPs)", row.get(DATA_KEYS["RESPONDENT_NAME"])))
    respondent_phone = s(overrides.get("Contact Number of the Respondent", format_af_phone(row.get(DATA_KEYS["RESPONDENT_PHONE"]))))
    respondent_email = s(overrides.get("Email Address of the Respondent", normalize_email_or_na_strict(row.get(DATA_KEYS["RESPONDENT_EMAIL"]))))

    cost_label = s(row.get(DATA_KEYS["PROJECT_COST_LABEL"])).lower()
    estimated_amount = s(row.get(DATA_KEYS["EST_COST_AMOUNT"]))
    contracted_amount = s(row.get(DATA_KEYS["CONTRACT_COST_AMOUNT"]))
    estimated_cost = estimated_amount if "estimated" in cost_label else ""
    contracted_cost = contracted_amount if "contract" in cost_label else ""

    project_status_val = overrides.get("Project Status", row.get(DATA_KEYS["PROJECT_STATUS_LABEL"]))
    project_status = three_option_checkbox_line(project_status_val, "Ongoing", "Completed", "Suspended")

    reason_delay = s(overrides.get("Reason for delay", na_if_empty(row.get("B8_Reasons_for_delay"))))

    progress_val = overrides.get("Project progress", row.get(DATA_KEYS["PROJECT_PROGRESS_LABEL"]))
    project_progress = three_option_checkbox_line(progress_val, "Ahead of Schedule", "On Schedule", "Running behind")

    contract_start = s(overrides.get("Contract Start Date", format_date_dd_mon_yyyy(row.get(DATA_KEYS["START_DATE"]))))
    contract_end = s(overrides.get("Contract End Date", format_date_dd_mon_yyyy(row.get(DATA_KEYS["END_DATE"]))))

    prev_phys = s(overrides.get("Previous Physical Progress (%)", row.get(DATA_KEYS["PREV_PROGRESS"])))
    curr_phys = s(overrides.get("Current Physical Progress (%)", row.get(DATA_KEYS["CURR_PROGRESS"])))

    cdc_code = s(overrides.get("CDC Code", row.get("A23_CDC_code", "")))
    donor_name = donor_upper_and_pipe(overrides.get("Donor Name", row.get(DATA_KEYS["DONOR_NAME"])))

    monitoring_report_no = s(overrides.get("Monitoring Report Number", row.get(DATA_KEYS["MONITORING_REPORT_NO"])))
    current_report_date = s(overrides.get("Date of Current Report", format_date_dd_mon_yyyy(row.get(DATA_KEYS["CURRENT_REPORT_DATE"]))))

    last_report_date = s(overrides.get("Date of Last Monitoring Report", format_date_dd_mon_yyyy(row.get(DATA_KEYS["PREV_REPORT_DATE"]))))
    sites_visited = s(overrides.get("Number of Sites Visited", row.get(DATA_KEYS["VISIT_NO"])))

    community_agreement = overrides.get(
        "community agreement - Is the community/user group agreed on the well site?",
        row.get(DATA_KEYS["COMMUNITY_AGREEMENT"]) or row.get("Community_agreement"),
    )
    work_safety = overrides.get("Is work_safety_considered -", row.get(DATA_KEYS["WORK_SAFETY"]))
    env_risk = overrides.get("environmental risk -", row.get(DATA_KEYS["ENV_RISK"]))

    # ---------- Main table ----------
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.style = "Table Grid"

    set_table_fixed_layout(table)
    set_table_borders(table, color_hex=BORDER_HEX)

    usable_w_in = _usable_width_inches(section)
    set_table_width_exact(table, Inches(usable_w_in))

    field_in = min(float(FIELD_COL_WIDTH_IN), max(1.0, usable_w_in - 1.0))
    detail_in = max(0.75, usable_w_in - field_in)

    field_w = Inches(field_in)
    detail_w = Inches(detail_in)

    # lock widths
    set_table_columns_exact(table, [field_in, detail_in])
    table.columns[0].width = field_w
    table.columns[1].width = detail_w

    # Header row
    hdr = table.rows[0]
    set_row_cant_split(hdr, cant_split=True)     # ✅ keep header intact
    set_repeat_table_header(hdr)                 # ✅ repeats on page 2

    for i, txt in enumerate(("Field", "Details")):
        cell = hdr.cells[i]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        shade_cell(cell, HEADER_FILL_HEX)
        set_cell_margins(cell, M_TOP, M_BOTTOM, M_LEFT, M_RIGHT)
        write_cell_text(cell, txt, font=BODY_FONT, size=BODY_SIZE, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)

    def _lock_widths_again() -> None:
        set_table_columns_exact(table, [field_in, detail_in])

    def add_row(field: str, value: Any) -> None:
        r = table.add_row()

        # ✅ IMPORTANT FIX:
        # Allow Word to split rows if needed to avoid excessive white space and early page breaks
        set_row_cant_split(r, cant_split=False)

        c0, c1 = r.cells[0], r.cells[1]
        c0.width = field_w
        c1.width = detail_w

        c0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        shade_cell(c0, FIELD_FILL_HEX)
        set_cell_margins(c0, M_TOP, M_BOTTOM, M_LEFT, M_RIGHT)
        write_cell_text(c0, field, font=BODY_FONT, size=BODY_SIZE, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)

        c1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_margins(c1, M_TOP, M_BOTTOM, M_LEFT, M_RIGHT)
        write_cell_text(c1, s(value), font=BODY_FONT, size=BODY_SIZE, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT)

        _lock_widths_again()

    def add_row_custom(field: str, renderer) -> None:
        r = table.add_row()

        # ✅ IMPORTANT FIX:
        # Allow split for potentially tall custom rows (inner tables / checkboxes)
        set_row_cant_split(r, cant_split=False)

        c0, c1 = r.cells[0], r.cells[1]
        c0.width = field_w
        c1.width = detail_w

        c0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        shade_cell(c0, FIELD_FILL_HEX)
        set_cell_margins(c0, M_TOP, M_BOTTOM, M_LEFT, M_RIGHT)
        write_cell_text(c0, field, font=BODY_FONT, size=BODY_SIZE, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)

        c1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_margins(c1, M_TOP, M_BOTTOM, M_LEFT, M_RIGHT)
        c1.text = ""
        renderer(c1)

        _lock_widths_again()

    # Rows
    add_row("Province", province)
    add_row("District", district)
    add_row("Village / Community", village)
    add_row("GPS points", f"{gps_lat}, {gps_lon}".strip().strip(","))
    add_row("Project Name", project_name)
    add_row("Date of Visit", visit_date)
    add_row("Name of the IP, Organization / NGO", ip_name)
    add_row("Name of the monitor Engineer", monitor_name)
    add_row("Email of the monitor engineer", monitor_email)
    add_row("Name of the respondent (Participant / UNICEF / IPs)", respondent_name)

    if respondent_sex_val is None:
        respondent_sex_val = overrides.get("Sex of Respondent", row.get(DATA_KEYS["RESPONDENT_SEX"]))

    add_row_custom(
        "Sex of Respondent",
        lambda cell: write_two_option_checkboxes(cell, respondent_sex_val, "Male", "Female", font_size=BODY_SIZE),
    )

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
    add_row("Monitoring Report Number", monitoring_report_no)
    add_row("Date of Current Report", current_report_date)
    add_row("Date of Last Monitoring Report", na_if_empty(last_report_date))
    add_row("Number of Sites Visited", sites_visited)

    add_row_custom("Available documents in the site", lambda cell: add_available_documents_inner_table(cell, row, DATA_KEYS))

    add_row_custom(
        "community agreement - Is the community/user group agreed on the well site?",
        lambda cell: write_yes_no_checkboxes(cell, community_agreement, font_size=BODY_SIZE, align=WD_ALIGN_PARAGRAPH.LEFT),
    )
    add_row_custom(
        "Is work_safety_considered -",
        lambda cell: write_yes_no_checkboxes(cell, work_safety, font_size=BODY_SIZE, align=WD_ALIGN_PARAGRAPH.LEFT),
    )
    add_row_custom(
        "environmental risk -",
        lambda cell: write_yes_no_checkboxes(cell, env_risk, font_size=BODY_SIZE, align=WD_ALIGN_PARAGRAPH.LEFT),
    )
