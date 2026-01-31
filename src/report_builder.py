# src/report_builder.py
import io
import os
import inspect
from typing import Optional, Dict, Any, List, Callable, Tuple

from docx import Document

from src.report_sections._hf import apply_header_footer
from src.report_sections.cover_page import add_cover_page
from src.report_sections.toc_page import add_toc_page

from src.report_sections._word_common import (
    set_page_a4,
    strip_heading_numbering,
    update_docx_fields_bytes,
)

from src.report_sections.general_project_information import add_general_project_information
from src.report_sections.executive_summary import add_executive_summary
from src.report_sections.data_collection_methods import add_data_collection_methods
from src.report_sections.work_progress_summary import add_work_progress_summary_during_visit
from src.report_sections.component_wise_key_observations import add_component_wise_key_observations_tool6
from src.report_sections.summary_of_findings import add_summary_of_findings_section6
from src.report_sections.conclusion import add_conclusion_section


# ============================================================
# Utils (fast)
# ============================================================
_ASSET_DIR_CACHE: Optional[Tuple[str, str]] = None
_ASSET_FILE_INDEX: Optional[Dict[str, str]] = None


def _doc_to_bytes(doc: Document) -> bytes:
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _project_root() -> str:
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def _asset_dirs() -> Tuple[str, str]:
    global _ASSET_DIR_CACHE
    if _ASSET_DIR_CACHE is not None:
        return _ASSET_DIR_CACHE

    root = _project_root()
    _ASSET_DIR_CACHE = (
        os.path.join(root, "assets", "images"),
        os.path.join(root, "assets", "Images"),
    )
    return _ASSET_DIR_CACHE


def _build_asset_index() -> Dict[str, str]:
    global _ASSET_FILE_INDEX
    if _ASSET_FILE_INDEX is not None:
        return _ASSET_FILE_INDEX

    idx: Dict[str, str] = {}
    for folder in _asset_dirs():
        if not os.path.isdir(folder):
            continue
        try:
            for f in os.listdir(folder):
                p = os.path.join(folder, f)
                if os.path.isfile(p):
                    idx[f.lower()] = p
        except Exception:
            # ignore folder read issues
            pass

    _ASSET_FILE_INDEX = idx
    return idx


def _assets_image_path(filename: str) -> Optional[str]:
    if not filename:
        return None
    return _build_asset_index().get(filename.lower())


def _extract_section5_activity_titles(component_observations: Optional[List[Dict[str, Any]]]) -> List[str]:
    component_observations = component_observations or []
    titles: List[str] = []

    for comp in component_observations:
        if not isinstance(comp, dict):
            continue
        raw = (comp.get("title") or "").strip() or (comp.get("comp_id") or "").strip()
        if raw:
            clean = strip_heading_numbering(raw)
            if clean:
                titles.append(clean)

    return titles


# ============================================================
# Compatibility caller (fast + safe)
# ============================================================
def _call_compat(fn: Callable, *args, **kwargs):
    """
    Fast robust caller:
      1) Try direct call
      2) If TypeError:
         - trim extra positional args
         - filter kwargs based on signature
         - retry
    """
    try:
        return fn(*args, **kwargs)
    except TypeError:
        pass

    trimmed_args = args
    accepted_kwargs: Dict[str, Any] = {}

    try:
        sig = inspect.signature(fn)
        params = sig.parameters

        # positional trimming
        max_positional = 0
        has_var_positional = False
        for p in params.values():
            if p.kind == inspect.Parameter.VAR_POSITIONAL:
                has_var_positional = True
                break
            if p.kind in (inspect.Parameter.POSITIONAL_ONLY, inspect.Parameter.POSITIONAL_OR_KEYWORD):
                max_positional += 1

        if not has_var_positional and len(args) > max_positional:
            trimmed_args = args[:max_positional]

        # kwargs filtering
        has_var_kw = any(p.kind == inspect.Parameter.VAR_KEYWORD for p in params.values())
        if has_var_kw:
            accepted_kwargs = dict(kwargs)
        else:
            accepted = set(params.keys())
            accepted_kwargs = {k: v for k, v in kwargs.items() if k in accepted}

    except Exception:
        trimmed_args = args
        accepted_kwargs = dict(kwargs)

    try:
        return fn(*trimmed_args, **accepted_kwargs)
    except TypeError:
        # last attempt: positional only
        return fn(*trimmed_args)


# ============================================================
# Base doc builder
# ============================================================
def _build_base_doc(
    *,
    row: Dict[str, Any],
    cover_image_bytes: Optional[bytes] = None,
    general_info_overrides: Optional[dict] = None,
    reserved_mm: int = 165,
    unicef_logo_filename: str = "Logo_of_UNICEF.png",
    act_logo_filename: str = "Logo_of_ACT.png",
    ppc_logo_filename: str = "Logo_of_PPC.png",
    toc_levels: str = "1-3",
) -> Document:
    """
    Base doc:
      - A4 settings
      - header/footer on all pages
      - cover page (page 1)
      - TOC page (page 2)
    """
    doc = Document()
    set_page_a4(doc.sections[0])

    # resolve logos (fast cached)
    unicef_logo_path = _assets_image_path(unicef_logo_filename)
    act_logo_path = _assets_image_path(act_logo_filename)
    ppc_logo_path = _assets_image_path(ppc_logo_filename)

    # Header/Footer for ALL sections
    apply_header_footer(
        doc,
        unicef_logo_path=unicef_logo_path,
        act_logo_path=act_logo_path,
        ppc_logo_path=ppc_logo_path,
    )

    # Cover page (page 1)
    _call_compat(
        add_cover_page,
        doc,
        row,
        cover_image_bytes,
        general_info_overrides=general_info_overrides,
        reserved_mm=reserved_mm,
    )

    # TOC page MUST be page 2
    add_toc_page(
        doc,
        toc_levels=toc_levels,
        include_hyperlinks=True,
        hide_page_numbers_in_web_layout=False,
    )

    return doc


# ============================================================
# Public APIs
# ============================================================
def build_any_tool_report_docx_bytes(
    *,
    row: Dict[str, Any],
    cover_image_bytes: Optional[bytes] = None,
    general_info_overrides: Optional[dict] = None,
    reserved_mm: int = 165,
    unicef_logo_filename: str = "Logo_of_UNICEF.png",
    act_logo_filename: str = "Logo_of_ACT.png",
    ppc_logo_filename: str = "Logo_of_PPC.png",
) -> bytes:
    """
    Generic builder (cover + TOC only).
    """
    doc = _build_base_doc(
        row=row,
        cover_image_bytes=cover_image_bytes,
        general_info_overrides=general_info_overrides,
        reserved_mm=reserved_mm,
        unicef_logo_filename=unicef_logo_filename,
        act_logo_filename=act_logo_filename,
        ppc_logo_filename=ppc_logo_filename,
        toc_levels="1-3",
    )

    docx_bytes = _doc_to_bytes(doc)
    # update fields (TOC, page numbers, etc.)
    try:
        docx_bytes = update_docx_fields_bytes(docx_bytes)
    except Exception:
        pass
    return docx_bytes


def build_tool6_full_report_docx(
    *,
    row: Dict[str, Any],
    cover_image_bytes: Optional[bytes] = None,
    general_info_overrides: Optional[dict] = None,
    component_observations: Optional[List[Dict[str, Any]]] = None,
    photo_bytes: Optional[Dict[str, bytes]] = None,
    photo_field_map: Optional[Dict[str, str]] = None,
    reserved_mm: int = 165,
    unicef_logo_filename: str = "Logo_of_UNICEF.png",
    act_logo_filename: str = "Logo_of_ACT.png",
    ppc_logo_filename: str = "Logo_of_PPC.png",
    severity_by_no: Optional[Dict[int, str]] = None,
    severity_by_finding: Optional[Dict[str, str]] = None,
    add_legend: bool = True,
    **kwargs,
) -> bytes:
    """
    Tool 6 full report builder.

    Notes:
      - Section 4 MUST NOT add a page break (it continues from section 3).
      - Severity mapping is passed to Section 6.
      - Field updates happen at the end.
    """
    _ = kwargs  # pipeline compatibility

    general_info_overrides = general_info_overrides or {}
    component_observations = component_observations or []
    photo_bytes = photo_bytes or {}
    photo_field_map = photo_field_map or {}
    severity_by_no = severity_by_no or {}
    severity_by_finding = severity_by_finding or {}

    # Base (Cover + TOC)
    doc = _build_base_doc(
        row=row,
        cover_image_bytes=cover_image_bytes,
        general_info_overrides=general_info_overrides,
        reserved_mm=reserved_mm,
        unicef_logo_filename=unicef_logo_filename,
        act_logo_filename=act_logo_filename,
        ppc_logo_filename=ppc_logo_filename,
        toc_levels="1-3",
    )

    # Section 1
    _call_compat(
        add_general_project_information,
        doc,
        row=row,
        overrides=general_info_overrides,
        respondent_sex_val=general_info_overrides.get("Sex of Respondent"),
    )

    # Section 2
    _call_compat(
        add_executive_summary,
        doc,
        row=row,
        overrides=general_info_overrides,
    )

    # Section 3
    _call_compat(
        add_data_collection_methods,
        doc,
        row=row,
        overrides=general_info_overrides,
    )

    # Section 4 (activities from section 5 titles)
    section5_titles = _extract_section5_activity_titles(component_observations)
    _call_compat(
        add_work_progress_summary_during_visit,
        doc,
        activity_titles_from_section5=section5_titles,
        title_text="4.    Work Progress Summary during the Visit.",
    )

    # Section 5
    _call_compat(
        add_component_wise_key_observations_tool6,
        doc,
        component_observations=component_observations,
        photo_bytes=photo_bytes,
        photo_field_map=photo_field_map,
    )

    # Section 6 (pass severity mappings)
    _call_compat(
        add_summary_of_findings_section6,
        doc,
        component_observations=component_observations,
        severity_by_no=severity_by_no,
        severity_by_finding=severity_by_finding,
        add_legend=bool(add_legend),
    )

    # Section 7 (Conclusion)
    _call_compat(
        add_conclusion_section,
        doc,
        conclusion_text=None,          # default text will appear
        key_points=None,
        recommendations_summary=None,
        section_no="7",
    )

    docx_bytes = _doc_to_bytes(doc)

    # update fields (TOC, page numbers, etc.)
    try:
        docx_bytes = update_docx_fields_bytes(docx_bytes)
    except Exception:
        pass

    return docx_bytes
