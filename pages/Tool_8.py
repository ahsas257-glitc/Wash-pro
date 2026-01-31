# pages/Tool_8.py
from __future__ import annotations

import os
import re
from io import BytesIO
from typing import Optional, Dict, Tuple, List
from uuid import uuid4
from urllib.parse import urlparse

import streamlit as st
import numpy as np
import pandas as pd
import requests
from PIL import Image

# ============================================================
# Internal project imports (DO NOT REMOVE)
# ============================================================
from src.config import GOOGLE_SHEET_ID, TPM_COL
from src.data_processing import get_row_cached
from src.report_builder import build_tool6_full_report_docx
from src.ui.wizard import Wizard, WizardConfig
from src.integrations.surveycto_client import (
    surveycto_login_ui,
    load_auth_state,
    is_logged_in,
    surveycto_request,
)

# ✅ base_tool_ui only (no sticky_open/sticky_close)
from design.components.base_tool_ui import (
    topbar,
    card_open,
    card_close,
    kv,
    status_card,
    table_card_open,
    table_card_close,
)
from design.components.global_ui import apply_global_background

# ============================================================
# Page config (ONLY ONCE)
# ============================================================
st.set_page_config(page_title="Tool 8 — WASH Report Generator", layout="wide")
apply_global_background("assets/images/Logo_of_PPC.png")
project_root = os.path.dirname(os.path.dirname(__file__))

# ============================================================
# Optional: SurveyCTO SDK (kept optional)
# ============================================================
try:
    import pysurveycto  # pip install pysurveycto
    _HAS_PYSURVEYCTO = True
except Exception:
    pysurveycto = None
    _HAS_PYSURVEYCTO = False

# ============================================================
# Google Sheet Column Mapping (KEEP AS IS)
# ============================================================
COL = {
    "PROVINCE": "A01_Province",
    "DISTRICT": "A02_District",
    "VILLAGE": "Village",
    "GPS_LAT": "GPS_1-Latitude",
    "GPS_LON": "GPS_1-Longitude",
    "STARTTIME": "starttime",
    "ACTIVITY_NAME": "Activity_Name",
    "PROJECT_STATUS": "Project_Status",
    "DELAY_REASON": "B8_Reasons_for_delay",
    "PRIMARY_PARTNER": "Primary_Partner_Name",
    "MONITOR_NAME": "A07_Monitor_name",
    "MONITOR_EMAIL": "A12_Monitor_email",
    "RESP_NAME": "A08_Respondent_name",
    "RESP_SEX_LABEL": "A09_Respondent_sex",
    "RESP_PHONE": "A10_Respondent_phone",
    "RESP_EMAIL": "A11_Respondent_email",
    "CDC_CODE": "A23_CDC_code",
    "DONOR_NAME": "A24_Donor_name",
    "REPORT_NUMBER": "A25_Monitoring_report_number",
    "CURRENT_REPORT_DATE": "A20_Current_report_date",
    "VISIT_NUMBER": "A26_Visit_number",
}

def col(row: dict, key: str, default=""):
    return (row or {}).get(COL.get(key, ""), default)

def gps_points_from_row(row: dict) -> str:
    lat = str(col(row, "GPS_LAT", "") or "").strip()
    lon = str(col(row, "GPS_LON", "") or "").strip()
    return f"{lat}, {lon}".strip().strip(",")

# ============================================================
# Small helpers (SAFE & FAST)
# ============================================================
def safe_str(x) -> str:
    return "" if x is None else str(x)

def ensure_http(url: str) -> str:
    u = (url or "").strip()
    return u if u.startswith("http") else ""

def na_if_empty_ui(raw) -> str:
    s0 = safe_str(raw).strip()
    return s0 if s0 else "N/A"

def format_af_phone_ui(raw) -> str:
    s0 = re.sub(r"\D+", "", safe_str(raw))
    if not s0:
        return ""
    if s0.startswith("0"):
        s0 = s0[1:]
    if s0.startswith("93"):
        return f"+{s0}"
    return f"+93{s0}"

def enforce_single_cover(selections: Dict[str, str]) -> Dict[str, str]:
    covers = [u for u, p in selections.items() if p == "Cover Page"]
    if len(covers) <= 1:
        return selections
    keep = covers[-1]
    for u in covers[:-1]:
        selections[u] = "Not selected"
    selections[keep] = "Cover Page"
    return selections

# ============================================================
# Image safety helpers
# ============================================================
def _looks_like_html(data: bytes) -> bool:
    head = (data or b"")[:300].lower()
    return b"<html" in head or b"<!doctype" in head

def _to_clean_png_bytes(img_bytes: bytes) -> bytes:
    img = Image.open(BytesIO(img_bytes)).convert("RGB")
    out = BytesIO()
    img.save(out, format="PNG", optimize=True)
    return out.getvalue()

def cover_suitability(img: Image.Image):
    w, h = img.size
    ratio = w / max(h, 1)

    arr = np.asarray(img.convert("RGB")).astype(np.float32)
    gray = (0.299 * arr[:, :, 0] + 0.587 * arr[:, :, 1] + 0.114 * arr[:, :, 2])

    brightness = float(gray.mean())
    gx = np.diff(gray, axis=1)
    gy = np.diff(gray, axis=0)
    sharpness = float(np.var(gx) + np.var(gy))

    issues = []
    if w < 1200 or h < 700:
        issues.append("Low resolution (≥1200×700 recommended)")
    if ratio < 1.2:
        issues.append("Landscape photo recommended")
    if brightness < 60:
        issues.append("Too dark")
    if brightness > 200:
        issues.append("Too bright")
    if sharpness < 40:
        issues.append("May be blurry")

    return issues, {"w": w, "h": h, "ratio": ratio, "brightness": brightness, "sharpness": sharpness}

# ============================================================
# SurveyCTO helpers (SINGLE SOURCE OF AUTH)
# ============================================================
def _is_scto_view_attachment(url: str) -> bool:
    return "surveycto.com/view/submission-attachment" in (url or "").lower()

def _is_surveycto_url(url: str) -> bool:
    try:
        host = (urlparse(url).netloc or "").lower()
        return host.endswith("surveycto.com")
    except Exception:
        return False

def _url_to_scto_path(url: str) -> str:
    """
    Convert full SurveyCTO URL to path for surveycto_request.
    surveycto_request expects a *path* appended to BASE_URL inside surveycto_client.py.
    """
    p = urlparse(url)
    path = (p.path or "").lstrip("/")
    if p.query:
        path = f"{path}?{p.query}"
    return path

def get_scto_client():
    """
    Optional pysurveycto object. We keep it because some deployments work best
    for /view/submission-attachment via SDK.
    """
    if not _HAS_PYSURVEYCTO:
        return None

    load_auth_state()
    user = st.session_state.get("scto_username", "").strip()
    pwd = st.session_state.get("scto_password", "").strip()
    if not user or not pwd:
        return None

    try:
        # Server name is inside surveycto_client.py; keep consistent:
        return pysurveycto.SurveyCTOObject("act4performance", user, pwd)
    except Exception:
        return None

# ============================================================
# Cached media loaders (PERFORMANCE CRITICAL)
# ============================================================
@st.cache_data(show_spinner=False, ttl=3600)
def scto_get_attachment_bytes(url: str, username: str) -> Optional[bytes]:
    """
    Try via pysurveycto if available for submission-attachment links.
    """
    scto = get_scto_client()
    if scto is None:
        return None
    try:
        data = scto.get_attachment(url)
        if not data or _looks_like_html(data):
            return None
        return data
    except Exception:
        return None

def _plain_http_get(url: str, *, timeout: int) -> requests.Response:
    return requests.get(url, timeout=timeout, allow_redirects=True)

def _scto_http_get(url: str, *, timeout: int) -> requests.Response:
    """
    All SurveyCTO GET requests MUST go through surveycto_request,
    so 401/403 auto resets login state.
    """
    path = _url_to_scto_path(url)
    # surveycto_request builds BASE_URL + path internally
    return surveycto_request("GET", path, timeout=timeout)

@st.cache_data(show_spinner=False, ttl=3600)
def fetch_image_cached(url: str, username: str) -> Tuple[bool, Optional[bytes], str]:
    """
    - SurveyCTO url => surveycto_request (single auth source)
    - submission-attachment => try pysurveycto first, then fallback to surveycto_request
    - non SurveyCTO => normal requests.get
    """
    try:
        if not url or not url.startswith("http"):
            return False, None, "Invalid URL"

        if _is_surveycto_url(url):
            if _is_scto_view_attachment(url):
                b = scto_get_attachment_bytes(url, username)
                if b:
                    try:
                        return True, _to_clean_png_bytes(b), "OK"
                    except Exception:
                        return False, None, "Invalid/unsupported image data (SDK)"

            r = _scto_http_get(url, timeout=25)
            if r.status_code >= 400:
                return False, None, f"HTTP {r.status_code}"
            if _looks_like_html(r.content):
                return False, None, "HTML response (auth required)"
            try:
                return True, _to_clean_png_bytes(r.content), "OK"
            except Exception:
                return False, None, "Invalid/unsupported image data"

        # Non SurveyCTO
        r = _plain_http_get(url, timeout=25)
        if r.status_code >= 400:
            return False, None, f"HTTP {r.status_code}"
        if _looks_like_html(r.content):
            return False, None, "HTML response"
        try:
            return True, _to_clean_png_bytes(r.content), "OK"
        except Exception:
            return False, None, "Invalid/unsupported image data"

    except Exception as e:
        return False, None, str(e)

@st.cache_data(show_spinner=False, ttl=3600)
def fetch_audio_cached(url: str, username: str) -> Tuple[bool, Optional[bytes], str, str]:
    """
    - SurveyCTO url => surveycto_request
    - submission-attachment => try pysurveycto first
    - non SurveyCTO => normal requests.get
    """
    try:
        if not url or not url.startswith("http"):
            return False, None, "Invalid URL", ""

        if _is_surveycto_url(url):
            if _is_scto_view_attachment(url):
                b = scto_get_attachment_bytes(url, username)
                if b:
                    return True, b, "OK", "audio/aac"

            r = _scto_http_get(url, timeout=35)
            if r.status_code >= 400:
                return False, None, f"HTTP {r.status_code}", ""
            if _looks_like_html(r.content):
                return False, None, "HTML response", ""
            mime = (r.headers.get("Content-Type") or "audio/aac").split(";")[0]
            return True, r.content, "OK", mime

        # Non SurveyCTO
        r = _plain_http_get(url, timeout=35)
        if r.status_code >= 400:
            return False, None, f"HTTP {r.status_code}", ""
        if _looks_like_html(r.content):
            return False, None, "HTML response", ""
        mime = (r.headers.get("Content-Type") or "audio/aac").split(";")[0]
        return True, r.content, "OK", mime

    except Exception as e:
        return False, None, str(e), ""

def _cache_user_key() -> str:
    load_auth_state()
    return (st.session_state.get("scto_username") or "").strip() or "anon"

def fetch_image(url: str) -> Tuple[bool, Optional[bytes], str]:
    return fetch_image_cached(url, username=_cache_user_key())

def fetch_audio(url: str) -> Tuple[bool, Optional[bytes], str, str]:
    return fetch_audio_cached(url, username=_cache_user_key())

# ============================================================
# Media normalization (KEEP)
# ============================================================
def normalize_media_url(url: str) -> str:
    u = (url or "").strip()
    if not u.startswith("http"):
        return u

    m = re.search(r"drive\.google\.com\/file\/d\/([^\/]+)\/", u)
    if m:
        fid = m.group(1)
        return f"https://drive.google.com/uc?export=download&id={fid}"

    m = re.search(r"drive\.google\.com\/open\?id=([^&]+)", u)
    if m:
        fid = m.group(1)
        return f"https://drive.google.com/uc?export=download&id={fid}"

    return u

def extract_photo_links(row: dict) -> List[Dict[str, str]]:
    links: List[Dict[str, str]] = []
    for k, v in (row or {}).items():
        sv = safe_str(v).strip()
        if not sv.startswith("http"):
            continue

        low = sv.lower()
        is_img = any(ext in low for ext in [".jpg", ".jpeg", ".png", ".webp"])
        is_drive = ("drive.google.com" in low) or ("googleusercontent.com" in low)
        is_scto = "surveycto.com/view/submission-attachment" in low

        if is_img or is_drive or is_scto:
            links.append({"field": k, "url": normalize_media_url(sv)})

    uniq: List[Dict[str, str]] = []
    seen = set()
    for it in links:
        if it["url"] in seen:
            continue
        seen.add(it["url"])
        uniq.append(it)
    return uniq

def extract_audio_links(row: dict) -> List[Dict[str, str]]:
    links: List[Dict[str, str]] = []
    for k, v in (row or {}).items():
        sv = safe_str(v).strip()
        if not sv.startswith("http"):
            continue

        low = sv.lower()
        is_audio = any(ext in low for ext in [".aac", ".mp3", ".wav", ".m4a", ".ogg", ".opus"])
        is_drive = ("drive.google.com" in low) or ("googleusercontent.com" in low)
        is_scto = "surveycto.com/view/submission-attachment" in low

        if is_audio or (is_scto and ("audio" in k.lower() or "voice" in k.lower() or "record" in k.lower())):
            links.append({"field": k, "url": normalize_media_url(sv)})

    uniq: List[Dict[str, str]] = []
    seen = set()
    for it in links:
        if it["url"] in seen:
            continue
        seen.add(it["url"])
        uniq.append(it)
    return uniq

# ============================================================
# Wizard steps
# ============================================================
STEPS = [
    "1) Cover Photo",
    "2) General Info",
    "3) Photos (Findings / Observations)",
    "4) Audio",
    "5) Components",
    "6) Generate DOCX",
]

# ============================================================
# Session State Initialization (CRITICAL FOR PERFORMANCE)
# ============================================================
def init_tool_state():
    st.session_state.setdefault("general_info_overrides", {})

    st.session_state.setdefault("photo_selections", {})
    st.session_state.setdefault("photo_bytes", {})
    st.session_state.setdefault("photo_field", {})
    st.session_state.setdefault("cover_upload_bytes", None)

    st.session_state.setdefault("audio_selections", {})
    st.session_state.setdefault("audio_bytes", {})
    st.session_state.setdefault("audio_field", {})
    st.session_state.setdefault("audio_notes", {})
    st.session_state.setdefault("audio_mime", {})

    st.session_state.setdefault("components_list", [])
    st.session_state.setdefault("component_inputs", {})
    st.session_state.setdefault("component_observations", [])

    st.session_state.setdefault("executive_summary", "")
    st.session_state.setdefault("data_collection", "")
    st.session_state.setdefault("work_progress", "")
    st.session_state.setdefault("findings", "")
    st.session_state.setdefault("conclusion", "")

init_tool_state()

# ============================================================
# DOCX generation (central)
# ============================================================
def _resolve_cover_bytes() -> Optional[bytes]:
    selections = st.session_state.get("photo_selections", {})
    photo_bytes = st.session_state.get("photo_bytes", {})

    cover_urls = [u for u, p in selections.items() if p == "Cover Page"]
    cover_bytes = photo_bytes.get(cover_urls[0]) if cover_urls else None

    if cover_bytes is None and st.session_state.get("cover_upload_bytes") is not None:
        cover_bytes = st.session_state["cover_upload_bytes"]

    return cover_bytes

def _generate_docx(
    row: dict,
    *,
    unicef_logo_path: Optional[str],
    act_logo_path: Optional[str],
    ppc_logo_path: Optional[str],
    tpm_id: str,
) -> bool:
    cover_bytes = _resolve_cover_bytes()
    if cover_bytes is None:
        st.session_state["tool8_docx_bytes"] = None
        return False

    selections = st.session_state.get("photo_selections", {})
    photo_bytes = st.session_state.get("photo_bytes", {})

    docx_bytes = build_tool6_full_report_docx(
        row=row,
        cover_image_bytes=cover_bytes,
        unicef_logo_path=unicef_logo_path,
        act_logo_path=act_logo_path,
        ppc_logo_path=ppc_logo_path,
        general_info_overrides=st.session_state.get("general_info_overrides", {}),
        executive_summary_text=st.session_state.get("executive_summary", ""),
        data_collection_text=st.session_state.get("data_collection", ""),
        work_progress_text=st.session_state.get("work_progress", ""),
        component_observations=st.session_state.get("component_observations", []),
        findings_text=st.session_state.get("findings", ""),
        conclusion_text=st.session_state.get("conclusion", ""),
        photo_selections=selections,
        photo_bytes=photo_bytes,
        photo_field_map=st.session_state.get("photo_field", {}),
    )

    st.session_state["tool8_docx_bytes"] = docx_bytes
    return True

# ============================================================
# Wizard instance
# ============================================================
wiz = Wizard(WizardConfig(tool_name="Tool 8", steps=STEPS, key_prefix="tool8"))

# ============================================================
# Top bar
# ============================================================
topbar(
    title="Tool 8 — Report Generator",
    subtitle="Auto-filled from Google Sheet + SurveyCTO attachments",
    right_chip="WASH • UNICEF",
)

# ============================================================
# TPM selection (GUARD)
# ============================================================
tool_name = "Tool 8"
tpm_id = st.session_state.get("tpm_id")

if not tpm_id:
    st.warning("No TPM ID selected. Please go back to Home and select a TPM ID.")
    st.stop()

with st.container(border=True):
    st.info(f"Selected TPM ID: {tpm_id}")

# ============================================================
# Sidebar — SurveyCTO Login
# ============================================================
with st.sidebar:
    logged_in = surveycto_login_ui(in_sidebar=False)  # چون همینجا داخل sidebar هستیم

if not logged_in:
    st.warning("Please login via the sidebar to download images and audio from SurveyCTO")
    # stop نکن؛ شاید کاربر فقط Google Sheet را بخواهد.

# ============================================================
# Load Google Sheet data (ONCE)
# ============================================================
@st.cache_data(show_spinner=False, ttl=600)
def _load_tool8_row(sheet_id: str, tool: str, tpm_value: str, tpm_col: str):
    return get_row_cached(sheet_id, tool, tpm_id=tpm_value, tpm_col=tpm_col)

row = _load_tool8_row(GOOGLE_SHEET_ID, tool_name, tpm_id, TPM_COL)

if not row:
    st.error("The selected TPM ID was not found in the Tool 8 worksheet.")
    st.stop()

# ============================================================
# Logos (safe paths)
# ============================================================
def _safe_logo(path: str) -> Optional[str]:
    return path if path and os.path.exists(path) else None

unicef_logo_path = _safe_logo(os.path.join(project_root, "assets/images/Logo_of_UNICEF.png"))
act_logo_path = _safe_logo(os.path.join(project_root, "assets/images/Logo_of_ACT.png"))
ppc_logo_path = _safe_logo(os.path.join(project_root, "assets/images/Logo_of_PPC.png"))

# ============================================================
# Defaults (from Google Sheet)
# ============================================================
defaults = {
    "Province": col(row, "PROVINCE", ""),
    "District": col(row, "DISTRICT", ""),
    "Village / Community": col(row, "VILLAGE", ""),
    "GPS points": gps_points_from_row(row),
    "Project Name": col(row, "ACTIVITY_NAME", ""),
    "Date of Visit": safe_str(col(row, "STARTTIME", "")).split(" ")[0],

    "Name of the IP, Organization / NGO": col(row, "PRIMARY_PARTNER", ""),
    "Name of the monitor Engineer": col(row, "MONITOR_NAME", ""),
    "Email of the monitor engineer": col(row, "MONITOR_EMAIL", ""),

    "Name of the respondent (Participant / UNICEF / IPs)": col(row, "RESP_NAME", ""),
    "Sex of Respondent": col(row, "RESP_SEX_LABEL", ""),
    "Contact Number of the Respondent": format_af_phone_ui(col(row, "RESP_PHONE", "")),
    "Email Address of the Respondent": na_if_empty_ui(col(row, "RESP_EMAIL", "")),

    "Project Status": col(row, "PROJECT_STATUS", ""),
    "Reason for delay": na_if_empty_ui(col(row, "DELAY_REASON", "")),
    "CDC Code": col(row, "CDC_CODE", ""),
    "Donor Name": col(row, "DONOR_NAME", ""),

    "Monitoring Report Number": col(row, "REPORT_NUMBER", ""),
    "Date of Current Report": safe_str(col(row, "CURRENT_REPORT_DATE", "")).split(" ")[0],
    "Number of Sites Visited": col(row, "VISIT_NUMBER", ""),
}

hints = {
    "CDC Code": "Verify against official documentation.",
    "Monitoring Report Number": "Verify before final generation.",
    "Contact Number of the Respondent": "Auto-formatted to +93.",
    "Email Address of the Respondent": "If empty, DOCX shows N/A.",
    "Reason for delay": "If empty, DOCX shows N/A.",
}

if not st.session_state["general_info_overrides"]:
    st.session_state["general_info_overrides"] = {k: str(v) for k, v in defaults.items()}

# ============================================================
# Collect media links (FAST)
# ============================================================
photos = extract_photo_links(row)
audios = extract_audio_links(row)

photo_label_by_url: Dict[str, str] = {}
for i, p in enumerate(photos or [], start=1):
    u = ensure_http(p.get("url", ""))
    f = p.get("field", "Photo")
    if u:
        photo_label_by_url[u] = f"{i:02d}. {f}"

all_photo_urls = list(photo_label_by_url.keys())

# ============================================================
# Wizard header
# ============================================================
wiz.header()
step = wiz.step_idx

# -------------------------------------------------
# STEP 1 — Cover Photo
# -------------------------------------------------
if step == 0:
    box = st.container(border=True)
    with box:
        card_open("Cover Photo Selection", subtitle="Select ONE photo for the cover page", variant="lg-variant-cyan")

        if not all_photo_urls:
            status_card("No photos found", "No image URLs detected for this TPM ID.", level="error")
            card_close()

            clicked_back, clicked_right = wiz.nav(
                can_next=False,
                back_label="Back",
                next_label="Next",
                generate_label="Generate",
            )
            if clicked_back or clicked_right:
                st.rerun()
            st.stop()

        selected_url = st.selectbox(
            "Choose cover image",
            options=all_photo_urls,
            format_func=lambda u: photo_label_by_url.get(u, u),
        )
        st.session_state["photo_field"][selected_url] = photo_label_by_url.get(selected_url, "Cover Photo")

        ok, data, msg = fetch_image(selected_url)
        if ok and data:
            st.image(data, use_container_width=True)
            st.session_state["photo_bytes"][selected_url] = data
            st.session_state["photo_selections"][selected_url] = "Cover Page"
            st.session_state["photo_selections"] = enforce_single_cover(st.session_state["photo_selections"])

            try:
                issues, _meta = cover_suitability(Image.open(BytesIO(data)))
                if issues:
                    status_card("Cover selected (with warnings)", " • ".join(issues), level="warning")
                else:
                    status_card("Cover selected", "This image will be used as report cover.", level="success")
            except Exception:
                status_card("Cover selected", "This image will be used as report cover.", level="success")
        else:
            status_card("Failed to load cover", msg or "Unknown error.", level="error")

        st.markdown("### Optional: Upload Cover Manually")
        up = st.file_uploader("Upload cover image", type=["jpg", "jpeg", "png"], key="cover_upload")
        if up:
            b = up.read()
            try:
                b = _to_clean_png_bytes(b)
            except Exception:
                pass
            st.session_state["cover_upload_bytes"] = b
            st.image(b, use_container_width=True)
            status_card("Uploaded cover saved", "This will be used if online cover fails.", level="success")

        card_close()

    selections = st.session_state.get("photo_selections", {})
    photo_bytes = st.session_state.get("photo_bytes", {})
    cover_urls = [u for u, p in selections.items() if p == "Cover Page"]
    cover_ok = ((cover_urls and cover_urls[0] in photo_bytes) or (st.session_state.get("cover_upload_bytes") is not None))

    clicked_back, clicked_right = wiz.nav(
        can_next=cover_ok,
        back_label="Back",
        next_label="Next",
        generate_label="Generate",
    )
    if clicked_back or clicked_right:
        st.rerun()
    st.stop()

# -------------------------------------------------
# STEP 2 — General Info
# -------------------------------------------------
if step == 1:
    st.subheader("General Project Information")

    box = st.container(border=True)
    with box:
        card_open("Review & Edit Information", subtitle="Values are auto-filled from Google Sheet", variant="lg-variant-green")
        tabs = st.tabs(["Project", "Respondent", "Monitoring", "Status / Other"])

        def _input(field: str):
            cur = st.session_state["general_info_overrides"].get(field, defaults.get(field, ""))
            st.session_state["general_info_overrides"][field] = st.text_input(field, value=str(cur))
            hint = hints.get(field, "")
            if hint:
                st.caption(hint)

        with tabs[0]:
            for f in ["Province", "District", "Village / Community", "GPS points", "Project Name", "Date of Visit"]:
                _input(f)

        with tabs[1]:
            for f in ["Name of the respondent (Participant / UNICEF / IPs)", "Sex of Respondent", "Contact Number of the Respondent", "Email Address of the Respondent"]:
                _input(f)

        with tabs[2]:
            for f in ["Name of the IP, Organization / NGO", "Name of the monitor Engineer", "Email of the monitor engineer", "Monitoring Report Number", "Date of Current Report", "Number of Sites Visited"]:
                _input(f)

        with tabs[3]:
            for f in ["Project Status", "Reason for delay", "CDC Code", "Donor Name"]:
                _input(f)

        card_close()

    status_card("Information saved", "Edits are stored and will be used in the report.", level="success")

    clicked_back, clicked_right = wiz.nav(
        can_next=True,
        back_label="Back",
        next_label="Next",
        generate_label="Generate",
    )
    if clicked_back or clicked_right:
        st.rerun()
    st.stop()

# -------------------------------------------------
# STEP 3 — Photos (Findings / Observations)
# -------------------------------------------------
if step == 2:
    st.subheader("Project Photos")

    box = st.container(border=True)
    with box:
        card_open(
            "Photo Classification",
            subtitle="Assign photos to Findings or Observations (Cover is managed in Step 1)",
            variant="lg-variant-cyan",
        )

        if not all_photo_urls:
            status_card("No photos available", "No image URLs were detected.", level="info")
            card_close()

            clicked_back, clicked_right = wiz.nav(
                can_next=True,
                back_label="Back",
                next_label="Next",
                generate_label="Generate",
            )
            if clicked_back or clicked_right:
                st.rerun()
            st.stop()

        selected_url = st.selectbox(
            "Select a photo",
            options=all_photo_urls,
            format_func=lambda u: photo_label_by_url.get(u, u),
            key="step3_photo_pick",
        )

        current = st.session_state["photo_selections"].get(selected_url, "Not selected")
        if current == "Cover Page":
            current = "Not selected"

        purpose_options = ["Not selected", "Findings", "Observations"]
        purpose = st.selectbox(
            "Assign purpose",
            purpose_options,
            index=purpose_options.index(current) if current in purpose_options else 0,
            key="step3_purpose",
        )
        st.session_state["photo_selections"][selected_url] = purpose

        ok, data, msg = fetch_image(selected_url)
        if ok and data:
            st.image(data, use_container_width=True)
            st.session_state["photo_bytes"][selected_url] = data
        else:
            status_card("Photo load failed", msg or "Unknown error.", level="error")

        st.markdown("### Selected summary")
        any_sel = False
        for u, p0 in st.session_state["photo_selections"].items():
            if p0 in ("Findings", "Observations"):
                any_sel = True
                flag = "✅" if u in st.session_state["photo_bytes"] else "⚠️"
                st.write(f"{flag} **{p0}** — {photo_label_by_url.get(u, u)}")
        if not any_sel:
            st.caption("No photos assigned yet.")

        card_close()

    clicked_back, clicked_right = wiz.nav(
        can_next=True,
        back_label="Back",
        next_label="Next",
        generate_label="Generate",
    )
    if clicked_back or clicked_right:
        st.rerun()
    st.stop()

# -------------------------------------------------
# STEP 4 — Audio
# -------------------------------------------------
if step == 3:
    st.subheader("Project Audio")

    box = st.container(border=True)
    with box:
        card_open(
            "Audio Player",
            subtitle="Select audio, load, play, and store notes (optional).",
            variant="lg-variant-purple",
        )

        if not audios:
            status_card("No audio links found", "No audio URLs detected for this TPM ID.", level="info")
            card_close()

            clicked_back, clicked_right = wiz.nav(
                can_next=True,
                back_label="Back",
                next_label="Next",
                generate_label="Generate",
            )
            if clicked_back or clicked_right:
                st.rerun()
            st.stop()

        options = []
        for i, a in enumerate(audios, start=1):
            field = a.get("field", "Unknown field")
            url = ensure_http(a.get("url", ""))
            if url:
                options.append((f"{i:02d}. {field}", url, field))

        if not options:
            status_card("No valid audio URLs", "Audio fields exist but none contain valid http links.", level="warning")
            card_close()

            clicked_back, clicked_right = wiz.nav(
                can_next=True,
                back_label="Back",
                next_label="Next",
                generate_label="Generate",
            )
            if clicked_back or clicked_right:
                st.rerun()
            st.stop()

        labels = [x[0] for x in options]
        pick = st.selectbox("Select an audio", labels, index=0, key="step4_audio_pick")
        _, sel_url, sel_field = next(x for x in options if x[0] == pick)
        st.session_state["audio_field"][sel_url] = sel_field

        purpose_options = ["Not selected", "Evidence / Findings", "Observation", "Other"]
        cur_purpose = st.session_state["audio_selections"].get(sel_url, "Not selected")
        purpose = st.selectbox(
            "Assign purpose",
            purpose_options,
            index=purpose_options.index(cur_purpose) if cur_purpose in purpose_options else 0,
            key="step4_audio_purpose",
        )
        st.session_state["audio_selections"][sel_url] = purpose

        ok, data, msg, mime = fetch_audio(sel_url)
        if ok and data:
            st.session_state["audio_bytes"][sel_url] = data
            st.session_state["audio_mime"][sel_url] = (mime or "audio/aac")

            status_card("Audio loaded", "You can play it and store notes below.", level="success")
            st.audio(data, format=st.session_state["audio_mime"].get(sel_url, "audio/aac"), start_time=0)

            cur_note = st.session_state["audio_notes"].get(sel_url, "")
            note = st.text_area(
                "Notes / Transcript (optional)",
                value=cur_note,
                height=130,
                key="step4_audio_note",
                placeholder="Write key points you heard to use in reporting.",
            )
            st.session_state["audio_notes"][sel_url] = note
        else:
            status_card("Audio could not be loaded", f"Reason: {msg or 'Unknown error.'}", level="error")
            st.info("Login to SurveyCTO in sidebar. If already logged in, your account may not have permission.")

        card_close()

    clicked_back, clicked_right = wiz.nav(
        can_next=True,
        back_label="Back",
        next_label="Next",
        generate_label="Generate",
    )
    if clicked_back or clicked_right:
        st.rerun()
    st.stop()

# -------------------------------------------------
# STEP 5 — Components (Dynamic)
# -------------------------------------------------
DEFAULT_COMPONENT_TITLES = [
    ("Construction of bore well and well protection structure", "Recommendations for bore well / well protection"),
    ("Supply and installation of the solar system", "Recommendations for solar system"),
    ("Construction of 60 m3 reservoir", "Recommendations for construction of reservoir"),
    ("Construction of 5 m3 reservoir for School", "Recommendations for school reservoir"),
    ("Construction of boundary wall", "Recommendations for boundary wall"),
    ("Construction of guard room and latrine", "Recommendations for guard room and latrine"),
    ("Construction of Stand taps", "Recommendations for stand taps"),
]

def _reindex_components():
    comps = st.session_state.get("components_list", [])
    for i, c in enumerate(comps, start=1):
        c["comp_id"] = f"5.{i}"
    st.session_state["components_list"] = comps

def _remove_component(uid: str):
    st.session_state["components_list"] = [c for c in st.session_state.get("components_list", []) if c.get("uid") != uid]
    st.session_state.setdefault("component_inputs", {})
    st.session_state["component_inputs"].pop(uid, None)
    for k in [f"{uid}_photos", f"{uid}_obs", f"{uid}_reco", f"{uid}_major_table_editor"]:
        st.session_state.pop(k, None)
    _reindex_components()
    st.session_state["component_observations"] = []

if step == 4:
    if not st.session_state["components_list"]:
        st.session_state["components_list"] = [
            {"uid": str(uuid4()), "comp_id": "TEMP", "title": t, "reco_title": r}
            for (t, r) in DEFAULT_COMPONENT_TITLES
        ]
        _reindex_components()

    st.subheader("5. Project Component-Wise Key Observations")

    ctrl = st.container(border=True)
    with ctrl:
        c1, c2, c3 = st.columns([2, 2, 1])
        with c1:
            new_title = st.text_input("Add new component — Title", placeholder="e.g., Construction of ...", key="new_comp_title")
        with c2:
            new_reco = st.text_input("Recommendation Title", placeholder="e.g., Recommendations for ...", key="new_comp_reco")
        with c3:
            if st.button("Add", use_container_width=True, key="add_component_btn"):
                if new_title.strip():
                    st.session_state["components_list"].append({
                        "uid": str(uuid4()),
                        "comp_id": "TEMP",
                        "title": new_title.strip(),
                        "reco_title": (new_reco.strip() or "Recommendations:"),
                    })
                    _reindex_components()
                    st.rerun()
                else:
                    status_card("Title is required", "Please enter a component title to add it.", level="warning")

    box5 = st.container(border=True)
    with box5:
        card_open(
            "Complete components",
            subtitle="For each component: select up to 3 photos, write Observation, Major findings table, and Recommendations.",
            variant="lg-variant-green",
        )

        for comp in st.session_state["components_list"]:
            uid = comp["uid"]
            comp_id = comp["comp_id"]
            title = comp["title"]
            reco_title = comp["reco_title"]

            st.session_state["component_inputs"].setdefault(uid, {"photos": [], "observation": "", "major_table": [], "reco": ""})
            data = st.session_state["component_inputs"][uid]

            h1, h2 = st.columns([8, 2])
            with h1:
                st.markdown(f"### {comp_id}. {title}")
            with h2:
                if st.button("Remove", key=f"rm_{uid}", use_container_width=True):
                    _remove_component(uid)
                    st.rerun()

            with st.expander("Open / Edit", expanded=False):
                sel = st.multiselect(
                    "Select photos (max 3)",
                    options=all_photo_urls,
                    default=data.get("photos", []),
                    format_func=lambda u: photo_label_by_url.get(u, u),
                    key=f"{uid}_photos",
                )[:3]

                if sel:
                    cols = st.columns(len(sel))
                    for idx, u in enumerate(sel):
                        if u not in st.session_state["photo_bytes"]:
                            ok_img, b, _msg = fetch_image(u)
                            if ok_img and b:
                                st.session_state["photo_bytes"][u] = b
                        with cols[idx]:
                            if u in st.session_state["photo_bytes"]:
                                st.image(st.session_state["photo_bytes"][u], use_container_width=True)
                            else:
                                st.caption("⚠️ Not loaded. Try again or check permissions.")

                obs = st.text_area(
                    "Observation / Description",
                    value=data.get("observation", ""),
                    height=120,
                    key=f"{uid}_obs",
                )

                table_card_open("Major findings (table format)")
                st.caption("Max 10 rows • NO auto-numbered")

                default_rows = data.get("major_table") or [{"NO": 1, "Findings": "", "Compliance": "Yes", "Photo": ""}]
                df = pd.DataFrame(default_rows)
                for coln in ["NO", "Findings", "Compliance", "Photo"]:
                    if coln not in df.columns:
                        df[coln] = ""

                edited = st.data_editor(
                    df,
                    use_container_width=True,
                    num_rows="dynamic",
                    hide_index=True,
                    column_config={
                        "NO": st.column_config.NumberColumn("NO", min_value=1, step=1),
                        "Findings": st.column_config.TextColumn("Findings", width="large"),
                        "Compliance": st.column_config.SelectboxColumn("Compliance", options=["Yes", "No"]),
                        "Photo": st.column_config.SelectboxColumn(
                            "Photo",
                            options=[""] + all_photo_urls,
                            format_func=lambda u: photo_label_by_url.get(u, "—") if u else "—",
                        ),
                    },
                    key=f"{uid}_major_table_editor",
                )
                table_card_close()

                edited = edited.head(10).copy()
                edited["NO"] = list(range(1, len(edited) + 1))
                major_table = edited.to_dict(orient="records")

                reco = st.text_area(
                    reco_title + ":",
                    value=data.get("reco", ""),
                    height=100,
                    key=f"{uid}_reco",
                )

                st.session_state["component_inputs"][uid] = {
                    "photos": sel,
                    "observation": obs,
                    "major_table": major_table,
                    "reco": reco,
                }

        card_close()

    sections_out = []
    for comp in st.session_state["components_list"]:
        uid = comp["uid"]
        comp_id = comp["comp_id"]
        title = comp["title"]
        reco_title = comp["reco_title"]
        d = st.session_state["component_inputs"].get(uid, {})

        sections_out.append({
            "comp_id": comp_id,
            "title": f"{comp_id}. {title}:",
            "paragraphs": [d.get("observation", "").strip()] if d.get("observation", "").strip() else [],
            "subsections": [
                {"title": "Major findings:", "major_table": d.get("major_table", [])},
                {"title": reco_title + ":", "paragraphs": [d.get("reco", "").strip()] if d.get("reco", "").strip() else []},
            ],
            "photos": d.get("photos", []),
        })

    st.session_state["component_observations"] = sections_out

    clicked_back, clicked_right = wiz.nav(
        can_next=True,
        back_label="Back",
        next_label="Next",
        generate_label="Generate",
    )
    if clicked_back or clicked_right:
        st.rerun()
    st.stop()

# -------------------------------------------------
# STEP 6 — Generate DOCX
# -------------------------------------------------
if step == 5:
    st.subheader("Generate Report (DOCX)")

    box = st.container(border=True)
    with box:
        card_open(
            "Generate DOCX",
            subtitle="Go back to review steps, or press Generate to build the DOCX.",
            variant="lg-variant-cyan",
        )

        cover_bytes = _resolve_cover_bytes()
        can_generate = (cover_bytes is not None)

        if not can_generate:
            status_card("No usable cover photo", "Go back to Step 1 and select/upload a cover photo.", level="error")

        existing = st.session_state.get("tool8_docx_bytes")
        if existing:
            status_card("DOCX already generated", "You can download it below or generate again.", level="success")
            st.download_button(
                "Download Report (DOCX)",
                data=existing,
                file_name=f"Tool8_Report_{tpm_id}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

        card_close()

    def _on_generate():
        ok = _generate_docx(
            row=row,
            unicef_logo_path=unicef_logo_path,
            act_logo_path=act_logo_path,
            ppc_logo_path=ppc_logo_path,
            tpm_id=tpm_id,
        )
        if ok:
            st.success("Report generated. Scroll to download.")
        else:
            st.error("Failed: cover photo is missing.")
        st.rerun()

    clicked_back, clicked_right = wiz.nav(
        can_next=can_generate,
        back_label="Back",
        next_label="Next",
        generate_label="Generate",
        on_generate=_on_generate,
    )
    if clicked_back or clicked_right:
        st.rerun()
    st.stop()
