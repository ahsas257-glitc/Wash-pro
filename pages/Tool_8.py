# pages/Tool_8.py
import os
import re
from io import BytesIO
from typing import Optional, Dict, Tuple, List

import streamlit as st
from PIL import Image
import numpy as np
import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
from uuid import uuid4

from src.config import GOOGLE_SHEET_ID, TPM_COL
from src.data_processing import open_worksheet, get_all_records, find_by_tpm_id
from src.report_builder import build_tool6_full_report_docx

from design.components.tool8_ui import(
    topbar,
    card_open,
    card_close,
    kv,
    sticky_open,
    sticky_close,
    status_card,
    table_card_open,
    table_card_close,
)
from design.components.global_ui import apply_global_background

apply_global_background("assets/images/Logo_of_PPC.png")

# =============================
# Optional: pysurveycto (API attachments)
# =============================
try:
    import pysurveycto  # pip install pysurveycto
    _HAS_PYSURVEYCTO = True
except Exception:
    pysurveycto = None
    _HAS_PYSURVEYCTO = False

# =============================
# UI / Design
# =============================
st.set_page_config(page_title="Tool 8", layout="wide")

project_root = os.path.dirname(os.path.dirname(__file__))

topbar(
    title="Tool 8 — Report Generator",
    subtitle="Auto-filled from Google Sheet + attachments from SurveyCTO",
    right_chip="WASH  UNICEF",
)

# =============================
# Google Sheet Column Mapping
# =============================
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
    return f"{lat}, {lon}".strip().strip(",").strip()

# =============================
# Small helpers
# =============================
def safe_str(x) -> str:
    return "" if x is None else str(x)

def ensure_http(url: str) -> str:
    u = (url or "").strip()
    return u if u.startswith("http") else ""

def format_af_phone_ui(raw) -> str:
    s0 = re.sub(r"\D+", "", safe_str(raw).strip())
    if not s0:
        return ""
    if s0.startswith("0"):
        s0 = s0[1:]
    if s0.startswith("93"):
        return f"+{s0}"
    return f"+93{s0}"

def na_if_empty_ui(raw) -> str:
    s0 = safe_str(raw).strip()
    return s0 if s0 else "N/A"

def enforce_single_cover(selections: Dict[str, str]) -> Dict[str, str]:
    covers = [u for u, p in selections.items() if p == "Cover Page"]
    if len(covers) <= 1:
        return selections
    keep = covers[-1]
    for u in covers[:-1]:
        selections[u] = "Not selected"
    return selections

def _looks_like_html(data: bytes) -> bool:
    head = (data or b"")[:400].lower()
    return head.startswith(b"<!doctype html") or b"<html" in head or b"<head" in head

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
        issues.append("Low resolution (≥ 1200×700 recommended).")
    if ratio < 1.2:
        issues.append("Landscape photo recommended.")
    if brightness < 60:
        issues.append("Too dark.")
    if brightness > 200:
        issues.append("Too bright.")
    if sharpness < 40:
        issues.append("May be blurry.")

    return issues, {"w": w, "h": h, "ratio": ratio, "brightness": brightness, "sharpness": sharpness}

# =============================
# SurveyCTO Auth + Media loaders
# =============================
def get_auth() -> Optional[HTTPBasicAuth]:
    user = st.secrets.get("SURVEYCTO_USER", "") if hasattr(st, "secrets") else ""
    pwd = st.secrets.get("SURVEYCTO_PASS", "") if hasattr(st, "secrets") else ""
    user = user or st.session_state.get("scto_user", "")
    pwd = pwd or st.session_state.get("scto_pass", "")
    return HTTPBasicAuth(user, pwd) if user and pwd else None

def _is_scto_view_attachment(url: str) -> bool:
    u = (url or "").lower()
    return "surveycto.com/view/submission-attachment" in u

SCTO_SERVER = "act4performance"
FORM_ID = "wash06_solar_water_supply_V2"

def get_scto_client():
    if not _HAS_PYSURVEYCTO:
        return None
    user = st.session_state.get("scto_user", "").strip()
    pwd = st.session_state.get("scto_pass", "").strip()
    if not user or not pwd:
        return None
    try:
        return pysurveycto.SurveyCTOObject(SCTO_SERVER, user, pwd)
    except Exception:
        return None

@st.cache_data(show_spinner=False, ttl=3600)
def scto_get_attachment_bytes(url: str, user_key: str) -> Optional[bytes]:
    scto = get_scto_client()
    if scto is None:
        return None
    try:
        b = scto.get_attachment(url)
        if not b:
            return None
        if _looks_like_html(b):
            return None
        return b
    except Exception:
        return None

@st.cache_data(show_spinner=False, ttl=3600)
def fetch_image_cached(url: str, user_key: str) -> Tuple[bool, Optional[bytes], str]:
    auth = get_auth()
    if auth is None:
        return False, None, "Missing SurveyCTO credentials."

    if _is_scto_view_attachment(url):
        b = scto_get_attachment_bytes(url, user_key=user_key or "user")
        if b:
            try:
                clean = _to_clean_png_bytes(b)
                return True, clean, "OK"
            except Exception:
                pass

    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
        "Referer": "https://act4performance.surveycto.com/",
    }
    try:
        r = requests.get(url, headers=headers, timeout=25, allow_redirects=True, auth=auth)
        ctype = (r.headers.get("Content-Type") or "").lower()
        if r.status_code >= 400:
            return False, None, f"HTTP {r.status_code}"
        data = r.content or b""
        if "text/html" in ctype or _looks_like_html(data):
            return False, None, "Auth/session required (HTML response)."
        clean = _to_clean_png_bytes(data)
        return True, clean, "OK"
    except Exception as e:
        return False, None, str(e)

@st.cache_data(show_spinner=False, ttl=3600)
def fetch_audio_cached(url: str, user_key: str) -> Tuple[bool, Optional[bytes], str, str]:
    auth = get_auth()
    if auth is None:
        return False, None, "Missing SurveyCTO credentials.", ""

    if _is_scto_view_attachment(url):
        b = scto_get_attachment_bytes(url, user_key=user_key or "user")
        if b:
            ul = (url or "").lower()
            if ".mp3" in ul:
                mime = "audio/mpeg"
            elif ".wav" in ul:
                mime = "audio/wav"
            elif ".m4a" in ul:
                mime = "audio/mp4"
            elif ".ogg" in ul:
                mime = "audio/ogg"
            elif ".opus" in ul:
                mime = "audio/opus"
            else:
                mime = "audio/aac"
            return True, b, "OK", mime

    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "audio/*,*/*;q=0.8",
        "Referer": "https://act4performance.surveycto.com/",
    }
    try:
        r = requests.get(url, headers=headers, timeout=35, allow_redirects=True, auth=auth)
        ctype = (r.headers.get("Content-Type") or "").lower()
        if r.status_code >= 400:
            return False, None, f"HTTP {r.status_code}", ctype
        data = r.content or b""
        if "text/html" in ctype or _looks_like_html(data):
            return False, None, "Auth/session required (HTML response).", ctype

        if ctype.startswith("audio/"):
            mime = ctype.split(";")[0].strip()
        else:
            ul = (url or "").lower()
            if ".mp3" in ul:
                mime = "audio/mpeg"
            elif ".wav" in ul:
                mime = "audio/wav"
            elif ".m4a" in ul:
                mime = "audio/mp4"
            elif ".ogg" in ul:
                mime = "audio/ogg"
            elif ".opus" in ul:
                mime = "audio/opus"
            else:
                mime = "audio/aac"

        return True, data, "OK", mime
    except Exception as e:
        return False, None, str(e), ""

def fetch_image(url: str) -> Tuple[bool, Optional[bytes], str]:
    user_key = (st.secrets.get("SURVEYCTO_USER", "") if hasattr(st, "secrets") else "") or st.session_state.get("scto_user", "")
    return fetch_image_cached(url, user_key=user_key or "user")

def fetch_audio(url: str) -> Tuple[bool, Optional[bytes], str, str]:
    user_key = (st.secrets.get("SURVEYCTO_USER", "") if hasattr(st, "secrets") else "") or st.session_state.get("scto_user", "")
    return fetch_audio_cached(url, user_key=user_key or "user")

# =============================
# URL normalize (Drive only)
# =============================
def normalize_image_url(url: str) -> str:
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
    links = []
    for k, v in (row or {}).items():
        s0 = safe_str(v).strip()
        if not s0.startswith("http"):
            continue
        s_low = s0.lower()
        is_img = any(ext in s_low for ext in [".jpg", ".jpeg", ".png", ".webp"])
        is_drive = ("drive.google.com" in s_low) or ("googleusercontent.com" in s_low)
        is_scto = "surveycto.com/view/submission-attachment" in s_low
        if is_img or is_drive or is_scto:
            links.append({"field": k, "url": normalize_image_url(s0)})

    uniq, seen = [], set()
    for it in links:
        if it["url"] in seen:
            continue
        seen.add(it["url"])
        uniq.append(it)
    return uniq

def extract_audio_links(row: dict) -> List[Dict[str, str]]:
    links = []
    for k, v in (row or {}).items():
        s0 = safe_str(v).strip()
        if not s0.startswith("http"):
            continue
        s_low = s0.lower()
        is_audio = any(ext in s_low for ext in [".aac", ".mp3", ".wav", ".m4a", ".ogg", ".opus"])
        is_drive = ("drive.google.com" in s_low) or ("googleusercontent.com" in s_low)
        is_scto = "surveycto.com/view/submission-attachment" in s_low
        if is_audio or (is_scto and ("audio" in k.lower() or "voice" in k.lower() or "record" in k.lower())):
            links.append({"field": k, "url": normalize_image_url(s0)})

    uniq, seen = [], set()
    for it in links:
        if it["url"] in seen:
            continue
        seen.add(it["url"])
        uniq.append(it)
    return uniq

# =============================
# Wizard (Next/Back)
# =============================
STEPS = [
    "1) Cover Photo",
    "2) General Info",
    "3) Photos (Findings/Observations)",
    "4) Audio",
    "5) Section 5 (Components)",
    "6) Generate DOCX",
]

def init_state():
    st.session_state.setdefault("step_idx", 0)

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

    st.session_state.setdefault("executive_summary", "")
    st.session_state.setdefault("data_collection", "")
    st.session_state.setdefault("work_progress", "")
    st.session_state.setdefault("findings", "")
    st.session_state.setdefault("conclusion", "")

    st.session_state.setdefault("components_list", [])
    st.session_state.setdefault("component_inputs", {})
    st.session_state.setdefault("component_observations", [])

def go_next():
    if st.session_state["step_idx"] < len(STEPS) - 1:
        st.session_state["step_idx"] += 1

def go_back():
    if st.session_state["step_idx"] > 0:
        st.session_state["step_idx"] -= 1

def wizard_header():
    st.markdown("---")
    a, b = st.columns([7, 3])
    with a:
        st.subheader(STEPS[st.session_state["step_idx"]])
    with b:
        st.caption(f"Step {st.session_state['step_idx'] + 1} / {len(STEPS)}")

def wizard_nav(*, can_next: bool, next_label: str = "Next ➜"):
    c1, c2, c3 = st.columns([1, 6, 1])
    with c1:
        st.button("⬅ Back", on_click=go_back, disabled=(st.session_state["step_idx"] == 0), use_container_width=True)
    with c3:
        st.button(next_label, on_click=go_next, disabled=(not can_next), type="primary", use_container_width=True)

# =============================
# TPM selection
# =============================
init_state()

tool_name = "Tool 8"
tpm_id = st.session_state.get("tpm_id")
if not tpm_id:
    st.warning("No TPM ID found. Go back to Home and select a TPM ID.")
    st.stop()

sticky_open()
st.info(f"Selected TPM ID: {tpm_id}")
sticky_close()

# =============================
# Sidebar Login
# =============================
with st.sidebar:
    st.subheader("SurveyCTO Login")
    st.caption("")  #Hent for surveyCTO
    st.text_input("Username", key="scto_user")
    st.text_input("Password", type="password", key="scto_pass")
    st.caption("") #Hent for Bottom of CTO

# =============================
# Load data
# =============================
ws = open_worksheet(GOOGLE_SHEET_ID, tool_name)
records = get_all_records(ws)
row = find_by_tpm_id(records, tpm_id=tpm_id, tpm_col=TPM_COL)
if not row:
    st.error("The selected TPM ID was not found in the Tool 8 worksheet.")
    st.stop()

# =============================
# Defaults + hints
# =============================
defaults = {
    "Province": col(row, "PROVINCE", ""),
    "District": col(row, "DISTRICT", ""),
    "Village / Community": col(row, "VILLAGE", ""),
    "GPS points": gps_points_from_row(row),
    "Project Name": col(row, "ACTIVITY_NAME", ""),
    "Date of Visit": safe_str(col(row, "STARTTIME", "")).split(" ")[0]
    if safe_str(col(row, "STARTTIME", "")).strip()
    else "",
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
    "Date of Current Report": safe_str(col(row, "CURRENT_REPORT_DATE", "")).split(" ")[0]
    if safe_str(col(row, "CURRENT_REPORT_DATE", "")).strip()
    else "",
    "Number of Sites Visited": col(row, "VISIT_NUMBER", ""),
}

hints = {
    "CDC Code": "Verify against official report.",
    "Monitoring Report Number": "Verify against official report.",
    "Contact Number of the Respondent": "Auto-formatted to +93.",
    "Email Address of the Respondent": "If empty, DOCX shows N/A.",
    "Reason for delay": "If empty, DOCX shows N/A.",
}

# init overrides once
if not st.session_state["general_info_overrides"]:
    st.session_state["general_info_overrides"] = {k: str(v) for k, v in defaults.items()}

# =============================
# Collect media links (once)
# =============================
photos = extract_photo_links(row)
audios = extract_audio_links(row)

# Prepare photo labels (used in multiple steps)
photo_label_by_url: Dict[str, str] = {}
for i, p in enumerate(photos or [], start=1):
    u = ensure_http(p.get("url", ""))
    f = p.get("field", "Photo")
    if u:
        photo_label_by_url[u] = f"{i:02d}. {f}"
all_photo_urls = list(photo_label_by_url.keys())

# =============================
# Logos (paths passed to builder)
# =============================
unicef_logo_path = os.path.join(project_root, "assets", "images", "Logo_of_UNICEF.png")
act_logo_path = os.path.join(project_root, "assets", "images", "Logo_of_ACT.png")
ppc_logo_path = os.path.join(project_root, "assets", "images", "Logo_of_PPC.png")

unicef_logo_path = unicef_logo_path if os.path.exists(unicef_logo_path) else None
act_logo_path = act_logo_path if os.path.exists(act_logo_path) else None
ppc_logo_path = ppc_logo_path if os.path.exists(ppc_logo_path) else None

# =============================
# Render Wizard Steps
# =============================
wizard_header()
step = st.session_state["step_idx"]

# -------------------------------------------------
# STEP 1 — Cover Photo (ONLY PHOTOS)
# -------------------------------------------------
if step == 0:
    box = st.container(border=True)
    with box:
        card_open(
            "Cover Photo Selection",
            subtitle="Select ONE photo for the cover. Only photos are shown here (no audio).",
            variant="lg-variant-cyan",
        )

        if not all_photo_urls:
            status_card(
                title="No photo links found",
                description="No images detected in the dataset for this TPM ID.",
                pill="ERROR",
                level="error",
            )
            card_close()
            wizard_nav(can_next=False)
            st.stop()

        # single select cover
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

            # force cover
            st.session_state["photo_selections"][selected_url] = "Cover Page"
            st.session_state["photo_selections"] = enforce_single_cover(st.session_state["photo_selections"])

            # suitability
            try:
                img = Image.open(BytesIO(data)).convert("RGB")
                issues, m = cover_suitability(img)
                st.write(f"**Cover Check:** {m['w']}×{m['h']} | ratio={m['ratio']:.2f}")
                if issues:
                    status_card(
                        title="Cover warnings",
                        description=" • ".join(issues),
                        pill="WARNING",
                        level="warning",
                    )
            except Exception:
                pass

            status_card(
                title="Cover selected",
                description="Cover saved successfully",
                pill="SUCCESS",
                level="success",
            )
        else:
            status_card(
                title="Cover could not be loaded",
                description=f"Reason: {msg}",
                pill="ERROR",
                level="error",
            )

        st.markdown("### Fallback (Upload Cover)")
        up = st.file_uploader("Upload cover photo if needed", type=["jpg", "jpeg", "png"], key="cover_upl")
        if up is not None:
            b = up.read()
            try:
                b2 = _to_clean_png_bytes(b)
            except Exception:
                b2 = b
            st.session_state["cover_upload_bytes"] = b2
            st.image(b2, use_container_width=True)
            status_card(
                title="Uploaded cover saved",
                description="Upload cover will be used if online cover fails.",
                pill="SUCCESS",
                level="success",
            )

        card_close()

    # next condition
    selections = st.session_state.get("photo_selections", {})
    pb = st.session_state.get("photo_bytes", {})
    cover_urls = [u for u, p in selections.items() if p == "Cover Page"]
    cover_ok = (bool(cover_urls) and (cover_urls[0] in pb)) or (st.session_state.get("cover_upload_bytes") is not None)

    wizard_nav(can_next=cover_ok, next_label="Next ➜ General Info")
    st.stop()

# -------------------------------------------------
# STEP 2 — General Info (edit fields)
# -------------------------------------------------
if step == 1:
    st.subheader("General Project Information (Review & Edit)")

    box = st.container(border=True)
    with box:
        card_open("Edit fields", subtitle="Update if needed. Saved in session.", variant="lg-variant-green")

        tabs = st.tabs(["Project", "Respondent", "Monitoring", "Status/Other"])

        def _inp(field: str):
            cur = st.session_state["general_info_overrides"].get(field, str(defaults.get(field, "")))
            st.session_state["general_info_overrides"][field] = st.text_input(field, value=str(cur))
            st.caption(hints.get(field, " "))

        with tabs[0]:
            for f in ["Province", "District", "Village / Community", "GPS points", "Project Name", "Date of Visit"]:
                _inp(f)

        with tabs[1]:
            for f in [
                "Name of the respondent (Participant / UNICEF / IPs)",
                "Sex of Respondent",
                "Contact Number of the Respondent",
                "Email Address of the Respondent",
            ]:
                _inp(f)

        with tabs[2]:
            for f in [
                "Name of the IP, Organization / NGO",
                "Name of the monitor Engineer",
                "Email of the monitor engineer",
                "Monitoring Report Number",
                "Date of Current Report",
                "Number of Sites Visited",
            ]:
                _inp(f)

        with tabs[3]:
            for f in ["Project Status", "Reason for delay", "CDC Code", "Donor Name"]:
                _inp(f)

        card_close()

    status_card(
        title="Saved",
        description="Edits are stored and will be used in DOCX generation.",
        pill="SUCCESS",
        level="success",
    )

    wizard_nav(can_next=True, next_label="Next ➜ Photos")
    st.stop()

# -------------------------------------------------
# STEP 3 — Photos (Findings/Observations) — NO COVER HERE
# -------------------------------------------------
if step == 2:
    st.subheader("Project Photos (Findings / Observations)")

    box = st.container(border=True)
    with box:
        card_open(
            "Photo Manager",
            subtitle="Assign photos to Findings or Observations. Cover is not edited here.",
            variant="lg-variant-cyan",
        )

        if not all_photo_urls:
            status_card(
                title="No photo links found",
                description="No images detected in the dataset for this TPM ID.",
                pill="WARNING",
                level="warning",
            )
            card_close()
            wizard_nav(can_next=True, next_label="Next ➜ Audio")
            st.stop()

        selected_url = st.selectbox(
            "Select a photo",
            options=all_photo_urls,
            format_func=lambda u: photo_label_by_url.get(u, u),
            key="step3_photo_pick",
        )

        purpose_options = ["Not selected", "Findings", "Observations"]
        cur_purpose = st.session_state["photo_selections"].get(selected_url, "Not selected")
        if cur_purpose == "Cover Page":
            cur_purpose = "Not selected"  # protect cover from here

        purpose = st.selectbox(
            "Assign purpose",
            purpose_options,
            index=purpose_options.index(cur_purpose) if cur_purpose in purpose_options else 0,
            key="step3_purpose",
        )
        st.session_state["photo_selections"][selected_url] = purpose

        ok, data, msg = fetch_image(selected_url)
        if ok and data:
            st.image(data, use_container_width=True)
            st.session_state["photo_bytes"][selected_url] = data
            status_card(
                title="Saved",
                description="Photo bytes saved in session",
                pill="SUCCESS",
                level="success",
            )
        else:
            status_card(
                title="Failed to load",
                description=f"Reason: {msg}",
                pill="ERROR",
                level="error",
            )

        st.markdown("---")
        st.write("**Selected summary:**")
        any_sel = False
        for u, p in st.session_state["photo_selections"].items():
            if p in ["Findings", "Observations"]:
                any_sel = True
                flag = "✅" if u in st.session_state["photo_bytes"] else "⚠️"
                st.write(f"- {flag} **{p}** — {photo_label_by_url.get(u, 'Photo')}")
        if not any_sel:
            st.caption("No Findings/Observations photos selected yet.")

        card_close()

    wizard_nav(can_next=True, next_label="Next ➜ Audio")
    st.stop()

# -------------------------------------------------
# STEP 4 — Audio
# -------------------------------------------------
if step == 3:
    st.subheader("Project Audio (Play, Select, Save notes)")

    box = st.container(border=True)
    with box:
        card_open(
            "Audio Player",
            subtitle="Select audio, load, play, and store notes/transcript (optional).",
            variant="lg-variant-purple",
        )

        if not audios:
            status_card(
                title="No audio links found",
                description="If audio fields exist, ensure URLs are present (SurveyCTO attachment links).",
                pill="INFO",
                level="success",
            )
            card_close()
            wizard_nav(can_next=True, next_label="Next ➜ Section 5")
            st.stop()

        options = []
        for i, a in enumerate(audios, start=1):
            field = a.get("field", "Unknown field")
            url = ensure_http(a.get("url", ""))
            if url:
                options.append((f"{i:02d}. {field}", url, field))

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
            st.session_state["audio_mime"][sel_url] = mime or "audio/aac"

            status_card(
                title="Audio loaded successfully",
                description="You can play it now and store notes.",
                pill="SUCCESS",
                level="success",
            )

            st.audio(
                data,
                format=st.session_state["audio_mime"].get(sel_url, "audio/aac"),
                start_time=0,
            )

            cur_note = st.session_state["audio_notes"].get(sel_url, "")
            note = st.text_area(
                "Notes / Transcript (optional)",
                value=cur_note,
                height=120,
                key="step4_audio_note",
                placeholder="Write key points you heard to use in reporting.",
            )
            st.session_state["audio_notes"][sel_url] = note
        else:
            status_card(
                title="Audio could not be loaded",
                description=f"Reason: {msg}",
                pill="ERROR",
                level="error",
            )
            st.info("Login to SurveyCTO in sidebar. If already logged in, your account may not have permission.")

        card_close()

    wizard_nav(can_next=True, next_label="Next ➜ Section 5")
    st.stop()

# -------------------------------------------------
# STEP 5 — Section 5 Components (Dynamic)
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
    st.session_state["components_list"] = [
        c for c in st.session_state.get("components_list", []) if c.get("uid") != uid
    ]
    st.session_state.setdefault("component_inputs", {})
    st.session_state["component_inputs"].pop(uid, None)
    for k in [f"{uid}_photos", f"{uid}_obs", f"{uid}_reco", f"{uid}_major_table_editor"]:
        st.session_state.pop(k, None)
    _reindex_components()
    st.session_state["component_observations"] = []

if step == 4:
    # init default components once
    if not st.session_state["components_list"]:
        st.session_state["components_list"] = [
            {"uid": str(uuid4()), "comp_id": "TEMP", "title": t, "reco_title": r}
            for (t, r) in DEFAULT_COMPONENT_TITLES
        ]
        _reindex_components()

    st.subheader("5. Project Component-Wise Key Observations")

    # Controls add component
    ctrl = st.container(border=True)
    with ctrl:
        c1, c2, c3 = st.columns([2, 2, 1])
        with c1:
            new_title = st.text_input("Add new component — Title", placeholder="e.g., Construction of c.. unit", key="new_comp_title")
        with c2:
            new_reco = st.text_input("Recommendation Title", placeholder="e.g., Recommendations for c..", key="new_comp_reco")
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
                    status_card(
                        title="Title is required",
                        description="Please enter a component title to add it.",
                        pill="WARNING",
                        level="warning",
                    )

    box5 = st.container(border=True)
    with box5:
        card_open(
            "Complete components",
            subtitle="For each component: select up to 3 photos, write Observation, Major findings (table), and Recommendations.",
            variant="lg-variant-green",
        )

        for comp in st.session_state["components_list"]:
            uid = comp["uid"]
            comp_id = comp["comp_id"]
            title = comp["title"]
            reco_title = comp["reco_title"]

            st.session_state["component_inputs"].setdefault(
                uid, {"photos": [], "observation": "", "major_table": [], "reco": ""}
            )
            data = st.session_state["component_inputs"][uid]

            h1, h2 = st.columns([8, 2])
            with h1:
                st.markdown(f"### {comp_id}. {title}")
            with h2:
                if st.button("🗑 Remove", key=f"rm_{uid}", use_container_width=True):
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
                            ok, b, _ = fetch_image(u)
                            if ok and b:
                                st.session_state["photo_bytes"][u] = b
                        with cols[idx]:
                            if u in st.session_state["photo_bytes"]:
                                st.image(st.session_state["photo_bytes"][u], use_container_width=True)
                            else:
                                st.caption("⚠️ Not loaded. Load/save it in step 3 if needed.")

                obs = st.text_area("Observation / Description", value=data.get("observation", ""), height=120, key=f"{uid}_obs")

                table_card_open("Major findings (table format)", meta="Max 10 rows • NO auto-numbered")
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

                reco = st.text_area(reco_title + ":", value=data.get("reco", ""), height=100, key=f"{uid}_reco")

                st.session_state["component_inputs"][uid] = {
                    "photos": sel,
                    "observation": obs,
                    "major_table": major_table,
                    "reco": reco,
                }

        card_close()

    # Export section 5 structure for DOCX builder
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

    wizard_nav(can_next=True, next_label="Next ➜ Generate DOCX")
    st.stop()

# -------------------------------------------------
# STEP 6 — Generate DOCX
# -------------------------------------------------
if step == 5:
    st.subheader("Generate Report")

    box = st.container(border=True)
    with box:
        card_open(
            "Generate DOCX",
            subtitle="Uses edited values + selected media + section 5 data.",
            variant="lg-variant-cyan",
        )

        st.write("- Uses edited values (General Info)")
        st.write("- Uses Cover photo from Step 1")
        st.write("- Uses Photos (Findings/Observations) from Step 3")
        st.write("- Uses Audio notes (optional) from Step 4")
        st.write("- Uses Section 5 (Components) from Step 5")

        # Resolve cover bytes
        selections = st.session_state.get("photo_selections", {})
        photo_bytes = st.session_state.get("photo_bytes", {})

        cover_urls = [u for u, p in selections.items() if p == "Cover Page"]
        cover_bytes = photo_bytes.get(cover_urls[0]) if cover_urls else None
        if cover_bytes is None and st.session_state.get("cover_upload_bytes") is not None:
            cover_bytes = st.session_state["cover_upload_bytes"]

        if cover_bytes is None:
            status_card(
                title="No usable Cover photo",
                description="Go back to Step 1 and select a cover photo.",
                pill="ERROR",
                level="error",
            )
            card_close()
            wizard_nav(can_next=False, next_label="—")
            st.stop()

        if st.button("Generate DOCX Report", type="primary", use_container_width=True):
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

            status_card(
                title="Report generated successfully",
                description="Your DOCX is ready to download.",
                pill="SUCCESS",
                level="success",
            )

            st.download_button(
                "Download Report (DOCX)",
                data=docx_bytes,
                file_name=f"Tool8_Report_{tpm_id}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

        card_close()

    wizard_nav(can_next=True, next_label="Finish ✅")
    st.stop()
