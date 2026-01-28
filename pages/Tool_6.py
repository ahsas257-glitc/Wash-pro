# pages/Tool_6.py
import os
import re
from io import BytesIO
from typing import Optional, Dict, Tuple, List

import streamlit as st
from PIL import Image
import numpy as np
import requests
import pandas as pd
from uuid import uuid4

from src.config import GOOGLE_SHEET_ID, TPM_COL
from src.data_processing import open_worksheet, get_all_records, find_by_tpm_id
from src.report_builder import build_tool6_full_report_docx
from design.components.tool6_ui import topbar

from design.components.tool6_ui import (
    inject_tool6_design,
    card_open,
    card_close,
    kv,
    sticky_open,
    sticky_close,
)

# =============================
# UI / Design
# =============================
st.set_page_config(page_title="Tool 6", layout="wide")

project_root = os.path.dirname(os.path.dirname(__file__))

inject_tool6_design(
    project_root=project_root,
    background_image_rel="assets/images/Logo_of_PPC.png",
    noise_image_rel="assets/images/bg_noise.png",
    intensity=0.09,
    enable_parallax=True
)

st.title("Tool 6 — Report Generator")


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
    """
    Convert ANY valid image bytes to clean PNG bytes (prevents docx UnrecognizedImageError).
    """
    img = Image.open(BytesIO(img_bytes)).convert("RGB")
    out = BytesIO()
    img.save(out, format="PNG", optimize=True)
    return out.getvalue()


# =============================
# Public Media Loaders (No SurveyCTO / No Account)
# =============================
@st.cache_data(show_spinner=False, ttl=3600)
def fetch_audio_cached(url: str) -> Tuple[bool, Optional[bytes], str, str]:
    """
    Download audio from a public URL (Drive direct/download or direct file links).
    No authentication is used.
    """
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "audio/*,*/*;q=0.8",
    }

    try:
        r = requests.get(url, headers=headers, timeout=35, allow_redirects=True)
        ctype = (r.headers.get("Content-Type") or "").lower()

        if r.status_code >= 400:
            return False, None, f"HTTP {r.status_code}", ctype

        data = r.content or b""
        if "text/html" in ctype or _looks_like_html(data):
            return False, None, "Private/protected link (HTML response). Please upload file.", ctype

        # Normalize MIME
        if ctype.startswith("audio/"):
            mime = ctype.split(";")[0].strip()
        else:
            ul = (url or "").lower()
            if ".aac" in ul:
                mime = "audio/aac"
            elif ".mp3" in ul:
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


@st.cache_data(show_spinner=False, ttl=3600)
def fetch_image_cached(url: str) -> Tuple[bool, Optional[bytes], str]:
    """
    Download image from a public URL (Drive direct/download or direct image links).
    No authentication is used.
    """
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
    }

    try:
        r = requests.get(url, headers=headers, timeout=25, allow_redirects=True)
        ctype = (r.headers.get("Content-Type") or "").lower()

        if r.status_code >= 400:
            return False, None, f"HTTP {r.status_code}"

        data = r.content or b""
        if "text/html" in ctype or _looks_like_html(data):
            return False, None, "Private/protected link (HTML response). Please upload file."

        clean = _to_clean_png_bytes(data)
        return True, clean, "OK"

    except Exception as e:
        return False, None, str(e)


def fetch_audio(url: str) -> Tuple[bool, Optional[bytes], str, str]:
    return fetch_audio_cached(url)

def fetch_image(url: str) -> Tuple[bool, Optional[bytes], str]:
    return fetch_image_cached(url)


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
    """
    Only collects likely public image links (direct image or Google Drive).
    SurveyCTO attachment links are ignored (since no auth exists anymore).
    """
    links = []
    for k, v in (row or {}).items():
        s0 = safe_str(v).strip()
        if not s0.startswith("http"):
            continue

        s_low = s0.lower()
        is_img = any(ext in s_low for ext in [".jpg", ".jpeg", ".png", ".webp"])
        is_drive = ("drive.google.com" in s_low) or ("googleusercontent.com" in s_low)

        if is_img or is_drive:
            links.append({"field": k, "url": normalize_image_url(s0)})

    uniq, seen = [], set()
    for it in links:
        if it["url"] in seen:
            continue
        seen.add(it["url"])
        uniq.append(it)
    return uniq


def extract_audio_links(row: dict) -> List[Dict[str, str]]:
    """
    Only collects likely public audio links (direct audio or Google Drive).
    SurveyCTO attachment links are ignored (since no auth exists anymore).
    """
    links = []
    for k, v in (row or {}).items():
        s0 = safe_str(v).strip()
        if not s0.startswith("http"):
            continue
        s_low = s0.lower()

        is_audio = any(ext in s_low for ext in [".aac", ".mp3", ".wav", ".m4a", ".ogg", ".opus"])
        is_drive = ("drive.google.com" in s_low) or ("googleusercontent.com" in s_low)

        if is_audio or is_drive:
            links.append({"field": k, "url": normalize_image_url(s0)})

    uniq, seen = [], set()
    for it in links:
        if it["url"] in seen:
            continue
        seen.add(it["url"])
        uniq.append(it)
    return uniq


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
# TPM selection
# =============================
tool_name = "Tool 6"
tpm_id = st.session_state.get("tpm_id")
if not tpm_id:
    st.warning("No TPM ID found. Go back to Home and select a TPM ID.")
    st.stop()

# ✅ sticky bar via design module
sticky_open()
st.info(f"Selected TPM ID: {tpm_id}")
sticky_close()


# =============================
# Load data
# =============================
ws = open_worksheet(GOOGLE_SHEET_ID, tool_name)
records = get_all_records(ws)
row = find_by_tpm_id(records, tpm_id=tpm_id, tpm_col=TPM_COL)
if not row:
    st.error("The selected TPM ID was not found in the Tool 6 worksheet.")
    st.stop()


# =============================
# Audio state
# =============================
st.session_state.setdefault("audio_selections", {})
st.session_state.setdefault("audio_bytes", {})
st.session_state.setdefault("audio_field", {})
st.session_state.setdefault("audio_notes", {})
st.session_state.setdefault("audio_mime", {})


# =============================
# Defaults + hints
# =============================
defaults = {
    "Province": row.get("A01_Province", ""),
    "District": row.get("A02_District", ""),
    "Village / Community": row.get("Village", ""),
    "GPS points": f'{row.get("GPS_1-Latitude","")}, {row.get("GPS_1-Longitude","")}'.strip(", ").strip(),
    "Project Name": row.get("Activity_Name", ""),
    "Date of Visit": safe_str(row.get("starttime", "")).split(" ")[0] if safe_str(row.get("starttime", "")).strip() else "",
    "Name of the IP, Organization / NGO": row.get("Primary_Partner_Name", ""),
    "Name of the monitor Engineer": row.get("A07_Monitor_name", ""),
    "Email of the monitor engineer": row.get("A12_Monitor_email", ""),
    "Name of the respondent (Participant / UNICEF / IPs)": row.get("A08_Respondent_name", ""),
    "Sex of Respondent": row.get("A09_Respondent_sex", ""),
    "Contact Number of the Respondent": format_af_phone_ui(row.get("A10_Respondent_phone", "")),
    "Email Address of the Respondent": na_if_empty_ui(row.get("A11_Respondent_email", "")),
    "Project Status": row.get("Project_Status", ""),
    "Reason for delay": na_if_empty_ui(row.get("B8_Reasons_for_delay", "")),
    "CDC Code": row.get("A23_CDC_code", ""),
    "Donor Name": row.get("A24_Donor_name", ""),
    "Monitoring Report Number": row.get("A25_Monitoring_report_number", ""),
    "Date of Current Report": safe_str(row.get("A20_Current_report_date", "")).split(" ")[0] if safe_str(row.get("A20_Current_report_date", "")).strip() else "",
    "Number of Sites Visited": row.get("A26_Visit_number", ""),
}

hints = {
    "CDC Code": "Verify against official report.",
    "Monitoring Report Number": "Verify against official report.",
    "Contact Number of the Respondent": "Auto-formatted to +93.",
    "Email Address of the Respondent": "If empty, DOCX shows N/A.",
    "Reason for delay": "If empty, DOCX shows N/A.",
}


# =============================
# Overview
# =============================
st.subheader("Dataset Overview (Auto-filled)")
box = st.container(height=330, border=True)
with box:
    card_open("Key Project Info")
    c1, c2, c3, c4 = st.columns(4)
    with c1: kv("Province", defaults["Province"])
    with c2: kv("District", defaults["District"])
    with c3: kv("Village", defaults["Village / Community"])
    with c4: kv("Status", defaults["Project Status"])

    st.markdown("---")
    a, b, c = st.columns(3)
    with a: kv("Project Name", defaults["Project Name"])
    with b: kv("GPS", defaults["GPS points"])
    with c: kv("Date of Visit", defaults["Date of Visit"])
    card_close("Review values. Update in Edit if needed.")


# =============================
# Edit
# =============================
st.subheader("General Project Information (Review & Edit)")
st.session_state.setdefault("general_info_overrides", {k: str(v) for k, v in defaults.items()})

edit_box = st.container(height=520, border=True)
with edit_box:
    tabs = st.tabs(["Project", "Respondent", "Monitoring", "Status/Other"])

    def _inp(field: str):
        cur = st.session_state["general_info_overrides"].get(field, str(defaults.get(field, "")))
        st.session_state["general_info_overrides"][field] = st.text_input(field, value=str(cur))
        st.caption(hints.get(field, " "))

    with tabs[0]:
        card_open("Project")
        for f in ["Province", "District", "Village / Community", "GPS points", "Project Name", "Date of Visit"]:
            _inp(f)
        card_close("Keep location, GPS, and dates accurate.")

    with tabs[1]:
        card_open("Respondent")
        for f in ["Name of the respondent (Participant / UNICEF / IPs)", "Sex of Respondent",
                  "Contact Number of the Respondent", "Email Address of the Respondent"]:
            _inp(f)
        card_close("Phone is formatted to +93 automatically.")

    with tabs[2]:
        card_open("Monitoring")
        for f in ["Name of the IP, Organization / NGO", "Name of the monitor Engineer", "Email of the monitor engineer",
                  "Monitoring Report Number", "Date of Current Report", "Number of Sites Visited"]:
            _inp(f)
        card_close("Report number and dates are critical.")

    with tabs[3]:
        card_open("Status / Other")
        for f in ["Project Status", "Reason for delay", "CDC Code", "Donor Name"]:
            _inp(f)
        card_close("Use N/A when not applicable.")

st.success("Edits are saved for this session.")
# =============================
# Photos
# =============================
st.subheader("Project Photos (Preview, Select, Save to Report)")

photos = extract_photo_links(row)
st.session_state.setdefault("photo_selections", {})
st.session_state.setdefault("photo_bytes", {})
st.session_state.setdefault("photo_field", {})

photo_box = st.container(height=720, border=True)
with photo_box:
    if not photos:
        st.warning("No photo links found for this TPM ID.")
        st.caption("Ensure attachment URLs exist in the dataset (public links).")
    else:
        card_open("Photo Viewer")

        options = []
        for i, p in enumerate(photos, start=1):
            field = p.get("field", "Unknown field")
            url = ensure_http(p.get("url", ""))
            options.append((f"{i:02d}. {field}", url, field))

        labels = [x[0] for x in options]
        selected_label = st.selectbox("Select a photo", labels, index=0)
        _, sel_url, sel_field = next(x for x in options if x[0] == selected_label)
        st.session_state["photo_field"][sel_url] = sel_field

        purpose_options = ["Not selected", "Cover Page", "Findings", "Observations"]
        cur_purpose = st.session_state["photo_selections"].get(sel_url, "Not selected")
        purpose = st.selectbox(
            "Assign purpose",
            purpose_options,
            index=purpose_options.index(cur_purpose) if cur_purpose in purpose_options else 0,
        )
        st.session_state["photo_selections"][sel_url] = purpose
        st.session_state["photo_selections"] = enforce_single_cover(st.session_state["photo_selections"])
        st.caption("Only one Cover Page photo is allowed.")

        st.write(f"**Selected field:** `{sel_field}`")

        if not sel_url:
            st.error("Invalid photo URL.")
        else:
            ok, data, msg = fetch_image(sel_url)
            if ok and data:
                st.image(data, use_container_width=True)
                st.session_state["photo_bytes"][sel_url] = data
                st.success("Photo loaded and saved for the report.")

                if purpose == "Cover Page":
                    try:
                        img = Image.open(BytesIO(data)).convert("RGB")
                        issues, m = cover_suitability(img)
                        st.write(f"**Cover Check:** {m['w']}×{m['h']} | ratio={m['ratio']:.2f}")
                        if issues:
                            st.warning("Cover warnings:\n- " + "\n- ".join(issues))
                    except Exception:
                        pass
            else:
                # ✅ NO LINK SHOWN
                st.error("Photo could not be loaded.")
                st.caption(f"Reason: {msg}")
                st.info(
                    "This looks like a private/protected link. "
                    "Please upload the photo below as a fallback."
                )
                up = st.file_uploader(
                    "Upload this photo (fallback)",
                    type=["jpg", "jpeg", "png"],
                    key=f"upl_{sel_field}",
                )
                if up is not None:
                    b = up.read()
                    try:
                        b2 = _to_clean_png_bytes(b)
                    except Exception:
                        b2 = b
                    st.image(b2, use_container_width=True)
                    st.session_state["photo_bytes"][sel_url] = b2
                    st.success("Uploaded photo saved for the report.")

        st.markdown("---")
        st.write("**Selected summary:**")
        for u, p in st.session_state["photo_selections"].items():
            if p != "Not selected":
                flag = "✅" if u in st.session_state["photo_bytes"] else "⚠️"
                st.write(f"- {flag} **{p}** — `{st.session_state['photo_field'].get(u, 'Photo')}`")

        card_close("✅ Saved for DOCX. ⚠️ Selected but not saved yet (load or upload).")

# =============================
# Audio
# =============================
st.subheader("Project Audio (Play, Select, Save notes)")
audios = extract_audio_links(row)

audio_box = st.container(height=520, border=True)
with audio_box:
    if not audios:
        st.info("No audio links found for this TPM ID.")
        st.caption("If audio fields exist, ensure URLs are present (public links).")
    else:
        card_open("Audio Player")

        options = []
        for i, a in enumerate(audios, start=1):
            field = a.get("field", "Unknown field")
            url = ensure_http(a.get("url", ""))
            options.append((f"{i:02d}. {field}", url, field))

        labels = [x[0] for x in options]
        selected_label = st.selectbox("Select an audio", labels, index=0, key="audio_selectbox")
        _, sel_url, sel_field = next(x for x in options if x[0] == selected_label)

        st.session_state["audio_field"][sel_url] = sel_field

        purpose_options = ["Not selected", "Evidence / Findings", "Observation", "Other"]
        cur_purpose = st.session_state["audio_selections"].get(sel_url, "Not selected")
        purpose = st.selectbox(
            "Assign purpose",
            purpose_options,
            index=purpose_options.index(cur_purpose) if cur_purpose in purpose_options else 0,
            key="audio_purpose_selectbox",
        )
        st.session_state["audio_selections"][sel_url] = purpose

        st.write(f"**Selected field:** `{sel_field}`")

        if not sel_url:
            st.error("Invalid audio URL.")
        else:
            ok, data, msg, mime = fetch_audio(sel_url)
            if ok and data:
                st.session_state["audio_bytes"][sel_url] = data
                st.session_state["audio_mime"][sel_url] = mime or "audio/aac"

                st.success("Audio loaded successfully.")
                st.audio(
                    data,
                    format=st.session_state["audio_mime"].get(sel_url, "audio/aac"),
                    start_time=0,
                )

                note_key = f"audio_note_{sel_url}"
                cur_note = st.session_state["audio_notes"].get(sel_url, "")
                note = st.text_area(
                    "Notes / Transcript (optional)",
                    value=cur_note,
                    height=90,
                    key=note_key,
                    placeholder="Write key points you heard (findings/observation) to use in reporting.",
                )
                st.session_state["audio_notes"][sel_url] = note

            else:
                # ✅ NO LINK SHOWN
                st.error("Audio could not be loaded.")
                st.caption(f"Reason: {msg}")
                st.info(
                    "This looks like a private/protected link. "
                    "Please provide a public link or keep notes manually."
                )

        st.markdown("---")
        st.write("**Selected summary:**")
        any_sel = False
        for u, p in st.session_state["audio_selections"].items():
            if p != "Not selected":
                any_sel = True
                flag = "✅" if u in st.session_state["audio_bytes"] else "⚠️"
                fld = st.session_state["audio_field"].get(u, "Audio")
                st.write(f"- {flag} **{p}** — `{fld}`")
        if not any_sel:
            st.caption("No audios selected yet.")

        card_close("Loaded audios can be played and used to extract findings.")

# =============================
# Optional cover upload
# =============================
st.subheader("Cover Photo (Optional Upload)")
cover_photo = st.file_uploader("Upload cover photo if needed.", type=["jpg", "jpeg", "png"])
st.caption("Use upload only if the selected cover cannot be loaded.")

# =============================
# Logos
# =============================
unicef_logo_path = os.path.join(project_root, "assets", "images", "Logo_of_UNICEF.png")
act_logo_path = os.path.join(project_root, "assets", "images", "Logo_of_ACT.png")
ppc_logo_path = os.path.join(project_root, "assets", "images", "Logo_of_PPC.png")

unicef_logo_path = unicef_logo_path if os.path.exists(unicef_logo_path) else None
act_logo_path = act_logo_path if os.path.exists(act_logo_path) else None
ppc_logo_path = ppc_logo_path if os.path.exists(ppc_logo_path) else None

# =============================
# Placeholders
# =============================
st.session_state.setdefault("executive_summary", "")
st.session_state.setdefault("data_collection", "")
st.session_state.setdefault("work_progress", "")
st.session_state.setdefault("findings", "")
st.session_state.setdefault("conclusion", "")

# =============================
# Section 5 — Component-Wise Key Observations (Dynamic)
# =============================
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

if "components_list" not in st.session_state:
    st.session_state["components_list"] = [
        {"uid": str(uuid4()), "comp_id": "TEMP", "title": t, "reco_title": r}
        for (t, r) in DEFAULT_COMPONENT_TITLES
    ]
    _reindex_components()

st.session_state.setdefault("component_inputs", {})
st.session_state.setdefault("component_observations", [])

st.subheader("5. Project Component-Wise Key Observations")

ctrl = st.container(border=True)
with ctrl:
    c1, c2, c3 = st.columns([2, 2, 1])
    with c1:
        new_title = st.text_input("Add new component — Title", placeholder="e.g., Construction of chlorination unit")
    with c2:
        new_reco = st.text_input("Add new component — Recommendation title", placeholder="e.g., Recommendations for chlorination")
    with c3:
        if st.button("Add", use_container_width=True):
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
                st.warning("Title is required to add a component.")

photo_label_by_url: Dict[str, str] = {}
for i, p in enumerate(photos or [], start=1):
    u = ensure_http(p.get("url", ""))
    f = p.get("field", "Photo")
    if u:
        photo_label_by_url[u] = f"{i:02d}. {f}"
all_photo_urls = list(photo_label_by_url.keys())

box5 = st.container(height=820, border=True)
with box5:
    card_open("Complete components (you may add/remove)")
    st.caption("For each component: select up to 3 photos, write Observation, Major findings (table), and Recommendations.")

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
                            # ✅ NO LINK SHOWN
                            st.caption("⚠️ Not loaded. Load it from the Photo Viewer above or upload fallback there.")

            obs = st.text_area(
                "Observation / Description",
                value=data.get("observation", ""),
                height=120,
                key=f"{uid}_obs",
            )

            st.markdown("**Major findings (table format)**")

            default_rows = data.get("major_table") or [{"NO": 1, "Findings": "", "Compliance": "Yes", "Photo": ""}]
            df = pd.DataFrame(default_rows)
            for col in ["NO", "Findings", "Compliance", "Photo"]:
                if col not in df.columns:
                    df[col] = ""

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

            edited = edited.head(10).copy()
            edited["NO"] = list(range(1, len(edited) + 1))
            major_table = edited.to_dict(orient="records")

            reco = st.text_area(
                reco_title + ":",
                value=data.get("reco", ""),
                height=90,
                key=f"{uid}_reco",
            )

            st.session_state["component_inputs"][uid] = {
                "photos": sel,
                "observation": obs,
                "major_table": major_table,
                "reco": reco,
            }

    card_close("You can remove a component once, and it will disappear everywhere.")
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

# =============================
# Generate DOCX
# =============================
st.subheader("Generate Report")
gen = st.container(height=260, border=True)
with gen:
    card_open("Generate DOCX")
    st.write("- Uses edited values")
    st.write("- Uses selected photos (bytes) for Cover/Findings/Observations")
    st.write("- Uses component-wise photos + text from Section 5")
    st.caption("Make sure the Cover photo is saved (✅) or upload one.")

    if st.button("Generate DOCX Report", type="primary", use_container_width=True):
        selections = st.session_state.get("photo_selections", {})
        photo_bytes = st.session_state.get("photo_bytes", {})

        cover_urls = [u for u, p in selections.items() if p == "Cover Page"]
        cover_bytes = photo_bytes.get(cover_urls[0]) if cover_urls else None

        if cover_bytes is None and cover_photo is not None:
            try:
                cover_bytes = _to_clean_png_bytes(cover_photo.read())
            except Exception:
                cover_bytes = cover_photo.read()

        if cover_bytes is None:
            st.error("No usable Cover photo. Load/save a cover (✅) or upload one.")
            st.stop()

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

        st.success("Report generated successfully.")
        st.download_button(
            "Download Report (DOCX)",
            data=docx_bytes,
            file_name=f"Tool6_Report_{tpm_id}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    card_close()

