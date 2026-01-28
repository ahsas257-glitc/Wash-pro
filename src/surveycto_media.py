# src/surveycto_media.py
import re
from io import BytesIO
from typing import Optional, Tuple

import streamlit as st
import requests
from requests.auth import HTTPBasicAuth
from PIL import Image

# Optional pysurveycto
try:
    import pysurveycto
    _HAS_PYSURVEYCTO = True
except Exception:
    pysurveycto = None
    _HAS_PYSURVEYCTO = False


def _looks_like_html(data: bytes) -> bool:
    head = (data or b"")[:400].lower()
    return head.startswith(b"<!doctype html") or b"<html" in head or b"<head" in head


def _to_clean_png_bytes(img_bytes: bytes) -> bytes:
    img = Image.open(BytesIO(img_bytes)).convert("RGB")
    out = BytesIO()
    img.save(out, format="PNG", optimize=True)
    return out.getvalue()


def _is_scto_view_attachment(url: str) -> bool:
    u = (url or "").lower()
    return "surveycto.com/view/submission-attachment" in u


def get_scto_credentials() -> Tuple[str, str, str]:
    """
    Central place for SurveyCTO credentials.
    Priority: secrets -> session_state
    """
    user = (st.secrets.get("SURVEYCTO_USER", "") if hasattr(st, "secrets") else "").strip()
    pwd  = (st.secrets.get("SURVEYCTO_PASS", "") if hasattr(st, "secrets") else "").strip()
    server = (st.secrets.get("SURVEYCTO_SERVER", "act4performance") if hasattr(st, "secrets") else "act4performance")

    if not user:
        user = (st.session_state.get("scto_user", "") or "").strip()
    if not pwd:
        pwd = (st.session_state.get("scto_pass", "") or "").strip()

    return user, pwd, server


def get_auth() -> Optional[HTTPBasicAuth]:
    user, pwd, _ = get_scto_credentials()
    if user and pwd:
        return HTTPBasicAuth(user, pwd)
    return None


def get_scto_client():
    if not _HAS_PYSURVEYCTO:
        return None
    user, pwd, server = get_scto_credentials()
    if not user or not pwd:
        return None
    try:
        return pysurveycto.SurveyCTOObject(server, user, pwd)
    except Exception:
        return None


@st.cache_data(show_spinner=False, ttl=3600)
def scto_get_attachment_bytes(url: str, user_key: str) -> Optional[bytes]:
    scto = get_scto_client()
    if scto is None:
        return None
    try:
        b = scto.get_attachment(url)
        if not b or _looks_like_html(b):
            return None
        return b
    except Exception:
        return None


@st.cache_data(show_spinner=False, ttl=3600)
def fetch_image(url: str, user_key: str = "user") -> Tuple[bool, Optional[bytes], str]:
    """
    Fetch image from SurveyCTO (attachment link) first; fallback to requests.
    Returns (ok, png_bytes, message)
    """
    auth = get_auth()
    if auth is None:
        return False, None, "Missing SurveyCTO credentials."

    if _is_scto_view_attachment(url):
        b = scto_get_attachment_bytes(url, user_key=user_key)
        if b:
            try:
                return True, _to_clean_png_bytes(b), "OK"
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

        return True, _to_clean_png_bytes(data), "OK"

    except Exception as e:
        return False, None, str(e)


@st.cache_data(show_spinner=False, ttl=3600)
def fetch_audio(url: str, user_key: str = "user") -> Tuple[bool, Optional[bytes], str, str]:
    """
    Fetch audio from SurveyCTO (attachment link) first; fallback to requests.
    Returns (ok, bytes, message, mime)
    """
    auth = get_auth()
    if auth is None:
        return False, None, "Missing SurveyCTO credentials.", ""

    if _is_scto_view_attachment(url):
        b = scto_get_attachment_bytes(url, user_key=user_key)
        if b:
            # simple mime guess
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

        mime = ctype.split(";")[0].strip() if ctype.startswith("audio/") else "audio/aac"
        return True, data, "OK", mime

    except Exception as e:
        return False, None, str(e), ""
