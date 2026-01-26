# src/auth.py
from __future__ import annotations

import json
from pathlib import Path
from typing import Optional

import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

# Scopes needed for Google Sheets + Drive
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def _find_credentials_json() -> Optional[Path]:
    """
    Try to find credentials.json in common local paths.
    Returns path if found, else None.
    """
    candidates = [
        Path("credentials.json"),
        Path(".streamlit") / "credentials.json",
        Path("secrets") / "credentials.json",
    ]
    for p in candidates:
        if p.exists() and p.is_file():
            return p
    return None

def _creds_from_secrets() -> Credentials:
    """
    Build Google Credentials from Streamlit Cloud secrets.
    Expects a TOML section: [gcp_service_account]
    """
    if "gcp_service_account" not in st.secrets:
        raise FileNotFoundError(
            "credentials.json not found and st.secrets['gcp_service_account'] is missing. "
            "Add service account JSON to Streamlit Secrets under [gcp_service_account]."
        )

    info = dict(st.secrets["gcp_service_account"])

    # Some people store private_key without proper newlines; normalize
    pk = info.get("private_key", "")
    if pk and "\\n" in pk:
        info["private_key"] = pk.replace("\\n", "\n")

    return Credentials.from_service_account_info(info, scopes=SCOPES)

def _creds_from_file(path: Path) -> Credentials:
    info = json.loads(path.read_text(encoding="utf-8"))
    return Credentials.from_service_account_info(info, scopes=SCOPES)

def get_gspread_client():
    """
    Returns an authenticated gspread client.
    - Local: uses credentials.json if present
    - Streamlit Cloud: uses st.secrets["gcp_service_account"]
    """
    local_path = _find_credentials_json()

    if local_path:
        creds = _creds_from_file(local_path)
    else:
        creds = _creds_from_secrets()

    return gspread.authorize(creds)
