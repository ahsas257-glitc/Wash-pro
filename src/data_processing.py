# src/data_processing.py
from __future__ import annotations

import time
from typing import Any, Dict, List, Optional

import gspread
from gspread.exceptions import APIError

from src.auth import get_gspread_client


# -----------------------------------------
# Internal helpers
# -----------------------------------------
def _is_retryable_api_error(err: Exception) -> bool:
    """
    Decide if we should retry this error.
    Common transient cases:
      - 503 service unavailable
      - 429 too many requests / rate limit
      - 500 internal
      - 502 bad gateway
      - 504 gateway timeout
      - network-like errors sometimes appear as APIError text
    """
    msg = str(err)
    retry_markers = ["[503]", "[429]", "[500]", "[502]", "[504]", "service is currently unavailable"]
    return any(m in msg for m in retry_markers)


def _sleep_with_backoff(attempt: int, base_wait: float) -> None:
    """
    Exponential backoff: base_wait, 2*base_wait, 4*base_wait, ...
    attempt starts at 0.
    """
    wait = base_wait * (2**attempt)
    time.sleep(wait)


# -----------------------------------------
# Public API
# -----------------------------------------
def open_worksheet(
    sheet_id: str,
    worksheet_name: str,
    retries: int = 4,
    base_wait: float = 1.5,
) -> gspread.Worksheet:
    """
    Open a worksheet(tab) by name from a given Google Sheet ID.

    Features:
      - Uses your service account via get_gspread_client()
      - Retries transient Google API errors (503/429/5xx)
      - Exponential backoff to reduce rate-limit / burst issues

    Args:
      sheet_id: Google Sheet ID (the long id in the URL)
      worksheet_name: exact tab name (e.g., 'Tool 6')
      retries: number of attempts (default 4)
      base_wait: base seconds for exponential backoff

    Returns:
      gspread.Worksheet
    """
    gc = get_gspread_client()

    last_err: Optional[Exception] = None
    for attempt in range(retries):
        try:
            sh = gc.open_by_key(sheet_id)
            return sh.worksheet(worksheet_name)

        except APIError as e:
            last_err = e
            if _is_retryable_api_error(e) and attempt < retries - 1:
                _sleep_with_backoff(attempt, base_wait)
                continue
            raise

        except Exception as e:
            # Sometimes network/SSL issues come as generic exceptions.
            last_err = e
            if _is_retryable_api_error(e) and attempt < retries - 1:
                _sleep_with_backoff(attempt, base_wait)
                continue
            raise

    # Should never reach here, but just in case:
    if last_err:
        raise last_err
    raise RuntimeError("Failed to open worksheet for unknown reasons.")


def get_all_records(ws: gspread.Worksheet) -> List[Dict[str, Any]]:
    """
    Return all worksheet rows as a list of dictionaries (headers used as keys).

    Tip:
      If your sheet contains empty trailing rows/cols, get_all_records() is still safe.
    """
    return ws.get_all_records()


def list_tpm_ids(records: List[Dict[str, Any]], tpm_col: str = "TPM_ID") -> List[str]:
    """Build a unique, sorted list of TPM IDs from records."""
    ids: List[str] = []
    for r in records:
        v = str(r.get(tpm_col, "")).strip()
        if v:
            ids.append(v)
    return sorted(set(ids))


def find_by_tpm_id(
    records: List[Dict[str, Any]],
    tpm_id: str,
    tpm_col: str = "TPM_ID",
) -> Optional[Dict[str, Any]]:
    """Find the first record matching the provided TPM ID."""
    target = str(tpm_id).strip()
    for r in records:
        if str(r.get(tpm_col, "")).strip() == target:
            return r
    return None


# -----------------------------------------
# Optional: Streamlit-friendly caching hooks
# -----------------------------------------
def get_worksheet_cached(sheet_id: str, worksheet_name: str):
    """
    OPTIONAL helper: Use this inside Streamlit with st.cache_resource to reduce reruns.

    Example usage in Streamlit:
        import .streamlit as st
        from src.data_processing import get_worksheet_cached

        @st.cache_resource(show_spinner=False)
        def _ws(sheet_id, worksheet_name):
            return get_worksheet_cached(sheet_id, worksheet_name)

        ws = _ws(GOOGLE_SHEET_ID, selected_tool)

    We keep this function import-safe for non-Streamlit contexts.
    """
    return open_worksheet(sheet_id, worksheet_name)
