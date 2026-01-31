# pages/home.py
import streamlit as st

from src.config import GOOGLE_SHEET_ID, TPM_COL, TOOLS
from src.data_processing import open_worksheet, get_all_records, list_tpm_ids
from design.components.global_ui import apply_global_background

apply_global_background("assets/images/Logo_of_PPC.png")

# ---------------------------
# CSS (theme-aware) - does NOT force white/light
# ---------------------------
st.markdown(
    """
    <style>
      .block-container { padding-top: 2.2rem; }

      .login-wrap{
        max-width: 460px;
        margin: 0 auto;
      }

      .login-card{
        background: var(--secondary-background-color);
        border: 1px solid rgba(120,120,120,0.25);
        border-radius: 18px;
        padding: 22px 22px 18px 22px;
        box-shadow: 0 18px 60px rgba(0,0,0,0.25);
      }

      .login-title{
        font-size: 1.35rem;
        font-weight: 700;
        margin: 0 0 6px 0;
        letter-spacing: 0.2px;
        color: var(--text-color);
      }

      .login-sub{
        opacity: 0.85;
        margin: 0 0 16px 0;
        font-size: 0.95rem;
        line-height: 1.35rem;
        color: var(--text-color);
      }

      .mini-help{
        opacity: 0.7;
        font-size: 0.86rem;
        margin-top: 10px;
        color: var(--text-color);
      }

      div[data-baseweb="select"] { margin-bottom: 10px; }

      .btn-row{
        margin-top: 14px;
        display: flex;
        justify-content: center;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------
# Cached loader (FAST)
# ---------------------------
@st.cache_data(ttl=600, show_spinner=False)
def load_tpm_ids_cached(sheet_id: str, tool_name: str, tpm_col: str) -> list[str]:
    ws = open_worksheet(sheet_id, tool_name)
    records = get_all_records(ws)
    return list_tpm_ids(records, tpm_col=tpm_col)


def _safe_load_tpm_ids(tool_name: str):
    try:
        ids = load_tpm_ids_cached(GOOGLE_SHEET_ID, tool_name, TPM_COL)
        return ids, None
    except Exception:
        return [], "Failed to load TPM list. Please check Google Sheets access and configuration."


# ---------------------------
# UI Card
# ---------------------------
st.markdown('<div class="login-wrap">', unsafe_allow_html=True)
st.markdown('<div class="login-card">', unsafe_allow_html=True)

st.markdown('<div class="login-title">WASH Pro</div>', unsafe_allow_html=True)
st.markdown('<div class="login-sub">Select a tool and TPM ID to continue.</div>', unsafe_allow_html=True)

default_tool_index = TOOLS.index("Tool 6") if "Tool 6" in TOOLS else 0
selected_tool = st.selectbox("Tool", TOOLS, index=default_tool_index)

tpm_ids, load_error = _safe_load_tpm_ids(selected_tool)

selected_tpm_id = st.selectbox("TPM ID", options=[""] + tpm_ids)

if load_error:
    st.error(load_error)

login_disabled = (not selected_tpm_id) or bool(load_error)

st.markdown('<div class="btn-row">', unsafe_allow_html=True)
login_clicked = st.button("Continue", type="primary", disabled=login_disabled, use_container_width=True)
st.markdown("</div>", unsafe_allow_html=True)

st.markdown('<div class="mini-help">Tip: choose TPM ID first, then continue.</div>', unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------
# Navigate
# ---------------------------
if login_clicked:
    st.session_state["selected_tool"] = selected_tool
    st.session_state["tpm_id"] = selected_tpm_id

    page_file = selected_tool.replace(" ", "_") + ".py"
    st.switch_page(f"pages/{page_file}")
