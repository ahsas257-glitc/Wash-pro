# pages/home.py
import streamlit as st

from src.config import GOOGLE_SHEET_ID, TPM_COL, TOOLS
from src.data_processing import open_worksheet, get_all_records, list_tpm_ids


st.set_page_config(page_title="Home", layout="wide")
st.title("Home")

selected_tool = st.selectbox("Select a tool", TOOLS, index=TOOLS.index("Tool 6"))
import streamlit as st
st.write("Has gcp_service_account secrets?", "gcp_service_account" in st.secrets)

# Load TPM IDs from the selected worksheet
ws = open_worksheet(GOOGLE_SHEET_ID, selected_tool)
records = get_all_records(ws)
tpm_ids = list_tpm_ids(records, tpm_col=TPM_COL)

selected_tpm_id = st.selectbox(
    "Find or select a TPM ID (type to search)",
    options=[""] + tpm_ids,
)

if st.button("Search", type="primary", disabled=not selected_tpm_id):
    st.session_state["selected_tool"] = selected_tool
    st.session_state["tpm_id"] = selected_tpm_id

    # Map worksheet name 'Tool 6' -> page file 'Tool_6.py'
    page_file = selected_tool.replace(" ", "_") + ".py"
    st.switch_page(f"pages/{page_file}")

