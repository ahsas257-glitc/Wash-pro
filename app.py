# app.py
import os
import sys
import streamlit as st

# Ensure the project root is in Python path (important for Streamlit multi-page apps)
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

st.set_page_config(page_title="WASH TPM", layout="wide")
st.title("WASH TPM Reporting")
st.write("Use the sidebar to open Home and then Tool pages.")
