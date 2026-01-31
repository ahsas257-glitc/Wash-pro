# src/theme.py
from pathlib import Path
import base64
import streamlit as st

def _img_to_data_uri(path: str) -> str:
    p = Path(path)
    if not p.exists():
        return ""
    b64 = base64.b64encode(p.read_bytes()).decode("utf-8")
    ext = p.suffix.lower().replace(".", "")
    mime = "png" if ext == "png" else ext
    return f"url('data:image/{mime};base64,{b64}')"

def _read_css(path: str) -> str:
    p = Path(path)
    return p.read_text(encoding="utf-8") if p.exists() else ""

def inject_theme(project_root: str):
    # paths
    logo = Path(project_root) / "assets" / "images" / "Logo_of_PPC.png"
    css1 = Path(project_root) / "design" / "css" / "liquid_glass.css"
    css2 = Path(project_root) / "design" / "css" / "animations.css"

    logo_uri = _img_to_data_uri(str(logo))

    css = f"""
    <style>
    :root {{
      --bg-logo: {logo_uri};
    }}
    {_read_css(str(css1))}
    {_read_css(str(css2))}
    </style>
    """

    st.markdown(css, unsafe_allow_html=True)

def glass_card(title: str = "", subtitle: str = ""):
    st.markdown('<div class="glass-card fade-in">', unsafe_allow_html=True)
    if title:
        st.markdown(f"#### {title}")
    if subtitle:
        st.caption(subtitle)

def end_glass_card():
    st.markdown("</div>", unsafe_allow_html=True)

def sticky_info(html_text: str):
    st.markdown(f'<div class="stickybar fade-in">{html_text}</div>', unsafe_allow_html=True)
