import streamlit as st
import base64
from pathlib import Path

# تبدیل عکس به base64
def _b64_image(path: str) -> str:
    return base64.b64encode(Path(path).read_bytes()).decode("utf-8")

# استایل و طراحی کلی
def apply_global_background(
    logo_path: str = "assets/images/Logo_of_PPC.png",
    logo_opacity_light: float = 0.08,
    logo_opacity_dark: float = 0.15,
    intensity: float = 1.0,
):
    logo_layer = ""
    if Path(logo_path).exists():
        try:
            b64_logo = _b64_image(logo_path)
            logo_layer = f"""
            .stApp::before {{
                content: "";
                position: fixed;
                inset: 0;
                z-index: 0;
                background: url("data:image/png;base64,{b64_logo}") center/contain no-repeat;
                opacity: var(--logo-opacity);
                filter: var(--logo-filter);
                pointer-events: none;
                animation: bgLogoFloat 30s ease-in-out infinite;
            }}
            """
        except Exception as e:
            st.warning(f"⚠️ Logo could not be loaded: {e}")
    else:
        st.warning("⚠️ Logo file not found. Check `logo_path`.")

    css = f"""
    <style>
    html, body, .stApp {{
        margin: 0;
        padding: 0;
        width: 100%;
        height: 100%;
        font-family: "Segoe UI", sans-serif;
    }}

    html {{
        --font-base: 22px;
        --font-small: 0.95rem;
        --font-label: 1rem;
        --font-h1: 2.4rem;
        --font-h2: 1.9rem;
        --font-h3: 1.5rem;
    }}

    body, .stApp {{
        font-size: var(--font-base);
        line-height: 1.7;
    }}

    .stMarkdown, .stText, .stCaption, .stMarkdown p {{
        font-size: var(--font-base);
        line-height: 1.9;
    }}

    label, .stTextInput label {{
        font-size: var(--font-label);
    }}

    input, textarea, select {{
        font-size: var(--font-base);
    }}

    .stMarkdown h1 {{ font-size: var(--font-h1) !important; }}
    .stMarkdown h2 {{ font-size: var(--font-h2) !important; }}
    .stMarkdown h3 {{ font-size: var(--font-h3) !important; }}

    @media (max-width: 768px) {{
        html {{
            --font-base: 14px;
            --font-h1: 1.8rem;
            --font-h2: 1.4rem;
            --font-h3: 1.2rem;
        }}
    }}

    footer {{
        visibility: hidden;
    }}

    footer:after {{
        content: "Made by Shabeer Ahmad Ahsas";
        visibility: visible;
        display: block;
        text-align: center;
        font-size: 1.9rem;
        color: #409C9B;
        padding: 15px;
    }}

    .block-container {{
        padding: 2rem 1rem !important;
        max-width: 960px;
        margin: auto;
    }}

    .stApp {{
        background: var(--base-bg);
        overflow-x: hidden;
    }}

    .stApp::after {{
        content: "";
        position: fixed;
        inset: 0;
        z-index: -1;
        background:
            radial-gradient(800px 700px at 30% 20%, rgba(var(--c1),0.25), transparent 60%),
            radial-gradient(800px 650px at 80% 30%, rgba(var(--c2),0.2), transparent 58%);
        filter: blur(30px);
        opacity: {intensity};
        pointer-events: none;
    }}

    section[data-testid="stSidebar"],
    .main > div,
    div[data-testid="stExpander"],
    div[data-testid="stContainer"],
    div[data-testid="stVerticalBlock"] > div {{
        background: var(--card-bg) !important;
        border-radius: 16px;
        border: 1px solid var(--card-border);
        box-shadow: var(--card-shadow);
        padding: 1.5rem;
        margin-bottom: 1.5rem;
        transition: all 0.3s ease;
        backdrop-filter: blur(14px) saturate(1.1);
        animation: floatCard 6s ease-in-out infinite, fadeInUp 0.9s ease-out forwards;
        will-change: transform, opacity;
        opacity: 0;
    }}

    @keyframes floatCard {{
        0% {{
            transform: translateY(0px) scale(4);
            box-shadow: 0 20px 20px rgba(0,0,0,0.1);
        }}
        50% {{
            transform: translateY(-3px) scale(5.01);
            box-shadow: 0 30px 24px rgba(0,0,0,0.15);
        }}
        100% {{
            transform: translateY(0px) scale(1);
            box-shadow: 0 20px 20px rgba(0,0,0,0.1);
        }}
    }}

    @keyframes fadeInUp {{
        0% {{
            transform: translateY(20px);
            opacity: 0;
        }}
        100% {{
            transform: translateY(0);
            opacity: 1;
        }}
    }}

    @keyframes bgLogoFloat {{
        0% {{ transform: scale(1.05) rotate(0deg); }}
        50% {{ transform: scale(1.0) rotate(0.9deg); }}
        100% {{ transform: scale(1.05) rotate(0deg); }}
    }}

    button {{
        background: var(--btn-bg) !important;
        color: var(--btn-color) !important;
        border: none !important;
        border-radius: 12px !important;
        padding: .7rem 1.2rem !important;
        font-weight: bold;
        transition: transform .2s ease-in-out;
        font-size: var(--font-base);
    }}

    button:hover {{
        transform: scale(1.02);
        box-shadow: 0px 6px 16px rgba(0,0,0,0.2);
    }}

    html[data-theme="dark"] {{
        --base-bg: linear-gradient(180deg, #02030a, #0b0c12);
        --c1: 60,160,255;
        --c2: 255,120,50;

        --card-bg: rgba(12,13,27,0.45);
        --card-border: rgba(255,255,255,0.12);
        --card-shadow: 0 14px 40px rgba(0,0,0,0.6);

        --text-primary: #e5e9f0;
        --input-bg: rgba(30,30,42,0.55);
        --input-border: rgba(255,255,255,0.22);
        --input-text: #fff;
        --btn-bg: #2e74e1;
        --btn-color: #fafafa;
        --logo-opacity: {logo_opacity_dark};
        --logo-filter: brightness(1.15) contrast(1.1);
    }}

    html[data-theme="light"] {{
        --base-bg: linear-gradient(180deg, #eef2f7, #dfe6f2);
        --c1: 40,130,255;
        --c2: 255,140,55;

        --card-bg: rgba(255,255,255,0.6);
        --card-border: rgba(200,200,210,0.35);
        --card-shadow: 0 12px 24px rgba(0,0,0,0.1);

        --text-primary: #222;
        --input-bg: #ffffff;
        --input-border: #9ca3af;
        --input-text: #111;
        --btn-bg: #1f78d1;
        --btn-color: #fff;
        --logo-opacity: {logo_opacity_light};
        --logo-filter: brightness(1.25) contrast(1.05);
    }}

    table {{
        width: 100%;
        border-collapse: collapse;
        margin-top: 1rem;
        font-size: calc(var(--font-base) * 0.95);
        background-color: rgba(255, 255, 255, 0.05);
        backdrop-filter: blur(8px);
        border: 1px solid rgba(255,255,255,0.1);
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 4px 18px rgba(0,0,0,0.1);
    }}

    thead {{
        background-color: rgba(255, 255, 255, 0.08);
    }}

    th, td {{
        text-align: left;
        padding: 0.6rem 1rem;
        border-bottom: 1px solid rgba(255,255,255,0.08);
    }}

    tr:last-child td {{
        border-bottom: none;
    }}

    th {{
        font-weight: bold;
        color: var(--text-primary);
    }}

    td {{
        color: var(--text-primary);
    }}

    .sidebar-footer {{
        position: absolute;
        bottom: 1.2rem;
        left: 1rem;
        right: 1rem;
        font-size: 0.85rem;
        color: #ccc;
        text-align: center;
        opacity: 0.8;
    }}

    {logo_layer}
    </style>
    """

    st.markdown(css, unsafe_allow_html=True)
