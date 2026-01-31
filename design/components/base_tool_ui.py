from __future__ import annotations

import base64
import textwrap
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import streamlit as st


# ============================================================
# Internal helpers
# ============================================================
def _html(s: str) -> str:
    return textwrap.dedent(s).strip()


def _read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except Exception:
        return ""


def _guess_image_mime(ext: str) -> str:
    return {
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "png": "image/png",
        "webp": "image/webp",
        "svg": "image/svg+xml",
        "gif": "image/gif",
    }.get((ext or "").lower().lstrip("."), "image/png")


def _img_to_data_uri(path: Path) -> str:
    if not path or not path.exists():
        return ""
    data = base64.b64encode(path.read_bytes()).decode("utf-8")
    return f"data:{_guess_image_mime(path.suffix)};base64,{data}"


# ============================================================
# Config
# ============================================================
@dataclass(frozen=True)
class BaseToolDesignConfig:
    project_root: str = ""
    background_image_rel: str = ""
    background_opacity: float = 0.08
    background_size: str = "min(62vmin, 720px)"
    background_position: str = "50% 50%"


# ============================================================
# Public: inject design (CSS + JS once)
# ============================================================
def inject_base_tool_design(cfg: Optional[BaseToolDesignConfig] = None) -> None:
    key = "base_tool_ui__injected"
    if st.session_state.get(key):
        return
    st.session_state[key] = True

    cfg = cfg or BaseToolDesignConfig()
    root = Path(cfg.project_root) if cfg.project_root else Path(".")

    css_liquid = _read_text(root / "design/css/liquid_glass.css")
    css_anim = _read_text(root / "design/css/animations.css")
    bg_uri = _img_to_data_uri(root / cfg.background_image_rel) if cfg.background_image_rel else ""

    st.markdown(
        _html(f"""
<style>
{css_liquid}
{css_anim}

:root {{
  --bt-bg-image: {f'url("{bg_uri}")' if bg_uri else 'none'};
  --bt-bg-opacity: {cfg.background_opacity};
  --bt-bg-size: {cfg.background_size};
  --bt-bg-position: {cfg.background_position};
}}

[data-testid="stAppViewContainer"]::before {{
  content:"";
  position:fixed;
  inset:0;
  pointer-events:none;
  z-index:0;
  opacity:var(--bt-bg-opacity);
  background-image:var(--bt-bg-image);
  background-repeat:no-repeat;
  background-position:var(--bt-bg-position);
  background-size:var(--bt-bg-size);
}}

.block-container {{
  position:relative;
  z-index:10;
}}
</style>
        """),
        unsafe_allow_html=True,
    )

    _inject_theme_sync_js()
    _inject_tilt_js()


# ============================================================
# JS
# ============================================================
def _inject_theme_sync_js() -> None:
    st.markdown(
        _html("""
<script>
(function(){
  const html = window.parent.document.documentElement;
  const getBg = () => getComputedStyle(html).getPropertyValue('--background-color');
  const isDark = (c)=>{const m=c.match(/\\d+/g);return m && (0.2126*m[0]+0.7152*m[1]+0.0722*m[2])<128};
  const apply = ()=>html.setAttribute('data-theme', isDark(getBg())?'dark':'light');
  apply(); setInterval(apply,1500);
})();
</script>
        """),
        unsafe_allow_html=True,
    )


def _inject_tilt_js() -> None:
    st.markdown(
        _html("""
<script>
(function(){
  const root=window.parent.document;
  const init=()=>root.querySelectorAll('.js-tilt').forEach(c=>{
    if(c.dataset.tilt) return; c.dataset.tilt=1;
    c.onmousemove=e=>{
      const r=c.getBoundingClientRect();
      const x=(e.clientX-r.left)/r.width-.5;
      const y=(e.clientY-r.top)/r.height-.5;
      c.style.transform=`perspective(900px) rotateX(${-y*6}deg) rotateY(${x*8}deg)`;
    };
    c.onmouseleave=()=>c.style.transform='';
  });
  init(); new MutationObserver(init).observe(root,{childList:true,subtree:true});
})();
</script>
        """),
        unsafe_allow_html=True,
    )


# ============================================================
# UI blocks (shared by all tools)
# ============================================================
def topbar(title: str, subtitle: str = "", right_chip: str = "") -> None:
    st.markdown(
        _html(f"""
<div class="t6-topbar">
  <div>
    <div class="t6-topbar__title">{title}</div>
    <div class="t6-topbar__muted">{subtitle}</div>
  </div>
  {f'<div class="t6-chip">{right_chip}</div>' if right_chip else ''}
</div>
        """),
        unsafe_allow_html=True,
    )


def card_open(title: str = "", subtitle: str = "", variant: str = "") -> None:
    st.markdown(
        _html(f'<div class="glass-card lg-card lg-pad fade-in js-tilt {variant}">'),
        unsafe_allow_html=True,
    )
    if title:
        st.markdown(f"#### {title}")
    if subtitle:
        st.caption(subtitle)


def card_close() -> None:
    st.markdown("</div>", unsafe_allow_html=True)


def status_card(title: str, description: str = "", level: str = "success") -> None:
    cls = {"success":"lg-success","warning":"lg-warning","error":"lg-error","info":"lg-info"}.get(level,"lg-success")
    st.markdown(
        _html(f"""
<div class="glass-card lg-card lg-pad js-tilt {cls}">
  <strong>{title}</strong>
  <div style="opacity:.85;margin-top:6px">{description}</div>
</div>
        """),
        unsafe_allow_html=True,
    )


def table_card_open(title: str) -> None:
    st.markdown(_html(f'<div class="lg-table-wrap"><h4>{title}</h4><div>'), unsafe_allow_html=True)


def table_card_close() -> None:
    st.markdown("</div></div>", unsafe_allow_html=True)


def kv(label: str, value: str) -> None:
    st.markdown(_html(f'<div class="kv"><span>{label}</span><strong>{value}</strong></div>'), unsafe_allow_html=True)
