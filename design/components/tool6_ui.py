# design/components/tool6_ui.py
import base64
from pathlib import Path
import streamlit as st

# ---------------------------------
# UI: Topbar
# ---------------------------------
def topbar(title: str, subtitle: str = "", right_chip: str = "") -> None:
    st.markdown(
        f"""
<div class="t6-topbar">
  <div class="t6-topbar__left">
    <div>
      <div class="t6-topbar__title">{title}</div>
      <div class="t6-topbar__muted">{subtitle}</div>
    </div>
  </div>
  <div class="t6-topbar__right">
    {f'<div class="t6-chip">{right_chip}</div>' if right_chip else ''}
  </div>
</div>
        """,
        unsafe_allow_html=True,
    )

# ---------------------------------
# Internal helpers
# ---------------------------------
def _read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except Exception:
        return ""

def _guess_image_mime(ext: str) -> str:
    ext = (ext or "").lower().lstrip(".")
    if ext in ("jpg", "jpeg"):
        return "image/jpeg"
    if ext == "png":
        return "image/png"
    if ext == "webp":
        return "image/webp"
    if ext == "svg":
        return "image/svg+xml"
    if ext == "gif":
        return "image/gif"
    # fallback
    return "image/png"

def _img_to_data_uri(path: Path) -> str:
    if not path or not path.exists():
        return ""
    mime = _guess_image_mime(path.suffix)
    data = base64.b64encode(path.read_bytes()).decode("utf-8")
    return f"data:{mime};base64,{data}"

def _inject_theme_sync_js():
    """
    Sync Streamlit theme with CSS by setting:
    html[data-theme="dark|light"]
    """
    st.markdown(
        """
<script>
(function(){
  const rootDoc = window.parent.document;
  const html = rootDoc.documentElement;

  function parseColorToRGB(s){
    if(!s) return null;
    s = (s||"").trim();
    let m = s.match(/^rgba?\\((\\d+)\\s*,\\s*(\\d+)\\s*,\\s*(\\d+)/i);
    if(m) return {r:+m[1], g:+m[2], b:+m[3]};
    if(s[0] === '#'){
      let hex = s.slice(1);
      if(hex.length === 3) hex = hex.split('').map(c => c+c).join('');
      if(hex.length >= 6){
        return { r: parseInt(hex.slice(0,2),16),
                 g: parseInt(hex.slice(2,4),16),
                 b: parseInt(hex.slice(4,6),16) };
      }
    }
    return null;
  }

  function luminance(rgb){
    return (0.2126*rgb.r + 0.7152*rgb.g + 0.0722*rgb.b)/255;
  }

  function getStreamlitBg(){
    const styles = getComputedStyle(html);
    const bg = styles.getPropertyValue('--background-color') ||
               styles.getPropertyValue('--st-bg') ||
               styles.getPropertyValue('--st-background-color') ||
               '';
    return (bg || '').trim();
  }

  function decideTheme(){
    const bg = getStreamlitBg();
    const rgb = parseColorToRGB(bg);
    if(!rgb){
      return (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) ? 'dark' : 'light';
    }
    return luminance(rgb) < 0.5 ? 'dark' : 'light';
  }

  function applyTheme(){
    const t = decideTheme();
    if(html.getAttribute('data-theme') !== t){
      html.setAttribute('data-theme', t);
    }
  }

  applyTheme();
  const obs = new MutationObserver(() => applyTheme());
  obs.observe(html, { attributes: true, attributeFilter: ['style', 'class'] });

  // light touch: less CPU than 800ms
  setInterval(applyTheme, 1500);
})();
</script>
        """,
        unsafe_allow_html=True,
    )

def _inject_single_container_mark_js():
    """
    Mark ONLY one stAppViewContainer as the watermark host: [data-ppc-wm="1"]
    Prevents double watermark when Streamlit renders multiple containers.
    """
    st.markdown(
        """
<script>
(function(){
  const root = window.parent.document;

  function pickBest(nodes){
    for(const n of nodes){
      if(n.querySelector('.main, [data-testid="stSidebar"], [data-testid="stHeader"]')){
        return n;
      }
    }
    return nodes[0];
  }

  function mark(){
    const nodes = root.querySelectorAll('[data-testid="stAppViewContainer"]');
    if(!nodes || !nodes.length) return;

    const best = pickBest(nodes);
    nodes.forEach(n => n.removeAttribute('data-ppc-wm'));
    best.setAttribute('data-ppc-wm', '1');
  }

  mark();

  // observe DOM changes (no need to also do fast interval)
  const obs = new MutationObserver(mark);
  obs.observe(root.documentElement, { childList:true, subtree:true });
})();
</script>
        """,
        unsafe_allow_html=True,
    )

def _inject_watermark_parallax_js():
    """
    Parallax ONLY for watermark vars (no content transform!)
    Updates: --wm-par-x, --wm-par-y on the marked container.
    """
    st.markdown(
        """
<script>
(function(){
  const root = window.parent.document;

  const getHost = () =>
    root.querySelector('[data-testid="stAppViewContainer"][data-ppc-wm="1"]');

  const init = () => {
    const host = getHost();
    if(!host) return;

    if(host.dataset.wmParallaxInit === "1") return;
    host.dataset.wmParallaxInit = "1";

    let raf = 0;
    root.addEventListener("mousemove", (e) => {
      if(raf) return;
      raf = requestAnimationFrame(() => {
        raf = 0;
        const w = window.innerWidth || 1200;
        const h = window.innerHeight || 800;
        const cx = (e.clientX / w - 0.5);
        const cy = (e.clientY / h - 0.5);
        host.style.setProperty('--wm-par-x', (cx * 16).toFixed(2) + 'px');
        host.style.setProperty('--wm-par-y', (cy * 12).toFixed(2) + 'px');
      });
    }, {passive:true});
  };

  init();
  setTimeout(init, 350);
})();
</script>
        """,
        unsafe_allow_html=True,
    )

def _inject_tilt_js():
    st.markdown(
        """
<script>
(function(){
  const root = window.parent.document;

  const init = () => {
    const nodes = root.querySelectorAll(".js-tilt");
    nodes.forEach((card) => {
      if(card.dataset.tiltInit === "1") return;
      card.dataset.tiltInit = "1";

      card.addEventListener("mousemove", (e) => {
        const r = card.getBoundingClientRect();
        const x = (e.clientX - r.left) / r.width - 0.5;
        const y = (e.clientY - r.top) / r.height - 0.5;
        const rx = (-y * 6).toFixed(2);
        const ry = (x * 8).toFixed(2);
        card.style.transform =
          `perspective(900px) rotateX(${rx}deg) rotateY(${ry}deg) translateY(-2px)`;
      }, {passive:true});

      card.addEventListener("mouseleave", () => {
        card.style.transform =
          "perspective(900px) rotateX(0deg) rotateY(0deg) translateY(0px)";
      });
    });
  };

  init();
  setTimeout(init, 350);

  // re-init if Streamlit re-renders parts of DOM
  const obs = new MutationObserver(() => init());
  obs.observe(root.documentElement, { childList:true, subtree:true });
})();
</script>
        """,
        unsafe_allow_html=True,
    )

# ---------------------------------
# Public: inject_tool6_design
# ---------------------------------
def inject_tool6_design(
    project_root: str = "",
    background_image_rel: str = "assets/images/Logo_of_PPC.png",
    noise_image_rel: str = "",
    intensity: float = 0.10,
    enable_parallax: bool = True,
):
    root = Path(project_root) if project_root else Path(".")

    # Unified CSS (Light/Dark inside it)
    css_liquid = _read_text(root / "design" / "css" / "liquid_glass.css")
    css_anim = _read_text(root / "design" / "css" / "animations.css")

    bg_uri = _img_to_data_uri(root / background_image_rel) if background_image_rel else ""
    noise_uri = _img_to_data_uri(root / noise_image_rel) if noise_image_rel else ""

    st.markdown(
        f"""
<style>
{css_liquid}
{css_anim}

/* SAFE RESET: do NOT touch html/body/.stApp background */
[data-testid="stAppViewContainer"],
[data-testid="stMain"],
[data-testid="stSidebar"] {{
  background-image: none !important;
}}

/* Watermark vars */
:root {{
  --bg-logo: url("{bg_uri}");
  --bg-intensity: {float(intensity)};
  --wm-par-x: 0px;
  --wm-par-y: 0px;
  --wm-size: min(68vmin, 760px);
}}

[data-testid="stAppViewContainer"] {{
  position: relative;
}}

/* kill duplicates globally */
[data-testid="stAppViewContainer"]::before,
[data-testid="stAppViewContainer"]::after {{
  content: none !important;
}}

/* watermark only on the marked container */
[data-testid="stAppViewContainer"][data-ppc-wm="1"]::before {{
  content:"";
  position: fixed;
  inset: 0;
  z-index: 0;
  pointer-events: none;

  background-image: var(--bg-logo);
  background-repeat: no-repeat;
  background-position: 50% 50%;
  background-size: var(--wm-size);

  opacity: var(--bg-intensity);

  transform: translate3d(var(--wm-par-x), var(--wm-par-y), 0);
  will-change: transform, opacity;
}}

/* keep content above watermark */
.block-container {{
  position: relative;
  z-index: 10;
}}

/* optional noise layer (between watermark and content) */
{"".join([
f"""
.noise-layer {{
  position: fixed;
  inset: 0;
  z-index: 2;
  pointer-events: none;
  opacity: .06;
  background-image: url("{noise_uri}");
  background-repeat: repeat;
  mix-blend-mode: overlay;
  animation: noiseMove 6s steps(2) infinite;
}}
@keyframes noiseMove {{
  0%{{ transform: translate3d(0,0,0); }}
  25%{{ transform: translate3d(-1%,1%,0); }}
  50%{{ transform: translate3d(1%,-1%,0); }}
  75%{{ transform: translate3d(-1%,-1%,0); }}
  100%{{ transform: translate3d(0,0,0); }}
}}
""" if noise_uri else ""
])}
</style>
        """,
        unsafe_allow_html=True,
    )

    # ✅ inject layers ONCE (fix: previously duplicated)
    if noise_uri:
        st.markdown('<div class="noise-layer"></div>', unsafe_allow_html=True)

    st.markdown('<div class="t6-liquid-overlay"></div>', unsafe_allow_html=True)

    # order matters
    _inject_single_container_mark_js()
    _inject_theme_sync_js()

    if enable_parallax:
        _inject_watermark_parallax_js()

    _inject_tilt_js()

# ---------------------------------
# UI building blocks
# ---------------------------------
def card_open(title: str, subtitle: str = "") -> None:
    st.markdown('<div class="glass-card fade-in shimmer js-tilt">', unsafe_allow_html=True)
    if title:
        st.markdown(f"#### {title}")
    if subtitle:
        st.caption(subtitle)

def card_close(hint: str = "") -> None:
    if hint:
        st.caption(hint)
    st.markdown("</div>", unsafe_allow_html=True)

def kv(label: str, value) -> None:
    v = "" if value is None else str(value).strip()
    st.markdown(
        f'<div class="kv"><span class="label">{label}</span><span class="value">{v or "—"}</span></div>',
        unsafe_allow_html=True,
    )

def sticky_open() -> None:
    st.markdown('<div class="stickybar fade-in">', unsafe_allow_html=True)

def sticky_close() -> None:
    st.markdown("</div>", unsafe_allow_html=True)
