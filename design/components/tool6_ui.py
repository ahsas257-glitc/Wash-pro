# design/components/tool6_ui.py
import base64
from pathlib import Path
import streamlit as st

# ---------------------------------
# Internal helpers
# ---------------------------------
def _read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except Exception:
        return ""

def _img_to_data_uri(path: Path) -> str:
    if not path.exists():
        return ""
    ext = path.suffix.lower().replace(".", "")
    mime = "png" if ext == "png" else ext
    data = base64.b64encode(path.read_bytes()).decode("utf-8")
    return f"data:image/{mime};base64,{data}"

def _inject_parallax_js():
    st.markdown(
        """
<script>
(function(){
  const root = window.parent.document;
  const app = root.querySelector('[data-testid="stAppViewContainer"]');
  if(!app) return;

  if(app.dataset.parallaxInit === "1") return;
  app.dataset.parallaxInit = "1";

  root.addEventListener("mousemove", (e) => {
    const w = window.innerWidth || 1200;
    const h = window.innerHeight || 800;
    const cx = (e.clientX / w - 0.5);
    const cy = (e.clientY / h - 0.5);

    // subtle premium parallax
    app.style.transform = `translate3d(${cx*2}px, ${cy*2}px, 0)`;
  }, {passive:true});
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
        const rx = (-y * 7).toFixed(2);
        const ry = (x * 9).toFixed(2);
        card.style.transform =
          `perspective(900px) rotateX(${rx}deg) rotateY(${ry}deg) translateY(-3px)`;
      }, {passive:true});

      card.addEventListener("mouseleave", () => {
        card.style.transform =
          "perspective(900px) rotateX(0deg) rotateY(0deg) translateY(0px)";
      });
    });
  };

  // Streamlit rerenders: run twice
  init();
  setTimeout(init, 350);
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
    theme: str = "dark",  # ✅ NEW: "dark" یا "light"
    background_image_rel: str = "assets/images/Logo_of_PPC.png",
    noise_image_rel: str = "",
    intensity: float = 0.10,      # 0.08 تا 0.14 عالیه
    enable_parallax: bool = True,
):

    root = Path(project_root) if project_root else Path(".")

    # ✅ pick correct theme css
    theme = (theme or "dark").strip().lower()
    css_file = "liquid_glass_light.css" if theme == "light" else "liquid_glass_dark.css"

    css_liquid = _read_text(root / "design" / "css" / css_file)
    css_anim = _read_text(root / "design" / "css" / "animations.css")

    bg_uri = _img_to_data_uri(root / background_image_rel) if background_image_rel else ""
    noise_uri = _img_to_data_uri(root / noise_image_rel) if noise_image_rel else ""

    st.markdown(
        f"""
<style>
{css_liquid}
{css_anim}

/* ---- App variables ---- */
:root {{
  --bg-logo: url("{bg_uri}");
  --bg-intensity: {float(intensity)};
  /* اگر در liquid_glass_dark.css این تعریف شده باشد، همین استفاده می‌شود */
  --app-bg: var(--app-bg, #0b1220);
}}

/* ---- Background: single PPC logo fully visible + fallback color ---- */
[data-testid="stAppViewContainer"]::before {{
  content:"";
  position: fixed;
  inset: 0;
  z-index: 0;
  pointer-events: none;

  /* ✅ اگر لوگو کل صفحه را پوشش ندهد، همین رنگ نمایش می‌شود */
  background-color: var(--app-bg);

  /* ✅ لوگو کامل دیده شود */
  background-image: var(--bg-logo);
  background-repeat: no-repeat;
  background-position: center center;
  background-size: contain;              /* ✅ FULLY VISIBLE */

  opacity: var(--bg-intensity);

  /* ✅ حرکت آهسته و نرم */
  animation: bgSlowFloat 55s ease-in-out infinite;
  transform: translate3d(0,0,0);
  will-change: background-position;
}}

@keyframes bgSlowFloat {{
  0%   {{ background-position: 50% 50%; }}
  50%  {{ background-position: 52% 47%; }}
  100% {{ background-position: 50% 50%; }}
}}

/* ---- Ensure content is above background ---- */
.block-container {{
  position: relative;
  z-index: 2;
}}
</style>
        """,
        unsafe_allow_html=True,
    )

    # Optional noise overlay
    if noise_uri:
        st.markdown(
            f"""
<style>
.noise-layer {{
  position: fixed;
  inset: 0;
  z-index: 1;
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
</style>
<div class="noise-layer"></div>
            """,
            unsafe_allow_html=True,
        )

    if enable_parallax:
        _inject_parallax_js()

    _inject_tilt_js()

# ---------------------------------
# UI building blocks
# ---------------------------------
def card_open(title: str, subtitle: str = "") -> None:
    # ✅ tilt + shimmer + glass look
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
