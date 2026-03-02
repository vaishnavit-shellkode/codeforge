import os
import sys
import uuid
import base64

import streamlit as st
import streamlit.components.v1 as components

# -----------------------------------------------------------------------------
# Import your existing logic from image.py
# -----------------------------------------------------------------------------
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

try:
    from image import (
        prepare_slides,
        export_pptx,
        detect_mode,
        pick_theme,
        is_process_topic,
        reassemble_diagram,
        OUTPUT_DIR,
        THEME_COLORS,
    )
except ImportError as e:
    st.error(f"Import error while loading image.py: {e}")
    st.stop()


# -----------------------------------------------------------------------------
# Streamlit page config & base styles
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="AI PPT Studio",
    page_icon="📊",
    layout="wide",
)

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Outfit', sans-serif;
}

.stApp {
    background: #0f172a;
    background-image: 
        radial-gradient(at 0% 0%, rgba(79, 70, 229, 0.15) 0px, transparent 50%),
        radial-gradient(at 100% 100%, rgba(139, 92, 246, 0.15) 0px, transparent 50%);
}

.block-container {
    padding-top: 2rem;
    padding-bottom: 4rem;
}

/* Premium Glass Card */
.glass-card {
    background: rgba(30, 41, 59, 0.7);
    backdrop-filter: blur(12px);
    -webkit-backdrop-filter: blur(12px);
    border-radius: 24px;
    border: 1px solid rgba(255, 255, 255, 0.1);
    box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
    padding: 28px;
    margin-bottom: 24px;
    transition: transform 0.3s ease, border 0.3s ease;
}

.glass-card:hover {
    border: 1px solid rgba(99, 102, 241, 0.3);
}

/* Typography */
h1, h2, h3, h4, h5, h6 {
    color: #f8fafc !important;
    font-weight: 700 !important;
}

p, span, label {
    color: #94a3b8 !important;
}

/* Buttons */
.stButton>button {
    background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 100%);
    color: white !important;
    border-radius: 12px !important;
    border: none !important;
    padding: 0.6rem 1.5rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.025em !important;
    transition: all 0.3s ease !important;
    box-shadow: 0 10px 15px -3px rgba(79, 70, 229, 0.4) !important;
}

.stButton>button:hover {
    transform: translateY(-2px);
    box-shadow: 0 20px 25px -5px rgba(79, 70, 229, 0.5) !important;
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%);
}

/* Text Area */
.stTextArea textarea {
    background: rgba(15, 23, 42, 0.6) !important;
    color: #f1f5f9 !important;
    border-radius: 16px !important;
    border: 1px solid rgba(255, 255, 255, 0.1) !important;
    padding: 16px !important;
}

.stTextArea textarea:focus {
    border-color: #6366f1 !important;
    box-shadow: 0 0 0 2px rgba(99, 102, 241, 0.2) !important;
}

/* Slide Chips */
.slide-chip {
    border-radius: 16px;
    border: 1px solid rgba(255, 255, 255, 0.05);
    padding: 12px 16px;
    margin-bottom: 10px;
    background: rgba(30, 41, 59, 0.5);
    display: flex;
    justify-content: space-between;
    align-items: center;
    cursor: pointer;
    transition: all 0.2s ease;
}

.slide-chip:hover {
    background: rgba(51, 65, 85, 0.8);
    transform: translateX(4px);
}

.slide-chip.active {
    background: linear-gradient(135deg, rgba(79, 70, 229, 0.2), rgba(124, 58, 237, 0.2));
    border-color: #6366f1;
    box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
}

.active-slide-marker {
    width: 6px;
    height: 6px;
    background: #6366f1;
    border-radius: 50%;
    box-shadow: 0 0 10px #6366f1;
}

/* Stat Pill */
.stat-pill {
    background: rgba(99, 102, 241, 0.1);
    color: #818cf8;
    border: 1px solid rgba(99, 102, 241, 0.2);
    border-radius: 9999px;
    padding: 4px 12px;
    font-size: 11px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}

header[data-testid="stHeader"] {display: none;}
footer {visibility: hidden;}
</style>
""",
    unsafe_allow_html=True,
)


# -----------------------------------------------------------------------------
# Small utilities (base64 images + theme colors)
# -----------------------------------------------------------------------------
def _b64(path: str) -> str:
    if not path or not os.path.exists(path):
        return ""
    ext = os.path.splitext(path)[-1].lower().strip(".")
    mime = {"jpg": "jpeg", "jpeg": "jpeg", "png": "png", "webp": "webp"}.get(ext, "png")
    with open(path, "rb") as f:
        data = base64.b64encode(f.read()).decode()
    return f"data:image/{mime};base64,{data}"


def _colors(theme):
    try:
        return THEME_COLORS[theme]
    except Exception:
        return {"primary": "1E2761", "secondary": "CADCFC", "accent": "FFFFFF"}


def esc(t: str) -> str:
    return str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# -----------------------------------------------------------------------------
# Slide preview (read‑only, nice looking)
# -----------------------------------------------------------------------------
def render_slide_preview(slide, theme) -> None:
    """Render a single slide visually using HTML inside an iframe."""
    ct = slide.content_type.value
    c = slide.content
    col = _colors(theme)
    P = f"#{col['primary']}"
    S = f"#{col['secondary']}"

    if ct == "title":
        bg = _b64(c.get("domain_bg_path", ""))
        bg_css = (
            f"background:url('{bg}') center/cover no-repeat;"
            if bg
            else f"background:radial-gradient(circle at top left,{S}33,{P});"
        )
        body = f"""
        <div style="position:absolute;inset:0;{bg_css}">
          <div style="position:absolute;inset:0;background:linear-gradient(135deg,rgba(15,23,42,0.55),rgba(15,23,42,0.25));"></div>
          <div style="position:absolute;inset:0;display:flex;flex-direction:column;align-items:flex-start;justify-content:center;padding:8% 10%;text-align:left;">
            <div style="font-size:40px;font-weight:800;color:#fff;line-height:1.1;margin-bottom:12px;max-width:70%;">
              {esc(c.get("title", ""))}
            </div>
            <div style="width:70px;height:4px;background:{S};border-radius:999px;margin-bottom:14px;"></div>
            <div style="font-size:18px;color:{S};font-weight:500;max-width:60%;">
              {esc(c.get("subtitle", ""))}
            </div>
          </div>
        </div>
        """

    elif ct == "bullets":
        img = _b64(c.get("generated_image_path", ""))
        bullets_html = "".join(
            f"""
            <li style="margin-bottom:8px;">
              <span style="font-size:15px;color:#e5e7eb;line-height:1.5;">{esc(b)}</span>
            </li>
            """
            for b in c.get("bullets", [])
        )
        image_panel = (
            f'<div style="flex:1.1;border-radius:14px;overflow:hidden;background:#020617;"><img src="{img}" style="width:100%;height:100%;object-fit:cover;"/></div>'
            if img
            else ""
        )
        body = f"""
        <div style="position:absolute;inset:0;background:radial-gradient(circle at top left,#0b1120,#020617);">
          <div style="position:absolute;inset:0;padding:7% 8%;display:flex;flex-direction:column;">
            <div style="font-size:24px;font-weight:700;color:#f9fafb;margin-bottom:20px;">
              {esc(c.get("title", ""))}
            </div>
            <div style="display:flex;gap:26px;flex:1;align-items:flex-start;">
              <div style="flex:1.2;">
                <ul style="padding-left:1.2rem;margin:0;list-style:disc;">{bullets_html}</ul>
              </div>
              {image_panel}
            </div>
          </div>
        </div>
        """

    else:
        body = f"""
        <div style="position:absolute;inset:0;background:{P};display:flex;align-items:center;justify-content:center;">
          <div style="font-size:26px;font-weight:700;color:#fff;">
            {esc(c.get("title", slide.content_type.value.title()))}
          </div>
        </div>
        """

    html_doc = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<style>
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  html, body {{
    width:100%; height:100%;
    background:transparent;
    display:flex; align-items:center; justify-content:center;
  }}
  .frame {{
    width:100%;
    aspect-ratio:16/9;
    border-radius:18px;
    overflow:hidden;
    box-shadow:0 22px 70px rgba(15,23,42,0.55);
    font-family:'Inter',system-ui,sans-serif;
  }}
</style>
</head>
<body>
  <div class="frame">{body}</div>
</body>
</html>"""

    components.html(html_doc, height=420, scrolling=False)


# -----------------------------------------------------------------------------
# Session state
# -----------------------------------------------------------------------------
if "slides" not in st.session_state:
    st.session_state.slides = None
if "theme" not in st.session_state:
    st.session_state.theme = None
if "selected_idx" not in st.session_state:
    st.session_state.selected_idx = 0
if "ppt_bytes" not in st.session_state:
    st.session_state.ppt_bytes = None
if "last_prompt" not in st.session_state:
    st.session_state.last_prompt = ""


# -----------------------------------------------------------------------------
# HERO HEADER
# -----------------------------------------------------------------------------
hero_left, hero_right = st.columns([3, 1])

with hero_left:
    st.markdown(
        """
<div class="glass-card">
  <div style="display:flex;align-items:center;gap:12px;margin-bottom:12px;">
    <div class="stat-pill">GEN ALPHA ✨</div>
    <span style="font-size:12px;color:#64748b;">Enterprise AI Presentation Engine</span>
  </div>
  <div style="font-size:32px;font-weight:800;color:#f8fafc;margin-bottom:8px;letter-spacing:-0.02em;">
    Narrative into <span style="background:linear-gradient(135deg, #818cf8, #c084fc); -webkit-background-clip:text; -webkit-text-fill-color:transparent;">Visual Excellence.</span>
  </div>
  <div style="font-size:15px;color:#94a3b8;line-height:1.5;">
    Drop a topic or paste a full manuscript. Our model architectures your content into 
    professionally designed, theme-aware PowerPoint decks in seconds.
  </div>
</div>
""",
        unsafe_allow_html=True,
    )

with hero_right:
    status_color = "#34d399" if st.session_state.slides else "#6366f1"
    status_text = "Deck Ready" if st.session_state.slides else "System Idle"
    st.markdown(
        f"""
<div class="glass-card" style="text-align:center; padding: 24px;">
  <div style="font-size:11px; color:#64748b; text-transform:uppercase; letter-spacing:0.1em; margin-bottom:12px;">Engine Status</div>
  <div style="font-size:20px; font-weight:700; color:{status_color}; margin-bottom:4px;">
    {status_text}
  </div>
  <div style="height:4px; width:40px; background:{status_color}; border-radius:99px; margin: 12px auto 0;"></div>
</div>
""",
        unsafe_allow_html=True,
    )


st.markdown("")


# -----------------------------------------------------------------------------
# MAIN LAYOUT: left = input / controls, right = preview area
# -----------------------------------------------------------------------------
left, right = st.columns([1.2, 1.8])

# ----------------------------- LEFT PANEL ------------------------------------
with left:
    st.markdown('<div class="glass-card">', unsafe_allow_html=True)

    st.markdown("#### ✍️ Describe your presentation")

    user_input = st.text_area(
        "",
        height=180,
        label_visibility="collapsed",
        placeholder="Example: \"The modern lending technology landscape\" or paste full meeting notes / document here…",
        value=st.session_state.last_prompt,
    )

    col_a, col_b = st.columns(2)
    with col_a:
        slide_count_hint = ""
        if user_input.strip():
            wc = len(user_input.split())
            est = 10 if wc <= 40 else 12 if wc <= 200 else min(15, max(10, wc // 80))
            slide_count_hint = f"Estimated slides: **{est}**"
        st.markdown(f"**Slides**  \n{slide_count_hint or 'Will auto‑estimate based on length.'}")

    with col_b:
        if user_input.strip():
            mode = detect_mode(user_input)
            theme = pick_theme(user_input)
            is_proc = is_process_topic(user_input)
            mood = "Topic mode" if mode == "topic" else "Content mode"
            st.markdown(
                f"**Detection**  \n{mood} · Theme **{theme.value.replace('_',' ').title()}**"
                + (" · + Diagram" if is_proc else "")
            )
        else:
            st.markdown("**Detection**  \nWaiting for text…")

    st.markdown("---")

    generate_clicked = st.button("🚀 Generate deck", use_container_width=True, type="primary")

    if generate_clicked:
        if not user_input.strip():
            st.warning("Please type a topic or paste some content first.")
        else:
            st.session_state.last_prompt = user_input.strip()
            try:
                with st.spinner("Generating slides, icons & diagrams… this can take ~1 minute."):
                    slides, theme = prepare_slides(user_input.strip())
                st.session_state.slides = slides
                st.session_state.theme = theme
                st.session_state.selected_idx = 0
                st.session_state.ppt_bytes = None
                st.success("Deck created! Scroll to the right side to preview.")
            except Exception as e:
                st.error(f"Error during slide generation: {e}")

    st.markdown("</div>", unsafe_allow_html=True)


# ----------------------------- RIGHT PANEL -----------------------------------
with right:
    if not st.session_state.slides:
        st.markdown(
            """
<div class="glass-card" style="min-height:320px;display:flex;align-items:center;justify-content:center;flex-direction:column;">
  <div style="font-size:40px;margin-bottom:8px;">📊</div>
  <div style="font-size:18px;font-weight:600;color:#111827;margin-bottom:4px;">Your preview will appear here</div>
  <div style="font-size:13px;color:#6b7280;max-width:420px;text-align:center;">
    Describe your presentation on the left, then click <strong>Generate deck</strong>.
  </div>
</div>
""",
            unsafe_allow_html=True,
        )
    else:
        slides = st.session_state.slides
        theme = st.session_state.theme
        idx = st.session_state.selected_idx

        top_l, top_r = st.columns([1, 3])
        with top_l:
            st.markdown(
                """
<div class="glass-card" style="padding:14px 14px;max-height:430px;overflow-y:auto;">
  <div style="font-size:11px;font-weight:700;color:#6b7280;letter-spacing:.12em;text-transform:uppercase;margin-bottom:8px;">
    Slides
  </div>
""",
                unsafe_allow_html=True,
            )

            for i, sl in enumerate(slides):
                active = i == idx
                title = esc(sl.content.get("title", f"Slide {i+1}")[:50])
                ctype = sl.content_type.value.replace("_", " ").title()
                html_chip = f"""
<div class="slide-chip {'active' if active else ''}">
  <div style="display:flex; align-items:center; gap:12px;">
    <div style="font-size:14px; font-weight:700; color:{'#fff' if active else '#94a3b8'};">{(i+1):02d}</div>
    <div>
      <div style="font-size:11px; font-weight:600; color:{'#fff' if active else '#f1f5f9'};">{title}</div>
      <div style="font-size:9px; color:{'#a5b4fc' if active else '#64748b'}; text-transform:uppercase;">{ctype}</div>
    </div>
  </div>
  { '<div class="active-slide-marker"></div>' if active else '' }
</div>
"""
                st.markdown(html_chip, unsafe_allow_html=True)
                if st.button(f"View slide {i+1}", key=f"sel_{i}", use_container_width=True):
                    st.session_state.selected_idx = i
                    st.experimental_rerun()

            st.markdown("</div>", unsafe_allow_html=True)

        with top_r:
            st.markdown(
                """
<div class="glass-card" style="min-height:420px;">
""",
                unsafe_allow_html=True,
            )
            st.markdown(
                f"###### Preview · Slide {idx+1} of {len(slides)}",
                help="Use the list on the left to switch slides.",
            )
            render_slide_preview(slides[idx], theme)
            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("")

        # ----------------------- Compile & download row -----------------------
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        col1, col2, col3 = st.columns([2, 2, 2])

        with col1:
            if st.button("📦 Compile to PPTX", use_container_width=True):
                try:
                    with st.spinner("Exporting PowerPoint…"):
                        name = f"presentation_{uuid.uuid4().hex[:8]}.pptx"
                        st.session_state.ppt_bytes = export_pptx(slides, theme, filename=name)
                    st.success("PPTX compiled and ready to download.")
                except Exception as e:
                    st.error(f"Export error: {e}")

        with col2:
            if st.session_state.ppt_bytes:
                st.download_button(
                    "💾 Download PPTX",
                    data=st.session_state.ppt_bytes,
                    file_name="presentation.ai-generated.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )
            else:
                st.caption("Compile the deck first to enable download.")

        with col3:
            if st.button("🔁 Start over", use_container_width=True):
                st.session_state.slides = None
                st.session_state.theme = None
                st.session_state.selected_idx = 0
                st.session_state.ppt_bytes = None
                st.experimental_rerun()

        st.markdown("</div>", unsafe_allow_html=True)