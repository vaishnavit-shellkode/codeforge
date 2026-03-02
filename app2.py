import streamlit as st
import streamlit.components.v1 as components
import os, sys, base64, uuid

current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

try:
    from image import (
        prepare_slides, export_pptx, detect_mode, pick_theme, is_process_topic,
        reassemble_diagram,
        OUTPUT_DIR, ICONS_DIR, THEME_COLORS,
    )
except ImportError as e:
    st.error(f"Import error: {e}")
    st.stop()

st.set_page_config(page_title="AI PowerPoint Generator", page_icon="✨", layout="wide")

# ── Session state ──────────────────────────────────────────────────────────
for k, v in [("generated_slides",None),("generated_theme",None),
             ("target_data",None),("selected_slide_idx",0),("last_input","")]:
    if k not in st.session_state:
        st.session_state[k] = v

# ── CSS ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
*, body { font-family: 'Inter', sans-serif !important; }
.main { background: #f0f2f8; }
.block-container { padding: 1rem 2rem !important; }
.metric-card {
    background:white; padding:1rem; border-radius:10px;
    border-left:4px solid #6366f1; box-shadow:0 2px 8px rgba(0,0,0,0.06); margin-bottom:8px;
}
/* Thumbnail select buttons — thin, unobtrusive */
div[data-testid="column"] button[kind="secondary"] {
    background: #f8fafc !important; border: 1px solid #e2e8f0 !important;
    color: #64748b !important; font-size: 11px !important; padding: 3px 8px !important;
    border-radius: 5px !important; box-shadow: none !important; height: auto !important;
}
</style>
""", unsafe_allow_html=True)


# ── Utilities ────────────────────────────────────────────────────────────────
def _b64(path):
    if not path or not os.path.exists(path): return ""
    ext  = os.path.splitext(path)[-1].lower().strip(".")
    mime = {"jpg":"jpeg","jpeg":"jpeg","png":"png","webp":"webp"}.get(ext,"png")
    with open(path,"rb") as f: data = base64.b64encode(f.read()).decode()
    return f"data:image/{mime};base64,{data}"

def _save(up, dest, prefix="up"):
    ext = os.path.splitext(up.name)[-1].lower()
    if ext not in (".png",".jpg",".jpeg",".webp"): ext = ".png"
    os.makedirs(dest, exist_ok=True)
    path = os.path.join(dest, f"{prefix}_{uuid.uuid4().hex[:8]}{ext}")
    with open(path,"wb") as f: f.write(up.getbuffer())
    return path.replace("\\","/")

def _colors(theme):
    try: return THEME_COLORS[theme]
    except: return {"primary":"1E2761","secondary":"CADCFC","accent":"FFFFFF"}

def esc(t):
    return str(t).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE PREVIEW  — full HTML document rendered via components.html
# This is the ONLY way to correctly render base64 images in Streamlit
# ══════════════════════════════════════════════════════════════════════════════
def render_slide_preview(slide, theme):
    ct = slide.content_type.value
    c  = slide.content
    col = _colors(theme)
    P   = f"#{col['primary']}"
    S   = f"#{col['secondary']}"

    # ── build slide content HTML ─────────────────────────────────────────────
    if ct == "title":
        bg  = _b64(c.get("domain_bg_path",""))
        bg_css = f"background:url('{bg}') center/cover no-repeat;" if bg else f"background:{P};"
        body = f"""
        <div style="position:absolute;inset:0;{bg_css}">
          <div style="position:absolute;inset:0;background:rgba(0,0,0,0.5);"></div>
          <div style="position:absolute;inset:0;display:flex;flex-direction:column;
                      align-items:center;justify-content:center;padding:6% 10%;text-align:center;">
            <div style="font-size:38px;font-weight:800;color:#fff;text-shadow:0 2px 16px rgba(0,0,0,0.6);
                        line-height:1.2;margin-bottom:16px;">{esc(c.get('title',''))}</div>
            <div style="width:72px;height:4px;background:{S};border-radius:3px;margin-bottom:18px;"></div>
            <div style="font-size:18px;color:{S};font-weight:500;">{esc(c.get('subtitle',''))}</div>
          </div>
        </div>"""

    elif ct == "bullets":
        img = _b64(c.get("generated_image_path",""))
        buls = "".join(f"""
          <div style="display:flex;gap:10px;margin-bottom:10px;align-items:flex-start;">
            <div style="width:8px;height:8px;min-width:8px;border-radius:50%;background:{S};margin-top:6px;"></div>
            <div style="font-size:15px;color:#e2e8f0;line-height:1.5;">{esc(b)}</div>
          </div>""" for b in c.get("bullets",[]))
        img_panel = f'<div style="flex:1.2;border-radius:8px;overflow:hidden;"><img src="{img}" style="width:100%;height:100%;object-fit:cover;"/></div>' if img else ""
        txt_flex  = "0 0 50%" if img else "1"
        body = f"""
        <div style="position:absolute;inset:0;background:{P};">
          <div style="position:absolute;top:0;left:0;right:0;height:14%;background:rgba(0,0,0,0.2);
                      border-bottom:3px solid {S};display:flex;align-items:center;padding:0 4%;">
            <div style="font-size:22px;font-weight:700;color:#fff;">{esc(c.get('title',''))}</div>
          </div>
          <div style="position:absolute;top:14%;inset-inline:0;bottom:0;
                      display:flex;gap:20px;padding:22px 30px;align-items:center;">
            <div style="flex:{txt_flex};display:flex;flex-direction:column;justify-content:center;">{buls}</div>
            {img_panel}
          </div>
        </div>"""

    elif ct == "two_column":
        def cb(items):
            return "".join(f"""
              <div style="display:flex;gap:8px;margin-bottom:8px;align-items:flex-start;">
                <div style="width:7px;height:7px;min-width:7px;border-radius:50%;background:{S};margin-top:5px;"></div>
                <div style="font-size:13px;color:#e2e8f0;line-height:1.4;">{esc(b)}</div>
              </div>""" for b in items)
        body = f"""
        <div style="position:absolute;inset:0;background:#0b1428;">
          <div style="position:absolute;top:0;left:0;right:0;height:13%;background:{P};
                      border-bottom:3px solid {S};display:flex;align-items:center;padding:0 4%;">
            <div style="font-size:20px;font-weight:700;color:#fff;">{esc(c.get('title',''))}</div>
          </div>
          <div style="position:absolute;top:13%;inset-inline:0;bottom:0;
                      display:flex;gap:12px;padding:18px 22px;">
            <div style="flex:1;background:{P};border-radius:8px;padding:18px 20px;">
              <div style="font-size:14px;font-weight:700;color:{S};margin-bottom:12px;
                          padding-bottom:8px;border-bottom:2px solid {S}66;">{esc(c.get('left_heading',''))}</div>
              {cb(c.get('left_bullets',[]))}
            </div>
            <div style="flex:1;background:{P};border-radius:8px;padding:18px 20px;">
              <div style="font-size:14px;font-weight:700;color:{S};margin-bottom:12px;
                          padding-bottom:8px;border-bottom:2px solid {S}66;">{esc(c.get('right_heading',''))}</div>
              {cb(c.get('right_bullets',[]))}
            </div>
          </div>
        </div>"""

    elif ct == "stat_callout":
        cards = "".join(f"""
          <div style="flex:1;background:rgba(255,255,255,0.07);border:2px solid {S};
                      border-radius:10px;padding:18px 10px;text-align:center;margin:0 5px;">
            <div style="font-size:46px;font-weight:800;color:{S};line-height:1;">{esc(s.get('value',''))}</div>
            <div style="font-size:13px;color:rgba(255,255,255,0.75);margin-top:10px;">{esc(s.get('label',''))}</div>
          </div>""" for s in c.get("stats",[])[:4])
        body = f"""
        <div style="position:absolute;inset:0;background:{P};">
          <div style="position:absolute;top:0;left:0;right:0;height:13%;border-bottom:3px solid {S};
                      display:flex;align-items:center;padding:0 4%;">
            <div style="font-size:20px;font-weight:700;color:#fff;">{esc(c.get('title',''))}</div>
          </div>
          <div style="position:absolute;top:13%;inset-inline:0;bottom:12%;
                      display:flex;align-items:center;padding:22px 28px;">{cards}</div>
          <div style="position:absolute;bottom:3%;left:0;right:0;text-align:center;">
            <div style="font-size:13px;color:{S};font-style:italic;">{esc(c.get('body_text',''))}</div>
          </div>
        </div>"""

    elif ct == "timeline":
        evts = c.get("events",[])[:6]
        ehtml = "".join(f"""
          <div style="flex:1;text-align:center;padding:0 4px;">
            <div style="font-size:13px;font-weight:700;color:{S};margin-bottom:16px;">{esc(e.get('year',''))}</div>
            <div style="width:13px;height:13px;border-radius:50%;background:{S};
                        margin:0 auto 16px;border:3px solid rgba(255,255,255,0.25);"></div>
            <div style="font-size:12px;font-weight:600;color:#fff;margin-bottom:5px;">{esc(e.get('label',''))}</div>
            <div style="font-size:11px;color:rgba(255,255,255,0.6);line-height:1.4;">{esc(e.get('detail',''))}</div>
          </div>""" for e in evts)
        body = f"""
        <div style="position:absolute;inset:0;background:{P};">
          <div style="position:absolute;top:0;left:0;right:0;height:13%;border-bottom:3px solid {S};
                      display:flex;align-items:center;padding:0 4%;">
            <div style="font-size:20px;font-weight:700;color:#fff;">{esc(c.get('title',''))}</div>
          </div>
          <div style="position:absolute;top:37%;left:4%;right:4%;height:2px;background:rgba(255,255,255,0.1);"></div>
          <div style="position:absolute;top:13%;inset-inline:0;bottom:0;
                      display:flex;align-items:center;padding:14px 28px;">{ehtml}</div>
        </div>"""

    elif ct == "table":
        hdrs = c.get("headers",[])
        rows = c.get("rows",[])
        th = "".join(f"<th style='padding:9px 14px;font-size:12px;font-weight:700;color:{P};"
                     f"background:{S};text-align:left;border:none;'>{esc(h)}</th>" for h in hdrs)
        tr = "".join("<tr>"+"".join(
            f"<td style='padding:8px 14px;font-size:11px;color:#e2e8f0;"
            f"border-bottom:1px solid rgba(255,255,255,0.07);'>{esc(cell)}</td>"
            for cell in row)+"</tr>" for row in rows)
        body = f"""
        <div style="position:absolute;inset:0;background:{P};">
          <div style="position:absolute;top:0;left:0;right:0;height:13%;border-bottom:3px solid {S};
                      display:flex;align-items:center;padding:0 4%;">
            <div style="font-size:20px;font-weight:700;color:#fff;">{esc(c.get('title',''))}</div>
          </div>
          <div style="position:absolute;top:13%;inset-inline:0;bottom:0;padding:18px 28px;overflow:auto;">
            <table style="width:100%;border-collapse:collapse;">
              <thead><tr>{th}</tr></thead><tbody>{tr}</tbody>
            </table>
          </div>
        </div>"""

    elif ct == "quote":
        body = f"""
        <div style="position:absolute;inset:0;background:{P};">
          <div style="position:absolute;top:0;left:0;right:0;height:13%;border-bottom:3px solid {S};
                      display:flex;align-items:center;padding:0 4%;">
            <div style="font-size:20px;font-weight:700;color:#fff;">{esc(c.get('title',''))}</div>
          </div>
          <div style="position:absolute;top:13%;inset-inline:0;bottom:0;
                      display:flex;flex-direction:column;align-items:center;justify-content:center;padding:30px 70px;">
            <div style="font-size:64px;color:{S};font-weight:800;line-height:0.6;margin-bottom:14px;">"</div>
            <div style="font-size:20px;color:#fff;font-style:italic;line-height:1.6;
                        text-align:center;margin-bottom:22px;">{esc(c.get('quote',''))}</div>
            <div style="width:48px;height:3px;background:{S};border-radius:2px;margin-bottom:14px;"></div>
            <div style="font-size:14px;color:{S};font-weight:600;">{esc(c.get('attribution',''))}</div>
          </div>
        </div>"""

    elif ct == "diagram":
        img = _b64(c.get("diagram_image",""))
        if img:
            body = f'<div style="position:absolute;inset:0;background:#0d1117;"><img src="{img}" style="width:100%;height:100%;object-fit:contain;display:block;"/></div>'
        else:
            body = f"""
            <div style="position:absolute;inset:0;background:{P};display:flex;flex-direction:column;
                        align-items:center;justify-content:center;gap:14px;">
              <div style="font-size:52px;">🔷</div>
              <div style="color:rgba(255,255,255,0.45);font-size:15px;text-align:center;">
                Diagram image not ready<br/>Edit steps → Re-Assemble Diagram
              </div>
            </div>"""

    elif ct == "thank_you":
        bg  = _b64(c.get("domain_bg_path",""))
        bg_css = f"background:url('{bg}') center/cover no-repeat;" if bg else f"background:{P};"
        body = f"""
        <div style="position:absolute;inset:0;{bg_css}">
          <div style="position:absolute;inset:0;background:rgba(0,0,0,0.52);"></div>
          <div style="position:absolute;inset:0;display:flex;flex-direction:column;
                      align-items:center;justify-content:center;padding:6% 10%;text-align:center;">
            <div style="font-size:54px;font-weight:800;color:#fff;margin-bottom:18px;
                        text-shadow:0 2px 20px rgba(0,0,0,0.5);">{esc(c.get('title','Thank You'))}</div>
            <div style="width:80px;height:4px;background:{S};border-radius:2px;margin-bottom:18px;"></div>
            <div style="font-size:18px;color:{S};margin-bottom:10px;">{esc(c.get('message',''))}</div>
            <div style="font-size:13px;color:rgba(255,255,255,0.5);">{esc(c.get('contact',''))}</div>
          </div>
        </div>"""
    else:
        body = f'<div style="position:absolute;inset:0;background:{P};display:flex;align-items:center;justify-content:center;"><div style="font-size:28px;color:#fff;font-weight:700;">{esc(c.get("title",""))}</div></div>'

    # ── Wrap in a full standalone HTML doc — this is key for base64 images ────
    html_doc = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
  *{{margin:0;padding:0;box-sizing:border-box;}}
  html,body{{width:100%;height:100%;background:#1e293b;
             display:flex;align-items:center;justify-content:center;overflow:hidden;}}
  .frame{{width:100%;aspect-ratio:16/9;position:relative;border-radius:10px;
          overflow:hidden;box-shadow:0 16px 52px rgba(0,0,0,0.4);
          font-family:'Inter',system-ui,sans-serif;}}
</style>
</head>
<body>
  <div class="frame">{body}</div>
</body></html>"""

    # ── Render via components.html so images actually display ─────────────────
    components.html(html_doc, height=430, scrolling=False)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE THUMBNAIL
# ─────────────────────────────────────────────────────────────────────────────
def _thumb_html(slide, theme, i, active):
    ct  = slide.content_type.value
    c   = slide.content
    col = _colors(theme)
    P   = f"#{col['primary']}"
    S   = f"#{col['secondary']}"
    icons = {"title":"🎯","bullets":"📋","two_column":"▪️","stat_callout":"📊",
             "timeline":"📅","table":"🗂️","quote":"💬","diagram":"🔷","thank_you":"🙏"}
    icon  = icons.get(ct,"📄")
    title = esc(c.get("title",f"Slide {i+1}")[:28])
    bd    = "2.5px solid #6366f1" if active else "1.5px solid #e2e8f0"
    sh    = "0 4px 14px rgba(99,102,241,0.28)" if active else "0 1px 4px rgba(0,0,0,0.07)"
    bg    = "#eef2ff" if active else "white"
    return f"""
    <div style="background:{bg};border-radius:9px;border:{bd};padding:7px;
                margin-bottom:2px;box-shadow:{sh};">
      <div style="font-size:9px;color:#94a3b8;font-weight:700;letter-spacing:0.5px;margin-bottom:3px;">
        SLIDE {i+1}
      </div>
      <div style="background:{P};border-radius:5px;padding:7px 9px;display:flex;
                  align-items:center;gap:7px;min-height:36px;">
        <span style="font-size:13px;flex-shrink:0;">{icon}</span>
        <span style="font-size:9.5px;color:white;font-weight:600;line-height:1.3;
                     overflow:hidden;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;">
          {title}
        </span>
      </div>
      <div style="font-size:8px;color:#6366f1;margin-top:3px;text-align:center;font-weight:700;">
        {ct.replace('_',' ').upper()}
      </div>
    </div>"""


# ─────────────────────────────────────────────────────────────────────────────
# EDIT PANEL
# ─────────────────────────────────────────────────────────────────────────────
def render_edit_panel(slide, theme):
    ct  = slide.content_type.value
    c   = slide.content
    sn  = slide.slide_number

    # Reusable image upload widget
    def img_widget(field, label, dest, prefix):
        cur = c.get(field,"")
        with st.expander(f"🖼️ {label}"):
            if cur and os.path.exists(cur):
                st.image(cur, width=190, caption="Current")
                if st.button("🗑️ Remove", key=f"rm_{field}_{sn}"):
                    c.pop(field,None); st.rerun()
            else:
                st.caption("No image attached.")
            st.write("**Upload new image** (PNG / JPG / WEBP):")
            up = st.file_uploader(f"img_{field}_{sn}", type=["png","jpg","jpeg","webp"],
                                  label_visibility="collapsed", key=f"up_{field}_{sn}")
            if up:
                saved = _save(up, dest, prefix)
                c[field] = saved
                st.image(saved, width=190, caption="✅ Saved")
                st.rerun()

    if ct == "title":
        c["title"]    = st.text_input("Title",    value=c.get("title",""),    key=f"t_{sn}")
        c["subtitle"] = st.text_input("Subtitle", value=c.get("subtitle",""), key=f"s_{sn}")
        img_widget("domain_bg_path", "Background / Hero Image", OUTPUT_DIR, f"bg{sn}")

    elif ct == "bullets":
        c["title"] = st.text_input("Title", value=c.get("title",""), key=f"bt_{sn}")
        st.markdown("**Bullets**")
        for i,b in enumerate(c.get("bullets",[])):
            c["bullets"][i] = st.text_input(f"• {i+1}", value=b, key=f"b_{sn}_{i}")
        ca,cb = st.columns(2)
        with ca:
            if st.button("➕ Add", key=f"addb_{sn}"):
                c.setdefault("bullets",[]).append("New point"); st.rerun()
        with cb:
            if len(c.get("bullets",[])) > 1:
                if st.button("➖ Remove last", key=f"rmb_{sn}"):
                    c["bullets"].pop(); st.rerun()
        img_widget("generated_image_path", "Side Image (optional)", OUTPUT_DIR, f"slide{sn}")

    elif ct == "two_column":
        c["title"] = st.text_input("Title", value=c.get("title",""), key=f"2t_{sn}")
        c1,c2 = st.columns(2)
        with c1:
            st.markdown("**Left**")
            c["left_heading"] = st.text_input("Heading", value=c.get("left_heading",""), key=f"lh_{sn}")
            for i,b in enumerate(c.get("left_bullets",[])):
                c["left_bullets"][i] = st.text_input(f"L{i+1}", value=b, key=f"lb_{sn}_{i}")
        with c2:
            st.markdown("**Right**")
            c["right_heading"] = st.text_input("Heading", value=c.get("right_heading",""), key=f"rh_{sn}")
            for i,b in enumerate(c.get("right_bullets",[])):
                c["right_bullets"][i] = st.text_input(f"R{i+1}", value=b, key=f"rb_{sn}_{i}")
        img_widget("generated_image_path", "Side Image (optional)", OUTPUT_DIR, f"slide{sn}")

    elif ct == "stat_callout":
        c["title"]     = st.text_input("Title",           value=c.get("title",""),     key=f"sct_{sn}")
        c["body_text"] = st.text_input("Supporting text", value=c.get("body_text",""), key=f"scb_{sn}")
        st.markdown("**Stats**")
        for i,s in enumerate(c.get("stats",[])):
            a,b = st.columns(2)
            with a: c["stats"][i]["value"] = st.text_input(f"Value {i+1}", value=s.get("value",""), key=f"sv_{sn}_{i}")
            with b: c["stats"][i]["label"] = st.text_input(f"Label {i+1}", value=s.get("label",""), key=f"sl_{sn}_{i}")

    elif ct == "timeline":
        c["title"] = st.text_input("Title", value=c.get("title",""), key=f"tlt_{sn}")
        st.markdown("**Events**")
        for i,e in enumerate(c.get("events",[])):
            ec1,ec2,ec3 = st.columns([1,2,3])
            with ec1: c["events"][i]["year"]   = st.text_input("Year",   value=e.get("year",""),   key=f"ey_{sn}_{i}")
            with ec2: c["events"][i]["label"]  = st.text_input("Label",  value=e.get("label",""),  key=f"el_{sn}_{i}")
            with ec3: c["events"][i]["detail"] = st.text_input("Detail", value=e.get("detail",""), key=f"ed_{sn}_{i}")

    elif ct == "table":
        c["title"] = st.text_input("Title", value=c.get("title",""), key=f"tbt_{sn}")
        st.markdown("**Headers**")
        for i,h in enumerate(c.get("headers",[])):
            c["headers"][i] = st.text_input(f"Col {i+1}", value=h, key=f"th_{sn}_{i}")
        st.markdown("**Rows**")
        for ri,row in enumerate(c.get("rows",[])):
            rc = st.columns(len(row))
            for ci,cell in enumerate(row):
                with rc[ci]:
                    c["rows"][ri][ci] = st.text_input("", value=cell, key=f"tr_{sn}_{ri}_{ci}", label_visibility="collapsed")

    elif ct == "quote":
        c["title"]       = st.text_input("Title",       value=c.get("title",""),       key=f"qt_{sn}")
        c["quote"]       = st.text_area("Quote",        value=c.get("quote",""),        key=f"qq_{sn}", height=90)
        c["attribution"] = st.text_input("Attribution", value=c.get("attribution",""), key=f"qa_{sn}")

    elif ct == "diagram":
        _diagram_editor(slide)

    elif ct == "thank_you":
        c["title"]   = st.text_input("Title",   value=c.get("title","Thank You"), key=f"tyt_{sn}")
        c["message"] = st.text_area("Message",  value=c.get("message",""),        key=f"tym_{sn}", height=70)
        c["contact"] = st.text_input("Contact", value=c.get("contact",""),        key=f"tyc_{sn}")
        img_widget("domain_bg_path", "Background / Hero Image", OUTPUT_DIR, f"bg{sn}")

    if ct != "diagram":
        lmap = {"title":["default","title_left","title_dark"],
                "bullets":["default","bullets_box"],
                "two_column":["default","two_column_cards"]}
        layouts = lmap.get(ct,[])
        if len(layouts) >= 2:
            st.markdown("**Layout**")
            lcols = st.columns(len(layouts))
            for i,ly in enumerate(layouts):
                with lcols[i]:
                    sel = (slide.layout_override==ly) or (not slide.layout_override and ly=="default")
                    if st.button(("✅ " if sel else "")+ly.replace("_"," ").title(),
                                 key=f"lay_{sn}_{ly}", use_container_width=True, disabled=sel):
                        slide.layout_override = ly if ly != "default" else ""; st.rerun()

    if c.get("speaker_notes"):
        with st.expander("🎙️ Speaker Notes"):
            c["speaker_notes"] = st.text_area("", value=c.get("speaker_notes",""),
                                               key=f"spk_{sn}", height=70, label_visibility="collapsed")


# ─────────────────────────────────────────────────────────────────────────────
# DIAGRAM EDITOR
# ─────────────────────────────────────────────────────────────────────────────
def _diagram_editor(slide):
    c     = slide.content
    steps = c.get("steps",[])
    sn    = slide.slide_number

    c["title"] = st.text_input("Diagram Title", value=c.get("title","Process Flow"), key=f"dt_{sn}")

    st.markdown("**Layout**")
    lc1,lc2,lc3 = st.columns(3)
    for col_w, ly in zip([lc1,lc2,lc3], ["diagram","diagram_grid","diagram_vertical"]):
        with col_w:
            sel = (slide.layout_override==ly) or (not slide.layout_override and ly=="diagram")
            if st.button(("✅ " if sel else "")+ly.replace("_"," ").title(),
                         key=f"dlay_{sn}_{ly}", use_container_width=True, disabled=sel):
                slide.layout_override = ly
                imgs = c.get("diagram_images",{})
                if ly in imgs: c["diagram_image"] = imgs[ly]
                st.rerun()

    if not steps:
        st.info("No steps yet."); return

    st.markdown("---")
    st.markdown("**Edit Steps** — change icon (generate or upload), label, description")

    for row_start in range(0, len(steps), 3):
        row   = steps[row_start:row_start+3]
        cols  = st.columns(3)
        for ci, step in enumerate(row):
            ai   = row_start + ci
            snum = step.get("step_number", ai+1)
            with cols[ci]:
                # ── Step badge ────────────────────────────────────────────────
                st.markdown(
                    f"<div style='background:linear-gradient(135deg,#1E2761,#4f46e5);color:white;"
                    f"border-radius:8px 8px 0 0;padding:6px 10px;font-weight:700;font-size:12px;"
                    f"text-align:center;'>Step {snum}</div>",
                    unsafe_allow_html=True)

                # ── Label & description ────────────────────────────────────────
                steps[ai]["label"] = st.text_input(
                    "Label", value=step.get("label",""),
                    key=f"lbl_{sn}_{ai}", placeholder="Short title")
                steps[ai]["description"] = st.text_area(
                    "Description", value=step.get("description",""),
                    key=f"dsc_{sn}_{ai}", height=68, placeholder="One sentence…")

                # ── Step management: text-labeled buttons (no arrow symbols) ──
                st.markdown("<div style='display:flex;gap:4px;flex-wrap:wrap;margin-top:2px;'>",
                            unsafe_allow_html=True)
                m1,m2,m3,m4 = st.columns(4)
                with m1:
                    if ai > 0:
                        if st.button("Move Left", key=f"ml_{sn}_{ai}", use_container_width=True):
                            steps[ai-1],steps[ai] = steps[ai],steps[ai-1]
                            for j,s in enumerate(steps): s["step_number"]=j+1
                            c["steps"]=steps; st.rerun()
                with m2:
                    if ai < len(steps)-1:
                        if st.button("Move Right", key=f"mr_{sn}_{ai}", use_container_width=True):
                            steps[ai],steps[ai+1] = steps[ai+1],steps[ai]
                            for j,s in enumerate(steps): s["step_number"]=j+1
                            c["steps"]=steps; st.rerun()
                with m3:
                    if st.button("Insert After", key=f"ins_{sn}_{ai}", use_container_width=True):
                        steps.insert(ai+1,{"step_number":ai+2,"label":"New Step",
                            "description":"Describe this step.","icon_prompt":"","icon_path":""})
                        for j,s in enumerate(steps): s["step_number"]=j+1
                        c["steps"]=steps; st.rerun()
                with m4:
                    if len(steps)>1:
                        if st.button("Delete", key=f"del_{sn}_{ai}", use_container_width=True):
                            steps.pop(ai)
                            for j,s in enumerate(steps): s["step_number"]=j+1
                            c["steps"]=steps; st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)
                st.divider()

    c["steps"] = steps

    st.markdown("---")
    if st.button("🖼️ Re-Assemble Diagram", key=f"reas_{sn}",
                 use_container_width=True, type="primary"):
        with st.spinner("Re-assembling diagram PNG…"):
            nps = reassemble_diagram(steps_data=steps,
                                     title=c.get("title","Process Flow"),
                                     user_query=st.session_state.get("last_input",""))
        c["diagram_image"]  = nps["diagram"]
        c["diagram_images"] = nps
        sel = slide.layout_override or "diagram"
        if sel in nps: c["diagram_image"] = nps[sel]
        st.success("✅ Diagram updated!"); st.rerun()



# ─────────────────────────────────────────────────────────────────────────────
# PAGE HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="background:linear-gradient(135deg,#1E2761,#4f46e5);padding:16px 24px;
            border-radius:12px;margin-bottom:16px;display:flex;align-items:center;gap:14px;">
  <span style="font-size:1.8rem;">✨</span>
  <div>
    <div style="font-size:1.4rem;font-weight:800;color:white;">AI PowerPoint Generator</div>
    <div style="font-size:0.85rem;color:rgba(255,255,255,0.65);">Turn ideas into professional presentations instantly</div>
  </div>
</div>""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# INPUT PAGE
# ─────────────────────────────────────────────────────────────────────────────
if not st.session_state.generated_slides:
    st.markdown("### 📝 What do you want to present?")
    user_input = st.text_area("", height=200, label_visibility="collapsed",
        placeholder="Type a topic (e.g. 'AI in healthcare') or paste your full content…")

    if user_input.strip():
        mode_  = detect_mode(user_input)
        thm_   = pick_theme(user_input)
        hasp   = is_process_topic(user_input)
        wc     = len(user_input.split())
        nsl    = 10 if wc<=40 else 12 if wc<=200 else min(15,max(10,wc//80))
        m1,m2,m3 = st.columns(3)
        m1.markdown(f'<div class="metric-card"><small style="color:#64748b;font-weight:700;text-transform:uppercase;font-size:11px;">Mode</small><br/><strong>{"Topic Gen" if mode_=="topic" else "Content"}</strong></div>', unsafe_allow_html=True)
        m2.markdown(f'<div class="metric-card"><small style="color:#64748b;font-weight:700;text-transform:uppercase;font-size:11px;">Slides</small><br/><strong>{nsl}{" + Diagram" if hasp else ""}</strong></div>', unsafe_allow_html=True)
        m3.markdown(f'<div class="metric-card"><small style="color:#64748b;font-weight:700;text-transform:uppercase;font-size:11px;">Theme</small><br/><strong>{thm_.value.replace("_"," ").title()}</strong></div>', unsafe_allow_html=True)
        st.write("")

    if st.button("🚀 Generate Presentation", use_container_width=True):
        if not user_input.strip():
            st.warning("Please enter a topic or content.")
        else:
            try:
                with st.spinner("🤖 Generating slides, icons, and diagrams… (~1 min)"):
                    slides, theme = prepare_slides(user_input.strip())
                    st.session_state.generated_slides   = slides
                    st.session_state.generated_theme    = theme
                    st.session_state.last_input         = user_input.strip()
                    st.session_state.selected_slide_idx = 0
                    st.session_state.target_data        = None
                st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")
                st.info("Check AWS Bedrock credentials and node.js / pptxgenjs setup.")


# ─────────────────────────────────────────────────────────────────────────────
# EDITOR PAGE
# ─────────────────────────────────────────────────────────────────────────────
else:
    slides = st.session_state.generated_slides
    theme  = st.session_state.generated_theme

    tl, tr = st.columns([5,1])
    with tl: st.markdown(f"**{len(slides)} slides ready** — click a slide to select")
    with tr:
        if st.button("🔄 Start Over", use_container_width=True):
            for k in ["generated_slides","generated_theme","target_data"]:
                st.session_state[k] = None
            st.session_state.selected_slide_idx = 0; st.rerun()

    # Two-column layout: thumbnails + preview (no edit panel)
    col_nav, col_prev = st.columns([1, 4])

    # ── LEFT: thumbnails ──────────────────────────────────────────────────────
    with col_nav:
        st.markdown("<p style='font-size:10px;font-weight:700;color:#94a3b8;letter-spacing:1px;margin:0 0 8px;'>SLIDES</p>", unsafe_allow_html=True)
        for i, sl in enumerate(slides):
            active = (i == st.session_state.selected_slide_idx)
            st.markdown(_thumb_html(sl, theme, i, active), unsafe_allow_html=True)
            if st.button(f"Select slide {i+1}", key=f"nav_{i}", use_container_width=True):
                st.session_state.selected_slide_idx = i; st.rerun()

    # ── RIGHT: visual preview ────────────────────────────────────────────────
    with col_prev:
        idx   = st.session_state.selected_slide_idx
        slide = slides[idx]
        ct    = slide.content_type.value

        st.markdown(
            f"<p style='font-size:10px;font-weight:700;color:#94a3b8;letter-spacing:1px;margin:0 0 8px;'>"
            f"PREVIEW — SLIDE {idx+1} / {len(slides)}</p>",
            unsafe_allow_html=True)

        # ✅ KEY FIX: use components.html() — not st.markdown — for proper rendering
        render_slide_preview(slide, theme)

        # ── Navigation (Prev / label / Next) — text labels, no arrow symbols ──
        nl, nm, nr = st.columns([1, 3, 1])
        with nl:
            if idx > 0:
                if st.button("Prev", key="prev_btn", use_container_width=True):
                    st.session_state.selected_slide_idx -= 1; st.rerun()
        with nm:
            st.markdown(
                f"<div style='text-align:center;font-size:11px;color:#94a3b8;padding:10px 0;'>"
                f"Slide {idx+1} of {len(slides)} &nbsp;•&nbsp; "
                f"<b style='color:#6366f1;'>{ct.replace('_',' ').upper()}</b></div>",
                unsafe_allow_html=True)
        with nr:
            if idx < len(slides)-1:
                if st.button("Next", key="next_btn", use_container_width=True):
                    st.session_state.selected_slide_idx += 1; st.rerun()

    # ── Compile & Download ────────────────────────────────────────────────────
    st.divider()
    d1,d2,d3 = st.columns([2,2,2])
    with d1:
        if st.button("Compile All Slides to PPTX", use_container_width=True):
            with st.spinner("Compiling PowerPoint…"):
                try:
                    pid = uuid.uuid4().hex[:8]
                    st.session_state.target_data = export_pptx(
                        slides, theme, filename=f"presentation_{pid}.pptx")
                    st.success("✅ Ready!")
                except Exception as e:
                    st.error(f"Error: {e}")
    with d2:
        if st.session_state.target_data:
            st.download_button("💾 Download PPTX",
                data=st.session_state.target_data, file_name="presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True)
    with d3:
        st.markdown(
            f"<div style='text-align:center;padding:10px 0;font-size:12px;color:#94a3b8;'>"
            f"Theme: <b style='color:#6366f1;'>{theme.value.replace('_',' ').title()}</b></div>",
            unsafe_allow_html=True)