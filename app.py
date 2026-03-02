import streamlit as st
import os
import sys

# Change default directory to current file location to ensure imports work correctly, if needed
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

try:
    from image import (
        prepare_slides, export_pptx, detect_mode, pick_theme, is_process_topic,
        regenerate_step_icon, reassemble_diagram,
    )
    import uuid
except ImportError as e:
    st.error(f"Error importing backend modules from image.py: {e}")
    st.stop()

# Configure page
st.set_page_config(
    page_title="AI PowerPoint Generator", 
    page_icon="✨", 
    layout="centered"
)

# Initialize Session State
if "generated_slides" not in st.session_state:
    st.session_state.generated_slides = None
if "generated_theme" not in st.session_state:
    st.session_state.generated_theme = None
if "target_data" not in st.session_state:
    st.session_state.target_data = None

# Custom CSS for a modern, sleek interface
st.markdown("""
<style>
    .main {
        background-color: #f7f9fc;
    }
    .stTextArea textarea {
        border-radius: 12px;
        border: 1px solid #d1d5db;
        padding: 15px;
        font-size: 16px;
        box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.02);
    }
    .stButton>button {
        border-radius: 12px;
        background: linear-gradient(135deg, #1E2761 0%, #2a357f 100%);
        color: white;
        padding: 12px 24px;
        font-size: 16px;
        font-weight: 600;
        border: none;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        transition: all 0.3s ease;
        width: 100%;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 12px rgba(0,0,0,0.15);
        color: #ffffff;
    }
    .title-box {
        text-align: center;
        padding: 2rem 0 1rem 0;
    }
    .title-box h1 {
        color: #1E2761;
        font-weight: 800;
        letter-spacing: -0.5px;
        margin-bottom: 0.5rem;
    }
    .title-box p {
        color: #64748b;
        font-size: 1.1rem;
        max-width: 600px;
        margin: 0 auto;
    }
    .download-card {
        background: white;
        border-radius: 16px;
        padding: 2rem;
        text-align: center;
        box-shadow: 0 10px 25px rgba(0,0,0,0.05);
        border: 1px solid #e2e8f0;
        margin-top: 2rem;
    }
    .metric-card {
        background: white;
        padding: 1rem;
        border-radius: 10px;
        border-left: 4px solid #1E2761;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }
</style>
""", unsafe_allow_html=True)

# Header Section
st.markdown('''
<div class="title-box">
    <h1>✨ AI PowerPoint Generator</h1>
    <p>Turn your ideas, topics, or full articles into beautiful, professional presentations in seconds.</p>
</div>
''', unsafe_allow_html=True)

# Main Input
st.markdown("### Tell me what you want to present")
user_input = st.text_area(
    "Topic or Content", 
    height=240, 
    placeholder="Example: Explain the complete loan processing workflow followed by a modern lending tech platform. Cover how a loan is applied for, evaluated, underwritten, approved, and disbursed...",
    label_visibility="collapsed"
)

# Live Analysis Metrics
if user_input.strip():
    st.markdown("#### Presentation Blueprint")
    
    col1, col2, col3 = st.columns(3)
    
    # Run the same logic as the backend to give users a preview
    mode = detect_mode(user_input)
    theme = pick_theme(user_input)
    has_process = is_process_topic(user_input)
    
    word_count = len(user_input.split())
    if word_count <= 40:
        num_slides = 10
    elif word_count <= 200:
        num_slides = 12
    else:
        num_slides = min(15, max(10, word_count // 80))
        
    mode_text = "Topic Generation" if mode == 'topic' else "Content Structuring"
    slides_text = f"{num_slides}" + (" + Flow Diagram" if has_process else "")
    theme_text = theme.value.replace('_', ' ').title()
    
    col1.markdown(f'''
    <div class="metric-card">
        <small style="color: #64748b; text-transform: uppercase; font-weight: bold;">Mode</small><br/>
        <strong style="color: #1e293b; font-size: 1.1rem;">{mode_text}</strong>
    </div>
    ''', unsafe_allow_html=True)
    
    col2.markdown(f'''
    <div class="metric-card">
        <small style="color: #64748b; text-transform: uppercase; font-weight: bold;">Est. Slides</small><br/>
        <strong style="color: #1e293b; font-size: 1.1rem;">{slides_text}</strong>
    </div>
    ''', unsafe_allow_html=True)
    
    col3.markdown(f'''
    <div class="metric-card">
        <small style="color: #64748b; text-transform: uppercase; font-weight: bold;">AI Theme Match</small><br/>
        <strong style="color: #1e293b; font-size: 1.1rem;">{theme_text}</strong>
    </div>
    ''', unsafe_allow_html=True)

st.write("") # Spacer

# Generation triggers
if st.button("Generate Outline & Content 🚀", use_container_width=True):
    if not user_input.strip():
        st.warning("Please enter a topic or some content to generate the presentation.")
    else:
        try:
            with st.spinner("🤖 AI is thinking, generating content, and rendering slides... This might take a minute."):
                slides, theme = prepare_slides(user_input.strip())
                st.session_state.generated_slides = slides
                st.session_state.generated_theme = theme
                st.session_state.last_input = user_input.strip()
                st.session_state.target_data = None # reset data if new generation
                st.rerun()
        except Exception as e:
            st.error(f"**An error occurred during generation:**\n\n`{str(e)}`")
            st.info("Check your AWS/Bedrock setup, environment variables, or node.js/pptxgen dependencies as configured in `image.py`.")

if st.session_state.generated_slides:
    st.balloons()
    
    # Render Preview
    st.divider()
    st.markdown("### 📽️ Generated Presentation Preview & Layout Selection")
    
    if "selected_slide_idx" not in st.session_state or st.session_state.selected_slide_idx >= len(st.session_state.generated_slides):
        st.session_state.selected_slide_idx = 0
        
    col_nav, col_main = st.columns([1, 3])
    
    with col_nav:
        st.markdown("<h4 style='color:#1E2761; margin-bottom:1rem;'>📑 Slides</h4>", unsafe_allow_html=True)
        for i, slide in enumerate(st.session_state.generated_slides):
            t = slide.content.get('title', f"Slide {i+1}")
            # Emulate a thumbnail selector button
            btn_style = "primary" if i == st.session_state.selected_slide_idx else "secondary"
            if st.button(f"#{i+1} : {t[:20]}", key=f"nav_{i}", use_container_width=True, type=btn_style):
                st.session_state.selected_slide_idx = i
                st.rerun()

    with col_main:
        slide = st.session_state.generated_slides[st.session_state.selected_slide_idx]
        with st.container(border=True):
            st.markdown(f"#### Slide {slide.slide_number}: *{slide.content.get('title', 'Presentation')}*")
            
            # Simple rendering of content
            content = slide.content
            ct = slide.content_type.value
            
            # --- OVERRIDE TEMPLATE SELECTOR ---
            if ct == "diagram":
                layouts = ["diagram", "diagram_grid", "diagram_vertical"]
            else:
                layouts = ["default"]
                if ct == "title":
                    # Requirement: 5-6 design templates for slide 1
                    layouts.extend(["title_left", "title_dark", "title_split", "title_clean", "title_accent"])
                elif ct == "bullets":
                    layouts.extend(["bullets_box"])
                elif ct == "two_column":
                    layouts.extend(["two_column_cards"])
                
            st.divider()
            st.markdown("##### 🎨 Choose Design Template")
            cols = st.columns(len(layouts))
            
            for idx, layout in enumerate(layouts):
                with cols[idx]:
                    is_selected = (slide.layout_override == layout) or (slide.layout_override == "" and layout == "default")
                    
                    # Generate CSS mini-thumbnail mockups depending on layout
                    html_mockup = ""
                    if layout == "default" and ct in ["title", "bullets"]:
                        html_mockup = "<div style='height:40px; background:#f1f5f9; border:1px solid #cbd5e1; border-radius:4px; padding:4px;'><div style='height:8px; background:#1e293b; margin-bottom:4px; border-radius:2px;'></div><div style='height:3px; background:#94a3b8; width:80%; margin-bottom:3px;'></div><div style='height:3px; background:#94a3b8; width:60%;'></div></div>"
                    elif layout == "title_left":
                        html_mockup = "<div style='height:40px; background:#f1f5f9; border:1px solid #cbd5e1; border-radius:4px; display:flex;'><div style='width:30%; background:#1e293b; height:100%; border-radius:3px 0 0 3px;'></div><div style='flex:1; padding:4px;'><div style='height:5px; background:#64748b; width:70%; margin-bottom:4px;'></div><div style='height:3px; background:#94a3b8; width:40%;'></div></div></div>"
                    elif layout == "title_dark":
                        html_mockup = "<div style='height:40px; background:#1e293b; border-radius:4px; border:1.5px solid #cbd5e1; display:flex; flex-direction:column; justify-content:center; align-items:center;'><div style='height:6px; background:#f8fafc; width:60%; margin-bottom:4px; border-radius:2px;'></div><div style='height:3px; background:#94a3b8; width:30%;'></div></div>"
                    elif layout == "title_split":
                        html_mockup = "<div style='height:40px; background:#f1f5f9; border:1px solid #cbd5e1; border-radius:4px; display:flex;'><div style='width:40%; background:#1e293b; height:100%; border-radius:3px 0 0 3px;'></div><div style='flex:1; padding:4px; display:flex; align-items:center;'><div style='height:8px; background:#3b82f6; width:90%;'></div></div></div>"
                    elif layout == "title_clean":
                        html_mockup = "<div style='height:40px; background:#ffffff; border:1px solid #e2e8f0; border-radius:4px; padding:4px; display:flex; flex-direction:column; justify-content:center; align-items:center;'><div style='height:2px; background:#1e293b; width:30%; margin-bottom:4px;'></div><div style='height:6px; background:#1e293b; width:70%;'></div></div>"
                    elif layout == "title_accent":
                        html_mockup = "<div style='height:40px; background:#f1f5f9; border:1px solid #cbd5e1; border-radius:4px; display:flex; flex-direction:column; justify-content:space-between;'><div style='height:4px; background:#1e293b;'></div><div style='height:8px; background:#1e293b; width:60%; align-self:center;'></div><div style='height:4px; background:#1e293b;'></div></div>"
                    elif layout == "bullets_box":
                        html_mockup = "<div style='height:40px; background:#f1f5f9; border:1.5px solid #1e293b; border-radius:4px; padding:4px;'><div style='height:5px; background:#1e293b; width:40%; margin-bottom:4px;'></div><div style='height:2px; background:#94a3b8; width:70%; margin-bottom:3px;'></div><div style='height:2px; background:#94a3b8; width:50%;'></div></div>"
                    elif layout == "two_column_cards":
                        html_mockup = "<div style='height:40px; background:#e2e8f0; border-radius:4px; padding:3px; display:flex; gap:3px;'><div style='flex:1; background:white; border:1px solid #cbd5e1; border-radius:2px; padding:2px;'><div style='height:3px; background:#1e293b; margin-bottom:2px;'></div><div style='height:2px; background:#94a3b8; width:70%;'></div></div><div style='flex:1; background:white; border:1px solid #cbd5e1; border-radius:2px; padding:2px;'><div style='height:3px; background:#1e293b; margin-bottom:2px;'></div><div style='height:2px; background:#94a3b8; width:70%;'></div></div></div>"
                    elif layout == "diagram":
                        html_mockup = "<div style='height:40px; background:#f1f5f9; border:1px solid #cbd5e1; border-radius:4px; display:flex; gap:1.5px; justify-content:center; align-items:center;'><div style='width:8px; height:8px; border-radius:4px; background:#1e293b;'></div><div style='width:6px; height:1.5px; background:#94a3b8;'></div><div style='width:8px; height:8px; border-radius:4px; background:#1e293b;'></div><div style='width:6px; height:1.5px; background:#94a3b8;'></div><div style='width:8px; height:8px; border-radius:4px; background:#1e293b;'></div></div>"
                    else:
                        html_mockup = "<div style='height:40px; background:#f1f5f9; border-radius:4px; border:1px solid #cbd5e1; padding:4px;'><div style='height:6px; background:#1e293b; width:30%; margin-bottom:4px;'></div><div style='height:3px; background:#94a3b8; width:80%;'></div></div>"
                    
                    border_color = "#3b82f6" if is_selected else "transparent"
                    st.markdown(f"<div style='border:2px solid {border_color}; border-radius:6px; overflow:hidden;'>{html_mockup}</div>", unsafe_allow_html=True)
                    
                    btn_label = "✅" if is_selected else layout.replace('title_', '').replace('_', ' ').title()
                    if st.button(btn_label, key=f"btn_{slide.slide_number}_{layout}", use_container_width=True, disabled=is_selected):
                        if ct == "diagram":
                            slide.layout_override = layout
                            img_dict = slide.content.get("diagram_images", {})
                            if layout in img_dict:
                                slide.content["diagram_image"] = img_dict[layout]
                        else:
                            slide.layout_override = layout if layout != "default" else ""
                        st.rerun()
            
            st.divider()
            
            # --- CONTENT PREVIEW & EDIT ---
            if ct == "title":
                content["title"] = st.text_input("Title", value=content.get("title", ""), key=f"edit_title_{slide.slide_number}")
                content["subtitle"] = st.text_input("Subtitle", value=content.get("subtitle", ""), key=f"edit_subtitle_{slide.slide_number}")
            
            elif ct == "bullets":
                if content.get("generated_image_path") and os.path.exists(content["generated_image_path"]):
                    col_t, col_i = st.columns([2, 5])
                    with col_t:
                        content["title"] = st.text_input("Box Title", value=content.get("title", ""), key=f"edit_btitle_{slide.slide_number}")
                        st.markdown("**Bullet Points**")
                        for i, b in enumerate(content.get("bullets", [])):
                            content["bullets"][i] = st.text_area(f"Bullet {i+1}", value=b, height=68, key=f"edit_bull_{slide.slide_number}_{i}", label_visibility="collapsed")
                    with col_i:
                        st.image(content["generated_image_path"], use_container_width=True)
                else:
                    content["title"] = st.text_input("Box Title", value=content.get("title", ""), key=f"edit_btitle_{slide.slide_number}")
                    st.markdown("**Bullet Points**")
                    for i, b in enumerate(content.get("bullets", [])):
                        content["bullets"][i] = st.text_input(f"Bullet {i+1}", value=b, key=f"edit_bull_{slide.slide_number}_{i}", label_visibility="collapsed")
            
            elif ct == "two_column":
                content["title"] = st.text_input("Slide Title", value=content.get("title", ""), key=f"edit_2ctitle_{slide.slide_number}")
                if content.get("generated_image_path") and os.path.exists(content["generated_image_path"]):
                    c1, c2, c3 = st.columns([1.5, 1.5, 2.5])
                    with c1:
                        content["left_heading"] = st.text_input("Left Heading", value=content.get('left_heading', ''), key=f"edit_lh_{slide.slide_number}")
                        for i, b in enumerate(content.get("left_bullets", [])):
                            content["left_bullets"][i] = st.text_area(f"L{i}", value=b, height=68, key=f"edit_lb_{slide.slide_number}_{i}", label_visibility="collapsed")
                    with c2:
                        content["right_heading"] = st.text_input("Right Heading", value=content.get('right_heading', ''), key=f"edit_rh_{slide.slide_number}")
                        for i, b in enumerate(content.get("right_bullets", [])):
                            content["right_bullets"][i] = st.text_area(f"R{i}", value=b, height=68, key=f"edit_rb_{slide.slide_number}_{i}", label_visibility="collapsed")
                    with c3:
                        st.image(content["generated_image_path"], use_container_width=True)
                else:
                    c1, c2 = st.columns(2)
                    with c1:
                        content["left_heading"] = st.text_input("Left Heading", value=content.get('left_heading', ''), key=f"edit_lh_{slide.slide_number}")
                        for i, b in enumerate(content.get("left_bullets", [])):
                            content["left_bullets"][i] = st.text_input(f"L{i}", value=b, key=f"edit_lb_{slide.slide_number}_{i}", label_visibility="collapsed")
                    with c2:
                        content["right_heading"] = st.text_input("Right Heading", value=content.get('right_heading', ''), key=f"edit_rh_{slide.slide_number}")
                        for i, b in enumerate(content.get("right_bullets", [])):
                            content["right_bullets"][i] = st.text_input(f"R{i}", value=b, key=f"edit_rb_{slide.slide_number}_{i}", label_visibility="collapsed")
            
            elif ct == "stat_callout":
                st.subheader(content.get("title", ""))
                stats = content.get("stats", [])
                if stats:
                    col_stats = st.columns(len(stats))
                    for i, stat in enumerate(stats):
                        col_stats[i].metric(label=stat.get("label", ""), value=stat.get("value", ""))
                if content.get("body_text"):
                    st.info(content.get("body_text"))
            
            elif ct == "diagram":
                # ── Diagram title ─────────────────────────────────────────────
                content["title"] = st.text_input(
                    "Diagram Title", value=content.get("title", ""),
                    key=f"edit_diag_title_{slide.slide_number}"
                )

                steps = content.get("steps", [])
                
                # ── Current assembled diagram preview ─────────────────────────
                img_path = content.get("diagram_image")
                if img_path and os.path.exists(img_path):
                    st.image(img_path, use_container_width=True)
                else:
                    st.warning("Diagram image not found.")

                if steps:
                    st.markdown("---")
                    st.markdown(
                        "<h4 style='margin-bottom:4px;'>✏️ Edit Flow Steps</h4>"
                        "<p style='color:#64748b;font-size:0.88rem;margin-top:0;'>"
                        "Edit the icon image, label, or description of each step. "
                        "Click <b>🔄 Regen Icon</b> to create a new AI icon, then "
                        "<b>🖼️ Re-Assemble Diagram</b> to apply all changes.</p>",
                        unsafe_allow_html=True
                    )
                    
                    # Render steps in rows of 3
                    COLS_PER_ROW = 3
                    for row_start in range(0, len(steps), COLS_PER_ROW):
                        row_steps = steps[row_start : row_start + COLS_PER_ROW]
                        cols = st.columns(COLS_PER_ROW)
                        for col_idx, step in enumerate(row_steps):
                            i     = row_start + col_idx
                            snum  = step.get("step_number", i + 1)
                            with cols[col_idx]:
                                # Card container
                                st.markdown(
                                    f"<div style='background:#f8fafc;border:1.5px solid #e2e8f0;"
                                    f"border-radius:12px;padding:12px 10px 8px;text-align:center;"
                                    f"margin-bottom:4px;'>"
                                    f"<span style='background:#1e293b;color:white;border-radius:50%;"
                                    f"padding:2px 8px;font-size:0.75rem;font-weight:bold;'>{snum}</span>"
                                    f"</div>",
                                    unsafe_allow_html=True
                                )
                                # Icon image
                                icon_path = step.get("icon_path", "")
                                if icon_path and os.path.exists(icon_path):
                                    st.image(icon_path, use_container_width=True)
                                else:
                                    st.markdown(
                                        f"<div style='height:100px;background:#334155;border-radius:50px;"
                                        f"display:flex;align-items:center;justify-content:center;"
                                        f"color:white;font-size:2rem;margin-bottom:6px;'>{snum}</div>",
                                        unsafe_allow_html=True
                                    )
                                
                                # Regen icon button
                                if st.button("🔄 Regen Icon", key=f"regen_{slide.slide_number}_{i}",
                                             use_container_width=True):
                                    prompt = step.get("icon_prompt") or step.get("label", "business icon")
                                    with st.spinner(f"Generating icon…"):
                                        new_path = regenerate_step_icon(prompt, snum)
                                    steps[i]["icon_path"] = new_path.replace("\\", "/")
                                    content["steps"] = steps
                                    st.rerun()

                                # Label
                                steps[i]["label"] = st.text_input(
                                    "Label", value=step.get("label", ""),
                                    key=f"lbl_{slide.slide_number}_{i}",
                                    placeholder="Short step title"
                                )
                                # Description
                                steps[i]["description"] = st.text_area(
                                    "Description", value=step.get("description", ""),
                                    key=f"desc_{slide.slide_number}_{i}",
                                    height=90,
                                    placeholder="One sentence description"
                                )
                                # Icon prompt (collapsed in a small expander to keep UI clean)
                                with st.expander("🖊 Change icon prompt"):
                                    steps[i]["icon_prompt"] = st.text_input(
                                        "Icon prompt", value=step.get("icon_prompt", ""),
                                        key=f"prompt_{slide.slide_number}_{i}",
                                        label_visibility="collapsed"
                                    )

                    content["steps"] = steps

                    # ── Re-Assemble button ─────────────────────────────────────
                    st.markdown("---")
                    if st.button("🖼️ Re-Assemble Diagram from Edits",
                                 key=f"reassemble_{slide.slide_number}",
                                 use_container_width=True, type="primary"):
                        with st.spinner("Re-assembling all three diagram layouts from your edits…"):
                            new_paths = reassemble_diagram(
                                steps_data=steps,
                                title=content.get("title", "Process Flow"),
                                user_query=st.session_state.get("last_input", "")
                            )
                        content["diagram_image"]  = new_paths["diagram"]
                        content["diagram_images"] = new_paths
                        selected_layout = slide.layout_override or "diagram"
                        if selected_layout in new_paths:
                            content["diagram_image"] = new_paths[selected_layout]
                        st.success("✅ Diagram re-assembled!")
                        st.rerun()
                    
            elif ct == "timeline":
                st.subheader(content.get("title", ""))
                for e in content.get("events", []):
                    st.markdown(f"**{e.get('year', '')}:** {e.get('label', '')} (*{e.get('detail', '')}*)")
                    
            elif ct == "table":
                st.subheader(content.get("title", ""))
                import pandas as pd
                headers = content.get("headers", [])
                rows = content.get("rows", [])
                if headers and rows:
                    df = pd.DataFrame(rows, columns=headers)
                    st.table(df)
            
            elif ct == "quote":
                st.subheader(content.get("title", ""))
                st.markdown(f"> *{content.get('quote', '')}*")
                st.markdown(f"**— {content.get('attribution', '')}**")
                
            elif ct == "thank_you":
                st.title(content.get("title", "Thank You"))
                st.subheader(content.get("message", ""))
                st.caption(content.get("contact", ""))
                
            if content.get("speaker_notes"):
                st.divider()
                st.caption(f"**Speaker Notes:** {content.get('speaker_notes')}")

    st.markdown(f'''
    <div class="download-card" style="margin-bottom: 25px;">
        <h2 style="color: #1E2761; margin-top: 0;">🎉 Presentation Drafted!</h2>
        <p style="color: #64748b; margin-bottom: 25px;">Adjust the design templates for any slide, then compile your selections into a PPTX.</p>
    </div>
    ''', unsafe_allow_html=True)
    
    # Compile Button
    if st.button("Compile PPTX with Selected Designs ⚙️", use_container_width=True):
        with st.spinner("Compiling PowerPoint..."):
            try:
                pid = uuid.uuid4().hex[:8]
                st.session_state.target_data = export_pptx(
                    st.session_state.generated_slides, 
                    st.session_state.generated_theme, 
                    filename=f"presentation_{pid}.pptx"
                )
            except Exception as e:
                st.error(f"Error compiling PPTX: {e}")
                
    # Download Button
    if st.session_state.target_data:
        st.success("Compilation successful!")
        st.download_button(
            label="Download Final PPTX 💾",
            data=st.session_state.target_data,
            file_name="presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )
