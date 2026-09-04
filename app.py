import os
import time
import tempfile
import threading
import streamlit as st

import extractor
import doc_builder
import custom_template

# ── Page Config ──────────────────────────────────────────────
st.set_page_config(page_title="PaperEaseAI — Exam Paper Formatter", layout="wide")

# ── Custom CSS (blue theme, minimalistic) ───────────────────
st.markdown("""
<style>
/* ── Global overrides ── */
/* Radio button accent color is handled by .streamlit/config.toml primaryColor.
   No custom radio CSS needed — native Streamlit radio circles remain intact. */

/* Primary buttons */
.stButton > button[kind="primary"] {
    background-color: #2563EB !important;
    border-color: #2563EB !important;
    color: #fff !important;
}
.stButton > button[kind="primary"]:hover {
    background-color: #1D4ED8 !important;
    border-color: #1D4ED8 !important;
}
.stDownloadButton > button {
    background-color: #2563EB !important;
    border-color: #2563EB !important;
    color: #fff !important;
    font-size: 1.05rem !important;
}
.stDownloadButton > button:hover {
    background-color: #1D4ED8 !important;
    border-color: #1D4ED8 !important;
}

/* ── Hero ── */
.hero-wrap {
    padding: 3rem 0 1rem;
    text-align: center;
}
.hero-title {
    font-size: 3rem;
    font-weight: 800;
    color: #0F172A;
    margin: 0;
    line-height: 1.15;
}
.hero-accent { color: #2563EB; }
.hero-sub {
    font-size: 1.15rem;
    color: #64748B;
    margin: 1rem auto 0;
    max-width: 640px;
    line-height: 1.7;
}
.hero-divider {
    width: 56px;
    height: 3px;
    background: #2563EB;
    border-radius: 2px;
    margin: 2rem auto;
}

/* ── Feature cards ── */
.feat-card {
    background: #F8FAFC;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 1.5rem 1.25rem;
    text-align: center;
    transition: box-shadow .2s ease;
}
.feat-card:hover { box-shadow: 0 4px 16px rgba(37,99,235,.08); }
.feat-icon { font-size: 2rem; margin-bottom: .65rem; }
.feat-head { font-size: .95rem; font-weight: 700; color: #1E293B; margin-bottom: .3rem; }
.feat-body { font-size: .85rem; color: #64748B; line-height: 1.55; }

/* ── Workspace chrome ── */
.ws-head {
    font-size: 1.55rem;
    font-weight: 700;
    color: #0F172A;
    padding-bottom: .4rem;
    border-bottom: 2px solid #2563EB;
    display: inline-block;
    margin-bottom: .75rem;
}
.sec-label {
    font-size: .72rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: .1em;
    color: #2563EB;
    margin-bottom: .2rem;
}
.sec-title {
    font-size: 1.15rem;
    font-weight: 600;
    color: #1E293B;
    margin-bottom: 1rem;
}
.content-box {
    background: #fff;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 1.75rem;
    margin-bottom: 1.25rem;
}

/* ── Download area ── */
.dl-card {
    background: linear-gradient(135deg, #EFF6FF 0%, #DBEAFE 100%);
    border: 1px solid #BFDBFE;
    border-radius: 14px;
    padding: 2.25rem;
    text-align: center;
    margin: 1.25rem 0;
}
.dl-title { font-size: 1.35rem; font-weight: 700; color: #1E3A5F; margin-bottom: .3rem; }
.dl-sub   { font-size: .88rem; color: #64748B; }

/* ── Utility ── */
div[data-testid="stExpander"] {
    border: 1px solid #E2E8F0 !important;
    border-radius: 10px !important;
}
</style>
""", unsafe_allow_html=True)

# ── Session state defaults ──────────────────────────────────────
if "_started" not in st.session_state:
    st.session_state._started = False

# ── Hero Section (visible before user starts) ───────────────
if not st.session_state._started:
    # Extra spacing for hero layout
    st.markdown('<div style="height:1.5rem"></div>', unsafe_allow_html=True)
    st.markdown(
        '<p style="text-align:center;font-size:2.2rem;color:#2563EB;'
        'font-weight:700;letter-spacing:.02em;margin:0 0 .5rem;'
        'line-height:1.1;">PaperEaseAI</p>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<h1 style="text-align:center;font-size:2.8rem;font-weight:800;'
        'color:#0F172A;margin:0;line-height:1.2;">'
        'Your Exam Papers, <span style="color:#2563EB;">Perfectly Formatted</span>'
        "</h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p style="text-align:center;font-size:1.1rem;color:#64748B;'
        "max-width:620px;margin:1.2rem auto 0;line-height:1.75;"
        'letter-spacing:.01em;">'
        "Upload a handwritten photo, an unformatted document, or paste raw text. "
        "PaperEaseAI uses AI to extract, structure, and format your exam papers "
        "into clean, professional Word documents — in seconds."
        "</p>",
        unsafe_allow_html=True,
    )
    st.markdown('<div class="hero-divider"></div>', unsafe_allow_html=True)

    # Feature cards
    fc1, fc2, fc3 = st.columns([1, 1, 1])
    with fc1:
        st.markdown(
            '<div class="feat-card">'
            '<div class="feat-icon">📸</div>'
            '<div class="feat-head">Upload Anything</div>'
            '<div class="feat-body">Photos, unformatted .docx files, '
            "or pasted text — we handle it all.</div></div>",
            unsafe_allow_html=True,
        )
    with fc2:
        st.markdown(
            '<div class="feat-card">'
            '<div class="feat-icon">🤖</div>'
            '<div class="feat-head">AI-Powered Formatting</div>'
            '<div class="feat-body">Our AI extracts questions, marks, and '
            "structure automatically.</div></div>",
            unsafe_allow_html=True,
        )
    with fc3:
        st.markdown(
            '<div class="feat-card">'
            '<div class="feat-icon">📥</div>'
            '<div class="feat-head">Download &amp; Print</div>'
            '<div class="feat-body">Get a polished, print-ready Word document '
            "in seconds.</div></div>",
            unsafe_allow_html=True,
        )

    st.markdown('<div style="height:1.75rem"></div>', unsafe_allow_html=True)

    # CTA button (centered)
    cta_l, cta_m, cta_r = st.columns([1, 1.6, 1])
    with cta_m:
        st.button(
            "Format Your Paper",
            type="primary",
            use_container_width=True,
            on_click=lambda: st.session_state.update(_started=True),
        )

    st.markdown('<div style="height:3rem"></div>', unsafe_allow_html=True)
    st.stop()

# ── Auto-scroll to workspace after CTA click ────────────────
st.iframe(
    "data:text/html," + """<script>
(function(){
    var doc = window.parent.document;
    var tries = 0;
    var t = setInterval(function(){
        var el = doc.getElementById('ws-top');
        if (el) { el.scrollIntoView({behavior:'smooth',block:'start'}); clearInterval(t); }
        if (++tries > 40) clearInterval(t);
    }, 60);
})();
</script>""",
    height=1,
)

# ════════════════════════════════════════════════════════════
#  WORKSPACE
# ════════════════════════════════════════════════════════════
st.markdown('<div id="ws-top"></div>', unsafe_allow_html=True)

# Back button + header row
back_col, title_col = st.columns([0.5, 6])
with back_col:
    st.button(
        "← Back",
        key="back_to_hero",
        on_click=lambda: st.session_state.update(_started=False),
    )
with title_col:
    st.markdown(
        '<div class="ws-head">Formatting Studio</div>',
        unsafe_allow_html=True,
    )

uploaded_images, uploaded_docx, pasted_text = None, None, ""

# ── Section: Choose Template (inline, formerly sidebar) ─────
st.markdown(
    '<div class="sec-label">Step 1</div>'
    '<div class="sec-title">Choose Template</div>',
    unsafe_allow_html=True,
)

template_choice = st.radio(
    "Template style",
    ["Template 1 — Simple", "Template 2 — Modern", "Custom (upload school template)"],
    horizontal=True,
    label_visibility="collapsed",
)
custom_tpl_file = None
if template_choice.startswith("Custom"):
    st.caption(
        "Upload your school's exam header/letterhead as a .docx with "
        "placeholders like {{ school_name }}, {{ logo }}, {{ subject }}. "
        "We'll insert the extracted content directly below it."
    )
    custom_tpl_file = st.file_uploader(
        "School template (.docx)", type=["docx"]
    )

# ── Section: Add Exam Content ───────────────────────────────
st.markdown(
    '<div class="sec-label">Step 2</div>'
    '<div class="sec-title">Add Exam Content</div>',
    unsafe_allow_html=True,
)

upload_type = st.radio(
    "Input method",
    ["Handwritten photo(s)", "Typed but unformatted .docx", "Pasted plain text"],
    label_visibility="collapsed",
    horizontal=True,
)

if upload_type == "Handwritten photo(s)":
    uploaded_images = st.file_uploader(
        "Upload photo(s) of the exam paper (in order)",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
    )
elif upload_type == "Typed but unformatted .docx":
    uploaded_docx = st.file_uploader("Upload the .docx file", type=["docx"])
else:
    pasted_text = st.text_area(
        "Paste exam content",
        height=220,
        placeholder=(
            "e.g.\n\nPhysics Mid-Term — Class 9 — Total 50 marks — Time: 2 hrs\n"
            "Attempt all questions.\n\n"
            "Q1. Define velocity. (5)\n"
            "Q2 state newtons first law [5 marks]\n"
            "q3 explain difference between speed and velocity 10 marks"
        ),
    )

# ── Section: Exam Details ───────────────────────────────────
st.markdown(
    '<div class="sec-label">Step 3</div>'
    '<div class="sec-title">Exam Details</div>',
    unsafe_allow_html=True,
)

col1, col2, col3 = st.columns(3)
with col1:
    school_name = st.text_input("Institution name")
    subject = st.text_input("Subject")
with col2:
    class_name = st.text_input("Class/Grade")
    term = st.text_input("Term")
with col3:
    session = st.text_input("Session (e.g. 2025-26)")
    total_marks = st.text_input("Total marks")

time_allowed = st.text_input("Time allowed (e.g. 2 Hours)")
date_value = st.text_input("Date (optional — leave blank if hand-filled)")
logo_file = st.file_uploader("School logo (optional)", type=["png", "jpg", "jpeg"])

# ── Generate Button ─────────────────────────────────────────
st.markdown('<div style="height:.75rem"></div>', unsafe_allow_html=True)
gen_l, gen_m, gen_r = st.columns([1, 2, 1])
with gen_m:
    generate_clicked = st.button(
        "Generate Formatted Exam Docx",
        type="primary",
        use_container_width=True,
    )

# ── Processing & Output ─────────────────────────────────────
if generate_clicked:
    # ── Validation ──
    if not uploaded_images and not uploaded_docx and not pasted_text.strip():
        st.error("Please upload a photo/docx file or paste the exam text first.")
        st.stop()
    if template_choice.startswith("Custom") and not custom_tpl_file:
        st.error("Please upload a school template docx, or pick a built-in template.")
        st.stop()

    # Centered two-line status that updates in place
    status_box = st.empty()
    msg_box = st.empty()

    def _set_status(stage, sub=None):
        html = (
            '<div style="text-align:center;margin:.75rem 0;">'
            f'<p style="font-size:1.15rem;font-weight:700;color:#2563EB;'
            f'margin:0;">⏳ {stage}</p>'
        )
        if sub:
            html += (
                f'<p style="font-size:.9rem;color:#64748B;'
                f'margin:.3rem 0 0;">{sub}</p>'
            )
        html += "</div>"
        status_box.markdown(html, unsafe_allow_html=True)

    # ── Step 1: Extract content ─────────────────────────────
    # The (unchanged) extraction call runs in a worker thread so the
    # status line can keep cycling sub-messages while the AI works.
    _result = {}

    def _extract():
        try:
            if uploaded_images:
                image_data = [(img.read(), img.type) for img in uploaded_images]
                _result["data"] = extractor.extract_from_images(image_data)
            elif uploaded_docx:
                tmp_path = os.path.join(tempfile.gettempdir(), uploaded_docx.name)
                with open(tmp_path, "wb") as f:
                    f.write(uploaded_docx.read())
                _result["data"] = extractor.extract_from_docx(tmp_path)
            else:
                _result["data"] = extractor.extract_from_text(pasted_text)
        except Exception as e:
            _result["error"] = e

    _thread = threading.Thread(target=_extract)
    _thread.start()

    _sub_messages = ["Reading images…", "Analyzing content…", "Almost there…"]
    _i = 0
    while _thread.is_alive():
        _set_status(
            "Extracting Content From Images…", _sub_messages[_i % len(_sub_messages)]
        )
        _i += 1
        _thread.join(timeout=1.2)
    _thread.join()

    if "error" in _result:
        status_box.empty()
        st.error(f"AI extraction failed: {_result['error']}")
        st.stop()
    data = _result["data"]

    # ── Step 2: Build document ──────────────────────────────
    _set_status("Formatting Document Structure…")

    header_info = {
        "school_name": school_name,
        "subject": subject,
        "class_name": class_name,
        "term": term,
        "session": session,
        "total_marks": total_marks,
        "time": time_allowed,
        "date": date_value,
    }

    logo_path = None
    if logo_file:
        logo_path = os.path.join(tempfile.gettempdir(), logo_file.name)
        with open(logo_path, "wb") as f:
            f.write(logo_file.read())

    output_path = os.path.join(tempfile.gettempdir(), "formatted_exam.docx")

    try:
        if template_choice.startswith("Custom"):
            tpl_path = os.path.join(tempfile.gettempdir(), custom_tpl_file.name)
            with open(tpl_path, "wb") as f:
                f.write(custom_tpl_file.read())
            output_path = custom_template.build_custom_docx(
                tpl_path, header_info, logo_path, data, output_path
            )
        else:
            tkey = "template1" if "1" in template_choice else "template2"
            output_path = doc_builder.build_docx(
                data, header_info, logo_path, tkey, output_path
            )
    except Exception as e:
        status_box.empty()
        st.error(f"Document generation failed: {e}")
        st.stop()

    # ── Step 3: Finalize ────────────────────────────────────
    _set_status("Finalizing Your Document…")
    time.sleep(0.5)

    msg_box.success("Your formatted exam paper is ready!")
    time.sleep(0.4)

    # Clear the status line
    status_box.empty()
    msg_box.empty()

    # ── Download Section ────────────────────────────────────
    st.markdown(
        '<div class="dl-card">'
        '<div class="dl-title">📄 Your Exam Paper is Ready!</div>'
        '<div class="dl-sub">Download your professionally formatted Word document below.</div>'
        "</div>",
        unsafe_allow_html=True,
    )

    dl_l, dl_m, dl_r = st.columns([1, 1.5, 1])
    with dl_m:
        with open(output_path, "rb") as f:
            st.download_button(
                "Download Formatted Docx",
                f,
                file_name="formatted_exam.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )


