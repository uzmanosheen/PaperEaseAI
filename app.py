import os
import tempfile
import streamlit as st

import extractor
import doc_builder
import custom_template

st.set_page_config(page_title="AI Exam Paper Formatter", layout="wide")
st.title("📄 AI Exam Paper Formatter")
st.caption("Upload a handwritten or unformatted exam paper → get a clean, downloadable Word document.")

with st.sidebar:
    st.header("1. Template")
    template_choice = st.radio(
        "Choose a template",
        ["Template 1 — Simple", "Template 2 — Modern", "Custom (upload school template)"],
    )
    custom_tpl_file = None
    if template_choice.startswith("Custom"):
        custom_tpl_file = st.file_uploader(
            "Upload school template (.docx with {{ placeholders }})", type=["docx"]
        )

st.header("2. Upload exam content")
upload_type = st.radio("What are you uploading?", ["Handwritten photo(s)", "Typed but unformatted .docx"])

uploaded_images, uploaded_docx = None, None
if upload_type == "Handwritten photo(s)":
    uploaded_images = st.file_uploader(
        "Upload photo(s) of the exam paper (in order)", type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
    )
else:
    uploaded_docx = st.file_uploader("Upload the messy .docx file", type=["docx"])

st.header("3. Exam details")
col1, col2, col3 = st.columns(3)
with col1:
    school_name = st.text_input("School name")
    subject = st.text_input("Subject")
with col2:
    class_name = st.text_input("Class")
    term = st.text_input("Term")
with col3:
    session = st.text_input("Session (e.g. 2025-26)")
    total_marks = st.text_input("Total marks")
time_allowed = st.text_input("Time allowed (e.g. 2 Hours)")
logo_file = st.file_uploader("School logo (optional)", type=["png", "jpg", "jpeg"])

if st.button("🚀 Generate Formatted Paper", type="primary"):
    if not uploaded_images and not uploaded_docx:
        st.error("Please upload a photo or a docx file first.")
        st.stop()
    if template_choice.startswith("Custom") and not custom_tpl_file:
        st.error("Please upload a school template docx, or pick a built-in template.")
        st.stop()

    with st.spinner("Reading and understanding the exam paper with AI..."):
        try:
            if uploaded_images:
                image_data = [(img.read(), img.type) for img in uploaded_images]
                data = extractor.extract_from_images(image_data)
            else:
                tmp_path = os.path.join(tempfile.gettempdir(), uploaded_docx.name)
                with open(tmp_path, "wb") as f:
                    f.write(uploaded_docx.read())
                data = extractor.extract_from_docx(tmp_path)
        except Exception as e:
            st.error(f"AI extraction failed: {e}")
            st.stop()

    st.success("Exam content extracted!")
    with st.expander("Preview extracted data (JSON)"):
        st.json(data)

    header_info = {
        "school_name": school_name, "subject": subject, "class_name": class_name,
        "term": term, "session": session, "total_marks": total_marks, "time": time_allowed,
    }

    logo_path = None
    if logo_file:
        logo_path = os.path.join(tempfile.gettempdir(), logo_file.name)
        with open(logo_path, "wb") as f:
            f.write(logo_file.read())

    output_path = os.path.join(tempfile.gettempdir(), "formatted_exam.docx")

    with st.spinner("Building your Word document..."):
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
            st.error(f"Document generation failed: {e}")
            st.stop()

    st.success("✅ Your formatted exam paper is ready!")
    with open(output_path, "rb") as f:
        st.download_button(
            "⬇️ Download DOCX", f, file_name="formatted_exam.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )