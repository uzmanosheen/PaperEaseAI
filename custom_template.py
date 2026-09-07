# The "Custom" template flow: the school's own letterhead .docx doubles as a
# Jinja2 template ({{ school_name }}, {{ logo }}, ...). docxtpl fills in the
# header fields, then we reopen the file with python-docx and append the
# extracted questions using the same add_body() the built-in templates use.
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from docx import Document
import doc_builder


def build_custom_docx(school_template_path, header_info, logo_path, data, output_path="output_custom.docx"):
    tpl = DocxTemplate(school_template_path)

    context = {
        "school_name": header_info.get("school_name", ""),
        "subject": header_info.get("subject", ""),
        "class_name": header_info.get("class_name", ""),
        "term": header_info.get("term", ""),
        "session": header_info.get("session", ""),
        "total_marks": header_info.get("total_marks", ""),
        "time": header_info.get("time", ""),
        "date": header_info.get("date", ""),
        "exam_title": data.get("exam_title", ""),
    }
    if logo_path:
        # only set the key when a real image exists — InlineImage can't take
        # None, and a missing key just leaves the {{ logo }} placeholder blank
        context["logo"] = InlineImage(tpl, logo_path, width=Inches(1))

    tpl.render(context)
    tpl.save(output_path)

    # Reopen with python-docx and append the actual questions underneath
    # the school's letterhead — reusing the exact same body logic as the
    # built-in templates.
    doc = Document(output_path)
    doc_builder.add_body(doc, data, template="template1")
    doc.save(output_path)

    return output_path