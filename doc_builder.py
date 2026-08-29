from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

URDU_FONT = "Noto Nastaliq Urdu"   # free Google Font, works even if judges'
                                   # PC doesn't have Jameel Noori Nastaleeq
ENGLISH_FONT = "Times New Roman"


def set_rtl(paragraph):
    """Mark a paragraph as right-to-left (needed for Urdu)."""
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1")
    pPr.append(bidi)


def set_font(run, font_name, size=12, bold=False, urdu=False):
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    if urdu:
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.append(rFonts)
        rFonts.set(qn("w:cs"), font_name)
        rFonts.set(qn("w:ascii"), font_name)
        rFonts.set(qn("w:hAnsi"), font_name)
        lang = OxmlElement("w:lang")
        lang.set(qn("w:bidi"), "ur-PK")
        rPr.append(lang)


def shade(paragraph, color="D9D9D9"):
    """Light grey background behind a paragraph (used for Template 2 section headers)."""
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    pPr.append(shd)


def add_header_table(doc, header_info, logo_path=None):
    """The top block: logo + school name, then subject/class/term/session,
    then a name/marks/time row students fill in by hand."""
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    left_cell, right_cell = table.rows[0].cells

    if logo_path:
        p = left_cell.paragraphs[0]
        run = p.add_run()
        run.add_picture(logo_path, width=Inches(1))

    p = right_cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(header_info.get("school_name", ""))
    run.bold = True
    run.font.size = Pt(16)

    info_p = doc.add_paragraph()
    info_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fields = [
        f"Subject: {header_info.get('subject','')}",
        f"Class: {header_info.get('class_name','')}",
        f"Term: {header_info.get('term','')}",
        f"Session: {header_info.get('session','')}",
    ]
    info_p.add_run("   |   ".join(fields))

    fill_table = doc.add_table(rows=2, cols=3)
    fill_table.style = "Table Grid"
    labels = ["Name:", "Marks:", "Time:"]
    values = ["________________", header_info.get("total_marks", ""), header_info.get("time", "")]
    for i, lbl in enumerate(labels):
        fill_table.rows[0].cells[i].text = lbl
    for i, val in enumerate(values):
        fill_table.rows[1].cells[i].text = str(val)

    doc.add_paragraph()  # spacer


def add_body(doc, data, template="template1"):
    """The exam title, instructions, sections and questions. Shared by both
    built-in templates AND the custom-template flow (see custom_template.py)."""
    is_urdu = data.get("language", "english").lower() == "urdu"
    font = URDU_FONT if is_urdu else ENGLISH_FONT

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run(data.get("exam_title", ""))
    set_font(run, font, 15, bold=True, urdu=is_urdu)
    if is_urdu:
        set_rtl(title_p)

    if data.get("instructions"):
        label_p = doc.add_paragraph()
        r = label_p.add_run("ہدایات:" if is_urdu else "Instructions:")
        set_font(r, font, 11, bold=True, urdu=is_urdu)
        if is_urdu:
            set_rtl(label_p)
        for ins in data["instructions"]:
            b = doc.add_paragraph(style="List Bullet")
            r = b.add_run(ins)
            set_font(r, font, 11, urdu=is_urdu)
            if is_urdu:
                set_rtl(b)

    for sec in data.get("sections", []):
        sec_p = doc.add_paragraph()
        sec_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        marks_txt = f"  ({sec.get('section_marks','')})" if sec.get("section_marks") else ""
        r = sec_p.add_run(f"{sec.get('section_title','')}{marks_txt}")
        set_font(r, font, 13, bold=True, urdu=is_urdu)
        if template == "template2":
            shade(sec_p)
        if is_urdu:
            set_rtl(sec_p)

        for q in sec.get("questions", []):
            q_p = doc.add_paragraph()
            marks_str = f"   [{q.get('marks','')}]" if q.get("marks") else ""
            r = q_p.add_run(f"{q.get('number','')}. {q.get('text','')}{marks_str}")
            set_font(r, font, 12, urdu=is_urdu)
            if is_urdu:
                set_rtl(q_p)
            for sp in q.get("sub_parts", []):
                sp_p = doc.add_paragraph()
                sp_p.paragraph_format.left_indent = Cm(1)
                r2 = sp_p.add_run(sp)
                set_font(r2, font, 11, urdu=is_urdu)
                if is_urdu:
                    set_rtl(sp_p)


def build_docx(data, header_info, logo_path, template="template1", output_path="output.docx"):
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)

    add_header_table(doc, header_info, logo_path)
    add_body(doc, data, template)

    doc.save(output_path)
    return output_path