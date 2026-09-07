# Renders the final exam paper for the two built-in templates. python-docx
# only wraps a slice of Word's XML, so several helpers below write raw OOXML
# through lxml: table widths, cell shading, RTL runs, PAGE/NUMPAGES fields.
import re

from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

URDU_FONT = "Noto Nastaliq Urdu"   
ENGLISH_FONT = "Times New Roman"

# labels like "Subject: ..." that belong to the masthead. the header table
# already renders all of these, so strip_header_lines drops them from the
# extracted body text to avoid printing things twice
HEADER_LABEL_PATTERNS = [
    r"school\s*name\s*:", r"subject\s*:", r"class\s*:", r"term\s*:",
    r"session\s*:", r"total\s*marks\s*:", r"time\s*allowed\s*:", r"time\s*:",
    r"roll\s*(no|number)\s*:", r"student'?s?\s*name\s*:", r"name\s*:", r"date\s*:",
]

# Matches strings like "a) oxygen", "B) 9", " c)  none "
_MCQ_OPTION_RE = re.compile(r"^\s*([a-e])\)\s*(.*)$", re.IGNORECASE)


def strip_header_lines(text):
    """Remove any leftover masthead-style lines (e.g. 'Roll No: ___') from
    extracted text, since that info is rendered separately in the header table."""
    if not text:
        return text
    lines = text.split("\n")
    kept = [
        line for line in lines
        if not any(re.search(p, line, re.IGNORECASE) for p in HEADER_LABEL_PATTERNS)
    ]
    return "\n".join(kept).strip()


def sanitize_data(data):
    data = dict(data)  # shallow copy — section dicts stay shared, but data is only rendered once
    data["exam_title"] = strip_header_lines(data.get("exam_title", ""))
    data["instructions"] = [
        line for line in (strip_header_lines(i) for i in data.get("instructions", []))
        if line
    ]
    for sec in data.get("sections", []):
        cleaned = strip_header_lines(sec.get("section_title", ""))
        if cleaned:
            sec["section_title"] = cleaned
    return data


def _looks_like_mcq_options(sub_parts):
    """Return True when every sub_part looks like an MCQ option (a) ... b) ...)."""
    if not sub_parts or len(sub_parts) < 2:
        return False
    return all(_MCQ_OPTION_RE.match(sp) for sp in sub_parts)


def _format_mcq_options(sub_parts, usable_chars):
    """Format MCQ options inline.

    - Short option sets are placed on a single line with evenly spaced tabs.
    - Longer sets are split into two options per line, also tab-aligned.
    """
    options = []
    for sp in sub_parts:
        m = _MCQ_OPTION_RE.match(sp)
        letter = m.group(1).lower()
        text = m.group(2).strip()
        options.append((letter, text))

    rendered = [f"{letter}) {text}" for letter, text in options]

    # Account for separators between options.
    total_len = sum(len(r) for r in rendered) + (len(rendered) - 1) * 2

    use_single_line = total_len <= usable_chars

    if use_single_line:
        return use_single_line, rendered

    groups = [rendered[i:i + 2] for i in range(0, len(rendered), 2)]
    return use_single_line, groups


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
        # run.font.name only sets the latin typeface. Urdu is a "complex
        # script" in Word, so the font also has to go into w:cs or the text
        # comes out in some fallback font
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


def _set_cell_text(cell, text, bold=False, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    if align:
        p.alignment = align
    run = p.add_run(str(text))
    run.bold = bold
    return p


def _to_twips(value):
    """Convert a Length object or an integer of twips to integer twips."""
    return int(value.twips) if hasattr(value, "twips") else int(value)


def _set_cell_width(cell, width):
    """Set an exact width for a table cell (used by the borderless top section)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.append(tcW)
    tcW.set(qn("w:type"), "dxa")
    tcW.set(qn("w:w"), str(_to_twips(width)))


def _set_table_fixed_width(table, width):
    """Make a table exactly `width` wide with fixed column layout."""
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(_to_twips(width)))

    tblLayout = tblPr.find(qn("w:tblLayout"))
    if tblLayout is None:
        tblLayout = OxmlElement("w:tblLayout")
        tblPr.append(tblLayout)
    tblLayout.set(qn("w:type"), "fixed")


def _remove_table_borders(table):
    """Remove all borders from a table (for the invisible top-section layout)."""
    tblPr = table._tbl.tblPr
    tblBorders = tblPr.find(qn("w:tblBorders"))
    if tblBorders is None:
        tblBorders = OxmlElement("w:tblBorders")
        tblPr.append(tblBorders)
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = tblBorders.find(qn(f"w:{side}"))
        if border is None:
            border = OxmlElement(f"w:{side}")
            tblBorders.append(border)
        border.set(qn("w:val"), "none")
        border.set(qn("w:sz"), "0")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "auto")


def _add_top_section(doc, header_info, logo_path=None):
    """School logo on the left with the school name and session centered on
    the *full page width*, not just the space beside the logo.

    This is done with a three-column borderless table: the left and right
    columns are the same width, so the middle column (which holds the text)
    is perfectly centered on the page regardless of whether a logo is placed
    in the left column.
    """
    section = doc.sections[0]
    # python-docx length arithmetic returns EMUs, so compute widths directly in
    # twips to avoid accidentally writing EMU values into the XML width attrs.
    avail_width = (
        section.page_width.twips
        - (section.left_margin.twips if section.left_margin else 0)
        - (section.right_margin.twips if section.right_margin else 0)
    )

    logo_col_width = Inches(1.0).twips
    if avail_width < 2 * logo_col_width + Inches(2).twips:
        logo_col_width = Inches(0.5).twips
    text_col_width = avail_width - 2 * logo_col_width

    top_table = doc.add_table(rows=1, cols=3)
    _remove_table_borders(top_table)
    _set_table_fixed_width(top_table, avail_width)
    top_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    logo_cell = top_table.cell(0, 0)
    text_cell = top_table.cell(0, 1)
    right_cell = top_table.cell(0, 2)
    _set_cell_width(logo_cell, logo_col_width)
    _set_cell_width(text_cell, text_col_width)
    _set_cell_width(right_cell, logo_col_width)

    if logo_path:
        p = logo_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(logo_path, width=Inches(0.9))

    title_p = text_cell.paragraphs[0]
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run(header_info.get("school_name", "").upper())
    run.bold = True
    run.font.size = Pt(16)

    subtitle_parts = [p for p in [header_info.get("term", ""), header_info.get("session", "")] if p]
    if subtitle_parts:
        sub_p = text_cell.add_paragraph()
        sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = sub_p.add_run(f"({' — '.join(subtitle_parts)})")
        r2.bold = True 
        r2.font.size = Pt(12)


def add_header_table(doc, header_info, logo_path=None):
    """Default masthead: logo + centered school name/session above a 4-column
    Subject/Class/Name/Roll/Time/Marks/Date grid — matches the standard
    school exam header layout."""
    _add_top_section(doc, header_info, logo_path)

    table = doc.add_table(rows=5, cols=4)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    grid = [
        ("Subject:", header_info.get("subject", ""), "Class:", header_info.get("class_name", "")),
        ("Student's Name:", "", "Time Allowed:", header_info.get("time", "")),
        ("Roll Number:", "", "Total Marks:", header_info.get("total_marks", "")),
        ("Date:", "", "Obtained Marks:", ""),
        ("Invigilator Sign.", "", "Remarks:", ""),
    ]
    for row_idx, (l1, v1, l2, v2) in enumerate(grid):
        cells = table.rows[row_idx].cells
        _set_cell_text(cells[0], l1, bold=True)
        _set_cell_text(cells[1], v1)
        _set_cell_text(cells[2], l2, bold=True)
        _set_cell_text(cells[3], v2)

    doc.add_paragraph()


def _section_type_label(section):
    """Objective if any question has a/b/c/d options, otherwise subjective."""
    for q in section.get("questions", []):
        if _looks_like_mcq_options(q.get("sub_parts", [])):
            return "Objective Type"
    return "Subjective Type"


def add_body(doc, data, template="template1"):
    """The exam instructions, sections and questions. Shared by both
    built-in templates AND the custom-template flow (see custom_template.py)."""
    data = sanitize_data(data)
    is_urdu = data.get("language", "english").lower() == "urdu"
    font = URDU_FONT if is_urdu else ENGLISH_FONT

    # skip the exam title here; subject/class/etc already sit in the header table
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

    prev_type = None
    for sec in data.get("sections", []):
        # only print the type label when it changes between sections
        sec_type = _section_type_label(sec)
        if sec_type != prev_type:
            type_p = doc.add_paragraph()
            type_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = type_p.add_run(f"({sec_type})")
            set_font(r, font, 13, bold=True, urdu=is_urdu)
            if is_urdu:
                set_rtl(type_p)
            prev_type = sec_type

        sec_p = doc.add_paragraph()
        sec_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        marks_txt = f"  ({sec.get('section_marks','')})" if sec.get("section_marks") else ""
        r = sec_p.add_run(f"{sec.get('section_title','')}{marks_txt}")
        set_font(r, font, 13, bold=True, urdu=is_urdu)
        if template == "template2":
            shade(sec_p)
        if is_urdu:
            set_rtl(sec_p)

        for q in sec.get("questions", []):
            q_p = doc.add_paragraph()
            q_p.paragraph_format.left_indent = Cm(0.5)  # slight indent for question numbers
            marks_str = f"   [{q.get('marks','')}]" if q.get("marks") else ""
            r = q_p.add_run(f"{q.get('number','')}. {q.get('text','')}{marks_str}")
            set_font(r, font, 12, urdu=is_urdu)
            if is_urdu:
                set_rtl(q_p)
            sub_parts = q.get("sub_parts", [])
            if _looks_like_mcq_options(sub_parts):
                section = doc.sections[0]
                # page geometry comes out in twips; 567 twips = 1 cm
                content_width_cm = (
                    section.page_width.twips
                    - section.left_margin.twips
                    - section.right_margin.twips
                ) / 567.0
                # usable width after the option paragraph's left indent
                usable_width_cm = max(content_width_cm - 1.0, 6.0)
                # 11pt Times-like font ≈ 0.28 cm per average character
                usable_chars = int(usable_width_cm / 0.28)

                use_single_line, layout = _format_mcq_options(sub_parts, usable_chars)

                opt_p = doc.add_paragraph()
                opt_p.paragraph_format.left_indent = Cm(1)
                tabs = opt_p.paragraph_format.tab_stops

                if use_single_line:
                    # Evenly space N options across the usable width.
                    for i in range(1, len(layout)):
                        tabs.add_tab_stop(Cm(usable_width_cm * i / len(layout)))
                    text = "\t".join(layout)
                else:
                    # Two columns per line at 50% width.
                    tabs.add_tab_stop(Cm(usable_width_cm / 2))
                    text = "\n".join("\t".join(group) for group in layout)

                r2 = opt_p.add_run(text)
                set_font(r2, font, 11, urdu=is_urdu)
                if is_urdu:
                    set_rtl(opt_p)
            else:
                for sp in sub_parts:
                    sp_p = doc.add_paragraph()
                    sp_p.paragraph_format.left_indent = Cm(1)
                    r2 = sp_p.add_run(sp)
                    set_font(r2, font, 11, urdu=is_urdu)
                    if is_urdu:
                        set_rtl(sp_p)


def _add_fld_simple(paragraph, instruction, placeholder="1", bold=False,
                    color=None):
    """Insert a Word field (e.g. PAGE or NUMPAGES) into a paragraph."""
    fld_simple = OxmlElement("w:fldSimple")
    fld_simple.set(qn("w:instr"), instruction)

    # Include a cached result placeholder so the number is visible even before
    # Word updates fields; Word refreshes it to the actual page count on open.
    run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    if bold:
        rPr.append(OxmlElement("w:b"))
    if color:
        c = OxmlElement("w:color")
        c.set(qn("w:val"), color)
        rPr.append(c)
    run.append(rPr)

    t = OxmlElement("w:t")
    t.text = placeholder
    run.append(t)

    fld_simple.append(run)
    paragraph._p.append(fld_simple)


def _set_paragraph_bottom_border(paragraph, size=4, color="000000"):
    """Add a tight bottom border to a paragraph."""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    bottom = pBdr.find(qn("w:bottom"))
    if bottom is None:
        bottom = OxmlElement("w:bottom")
        pBdr.append(bottom)
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "0")
    bottom.set(qn("w:color"), color)


def _add_running_header(section, header_info):
    """Running header for page 2 onward. Page 1 keeps its full masthead."""
    section.different_first_page_header_footer = True

    header = section.header
    p = header.paragraphs[0]
    p.text = ""
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(0)

    subject = header_info.get("subject", "")
    class_name = header_info.get("class_name", "")
    # same twips -> cm conversion as in add_body (567 twips = 1 cm)
    content_width_cm = (
        section.page_width.twips
        - section.left_margin.twips
        - section.right_margin.twips
    ) / 567.0

    # Three tab stops at 25%, 50%, and 100% (right-aligned) of text width.
    p.paragraph_format.tab_stops.add_tab_stop(
        Cm(content_width_cm * 0.25), WD_TAB_ALIGNMENT.LEFT
    )
    p.paragraph_format.tab_stops.add_tab_stop(
        Cm(content_width_cm * 0.50), WD_TAB_ALIGNMENT.LEFT
    )
    p.paragraph_format.tab_stops.add_tab_stop(
        Cm(content_width_cm), WD_TAB_ALIGNMENT.RIGHT
    )

    blank = "___________"

    def add(text):
        run = p.add_run(text)
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        return run

    add(f"Name: {blank}")
    p.add_run("\t")
    add(f"Roll No. {blank}")
    p.add_run("\t")
    add(f"{subject}-{class_name}")
    add("\tPage ")
    _add_fld_simple(p, "PAGE \\* MERGEFORMAT", bold=True, color="000000")
    add(" of ")
    _add_fld_simple(p, "NUMPAGES \\* MERGEFORMAT", bold=True, color="000000")

    _set_paragraph_bottom_border(p)


def build_docx(data, header_info, logo_path, template="template1", output_path="output.docx"):
    """Assemble one full exam paper: running header, masthead, then the body."""
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)

    _add_running_header(section, header_info)
    add_header_table(doc, header_info, logo_path)
    add_body(doc, data, template)

    doc.save(output_path)
    return output_path