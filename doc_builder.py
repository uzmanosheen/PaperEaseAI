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

# Matches strings like "a) oxygen", "B) 9", " c)  none " (Latin markers),
# and Urdu papers' "(الف", "(ب)", "الف)" style markers.
_MCQ_OPTION_RE = re.compile(r"^\s*([a-eA-E])\s*[)\].:]\s*(.*)$")
_MCQ_OPTION_RE_UR = re.compile(
    r"^\s*([(\[]?)\s*(الف|ب|ج|د|ہ)\s*([)\]۔.:]?)(?:\s+(.*))?$"
)


def _match_mcq_option(sub_part):
    """Return (marker, text) when sub_part looks like an MCQ option, else None.
    Latin markers normalize to "a)" style; Urdu markers to "(الف)" style."""
    m = _MCQ_OPTION_RE.match(sub_part)
    if m:
        return f"{m.group(1).lower()})", m.group(2).strip()
    m = _MCQ_OPTION_RE_UR.match(sub_part)
    # bare single letters (e.g. "ب") are not options unless bracketed,
    # delimited, or followed by option text
    if m and (m.group(1) or m.group(3) or m.group(4)):
        return f"({m.group(2)})", (m.group(4) or "").strip()
    return None


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
    return all(_match_mcq_option(sp) for sp in sub_parts)


def _format_mcq_options(sub_parts, usable_chars):
    """Format MCQ options inline.

    - Short option sets are placed on a single line with evenly spaced tabs.
    - Longer sets are split into two options per line, also tab-aligned.
    """
    options = []
    for sp in sub_parts:
        marker, text = _match_mcq_option(sp)
        options.append((marker, text))

    # Bracketed (Urdu) markers get LRE/PDF so Word does not bidi-mirror
    # the parentheses into ")الف(" inside an RTL paragraph
    rendered = [
        f"\u202A{marker}\u202C {text}" if marker.startswith("(") else f"{marker} {text}"
        for marker, text in options
    ]

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
    bidi = pPr.find(qn("w:bidi"))
    if bidi is None:
        bidi = OxmlElement("w:bidi")
        # w:bidi must come before spacing/ind/jc in w:pPr — appended at the
        # end, some renderers silently drop it and digits/punctuation keep
        # their English (left-to-right) order inside the Urdu line
        for tag in ("w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
                    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap",
                    "w:jc", "w:rPr", "w:sectPr"):
            successor = pPr.find(qn(tag))
            if successor is not None:
                successor.addprevious(bidi)
                break
        else:
            pPr.append(bidi)
    bidi.set(qn("w:val"), "1")


def _rtl_isolate(text):
    """Wrap text in RLE/PDF so digits and punctuation keep their Urdu
    right-to-left positions even in renderers that apply bidi loosely."""
    return f"\u202B{text}\u202C" if text else text


def _ltr_isolate(text):
    """Wrap text in LRE/PDF so brackets and digits keep their written
    left-to-right orientation (used for the masthead's (term — session))."""
    return f"\u202A{text}\u202C" if text else text


def _tighten(paragraph, line_pt):
    """Compact spacing for Urdu lines — Nastaliq's tall default line box
    otherwise bloats the masthead and every table row."""
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.line_spacing = Pt(line_pt)


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


def _set_cell_text(cell, text, bold=False, align=None, urdu=False):
    cell.text = ""
    p = cell.paragraphs[0]
    if align:
        p.alignment = align
    value = str(text)
    if urdu and value:
        value = _rtl_isolate(value)
    run = p.add_run(value)
    if urdu:
        set_font(run, URDU_FONT, 11, bold=bold, urdu=True)
        # no paragraph-level bidi here — the RLE isolation above already
        # orders the text, and without it w:jc="right" unambiguously pins
        # the text to the physical right edge of the cell
        _tighten(p, 18)
    else:
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


# ── Urdu transliteration ─────────────────────────────────────
# Urdu papers get their masthead written in Urdu script. Text that is
# already in Urdu (or digits/punctuation) passes through untouched; Latin
# words are transliterated phonetically — the word list below covers the
# vocabulary that shows up in school names and exam headings
# ("XYZ Education System" -> "ایکس وائی زیڈ ایجوکیشن سسٹم",
# "Final Term" -> "فائنل ٹرم"), and words missing from it fall back to a
# grapheme-level mapping.
_URDU_WORDS = {
    "al": "ال", "the": "دی", "of": "آف", "and": "اینڈ", "for": "فور",
    "education": "ایجوکیشن", "educational": "ایجوکیشنل", "system": "سسٹم",
    "school": "اسکول", "schools": "اسکولز", "public": "پبلک",
    "grammar": "گرامر", "model": "ماڈل", "modern": "ماڈرن", "ideal": "آئیڈیل",
    "standard": "اسٹینڈرڈ", "national": "نیشنل", "international": "انٹرنیشنل",
    "academy": "اکیڈمی", "college": "کالج", "campus": "کیمپس",
    "branch": "برانچ", "high": "ہائی", "higher": "ہائر",
    "secondary": "سیکنڈری", "middle": "مڈل", "primary": "پرائمری",
    "science": "سائنس", "arts": "آرٹس", "commerce": "کامرس",
    "computer": "کمپیوٹر", "islamic": "اسلامک", "islamia": "اسلامیہ",
    "foundation": "فاؤنڈیشن", "trust": "ٹرسٹ", "group": "گروپ",
    "city": "سٹی", "star": "اسٹار", "bright": "برائٹ", "level": "لیول",
    "future": "فیوچر", "boys": "بوائز", "girls": "گرلز", "kids": "کڈز",
    "year": "ایئر", "years": "ایئرز",
    "pak": "پاک", "pakistan": "پاکستان", "studies": "اسٹڈیز",
    "general": "جنرل", "knowledge": "نالج", "social": "سوشل",
    "ali": "علی", "ahmed": "احمد", "ahmad": "احمد", "umar": "عمر",
    "umer": "عمر", "usman": "عثمان", "umair": "عمیر", "syed": "سید",
    "hassan": "حسن", "hussain": "حسین", "fatima": "فاطمہ", "aisha": "عائشہ",
    "final": "فائنل", "mid": "مڈ", "term": "ٹرم", "first": "فرسٹ",
    "second": "سیکنڈ", "third": "تھرڈ", "annual": "اینول",
    "fourth": "فورتھ", "fifth": "ففٹھ", "sixth": "سکستھ", "seventh": "سیونتھ",
    "eighth": "ایتھ", "ninth": "نائنتھ", "tenth": "ٹینتھ",
    "eleventh": "الیونتھ", "twelfth": "ٹولفتھ",
    "session": "سیشن", "exam": "ایگزام", "class": "کلاس",
    "hour": "گھنٹہ", "hours": "گھنٹے", "hr": "گھنٹہ", "hrs": "گھنٹے",
    "min": "منٹ", "mins": "منٹ", "minute": "منٹ", "minutes": "منٹ",
    "urdu": "اردو", "english": "انگلش", "physics": "فزکس",
    "chemistry": "کیمسٹری", "biology": "بائیولوجی", "math": "میتھ",
    "maths": "میتھ", "mathematics": "میتھمیٹکس", "islamiat": "اسلامیات",
    "arabic": "عربی", "geography": "جیوگرافی", "history": "ہسٹری",
    "drawing": "ڈرائنگ", "health": "ہیلتھ",
}

# Letters spelled out by name, for vowel-less acronym tokens like "XYZ"
_LETTER_NAMES = {
    "a": "اے", "b": "بی", "c": "سی", "d": "ڈی", "e": "ای", "f": "ایف",
    "g": "جی", "h": "ایچ", "i": "آئی", "j": "جے", "k": "کے", "l": "ایل",
    "m": "ایم", "n": "این", "o": "او", "p": "پی", "q": "کیو", "r": "آر",
    "s": "ایس", "t": "ٹی", "u": "یو", "v": "وی", "w": "ڈبلیو", "x": "ایکس",
    "y": "وائی", "z": "زیڈ",
}

# Multi-letter graphemes checked longest-first before single letters run.
_TRANSLIT_DIGRAPHS = [
    ("ation", "یشن"), ("ition", "یشن"), ("tion", "شن"), ("sion", "ژن"),
    ("ight", "ائٹ"), ("sure", "ژر"), ("ture", "چر"), ("ough", "و"),
    ("ite", "ائٹ"),
    ("ch", "چ"), ("sh", "ش"), ("th", "تھ"), ("ph", "ف"), ("wh", "و"),
    ("gh", "گ"), ("ck", "ک"), ("kn", "ن"), ("wr", "ر"), ("qu", "کو"),
    ("ng", "نگ"),
    ("ee", "ی"), ("ea", "ی"), ("oo", "و"), ("ou", "اؤ"), ("ow", "او"),
    ("oi", "وئ"), ("oy", "وے"), ("ay", "ے"), ("ai", "ی"), ("ey", "ی"),
]

_SINGLE_LETTERS = {
    "a": "ا", "b": "ب", "d": "ڈ", "f": "ف", "g": "گ", "h": "ہ", "j": "ج",
    "k": "ک", "l": "ل", "m": "م", "n": "ن", "p": "پ", "q": "ق", "r": "ر",
    "s": "س", "t": "ٹ", "v": "و", "w": "و", "x": "کس", "y": "ی", "z": "ز",
}

_VOWELS = "aeiou"


def _translit_graphemes(word):
    """Grapheme-level fallback for words missing from _URDU_WORDS."""
    out = []
    i, n = 0, len(word)
    while i < n:
        for src, ur in _TRANSLIT_DIGRAPHS:
            if word.startswith(src, i):
                out.append(ur)
                i += len(src)
                break
        else:
            ch = word[i]
            prev = word[i - 1] if i else ""
            nxt = word[i + 1] if i + 1 < n else ""
            both_sides_consonant = prev not in _VOWELS and nxt not in _VOWELS
            if ch == "c":
                out.append("س" if nxt in "eiy" else "ک")
            elif ch == "e":
                # silent at the end ("name" -> نام) and as a schwa between
                # consonants ("term" -> ٹرم)
                if i == 0:
                    out.append("ای")
                elif i != n - 1 and not both_sides_consonant:
                    out.append("ی")
            elif ch == "i":
                if i == 0:
                    out.append("ا")
                elif i == n - 1 or not both_sides_consonant:
                    out.append("ی")
            elif ch == "o":
                out.append("و")
            elif ch == "u":
                if i == 0:
                    out.append("ا")
                elif not both_sides_consonant:
                    out.append("و")
            else:
                out.append(_SINGLE_LETTERS.get(ch, ""))
            i += 1
    return "".join(out)


def _translit_word(word):
    """Transliterate one Latin word, preferring the known-word list."""
    low = word.lower()
    if low in _URDU_WORDS:
        return _URDU_WORDS[low]
    # vowel-less short tokens are acronyms — spell each letter by name
    if len(low) <= 4 and not any(v in low for v in _VOWELS):
        return " ".join(_LETTER_NAMES[c] for c in low)
    return _translit_graphemes(low)


# Ordinals: a standalone ordinal uses the classical Urdu word ("9th" ->
# نہم), one describing a following noun is transliterated from the English
# word ("1st year" -> فرسٹ ایئر, "1st Term" -> فرسٹ ٹرم), and anything above
# 10 just keeps the number ("11th" -> 11).
_URDU_ORDINALS = {
    1: "اول", 2: "دوم", 3: "سوم", 4: "چہارم", 5: "پنجم",
    6: "ششم", 7: "ہفتم", 8: "ہشتم", 9: "نہم", 10: "دہم",
}
_ORDINAL_WORDS = {
    1: "first", 2: "second", 3: "third", 4: "fourth", 5: "fifth",
    6: "sixth", 7: "seventh", 8: "eighth", 9: "ninth", 10: "tenth",
}


def _descriptive_ordinal(match):
    number = int(match.group(1))
    return _ORDINAL_WORDS[number] if number <= 10 else str(number)


def _standalone_ordinal(match):
    number = int(match.group(1))
    return _URDU_ORDINALS[number] if number <= 10 else str(number)


def transliterate_to_urdu(text):
    """Phonetically transliterate Latin text into Urdu script. Digits,
    punctuation and text already written in Urdu are left untouched."""
    if not text:
        return text
    text = str(text)
    text = re.sub(r"(\d+)\s*(?:st|nd|rd|th)\b(?=\s+[A-Za-z])",
                  _descriptive_ordinal, text, flags=re.IGNORECASE)
    text = re.sub(r"(\d+)\s*(?:st|nd|rd|th)\b", _standalone_ordinal, text,
                  flags=re.IGNORECASE)
    return re.sub(r"[A-Za-z]+", lambda m: _translit_word(m.group(0)), text)


def _add_top_section(doc, header_info, logo_path=None, is_urdu=False):
    """School logo beside the school name and session centered on the *full
    page width*, not just the space beside the logo.

    This is done with a three-column borderless table: the left and right
    columns are the same width, so the middle column (which holds the text)
    is perfectly centered on the page regardless of whether a logo is placed
    in a side column. Urdu papers read right-to-left, so there the logo sits
    to the right of the school name.
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

    left_cell = top_table.cell(0, 0)
    text_cell = top_table.cell(0, 1)
    right_cell = top_table.cell(0, 2)
    _set_cell_width(left_cell, logo_col_width)
    _set_cell_width(text_cell, text_col_width)
    _set_cell_width(right_cell, logo_col_width)

    logo_cell = right_cell if is_urdu else left_cell
    if logo_path:
        p = logo_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(logo_path, width=Inches(0.9))

    school_name = header_info.get("school_name", "")
    title_text = transliterate_to_urdu(school_name) if is_urdu else school_name.upper()
    title_p = text_cell.paragraphs[0]
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run(_rtl_isolate(title_text) if is_urdu else title_text)
    if is_urdu:
        set_font(title_run, URDU_FONT, 16, bold=True, urdu=True)
        set_rtl(title_p)
        _tighten(title_p, 26)
    else:
        title_run.bold = True
        title_run.font.size = Pt(16)

    # ordinals follow the header rules ("1st Term" -> فرسٹ ٹرم, standalone
    # "9th" -> نہم); the session digits stay as they were written
    subtitle_parts = [
        transliterate_to_urdu(part) if is_urdu else part
        for part in [header_info.get("term", ""), header_info.get("session", "")]
        if part
    ]
    if subtitle_parts:
        sub_p = text_cell.add_paragraph()
        sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subtitle_text = f"({' — '.join(subtitle_parts)})"
        # LRE (not RLE) so the brackets display as "( )" instead of getting
        # bidi-mirrored into ") ("
        r2 = sub_p.add_run(_ltr_isolate(subtitle_text) if is_urdu else subtitle_text)
        if is_urdu:
            set_font(r2, URDU_FONT, 12, bold=True, urdu=True)
            set_rtl(sub_p)
            _tighten(sub_p, 20)
        else:
            r2.bold = True
            r2.font.size = Pt(12)


def add_header_table(doc, header_info, logo_path=None, is_urdu=False):
    """Default masthead: logo + centered school name/session above a 4-column
    Subject/Class/Name/Roll/Time/Marks/Date grid — matches the standard
    school exam header layout.

    Urdu papers get the same grid mirrored for right-to-left reading: every
    row runs [label][fill-in cell][label][fill-in cell] from the right edge,
    with Urdu labels and transliterated values."""
    _add_top_section(doc, header_info, logo_path, is_urdu)

    table = doc.add_table(rows=5, cols=4)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    def header_value(key):
        value = header_info.get(key, "")
        return transliterate_to_urdu(value) if is_urdu else value

    if is_urdu:
        # mirrored grid: in every pair the label sits on the RIGHT and the
        # cell to its left stays blank for hand-filling
        grid = [
            (header_value("class_name"), "جماعت:", header_value("subject"), "مضمون:"),
            (header_value("time"), "وقت:", "", "نام طالب علم:"),
            (header_value("total_marks"), "کل نمبر:", "", "رول نمبر:"),
            ("", "حاصل کردہ نمبر:", "", "تاریخ:"),
            ("", "ریمارکس:", "", "دستخط نگران:"),
        ]
        bolds = (False, True, False, True)
        align = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        grid = [
            ("Subject:", header_value("subject"), "Class:", header_value("class_name")),
            ("Student's Name:", "", "Time Allowed:", header_value("time")),
            ("Roll Number:", "", "Total Marks:", header_value("total_marks")),
            ("Date:", "", "Obtained Marks:", ""),
            ("Invigilator Sign.", "", "Remarks:", ""),
        ]
        bolds = (True, False, True, False)
        align = None

    for row_idx, row_values in enumerate(grid):
        for cell, text, bold in zip(table.rows[row_idx].cells, row_values, bolds):
            _set_cell_text(cell, text, bold=bold, align=align, urdu=is_urdu)

    doc.add_paragraph()


# Section-type labels per language — Urdu papers use حصہ معروضی (objective)
# and حصہ انشائیہ (subjective).
_TYPE_LABELS = {
    "english": ("Objective Type", "Subjective Type"),
    "urdu": ("حصہ معروضی", "حصہ انشائیہ"),
}


def _section_type_label(section, is_urdu=False):
    """Objective if any question has a/b/c/d or الف/ب/ج options, else subjective."""
    objective, subjective = _TYPE_LABELS["urdu" if is_urdu else "english"]
    for q in section.get("questions", []):
        if _looks_like_mcq_options(q.get("sub_parts", [])):
            return objective
    return subjective


def add_body(doc, data, template="template1"):
    """The exam instructions, sections and questions. Shared by both
    built-in templates AND the custom-template flow (see custom_template.py)."""
    data = sanitize_data(data)
    is_urdu = data.get("language", "english").lower() == "urdu"
    font = URDU_FONT if is_urdu else ENGLISH_FONT
    # Nastaliq runs tall — Urdu body text is set 2pt smaller throughout
    sz = -2 if is_urdu else 0

    # skip the exam title here; subject/class/etc already sit in the header table
    if data.get("instructions"):
        label_p = doc.add_paragraph()
        r = label_p.add_run("ہدایات:" if is_urdu else "Instructions:")
        set_font(r, font, 11 + sz, bold=True, urdu=is_urdu)
        if is_urdu:
            set_rtl(label_p)
        for ins in data["instructions"]:
            b = doc.add_paragraph(style="List Bullet")
            r = b.add_run(ins)
            set_font(r, font, 11 + sz, urdu=is_urdu)
            if is_urdu:
                set_rtl(b)

    prev_type = None
    for sec in data.get("sections", []):
        # only print the type label when it changes between sections
        sec_type = _section_type_label(sec, is_urdu)
        if sec_type != prev_type:
            type_p = doc.add_paragraph()
            type_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = type_p.add_run(f"({sec_type})")
            set_font(r, font, 13 + sz, bold=True, urdu=is_urdu)
            if is_urdu:
                set_rtl(type_p)
            prev_type = sec_type

        sec_p = doc.add_paragraph()
        sec_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        marks_txt = f"  ({sec.get('section_marks','')})" if sec.get("section_marks") else ""
        sec_title = sec.get("section_title", "")
        if is_urdu:
            # main-question label in Urdu: "Q No. 1:" -> "سوال نمبر 1:"
            sec_title = re.sub(r"^\s*Q\s*No\.?\s*(\d+)\s*:", r"سوال نمبر \1:", sec_title)
        r = sec_p.add_run(f"{sec_title}{marks_txt}")
        set_font(r, font, 13 + sz, bold=True, urdu=is_urdu)
        if template == "template2":
            shade(sec_p)
        if is_urdu:
            set_rtl(sec_p)

        for q in sec.get("questions", []):
            q_p = doc.add_paragraph()
            q_p.paragraph_format.left_indent = Cm(0.5)  # slight indent for question numbers
            marks_str = f"   [{q.get('marks','')}]" if q.get("marks") else ""
            number = q.get("number", "")
            if is_urdu and number:
                # In an RTL paragraph the Unicode bidi algorithm flips a bare
                # leading number to the END of the line. LRE/PDF (U+202A /
                # U+202C) pins "1." to the start of the line — the RIGHT side
                # for Urdu — instead of it landing after the question text.
                q_text = f"\u202A{number}. \u202C{q.get('text','')}{marks_str}"
            else:
                q_text = f"{number}. {q.get('text','')}{marks_str}"
            r = q_p.add_run(q_text)
            set_font(r, font, 12 + sz, urdu=is_urdu)
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
                # 11pt Times ≈ 0.28 cm per average character; Noto Nastaliq
                # Urdu glyphs run noticeably wider, so budget more per char
                char_cm = 0.32 if is_urdu else 0.28
                usable_chars = int(usable_width_cm / char_cm)

                use_single_line, layout = _format_mcq_options(sub_parts, usable_chars)

                opt_p = doc.add_paragraph()
                opt_p.paragraph_format.left_indent = Cm(1)
                if is_urdu:
                    # Urdu options run compact — no gap after the options line
                    opt_p.paragraph_format.space_after = Pt(0)
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
                set_font(r2, font, 11 + sz, urdu=is_urdu)
                if is_urdu:
                    set_rtl(opt_p)
            else:
                for sp in sub_parts:
                    sp_p = doc.add_paragraph()
                    sp_p.paragraph_format.left_indent = Cm(1)
                    r2 = sp_p.add_run(sp)
                    set_font(r2, font, 11 + sz, urdu=is_urdu)
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

    # Urdu papers get an Urdu masthead and info grid (see add_header_table);
    # the running header and body keep their own language handling
    is_urdu = data.get("language", "english").lower() == "urdu"
    _add_running_header(section, header_info)
    add_header_table(doc, header_info, logo_path, is_urdu)
    add_body(doc, data, template)

    doc.save(output_path)
    return output_path