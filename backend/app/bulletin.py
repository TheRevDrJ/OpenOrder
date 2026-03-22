"""Generate a Word bulletin (.docx) from an OrderOfWorship."""

from datetime import date, timedelta
from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

from .models import OrderOfWorship
from . import church_config as cfg

RESOURCES = Path(__file__).parent.parent.parent / "resources"
OUTPUT_DIR = Path(__file__).parent.parent.parent / "output"
QR_CODE_PATH = RESOURCES / "images" / "bulletin_image2.png"

# Dark olive — only used for website link on cover
DARK_OLIVE = RGBColor(0x5A, 0x6B, 0x2E)

# Page content width (7.0 - 0.5 - 0.5 = 6.0 inches)
CONTENT_WIDTH = Inches(6.0)

FONT = "Georgia"

# Bullet character: ❖ (four diamond)
BULLET = "\u2756"


def _ordinal_suffix(day: int) -> str:
    if 11 <= day <= 13:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")


def _add_date_superscript(para, d: date, size=18, bold=True, suffix_text=""):
    """Add 'March 22nd, 2026' with superscript ordinal."""
    day = d.day
    suffix = _ordinal_suffix(day)

    r = para.add_run(f"{d.strftime('%B')} {day}")
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = bold

    r = para.add_run(suffix)
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.superscript = True

    r = para.add_run(f", {d.year}{suffix_text}")
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = bold


def _add_date_short_superscript(para, d: date, prefix="", suffix_text="",
                                 size=12, bold=True, italic=True):
    """Add 'SUNDAY, March 22nd' with superscript ordinal."""
    day = d.day
    suffix = _ordinal_suffix(day)

    r = para.add_run(f"{prefix}{d.strftime('%B')} {day}")
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic

    r = para.add_run(suffix)
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.superscript = True

    if suffix_text:
        r = para.add_run(suffix_text)
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic


def _p(doc, text="", size=12, bold=False, italic=False, align=None,
       before=0, after=0):
    """Add a paragraph. All text is black Georgia."""
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    if text:
        r = p.add_run(text)
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
    return p


def _worship_line(doc, left, right="", before=4, after=4):
    """Left-aligned bold label with optional right-aligned bold text."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)

    r = p.add_run(left)
    r.font.name = FONT
    r.font.size = Pt(12)
    r.font.bold = True

    if right:
        tab_stops = p.paragraph_format.tab_stops
        tab_stops.add_tab_stop(CONTENT_WIDTH, alignment=WD_TAB_ALIGNMENT.RIGHT)
        r = p.add_run(f"\t{right}")
        r.font.name = FONT
        r.font.size = Pt(12)
        r.font.bold = True
    return p


def _new_section(doc):
    """Add a new section with standard page setup."""
    s = doc.add_section()
    s.orientation = WD_ORIENT.LANDSCAPE
    s.page_width = Inches(7.0)
    s.page_height = Inches(8.5)
    s.left_margin = Inches(0.5)
    s.right_margin = Inches(0.5)
    s.top_margin = Inches(0.5)
    s.bottom_margin = Inches(0.5)
    return s


def generate_bulletin(order: OrderOfWorship) -> Path:
    """Generate a Word bulletin and return the file path."""
    doc = Document()

    # Default font
    style = doc.styles['Normal']
    style.font.name = FONT
    style.font.size = Pt(12)

    # --- Page setup ---
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(7.0)
    section.page_height = Inches(8.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)

    service_date = date.fromisoformat(order.date)
    C = WD_ALIGN_PARAGRAPH.CENTER

    # ===== PAGE 1: Cover =====

    # Theme image
    theme_path = OUTPUT_DIR / order.themeImageFilename if order.themeImageFilename else None
    if theme_path and theme_path.exists():
        p = doc.add_paragraph()
        p.alignment = C
        p.paragraph_format.space_after = Pt(12)
        r = p.add_run()
        r.add_picture(str(theme_path), width=Inches(4.8))

    # Date + time — size 20
    p = doc.add_paragraph()
    p.alignment = C
    p.paragraph_format.space_after = Pt(4)
    _add_date_superscript(p, service_date, size=20,
                           suffix_text=f"\u2022 {cfg.WORSHIP_TIME}")

    # Church name — size 20
    _p(doc, cfg.CHURCH_NAME, size=20, bold=True, align=C, after=0)

    # Website — size 20, dark olive
    p = _p(doc, "", align=C, after=8)
    r = p.add_run(cfg.CHURCH_WEBSITE)
    r.font.name = FONT
    r.font.size = Pt(20)
    r.font.bold = True
    r.font.color.rgb = DARK_OLIVE

    # Staff — size 12, name regular, title bold italic
    for staff in [order.speaker, order.worshipLeader]:
        p = doc.add_paragraph()
        p.alignment = C
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(f"{staff.name}, ")
        r.font.name = FONT
        r.font.size = Pt(12)
        r = p.add_run(staff.title)
        r.font.name = FONT
        r.font.size = Pt(12)
        r.font.bold = True
        r.font.italic = True

    _p(doc, "", after=4)  # spacer

    # Schedule, address, phone — size 12
    _p(doc, cfg.SCHEDULE_LINE, size=12, align=C, after=0)

    # Address with superscript "st"
    p = doc.add_paragraph()
    p.alignment = C
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("5505 S. 31")
    r.font.name = FONT
    r.font.size = Pt(12)
    r = p.add_run("st")
    r.font.name = FONT
    r.font.size = Pt(12)
    r.font.superscript = True
    r = p.add_run(" Street, Temple, TX, 76502")
    r.font.name = FONT
    r.font.size = Pt(12)

    _p(doc, cfg.PHONE, size=12, align=C, after=8)

    # Mission / Vision / Core Value — size 10, bold labels, NOT italic values
    for line in cfg.MISSION_STATEMENT.split('\n'):
        if ':' in line:
            label, value = line.split(':', 1)
            p = doc.add_paragraph()
            p.alignment = C
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(f"{label}: ")
            r.font.name = FONT
            r.font.size = Pt(10)
            r.font.bold = True
            r = p.add_run(value.strip())
            r.font.name = FONT
            r.font.size = Pt(10)

    # ===== PAGE 2: Worship Order (first half) =====
    _new_section(doc)

    # Date + season header — size 18
    season = f" \u2022 {order.serviceTitle}" if order.serviceTitle else ""
    p = doc.add_paragraph()
    p.alignment = C
    p.paragraph_format.space_after = Pt(8)
    _add_date_superscript(p, service_date, size=18, suffix_text=season)

    # ENTRANCE and GREETING — size 12
    _p(doc, "ENTRANCE and GREETING", size=12, bold=True, align=C, before=4, after=8)

    # Items — left aligned, size 12
    _p(doc, "PRELUDE", size=12, bold=True, before=2, after=2)
    _p(doc, "INFORMAL MOMENTS", size=12, bold=True, before=2, after=0)

    # Indented sub-items — size 12
    p = _p(doc, "Welcome, Announcements, & Registration of Attendance",
            size=12, before=0, after=0)
    p.paragraph_format.left_indent = Inches(0.8)
    p = _p(doc, "Birthdays, Anniversaries, & Celebrations",
            size=12, before=0, after=0)
    p.paragraph_format.left_indent = Inches(0.8)

    _p(doc, "PETITIONS FOR PRAYER", size=12, bold=True, before=4, after=8)

    # THANKSGIVING and PRAYER — section heading, size 14
    _p(doc, "THANKSGIVING and PRAYER", size=14, bold=True, align=C, before=8, after=8)

    # Praise Songs — size 12
    _p(doc, "* PRAISE SONGS & OFFERTORY", size=12, bold=True, before=4, after=6)

    praise1 = order.praiseHymn1.title if order.praiseHymn1 else "TBD"
    praise2 = order.praiseHymn2.title if order.praiseHymn2 else "TBD"
    _p(doc, praise1, size=12, italic=True, align=C, after=0)
    _p(doc, praise2, size=12, italic=True, align=C, after=6)

    # Offertory note — size 12, italic
    if order.praiseHymn2:
        _p(doc, f"\u201c{praise2}\u201d is our offertory this morning.",
           size=12, italic=True, align=C, after=0)
        _p(doc, "Please come to the altar at any time during the song",
           size=12, italic=True, align=C, after=0)
        _p(doc, "and place your offering into one of the plates.",
           size=12, italic=True, align=C, after=8)

    # Doxology
    dox_num = f"UMH #{order.doxology.number}" if order.doxology else ""
    _worship_line(doc, "* DOXOLOGY", dox_num)

    # Creed
    creed_num = f"UMH #{order.creed.number}" if order.creed else ""
    creed_title = order.creed.title if order.creed else ""
    _worship_line(doc, "TODAY\u2019S CREED", creed_num)
    if creed_title:
        _p(doc, creed_title, size=12, bold=True, italic=True, align=C, after=4)

    # Prayer Hymn
    prayer_num = f"UMH #{order.prayerHymn.number}" if order.prayerHymn else ""
    prayer_title = order.prayerHymn.title if order.prayerHymn else ""
    _worship_line(doc, "HYMN OF PRAYER", prayer_num)
    if prayer_title:
        _p(doc, prayer_title, size=12, italic=True, align=C, after=4)

    # Pastoral Prayer & Lord's Prayer
    _p(doc, "PASTORAL PRAYER", size=12, bold=True, before=4, after=0)
    if order.liturgicalPrayer:
        prayer_label = order.liturgicalPrayer.title
        # Extract the variant label for the bulletin
        if "895" in order.liturgicalPrayer.number:
            prayer_variant = "Trespasses"
        elif "896" in order.liturgicalPrayer.number:
            prayer_variant = "E.U.B."
        elif "894" in order.liturgicalPrayer.number:
            prayer_variant = "Ecumenical"
        else:
            prayer_variant = ""
        variant_text = f" (\u201c{prayer_variant}\u201d)" if prayer_variant else ""
        _p(doc, f"THE LORD\u2019S PRAYER{variant_text}",
           size=12, bold=True, after=4)
    else:
        _p(doc, "THE LORD\u2019S PRAYER",
           size=12, bold=True, after=4)

    # ===== PAGE 3: Worship Order (second half) =====
    _new_section(doc)

    # PROCLAMATION and RESPONSE — section heading, size 14
    _p(doc, "PROCLAMATION and RESPONSE", size=14, bold=True, align=C,
       before=4, after=8)

    # Scripture — size 12
    _worship_line(doc, "Scripture Lesson", order.scripture)

    # The Message
    speaker_short = order.speakerShortName or cfg.DEFAULT_SPEAKER_SHORT
    _worship_line(doc, "The Message", speaker_short)

    # Sermon title — size 12
    if order.sermonTitle:
        p = doc.add_paragraph()
        p.alignment = C
        p.paragraph_format.space_after = Pt(8)
        r = p.add_run(f"\u201c{order.sermonTitle}\u201d ")
        r.font.name = FONT
        r.font.size = Pt(12)
        r.font.bold = True
        if order.sermonSubtitle:
            r = p.add_run(f"({order.sermonSubtitle})")
            r.font.name = FONT
            r.font.size = Pt(12)
            r.font.bold = True
            r.font.italic = True

    # INVITATION and DISMISSAL — section heading, size 14
    _p(doc, "INVITATION and DISMISSAL", size=14, bold=True, align=C,
       before=8, after=6)

    _p(doc, "INVITATION TO CHRISTIAN DISCIPLESHIP", size=12, bold=True,
       before=2, after=2)

    # "Anyone who desires..." — size 10, bold italic, indented
    p = _p(doc, "", after=0)
    p.paragraph_format.left_indent = Inches(0.5)
    p.alignment = C
    r = p.add_run("Anyone who desires a relationship with Christ and Christ\u2019s church")
    r.font.name = FONT
    r.font.size = Pt(10)
    r.font.bold = True
    r.font.italic = True

    p = _p(doc, "", after=6)
    p.paragraph_format.left_indent = Inches(0.5)
    p.alignment = C
    r = p.add_run("is invited to come forward as we sing our closing hymn together.")
    r.font.name = FONT
    r.font.size = Pt(10)
    r.font.bold = True
    r.font.italic = True

    # Closing Hymn
    closing_num = f"UMH #{order.closingHymn.number}" if order.closingHymn else ""
    closing_title = order.closingHymn.title if order.closingHymn else ""
    _worship_line(doc, "* CLOSING HYMN", closing_num)
    if closing_title:
        _p(doc, closing_title, size=12, italic=True, align=C, after=6)

    # Benediction & Postlude
    _p(doc, "* BENEDICTION", size=12, bold=True, before=4, after=0)
    _p(doc, "* POSTLUDE", size=12, bold=True, after=6)

    _p(doc, "* Please stand, as you are able.", size=12, italic=True, after=12)

    # E-giving — text on left (size 14), QR on right, side by side
    if QR_CODE_PATH.exists():
        table = doc.add_table(rows=1, cols=2)
        table.autofit = True
        tbl = table._tbl
        tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            '  <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            '  <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            '  <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            '  <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            '  <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            '  <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            '</w:tblBorders>'
        )
        tblPr.append(borders)

        # Left cell: text — size 14
        left_cell = table.cell(0, 0)
        left_cell.width = Inches(3.5)
        p = left_cell.paragraphs[0]
        p.alignment = C
        r = p.add_run("You can scan this code with your phone\u2019s\ncamera app to donate to Oak Park UMC.")
        r.font.name = FONT
        r.font.size = Pt(14)
        r.font.bold = True

        p2 = left_cell.add_paragraph()
        p2.alignment = C
        p2.paragraph_format.space_before = Pt(6)
        r = p2.add_run("We couldn\u2019t do it without you!")
        r.font.name = FONT
        r.font.size = Pt(14)
        r.font.bold = True

        # Right cell: QR code
        right_cell = table.cell(0, 1)
        right_cell.width = Inches(2.0)
        p = right_cell.paragraphs[0]
        p.alignment = C
        r = p.add_run()
        r.add_picture(str(QR_CODE_PATH), width=Inches(1.5))

    # ===== PAGE 4: Calendar =====
    _new_section(doc)

    # UPCOMING CALENDAR EVENTS — size 14
    _p(doc, "UPCOMING CALENDAR EVENTS", size=14, bold=True, align=C, after=12)

    # Generate 4 Sundays
    for week in range(4):
        sunday = service_date + timedelta(weeks=week)

        # Date header — size 12, bold italic, superscript ordinal
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10 if week > 0 else 4)
        p.paragraph_format.space_after = Pt(4)
        _add_date_short_superscript(
            p, sunday,
            prefix=f"{sunday.strftime('%A').upper()}, ",
            size=12, bold=True, italic=True
        )

        # Events with ❖ bullet, tab-aligned
        events = [
            ("9:30 AM", "Sunday School (Library)"),
            ("9:30 AM", "CSL Studies Class (Education Wing)"),
            ("10:45 AM", "Worship"),
        ]
        for time_str, event_name in events:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.left_indent = Inches(0.3)

            tab_stops = p.paragraph_format.tab_stops
            tab_stops.add_tab_stop(Inches(0.6))
            tab_stops.add_tab_stop(Inches(2.2))

            r = p.add_run(f"{BULLET}  ")
            r.font.name = FONT
            r.font.size = Pt(12)

            r = p.add_run(f"{time_str}\t{event_name}")
            r.font.name = FONT
            r.font.size = Pt(12)

    # Save
    OUTPUT_DIR.mkdir(exist_ok=True)
    filename = f"{order.date} - Bulletin.docx"
    filepath = OUTPUT_DIR / filename
    doc.save(str(filepath))
    return filepath
