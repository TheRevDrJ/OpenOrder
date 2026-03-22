"""Generate a PowerPoint presentation (.pptx) from an OrderOfWorship."""

import json
import re
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor

from .models import OrderOfWorship, HymnRef

RESOURCES = Path(__file__).parent.parent.parent / "resources"
HYMNAL_DIR = Path(__file__).parent.parent.parent / "hymnal-json"
OUTPUT_DIR = Path(__file__).parent.parent.parent / "output"

# Slide dimensions: 13.333" x 7.500" (widescreen 16:9)
SLIDE_WIDTH = Emu(12_192_000)
SLIDE_HEIGHT = Emu(6_858_000)

# Backgrounds
HYMN_BG = RESOURCES / "slides" / "Background.png"
CREED_BG = RESOURCES / "slides" / "CreedBackground.png"

# Static slide images
ANNOUNCEMENTS_SLIDE = RESOURCES / "slides" / "AnnouncementsSlide.jpg"
CONCERNS_SLIDE = RESOURCES / "slides" / "ConcernsSlide.jpg"
OFFERING_SLIDE = RESOURCES / "slides" / "OfferingSlide.jpg"
PRAYER_SLIDE = RESOURCES / "slides" / "PrayerSlide.jpg"

# Colors
TITLE_COLOR = RGBColor(0x33, 0x33, 0x99)       # #333399 dark blue
NUMBER_BG_COLOR = RGBColor(0xB2, 0xB2, 0xB2)   # gray fill
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

# Fonts
FONT_NAME = "Times New Roman"

# Layout constants (from analysis of existing presentation)
TEXT_LEFT = Emu(1_524_000)           # 1.667 inches
TEXT_WIDTH = Emu(9_144_000)          # 10.000 inches
TITLE_TOP = Emu(304_800)            # 0.333 inches
NUM_BOX_LEFT = Emu(9_144_000)       # 10.000 inches
NUM_BOX_WIDTH = Emu(1_524_000)      # 1.667 inches
NUM_BOX_HEIGHT = Emu(762_000)       # 0.833 inches
BG_IMG_LEFT = Emu(2_667_000)        # 2.917 inches
BG_IMG_SIZE = Emu(6_858_000)        # 7.500 inches (square)

# Text margins
MARGIN_LR = Emu(91_440)
MARGIN_TB = Emu(45_720)


def _add_shadow(shape):
    """Add outer shadow effect to a shape for text readability."""
    from lxml import etree
    spPr = shape._element.spPr
    if spPr is None:
        from pptx.oxml.ns import qn
        spPr = etree.SubElement(shape._element, qn('a:spPr'))

    # Build effectLst with outer shadow
    nsmap = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    effectLst = etree.SubElement(spPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}effectLst')
    outerShdw = etree.SubElement(effectLst, '{http://schemas.openxmlformats.org/drawingml/2006/main}outerShdw')
    outerShdw.set('dist', '35921')
    outerShdw.set('dir', '2700000')
    outerShdw.set('algn', 'ctr')
    outerShdw.set('rotWithShape', '0')
    schemeClr = etree.SubElement(outerShdw, '{http://schemas.openxmlformats.org/drawingml/2006/main}schemeClr')
    schemeClr.set('val', 'bg1')


def _set_paragraph_spacing(paragraph, space_before_pct=50):
    """Set paragraph spacing before as percentage of font size."""
    from lxml import etree
    pPr = paragraph._p.get_or_add_pPr()
    spcBef = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}spcBef')
    spcPct = etree.SubElement(spcBef, '{http://schemas.openxmlformats.org/drawingml/2006/main}spcPct')
    spcPct.set('val', str(space_before_pct * 1000))


def _add_full_image_slide(prs, image_path):
    """Add a slide with a single full-bleed image."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    if image_path.exists():
        slide.shapes.add_picture(
            str(image_path), Emu(0), Emu(0),
            SLIDE_WIDTH, SLIDE_HEIGHT
        )
    return slide


def _add_hymn_background(slide):
    """Add the hymn background image (7.5x7.5 square, right-aligned)."""
    if HYMN_BG.exists():
        slide.shapes.add_picture(
            str(HYMN_BG), BG_IMG_LEFT, Emu(0),
            BG_IMG_SIZE, BG_IMG_SIZE
        )


def _add_creed_background(slide):
    """Add the creed/liturgy background image (full-bleed)."""
    if CREED_BG.exists():
        slide.shapes.add_picture(
            str(CREED_BG), Emu(0), Emu(0),
            SLIDE_WIDTH, SLIDE_HEIGHT
        )


def _load_hymn_data(ref: HymnRef) -> dict | None:
    """Load hymn JSON data from the hymnal-json directory."""
    source_dir = HYMNAL_DIR / ref.source
    # Try to find the file by number prefix
    for f in source_dir.glob("*.json"):
        data = json.load(open(f, encoding='utf-8'))
        if data.get('number') == ref.number:
            return data
    return None


def _is_title_line(line: str, hymn_title: str) -> bool:
    """Check if a line is a title/header (not lyrics)."""
    line_lower = line.lower().strip()
    title_lower = hymn_title.lower().strip()
    # Exact title match
    if line_lower == title_lower:
        return True
    # (Verse N) or (Refrain) or (N)
    if re.match(r'^\((?:verse\s+)?\d+\)$', line_lower):
        return True
    if line_lower == '(refrain)':
        return True
    return False


def _is_number_line(line: str, hymn_number: str) -> bool:
    """Check if a line is just the hymn number."""
    return line.strip() == hymn_number.strip()


def _is_attribution_line(line: str) -> bool:
    """Check if a line is a WORDS:/copyright attribution."""
    line_stripped = line.strip()
    if line_stripped.startswith('WORDS:'):
        return True
    if '©' in line_stripped or '\u00a9' in line_stripped:
        return True
    return False


def _is_refrain_label(line: str) -> bool:
    """Check if a line is just 'Refrain'."""
    return line.strip().lower() == 'refrain'


def _parse_hymn_slides(hymn_data: dict) -> list[dict]:
    """
    Parse hymn JSON into structured slide data.
    Returns list of dicts with keys:
      - type: 'first' | 'continuation' | 'refrain'
      - title: str (for first slides)
      - number: str (for first slides)
      - attribution: str (for first slides)
      - lyrics: list[str]
      - refrain: bool
      - verse_label: str (e.g. "(Verse 1)")
    """
    slides = hymn_data['slides']
    title = hymn_data['title']
    number = hymn_data['number']
    result = []

    for si, slide in enumerate(slides):
        lines = slide['lines']

        # Separate metadata from lyrics
        lyrics = []
        attribution = ""
        verse_label = ""
        is_refrain = False
        is_first = (si == 0)

        for line in lines:
            if _is_title_line(line, title):
                continue
            if _is_number_line(line, number):
                continue
            if _is_attribution_line(line):
                attribution = line.strip()
                continue
            if re.match(r'^\((?:verse\s+)?\d+\)$', line.strip(), re.IGNORECASE):
                verse_label = line.strip()
                continue
            if line.strip().lower() == '(refrain)':
                verse_label = "(Refrain)"
                continue
            if _is_refrain_label(line):
                is_refrain = True
                continue
            # Check for "FROM THE RITUAL..." type lines in services
            if line.strip().startswith('FROM THE'):
                attribution = line.strip()
                continue
            # Copyright line
            if line.strip().startswith('\u00a9') or line.strip().startswith('©'):
                attribution = line.strip()
                continue
            lyrics.append(line)

        # Also check for verse label at the start like "1. Amazing grace!"
        # We keep these as part of lyrics since they contain actual lyric text

        slide_info = {
            'type': 'first' if is_first else ('refrain' if is_refrain else 'continuation'),
            'title': title if is_first else '',
            'number': number if is_first else '',
            'attribution': attribution,
            'lyrics': lyrics,
            'refrain': is_refrain,
            'verse_label': verse_label,
        }
        result.append(slide_info)

    return result


def _create_hymn_first_slide(prs, slide_info: dict):
    """Create a hymn's first slide with title, number box, attribution, lyrics, and background."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank

    # 1. Title text box
    title_box = slide.shapes.add_textbox(
        TEXT_LEFT, TITLE_TOP, TEXT_WIDTH, Emu(500_000)
    )
    title_box.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    title_box.text_frame.margin_left = MARGIN_LR
    title_box.text_frame.margin_right = MARGIN_LR
    title_box.text_frame.margin_top = MARGIN_TB
    title_box.text_frame.margin_bottom = MARGIN_TB

    p = title_box.text_frame.paragraphs[0]
    p.text = slide_info['title']
    p.font.name = FONT_NAME
    p.font.size = Pt(40)
    p.font.color.rgb = TITLE_COLOR
    _set_paragraph_spacing(p, 50)
    _add_shadow(title_box)

    # 2. Hymn number box
    num_box = slide.shapes.add_textbox(
        NUM_BOX_LEFT, TITLE_TOP, NUM_BOX_WIDTH, NUM_BOX_HEIGHT
    )
    num_box.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    # Set gray background fill
    num_box.fill.solid()
    num_box.fill.fore_color.rgb = NUMBER_BG_COLOR

    p = num_box.text_frame.paragraphs[0]
    p.text = slide_info['number']
    p.font.name = FONT_NAME
    p.font.size = Pt(44)
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER
    _set_paragraph_spacing(p, 50)
    _add_shadow(num_box)

    # 3. Attribution line (if present)
    attr_top = Emu(1_200_000)  # ~1.3 inches
    if slide_info['attribution']:
        attr_box = slide.shapes.add_textbox(
            TEXT_LEFT, attr_top, TEXT_WIDTH, Emu(400_000)
        )
        attr_box.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        p = attr_box.text_frame.paragraphs[0]
        p.text = slide_info['attribution']
        p.font.name = FONT_NAME
        p.font.size = Pt(22)
        p.font.color.rgb = TITLE_COLOR
        _set_paragraph_spacing(p, 50)
        _add_shadow(attr_box)
        lyrics_top = Emu(1_700_000)  # below attribution
    else:
        lyrics_top = Emu(1_400_000)

    # 4. Lyrics text box
    if slide_info['lyrics']:
        lyrics_box = slide.shapes.add_textbox(
            TEXT_LEFT, lyrics_top, TEXT_WIDTH, Emu(500_000)
        )
        lyrics_box.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        lyrics_box.text_frame.margin_left = MARGIN_LR
        lyrics_box.text_frame.margin_right = MARGIN_LR
        lyrics_box.text_frame.margin_top = MARGIN_TB
        lyrics_box.text_frame.margin_bottom = MARGIN_TB

        # Use soft returns (\x0b) to join lines within one paragraph
        p = lyrics_box.text_frame.paragraphs[0]
        lyrics_text = '\x0b'.join(slide_info['lyrics'])
        p.text = lyrics_text
        p.font.name = FONT_NAME
        p.font.size = Pt(50)
        _set_paragraph_spacing(p, 50)
        _add_shadow(lyrics_box)

    # 5. Background image (added last so it's on top in z-order,
    #    matching original presentation structure)
    _add_hymn_background(slide)

    return slide


def _create_hymn_continuation_slide(prs, slide_info: dict):
    """Create a hymn continuation slide (lyrics only + background)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank

    if slide_info['lyrics']:
        # Calculate top position based on content amount
        num_lines = len(slide_info['lyrics'])
        if num_lines <= 3:
            top = Emu(1_800_000)  # center-ish for short content
        elif num_lines <= 5:
            top = Emu(1_200_000)
        else:
            top = Emu(800_000)

        lyrics_box = slide.shapes.add_textbox(
            TEXT_LEFT, top, TEXT_WIDTH, Emu(500_000)
        )
        lyrics_box.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        lyrics_box.text_frame.margin_left = MARGIN_LR
        lyrics_box.text_frame.margin_right = MARGIN_LR
        lyrics_box.text_frame.margin_top = MARGIN_TB
        lyrics_box.text_frame.margin_bottom = MARGIN_TB

        p = lyrics_box.text_frame.paragraphs[0]

        if slide_info['refrain']:
            # Add "Refrain" label as first paragraph, italic
            p.text = "Refrain"
            p.font.name = FONT_NAME
            p.font.size = Pt(40)
            p.font.italic = True
            _set_paragraph_spacing(p, 50)

            # Lyrics in second paragraph
            p2 = lyrics_box.text_frame.add_paragraph()
            lyrics_text = '\x0b'.join(slide_info['lyrics'])
            p2.text = lyrics_text
            p2.font.name = FONT_NAME
            p2.font.size = Pt(50)
            _set_paragraph_spacing(p2, 50)
        else:
            lyrics_text = '\x0b'.join(slide_info['lyrics'])
            p.text = lyrics_text
            p.font.name = FONT_NAME
            p.font.size = Pt(50)
            _set_paragraph_spacing(p, 50)

        _add_shadow(lyrics_box)

    _add_hymn_background(slide)
    return slide


def _create_liturgy_first_slide(prs, slide_info: dict, bg_path: Path):
    """Create a liturgy first slide (creed/Lord's Prayer) with title, number, and text."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Title
    title_box = slide.shapes.add_textbox(
        TEXT_LEFT, TITLE_TOP, TEXT_WIDTH, Emu(500_000)
    )
    title_box.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    p = title_box.text_frame.paragraphs[0]
    p.text = slide_info['title'].upper()
    p.font.name = FONT_NAME
    p.font.size = Pt(32)
    p.font.color.rgb = TITLE_COLOR
    _set_paragraph_spacing(p, 50)
    _add_shadow(title_box)

    # Number
    num_box = slide.shapes.add_textbox(
        NUM_BOX_LEFT, TITLE_TOP, NUM_BOX_WIDTH, NUM_BOX_HEIGHT
    )
    num_box.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    num_box.fill.solid()
    num_box.fill.fore_color.rgb = NUMBER_BG_COLOR
    p = num_box.text_frame.paragraphs[0]
    p.text = slide_info['number']
    p.font.name = FONT_NAME
    p.font.size = Pt(40)
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER
    _add_shadow(num_box)

    # Lyrics/text
    if slide_info['lyrics']:
        lyrics_box = slide.shapes.add_textbox(
            Emu(1_828_800), Emu(1_400_000),  # ~2.0 inches left
            Emu(8_534_400), Emu(500_000)      # ~9.333 inches wide
        )
        lyrics_box.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        lyrics_box.text_frame.margin_left = MARGIN_LR
        lyrics_box.text_frame.margin_right = MARGIN_LR

        p = lyrics_box.text_frame.paragraphs[0]
        # For liturgy, check for speaker labels (Pastor:, People:, etc.)
        _build_liturgy_text(lyrics_box.text_frame, slide_info['lyrics'])
        _add_shadow(lyrics_box)

    # Background (added last)
    if bg_path.exists():
        slide.shapes.add_picture(
            str(bg_path), Emu(0), Emu(0),
            SLIDE_WIDTH, SLIDE_HEIGHT
        )

    return slide


def _create_liturgy_continuation_slide(prs, slide_info: dict, bg_path: Path):
    """Create a liturgy continuation slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    if slide_info['lyrics']:
        num_lines = len(slide_info['lyrics'])
        top = Emu(800_000) if num_lines > 4 else Emu(1_400_000)

        lyrics_box = slide.shapes.add_textbox(
            Emu(1_828_800), top,
            Emu(8_534_400), Emu(500_000)
        )
        lyrics_box.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        lyrics_box.text_frame.margin_left = MARGIN_LR
        lyrics_box.text_frame.margin_right = MARGIN_LR

        _build_liturgy_text(lyrics_box.text_frame, slide_info['lyrics'])
        _add_shadow(lyrics_box)

    if bg_path.exists():
        slide.shapes.add_picture(
            str(bg_path), Emu(0), Emu(0),
            SLIDE_WIDTH, SLIDE_HEIGHT
        )

    return slide


def _build_liturgy_text(text_frame, lines: list[str]):
    """Build liturgy text with speaker labels in red italic."""
    # Join all lines with soft returns
    full_text = '\x0b'.join(lines)

    p = text_frame.paragraphs[0]

    # Check if any lines have speaker labels like "Pastor:", "People:", etc.
    speaker_pattern = re.compile(r'^(Pastor(?:\s+and\s+People)?|People|All|Leader|Pastor\s*:?)\s*:?\s*', re.IGNORECASE)

    # Simple approach: put all text, then format speaker labels
    # For now, use soft returns and consistent formatting
    has_speakers = any(speaker_pattern.match(line) for line in lines)

    if has_speakers:
        # Build with runs for speaker labels in red
        first = True
        for line in lines:
            if not first:
                # Add soft return
                from pptx.oxml.ns import qn
                from lxml import etree
                br = etree.SubElement(p._p, qn('a:br'))

            match = speaker_pattern.match(line)
            if match:
                # Speaker label in red italic
                speaker = match.group(0).strip()
                if not speaker.endswith(':'):
                    speaker += ':'
                rest = line[match.end():].strip()

                run = p.add_run()
                run.text = speaker + ' '
                run.font.name = FONT_NAME
                run.font.size = Pt(48)
                run.font.italic = True
                run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)

                if rest:
                    run2 = p.add_run()
                    run2.text = rest
                    run2.font.name = FONT_NAME
                    run2.font.size = Pt(48)
            else:
                run = p.add_run()
                run.text = line
                run.font.name = FONT_NAME
                run.font.size = Pt(48)

            first = False
    else:
        p.text = full_text
        p.font.name = FONT_NAME
        p.font.size = Pt(48)

    _set_paragraph_spacing(p, 50)


def _add_hymn_slides(prs, ref: HymnRef, bg_type: str = 'hymn'):
    """Add all slides for a hymn/creed/prayer."""
    hymn_data = _load_hymn_data(ref)
    if not hymn_data:
        return

    parsed = _parse_hymn_slides(hymn_data)
    bg_path = CREED_BG if bg_type == 'creed' else HYMN_BG

    for i, slide_info in enumerate(parsed):
        if bg_type == 'creed':
            if i == 0:
                _create_liturgy_first_slide(prs, slide_info, bg_path)
            else:
                _create_liturgy_continuation_slide(prs, slide_info, bg_path)
        else:
            if i == 0:
                _create_hymn_first_slide(prs, slide_info)
            else:
                _create_hymn_continuation_slide(prs, slide_info)


def _add_theme_slide(prs, theme_image_path: Path | None):
    """Add a theme/separator slide using the uploaded theme image."""
    if theme_image_path and theme_image_path.exists():
        _add_full_image_slide(prs, theme_image_path)
    else:
        # Blank slide as fallback
        prs.slides.add_slide(prs.slide_layouts[6])


def generate_slides(order: OrderOfWorship) -> Path:
    """Generate a PowerPoint presentation and return the file path."""
    prs = Presentation()

    # Set slide dimensions
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    # Theme image path
    theme_path = None
    if order.themeImageFilename:
        candidate = OUTPUT_DIR / order.themeImageFilename
        if candidate.exists():
            theme_path = candidate

    # ===== SLIDE ORDER (from VBA Main()) =====

    # 1. Theme image (full-bleed)
    _add_theme_slide(prs, theme_path)

    # 2. Announcements slide (static image)
    _add_full_image_slide(prs, ANNOUNCEMENTS_SLIDE)

    # 3. Concerns slide (static image)
    _add_full_image_slide(prs, CONCERNS_SLIDE)

    # 4. Theme slide (separator)
    _add_theme_slide(prs, theme_path)

    # 5. Praise Hymn 1 (Opening Hymn)
    if order.praiseHymn1:
        _add_hymn_slides(prs, order.praiseHymn1, bg_type='hymn')

    # 6. Offering slide (static image)
    _add_full_image_slide(prs, OFFERING_SLIDE)

    # 7. Praise Hymn 2 (Offertory Hymn)
    if order.praiseHymn2:
        _add_hymn_slides(prs, order.praiseHymn2, bg_type='hymn')

    # 8. Doxology
    if order.doxology:
        _add_hymn_slides(prs, order.doxology, bg_type='hymn')

    # 9. Theme slide (separator)
    _add_theme_slide(prs, theme_path)

    # 10. Creed (uses creed background)
    if order.creed:
        _add_hymn_slides(prs, order.creed, bg_type='creed')

    # 11. Theme slide (separator)
    _add_theme_slide(prs, theme_path)

    # 12. Prayer Hymn
    if order.prayerHymn:
        _add_hymn_slides(prs, order.prayerHymn, bg_type='hymn')

    # 13. Prayer slide (static image)
    _add_full_image_slide(prs, PRAYER_SLIDE)

    # 14. Lord's Prayer / Liturgical Prayer (creed background)
    if order.liturgicalPrayer:
        _add_hymn_slides(prs, order.liturgicalPrayer, bg_type='creed')

    # 15. Theme slide (separator)
    _add_theme_slide(prs, theme_path)

    # 16. Closing Hymn
    if order.closingHymn:
        _add_hymn_slides(prs, order.closingHymn, bg_type='hymn')

    # 17. Theme slide (separator, final)
    _add_theme_slide(prs, theme_path)

    # Save
    OUTPUT_DIR.mkdir(exist_ok=True)
    filename = f"{order.date} - Slides.pptx"
    filepath = OUTPUT_DIR / filename
    prs.save(str(filepath))
    return filepath
