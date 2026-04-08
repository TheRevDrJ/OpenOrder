"""
Generate a Word bulletin (.docx) from an OrderOfWorship by replacing
placeholders in a template document.

The template IS the theme — each church has their own .docx template
with their formatting, spacing, fonts, and layout. This code just
swaps the placeholders with real content and replaces the theme image.
"""

import copy
import io
import re
from datetime import date
from pathlib import Path

from docx import Document
from docx.shared import Inches

from .models import OrderOfWorship
from .paths import RESOURCES_DIR, OUTPUT_DIR

# Template path — this is the "theme"
TEMPLATE_PATH = RESOURCES_DIR / "bulletin" / "Template - Bulletin.docx"


def _ordinal_date(d: date) -> str:
    """Format a date like 'April 5th, 2026' with ordinal suffix."""
    day = d.day
    if 11 <= day <= 13:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    return f"{d.strftime('%B')} {day}{suffix}, {d.year}"


def _liturgical_prayer_label(order: OrderOfWorship) -> str:
    """Build the liturgical prayer label from the selected prayer."""
    if not order.liturgicalPrayer:
        return 'THE LORD\u2019S PRAYER'

    num = order.liturgicalPrayer.number
    title = order.liturgicalPrayer.title

    # Shorten the common Lord's Prayer variants
    if num in ('894', '895', '896'):
        return "THE LORD\u2019S PRAYER"

    return title.upper() if title else 'THE LORD\u2019S PRAYER'


def _replace_in_runs(paragraph, placeholder: str, replacement: str):
    """
    Replace a placeholder across paragraph runs, preserving formatting.

    Word splits text across multiple runs unpredictably. This finds the
    placeholder across run boundaries and replaces it while keeping the
    formatting of the first run that contains part of the placeholder.
    """
    # First, try simple per-run replacement
    for run in paragraph.runs:
        if placeholder in run.text:
            run.text = run.text.replace(placeholder, replacement)
            return True

    # If not found in a single run, search across runs
    full_text = ''.join(run.text for run in paragraph.runs)
    if placeholder not in full_text:
        return False

    # Find where the placeholder starts and ends across runs
    idx = full_text.index(placeholder)
    end_idx = idx + len(placeholder)

    # Rebuild runs: before placeholder, replacement, after placeholder
    char_pos = 0
    new_runs_text = []
    replacement_inserted = False

    for run in paragraph.runs:
        run_start = char_pos
        run_end = char_pos + len(run.text)

        if run_end <= idx:
            # Entirely before placeholder — keep as is
            new_runs_text.append(None)
        elif run_start >= end_idx:
            # Entirely after placeholder — keep as is
            new_runs_text.append(None)
        else:
            # This run overlaps with the placeholder
            before = run.text[:max(0, idx - run_start)]
            after = run.text[max(0, end_idx - run_start):]

            if not replacement_inserted:
                new_runs_text.append(before + replacement + after)
                replacement_inserted = True
            else:
                # Subsequent runs that were part of the placeholder
                new_runs_text.append(after if after else '')

        char_pos = run_end

    # Apply the new text to runs
    for run, new_text in zip(paragraph.runs, new_runs_text):
        if new_text is not None:
            run.text = new_text

    return True


def _replace_theme_image(doc: Document, image_path: Path):
    """Replace the first inline image in the document with the theme image."""
    if not image_path or not image_path.exists():
        return

    # Convert unsupported formats to PNG
    supported = {'.bmp', '.gif', '.jpg', '.jpeg', '.png', '.tiff', '.tif', '.wmf'}
    if image_path.suffix.lower() not in supported:
        try:
            from PIL import Image
            png_path = image_path.with_suffix('.png')
            if not png_path.exists():
                img = Image.open(image_path)
                img.save(png_path, 'PNG')
            image_path = png_path
        except Exception:
            return

    # Find the first inline shape (the theme image placeholder)
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if run._element.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip'):
                # Found an inline image — get its relationship ID
                blips = run._element.findall(
                    './/{http://schemas.openxmlformats.org/drawingml/2006/main}blip'
                )
                if blips:
                    blip = blips[0]
                    ns = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
                    rId = blip.get(f'{ns}embed')
                    if rId:
                        # Replace the image data in the relationship
                        part = doc.part
                        image_part = part.related_parts[rId]
                        with open(image_path, 'rb') as f:
                            image_part._blob = f.read()
                        return


def generate_bulletin(order: OrderOfWorship) -> Path:
    """Generate a Word bulletin by replacing placeholders in the template."""

    doc = Document(str(TEMPLATE_PATH))
    service_date = date.fromisoformat(order.date)

    # Build the replacement map
    replacements = {
        '{{DATE}}': _ordinal_date(service_date),
        '{{SERVICE_TITLE}}': order.serviceTitle or '',
        '{{OPENING_HYMN_TITLE}}': order.praiseHymn1.title if order.praiseHymn1 else 'TBD',
        '{{OFFERTORY_HYMN_TITLE}}': order.praiseHymn2.title if order.praiseHymn2 else 'TBD',
        '{{OFFERTORY_HYMN}}': order.praiseHymn2.title if order.praiseHymn2 else 'TBD',
        '{{DOX}}': order.doxology.number if order.doxology else '95',
        '{{CREED}}': order.creed.number if order.creed else '881',
        '{{CREED_TITLE}}': order.creed.title if order.creed else '',
        '{{PRAYER_HYMN_NUMBER}}': order.prayerHymn.number if order.prayerHymn else '',
        '{{PRAYER_HYMN_TITLE}}': order.prayerHymn.title if order.prayerHymn else '',
        '{{LITURGICAL_PRAYER}}': _liturgical_prayer_label(order),
        '{{SCRIPTURE}}': order.scripture or '',
        '{{SPEAKER}}': order.speakerShortName or '',
        '{{SERMON_TITLE}}': order.sermonTitle or '',
        '{{SERMON_SUBTITLE}}': order.sermonSubtitle or '',
        '{{CLOSING_HYMN_NUMBER}}': order.closingHymn.number if order.closingHymn else '',
        '{{CLOSING_HYMN_TITLE}}': order.closingHymn.title if order.closingHymn else '',
    }

    # Replace placeholders in all paragraphs
    for paragraph in doc.paragraphs:
        for placeholder, value in replacements.items():
            if placeholder in paragraph.text:
                _replace_in_runs(paragraph, placeholder, value)

    # Also check tables (QR code section might be in one)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for placeholder, value in replacements.items():
                        if placeholder in paragraph.text:
                            _replace_in_runs(paragraph, placeholder, value)

    # Replace theme image
    theme_path = OUTPUT_DIR / order.themeImageFilename if order.themeImageFilename else None
    _replace_theme_image(doc, theme_path)

    # Save
    OUTPUT_DIR.mkdir(exist_ok=True)
    filename = f"{order.date} - Bulletin.docx"
    filepath = OUTPUT_DIR / filename
    doc.save(str(filepath))
    return filepath
