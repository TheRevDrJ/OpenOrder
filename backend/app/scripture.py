"""Scripture fetcher — retrieves Bible text via API, caches locally, splits into slides."""

import json
import re
import textwrap
from pathlib import Path

import httpx
import pythonbible as bible

from .paths import SCRIPTURE_CACHE_DIR as CACHE_DIR
API_BASE = "https://bible.helloao.org/api"

# pythonbible Book enum -> USFM 3-letter code (used by AO Lab API)
BOOK_TO_USFM = {
    bible.Book.GENESIS: "GEN", bible.Book.EXODUS: "EXO", bible.Book.LEVITICUS: "LEV",
    bible.Book.NUMBERS: "NUM", bible.Book.DEUTERONOMY: "DEU", bible.Book.JOSHUA: "JOS",
    bible.Book.JUDGES: "JDG", bible.Book.RUTH: "RUT",
    bible.Book.SAMUEL_1: "1SA", bible.Book.SAMUEL_2: "2SA",
    bible.Book.KINGS_1: "1KI", bible.Book.KINGS_2: "2KI",
    bible.Book.CHRONICLES_1: "1CH", bible.Book.CHRONICLES_2: "2CH",
    bible.Book.EZRA: "EZR", bible.Book.NEHEMIAH: "NEH", bible.Book.ESTHER: "EST",
    bible.Book.JOB: "JOB", bible.Book.PSALMS: "PSA", bible.Book.PROVERBS: "PRO",
    bible.Book.ECCLESIASTES: "ECC", bible.Book.SONG_OF_SONGS: "SNG",
    bible.Book.ISAIAH: "ISA", bible.Book.JEREMIAH: "JER",
    bible.Book.LAMENTATIONS: "LAM", bible.Book.EZEKIEL: "EZK",
    bible.Book.DANIEL: "DAN", bible.Book.HOSEA: "HOS", bible.Book.JOEL: "JOL",
    bible.Book.AMOS: "AMO", bible.Book.OBADIAH: "OBA", bible.Book.JONAH: "JON",
    bible.Book.MICAH: "MIC", bible.Book.NAHUM: "NAM", bible.Book.HABAKKUK: "HAB",
    bible.Book.ZEPHANIAH: "ZEP", bible.Book.HAGGAI: "HAG",
    bible.Book.ZECHARIAH: "ZEC", bible.Book.MALACHI: "MAL",
    bible.Book.MATTHEW: "MAT", bible.Book.MARK: "MRK", bible.Book.LUKE: "LUK",
    bible.Book.JOHN: "JHN", bible.Book.ACTS: "ACT", bible.Book.ROMANS: "ROM",
    bible.Book.CORINTHIANS_1: "1CO", bible.Book.CORINTHIANS_2: "2CO",
    bible.Book.GALATIANS: "GAL", bible.Book.EPHESIANS: "EPH",
    bible.Book.PHILIPPIANS: "PHP", bible.Book.COLOSSIANS: "COL",
    bible.Book.THESSALONIANS_1: "1TH", bible.Book.THESSALONIANS_2: "2TH",
    bible.Book.TIMOTHY_1: "1TI", bible.Book.TIMOTHY_2: "2TI",
    bible.Book.TITUS: "TIT", bible.Book.PHILEMON: "PHM", bible.Book.HEBREWS: "HEB",
    bible.Book.JAMES: "JAS", bible.Book.PETER_1: "1PE", bible.Book.PETER_2: "2PE",
    bible.Book.JOHN_1: "1JN", bible.Book.JOHN_2: "2JN", bible.Book.JOHN_3: "3JN",
    bible.Book.JUDE: "JUD", bible.Book.REVELATION: "REV",
}

# Available translations (curated list for the UI)
# Format: (api_id, display_name, description)
TRANSLATIONS = [
    ("BSB", "BSB", "Berean Standard Bible (modern, free)"),
    ("eng_kjv", "KJV", "King James Version (1611)"),
    ("eng_asv", "ASV", "American Standard Version (1901)"),
    ("ENGWEBP", "WEB", "World English Bible (modern, public domain)"),
]

# Characters per line at 50pt Times New Roman on our slides (~28 chars)
CHARS_PER_LINE = 28
MAX_LINES_PER_SLIDE = 7


def parse_reference(ref_string: str) -> dict | None:
    """
    Parse a scripture reference string into structured data.
    Returns {book, book_name, chapter, start_verse, end_verse, usfm_book, verse_ids}
    or None if parsing fails.
    """
    try:
        refs = bible.get_references(ref_string)
        if not refs:
            return None

        verse_ids = bible.convert_references_to_verse_ids(refs)
        if not verse_ids:
            return None

        # Get book and chapter info from first verse
        first_ref = refs[0]
        book = first_ref.book
        book_name = book.title
        usfm = BOOK_TO_USFM.get(book)

        if not usfm:
            return None

        # Extract chapter and verse range from verse IDs
        # Verse ID format: BBCCCVVV (book=BB, chapter=CCC, verse=VVV)
        first_id = verse_ids[0]
        last_id = verse_ids[-1]

        # Extract chapters from verse IDs
        # Format: BBCCCVVV (book=BB, chapter=CCC, verse=VVV)
        chapters_needed = set()
        for vid in verse_ids:
            vid_str = str(vid)
            # Book is first 1-2 digits, chapter is next 3, verse is last 3
            ch = int(vid_str[-6:-3])
            chapters_needed.add(ch)

        start_chapter = min(chapters_needed)
        start_verse = int(str(first_id)[-3:])
        end_verse = int(str(last_id)[-3:])

        return {
            "reference": ref_string,
            "book": book,
            "book_name": book_name,
            "usfm_book": usfm,
            "chapters": sorted(chapters_needed),
            "start_chapter": start_chapter,
            "start_verse": start_verse,
            "end_verse": end_verse,
            "verse_ids": verse_ids,
        }
    except Exception:
        return None


def _cache_path(reference: str, translation: str) -> Path:
    """Get cache file path for a scripture reference."""
    # Sanitize the reference for use as a filename
    safe_ref = re.sub(r'[^\w\s\-]', '', reference).strip().replace(' ', '_')
    return CACHE_DIR / translation / f"{safe_ref}.json"


def _load_cached(reference: str, translation: str) -> dict | None:
    """Load cached scripture data if available."""
    path = _cache_path(reference, translation)
    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return None


def _save_cache(reference: str, translation: str, data: dict):
    """Save scripture data to cache."""
    path = _cache_path(reference, translation)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def _extract_verse_text(content_item: dict) -> str:
    """Extract plain text from an AO Lab API content item."""
    if content_item.get("type") != "verse":
        return ""

    parts = []
    for item in content_item.get("content", []):
        if isinstance(item, str):
            parts.append(item)
        elif isinstance(item, dict):
            # Nested content (e.g., poetry, quotes)
            for sub in item.get("content", []):
                if isinstance(sub, str):
                    parts.append(sub)
                elif isinstance(sub, dict):
                    for subsub in sub.get("content", []):
                        if isinstance(subsub, str):
                            parts.append(subsub)

    text = " ".join(parts).strip()
    # Clean up multiple spaces
    text = re.sub(r'\s+', ' ', text)
    return text


def fetch_scripture(reference: str, translation: str = "BSB") -> dict | None:
    """
    Fetch scripture text for a reference and translation.
    Returns {reference, translation, translation_name, verses: [{number, text}], slides: [...]}
    Uses cache if available.
    """
    # Check cache first
    cached = _load_cached(reference, translation)
    if cached:
        return cached

    # Parse the reference
    parsed = parse_reference(reference)
    if not parsed:
        return None

    # Fetch from API — one chapter at a time
    all_verses = []
    for chapter in parsed["chapters"]:
        url = f"{API_BASE}/{translation}/{parsed['usfm_book']}/{chapter}.json"
        try:
            r = httpx.get(url, timeout=15)
            if r.status_code != 200:
                return None
            data = r.json()
        except Exception:
            return None

        # Extract verses from chapter content
        chapter_content = data.get("chapter", {}).get("content", [])
        for item in chapter_content:
            if item.get("type") == "verse":
                verse_num = item.get("number")
                verse_text = _extract_verse_text(item)
                if verse_text:
                    # Build the full verse ID to check if it's in our range
                    # We need to check if this verse is in our requested range
                    all_verses.append({
                        "chapter": chapter,
                        "number": verse_num,
                        "text": verse_text,
                    })

    # Filter to only the requested verse range
    filtered_verses = []
    for v in all_verses:
        # Build verse ID to match against parsed verse_ids
        book_num = parsed["book"].value
        vid = int(f"{book_num}{v['chapter']:03d}{v['number']:03d}")
        if vid in parsed["verse_ids"]:
            filtered_verses.append(v)

    if not filtered_verses:
        return None

    # Get translation display name
    trans_name = translation
    for tid, display, desc in TRANSLATIONS:
        if tid == translation:
            trans_name = display
            break

    # Split into slides
    slides = split_into_slides(filtered_verses)

    result = {
        "reference": reference,
        "translation": translation,
        "translation_name": trans_name,
        "verses": filtered_verses,
        "slides": slides,
    }

    # Cache it
    _save_cache(reference, translation, result)

    return result


def _estimate_lines(text: str) -> int:
    """Estimate how many slide lines a piece of text will take."""
    # At 50pt Times New Roman, roughly 28 chars per line
    wrapped = textwrap.wrap(text, width=CHARS_PER_LINE)
    return max(len(wrapped), 1)


def split_into_slides(verses: list[dict], max_lines: int = MAX_LINES_PER_SLIDE) -> list[dict]:
    """
    Split verses into slide-sized chunks.
    Each slide is {lines: [str], verse_range: str, line_count: int}
    """
    slides = []
    current_lines = []
    current_verse_start = None
    current_verse_end = None
    current_line_count = 0

    for v in verses:
        verse_num = v["number"]
        verse_text = v["text"]

        # Format: "1 Then Jesus was led..."
        formatted = f"{verse_num} {verse_text}"

        # Estimate how many lines this verse takes
        est_lines = _estimate_lines(formatted)

        # If adding this verse would overflow, start a new slide
        if current_line_count + est_lines > max_lines and current_lines:
            slides.append({
                "lines": current_lines,
                "verse_range": f"{current_verse_start}-{current_verse_end}"
                    if current_verse_start != current_verse_end
                    else str(current_verse_start),
                "line_count": current_line_count,
            })
            current_lines = []
            current_line_count = 0
            current_verse_start = None

        # If a single verse is too long, split it across slides
        if est_lines > max_lines:
            wrapped = textwrap.wrap(formatted, width=CHARS_PER_LINE)
            for chunk_start in range(0, len(wrapped), max_lines):
                chunk = wrapped[chunk_start:chunk_start + max_lines]
                slides.append({
                    "lines": chunk,
                    "verse_range": str(verse_num),
                    "line_count": len(chunk),
                })
            current_verse_start = None
            current_verse_end = None
            continue

        # Add verse to current slide
        # Wrap it into display lines
        wrapped = textwrap.wrap(formatted, width=CHARS_PER_LINE)
        current_lines.extend(wrapped)
        current_line_count += len(wrapped)

        if current_verse_start is None:
            current_verse_start = verse_num
        current_verse_end = verse_num

    # Don't forget the last slide
    if current_lines:
        slides.append({
            "lines": current_lines,
            "verse_range": f"{current_verse_start}-{current_verse_end}"
                if current_verse_start != current_verse_end
                else str(current_verse_start),
            "line_count": current_line_count,
        })

    return slides


def get_available_translations() -> list[dict]:
    """Return list of available translations for the UI."""
    return [
        {"id": tid, "name": display, "description": desc}
        for tid, display, desc in TRANSLATIONS
    ]
