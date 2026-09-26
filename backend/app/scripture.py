"""Scripture fetcher — retrieves Bible text via API, caches locally, splits into slides."""

import json
import re
import textwrap
import time
from pathlib import Path

import httpx
import pythonbible as bible

from . import bible_local, paths
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


# ⛔⛔ THE CACHE EXPIRES, AND IT EXPIRES FOR EVERYTHING. @decision:gold 2026-09-25
# API.Bible's terms require cached scripture to be refreshed at least every 30
# days. That rule only reaches licensed text — but the policy is applied to ALL
# cached scripture on purpose. @decision:gold 2026-09-25
#   ⭐ TWO CACHE POLICIES IS A RULE SOMEBODY MAINTAINS; ONE IS A FACT. A split
#     by licence is a branch that will eventually be wrong for one of them.
#   ⭐ AND IT CAPS THE DISK. A runaway loop cannot quietly fill a drive
#     overnight, which is not worth arguing about because it costs nothing.
#   ⚠ IT DOES NOT COST OFFLINE USE, which is what an earlier version of this
#     comment claimed. The cache was never the offline guarantee — a machine
#     with no network cannot fetch an UNCACHED reference either, so the cache is
#     a latency saving, not availability. `pythonbible` is what ships text offline.
CACHE_DAYS = 30

# Their rule is fewer than 500 CONSECUTIVE verses held at a time. A single
# passage this large is a mistake rather than a service, so it is fetched and
# returned but never written down.
MAX_CACHED_VERSES = 500


def _runs(numbers: list[int]) -> list[tuple[int, int]]:
    """Split verse numbers into CONTIGUOUS runs: [15,16,17,25,26,27] -> [(15,17),(25,27)].

    ⛔⛔ TAKING min AND max SWALLOWS THE GAP. "John 14:15-17, 25-27" parses
    correctly to six verses, and asking for 15 through 27 quietly returns
    thirteen — the four verses in between were never wanted and nothing said so.
    ⚠ API.Bible will not take a comma-separated passage (HTTP 400), so each run
    is its own request and the results are stitched. @decision:gold 2026-09-25
    """
    if not numbers:
        return []
    out, start, prev = [], numbers[0], numbers[0]
    for n in numbers[1:]:
        if n != prev + 1:
            out.append((start, prev))
            start = n
        prev = n
    out.append((start, prev))
    return out


def _cache_path(reference: str, translation: str) -> Path:
    """Get cache file path for a scripture reference."""
    # Sanitize the reference for use as a filename
    safe_ref = re.sub(r'[^\w\s\-]', '', reference).strip().replace(' ', '_')
    return paths.SCRIPTURE_CACHE_DIR / translation / f"{safe_ref}.json"


def _load_cached(reference: str, translation: str) -> dict | None:
    """Cached scripture, if it is present AND still inside its 30 days."""
    path = _cache_path(reference, translation)
    if not path.exists():
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError):
        return None
    # ⚠ An entry written before the stamp existed has no _cached_at. Treat it as
    # expired rather than as fresh: failing towards a refetch is the safe side of
    # a licence rule, and it self-heals on first use.
    age = time.time() - data.get("_cached_at", 0)
    if age > CACHE_DAYS * 86400:
        path.unlink(missing_ok=True)
        return None
    return data


def _prune_expired(translation: str) -> int:
    """Delete every expired entry for one translation. Returns how many went.

    ⭐ Called on write rather than on a timer: the cache is only ever grown by a
    write, so that is the moment it can need pruning, and it needs no scheduler.
    """
    folder = paths.SCRIPTURE_CACHE_DIR / translation
    if not folder.is_dir():
        return 0
    cutoff = time.time() - CACHE_DAYS * 86400
    gone = 0
    for f in folder.glob("*.json"):
        try:
            stamped = json.loads(f.read_text(encoding="utf-8")).get("_cached_at", 0)
        except (OSError, ValueError):
            stamped = 0
        if stamped < cutoff:
            f.unlink(missing_ok=True)
            gone += 1
    return gone


def _save_cache(reference: str, translation: str, data: dict):
    """Write scripture to the cache, stamped, unless it is too large to hold."""
    if len(data.get("verses", [])) >= MAX_CACHED_VERSES:
        return
    _prune_expired(translation)
    path = _cache_path(reference, translation)
    path.parent.mkdir(parents=True, exist_ok=True)
    stamped = dict(data)
    stamped["_cached_at"] = time.time()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(stamped, f, indent=2, ensure_ascii=False)


def _extract_text_recursive(item) -> str:
    """Recursively extract plain text from any API content structure."""
    if isinstance(item, str):
        return item
    if isinstance(item, dict):
        # Direct text field (used in poetry/psalms: {"text": "...", "poem": 1})
        if "text" in item:
            return item["text"]
        # Nested content list
        if "content" in item:
            return " ".join(_extract_text_recursive(sub) for sub in item["content"])
    if isinstance(item, list):
        return " ".join(_extract_text_recursive(sub) for sub in item)
    return ""


def _extract_verse_text(content_item: dict) -> str:
    """Extract plain text from an AO Lab API content item."""
    if content_item.get("type") != "verse":
        return ""

    parts = []
    for item in content_item.get("content", []):
        text = _extract_text_recursive(item)
        if text:
            parts.append(text)

    text = " ".join(parts).strip()
    # Clean up multiple spaces
    text = re.sub(r'\s+', ' ', text)
    return text


# ⛔⛔ WHAT THE PERSON READS WHEN A TRANSLATION CANNOT BE FETCHED.
# @decision:gold 2026-09-25 · (sentence, tail). The sentence names the cause;
# the tail says why the substitute is the bundled text. ⚠ Never an API error
# string and never a vendor name where a plain cause will do — "no internet
# connection" and "could not reach API.Bible" are the same thing to the person
# holding the keyboard, and only one of them is in their language.
NO_CONNECTION = (
    "No internet connection, so this uses the Berean Standard Bible instead of {name}.",
    "BSB is built into OpenOrder and doesn't need a connection.",
)
BAD_KEY = (
    "Your API.Bible key wasn't accepted, so this uses the Berean Standard Bible instead of {name}.",
    "BSB is built into OpenOrder and doesn't need a key.",
)
NOT_IN_TRANSLATION = (
    "{name} doesn't have {ref}, so this uses the Berean Standard Bible.",
    "BSB is built into OpenOrder and always has the whole Bible.",
)


def _display_name(translation: str) -> str:
    """"NLT", not "apib:d6e14a625393b4da-01" — a message needs the name a
    person chose from the picker.

    ⚠ BEST EFFORT BY DESIGN. Resolving a licensed translation's name needs the
    catalogue, and the case this is called in is usually "the network is down",
    so it may be unavailable. Falling back to a plain phrase is better than
    printing an id at somebody.
    """
    for tid, display, _desc in TRANSLATIONS:
        if tid == translation:
            return display
    if translation.startswith("apib:"):
        try:
            from . import bible_api
            wanted_id = translation.split(":", 1)[1]
            for b in bible_api.list_bibles():
                if b["id"] == wanted_id:
                    return b["abbr"]
        except Exception:
            pass
    return "the translation you chose"


def _bsb_fallback(reference: str, wanted: str, why: str) -> dict | None:
    """Fall back to the bundled text when a fetch could not be made.

    ⛔⛔ A DECK WITH NO READING IN IT IS THE WORST OUTCOME, AND IT WAS THE OLD ONE.
    Generation returned a file, the toast said saved, and the scripture slides
    were simply absent — discoverable on a Sunday, after the sermon. Measured
    2026-09-25: with no wifi, or a key that bounces, three of the four
    translation paths wrote a 53-slide deck with the reading missing and no
    error anywhere. @decision:gold 2026-09-25

    ⭐ SO SUBSTITUTE RATHER THAN OMIT. The bundled text is on disk and cannot
    fail, so a service always has its reading.
    ⚠ AND SAY SO: the result carries a `fallback` the UI reports. A silent
    substitution would be its own lie — the deck would quietly be in a
    translation nobody chose.
    """
    if wanted == "BSB":
        return None
    got = fetch_scripture(reference, "BSB")
    if not got:
        return None
    got = dict(got)
    name = _display_name(wanted)
    # ⭐ A FINISHED SENTENCE, NOT AN ERROR CODE. This is read by whoever is
    # putting the service together, who does not care what a text service is.
    # ⚠ The tail differs per cause on purpose: it is what explains why the
    # substitute is THIS one, without ranking the translations at anybody.
    head, tail = why
    got["fallback"] = {
        "wanted": name,
        "why": head.replace("{name}", name).replace("{ref}", reference),
        "tail": tail,
    }
    return got


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

    # ⭐ A church's own licensed translation, through its own key.
    if translation.startswith("apib:"):
        from . import bible_api
        bible_id = translation.split(":", 1)[1]
        verses, fums, notice, shown_ref = [], "", "", ""
        try:
            shown_parts, exact = [], True
            for chapter in parsed["chapters"]:
                in_chapter = sorted(
                    int(str(v)[-3:]) for v in parsed["verse_ids"]
                    if int(str(v)[-6:-3]) == chapter
                )
                for first, last in _runs(in_chapter):
                    got = bible_api.fetch_passage(
                        bible_id, parsed["book"], chapter, first, last
                    )
                    if not got:
                        return _bsb_fallback(reference, translation, NOT_IN_TRANSLATION)
                    verses.extend(got["verses"])
                    fums = fums or got["fums"]
                    notice = notice or got["copyright"]
                    ref_back = got.get("reference", "")
                    shown_parts.append(ref_back)
                    # A run that came back wider than asked means the translation
                    # has no division there (a paraphrase). Then the title must
                    # say what is actually on the slide.
                    if ref_back and not ref_back.endswith(f"{first}-{last}") \
                       and not ref_back.endswith(f"{first}"):
                        exact = False
            shown_ref = "" if exact else "; ".join(p for p in shown_parts if p)
        except bible_api.BibleError as e:
            cause = BAD_KEY if "rejected that key" in str(e) else NO_CONNECTION
            return _bsb_fallback(reference, translation, cause)
        if not verses:
            return _bsb_fallback(reference, translation, NOT_IN_TRANSLATION)
        name = next((b["abbr"] for b in bible_api.list_bibles() if b["id"] == bible_id), "")
        result = {
            # ⭐ The reference THEY returned, so the title matches what is on the
            # slide. Falls back to what was asked when they say nothing.
            "reference": shown_ref or reference,
            "requested": reference,
            "translation": translation,
            "translation_name": name,
            "verses": verses,
            "slides": split_into_slides(verses),
            "source": "api.bible",
            # ⭐ Carried to the document so the notice can be printed, and to the
            # caller so the use can be reported to their fair-use system.
            "copyright": notice,
            "fums": fums,
        }
        _save_cache(reference, translation, result)
        return result

    # ⭐⭐ BSB COMES OFF THE DISK, NOT THE NETWORK. @decision:gold 2026-09-25
    # The default translation must work for a church with no connection. The
    # text is public domain and ships with the app, so there is nothing to fetch
    # and nothing to cache.
    # ⚠ Falls through to the API when the file is absent or the passage is not
    # wholly present — a partial reading must never reach a bulletin.
    if translation == "BSB" and bible_local.available():
        local = []
        for chapter in parsed["chapters"]:
            in_chapter = sorted(
                int(str(v)[-3:]) for v in parsed["verse_ids"]
                if int(str(v)[-6:-3]) == chapter
            )
            # ⛔ Each contiguous RUN, not first-to-last: a split reference like
            # "14:15-17, 25-27" must not quietly pick up 18 through 24.
            for first, last in _runs(in_chapter):
                got = bible_local.get_verses(parsed["book"], chapter, first, last)
                if not got:
                    local = []
                    break
                local.extend(got)
            if not local and in_chapter:
                break
        if local:
            return {
                "reference": reference,
                "translation": translation,
                "translation_name": "BSB",
                "verses": local,
                "slides": split_into_slides(local),
                "source": "local",
                # ⭐ Short form on purpose. The publisher's full dedication is two
                # sentences (bible_local.notice()); a slide foot gets the credit,
                # not the paragraph.
                "copyright": "Berean Standard Bible · public domain",
            }

    # Fetch from API — one chapter at a time
    all_verses = []
    for chapter in parsed["chapters"]:
        url = f"{API_BASE}/{translation}/{parsed['usfm_book']}/{chapter}.json"
        try:
            r = httpx.get(url, timeout=15)
            if r.status_code != 200:
                return _bsb_fallback(reference, translation, NO_CONNECTION)
            data = r.json()
        except Exception:
            # ⛔ The free translations are fetched, so no connection means no
            # reading — unless we substitute the one that ships with the app.
            return _bsb_fallback(reference, translation, NO_CONNECTION)

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
        return _bsb_fallback(reference, translation, NOT_IN_TRANSLATION)

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
    """Every translation the user can pick, ours first then their key's.

    ⭐ ONE LIST, TWO SOURCES. The four free texts are always there; anything a
    church's own API.Bible key unlocks is appended with an `apib:` prefix, which
    is how `fetch_scripture` knows where to send the request. ⛔ Never merged by
    display name — two different Bibles can share an abbreviation.
    """
    from . import bible_api
    out = [
        {"id": tid, "name": display, "description": desc, "source": "included"}
        for tid, display, desc in TRANSLATIONS
    ]
    # ⚠ A key that will not list must not take the picker down with it: the four
    # free texts still work, so a network problem degrades to fewer options.
    try:
        for b in bible_api.list_bibles():
            out.append({
                "id": f"apib:{b['id']}",
                "name": b["abbr"] or b["name"],
                "description": b["name"],
                "source": "api.bible",
                "copyright": b.get("copyright", ""),
            })
    except bible_api.BibleError:
        pass
    return out
