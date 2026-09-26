"""The Berean Standard Bible, on disk, so scripture works with no network.

⭐⭐ WHY THIS EXISTS: a church with no internet should still be able to use as
much of this software as possible. Until this, EVERY translation the app offered
was fetched over the network, including the default, and a service generated
offline simply had no scripture in it — no error, just a blank space in
something about to be printed. @decision:gold 2026-09-25

⭐ WE ARE ALLOWED TO SHIP IT. The BSB text was dedicated to the public domain on
2023-04-30 and its publisher states "licensing is not required for any use".
The file is `bible-text/bsb.txt`, BYTE-FOR-BYTE AS PUBLISHED — tab separated,
"Book C:V<TAB>text", with their own dedication on the first two lines.
⛔ Do not "clean up" or reformat it. Keeping the publisher's own file means the
provenance of every verse is checkable against bereanbible.com, and it means no
conversion step of ours can quietly corrupt a word of it.

CALLED BY: app.scripture (preferred over the network for BSB).
"""

import re
from pathlib import Path

import pythonbible as bible

from . import paths

# ⛔⛔ THE FILE'S BOOK NAMES ARE NOT pythonbible's, AND TWO OF THEM DIFFER.
# Measured, not assumed: 64 of 66 match and these two do not. Psalms is the most
# quoted book in worship, so without this map the failure would have been a
# blank space in a bulletin, silently, on the most likely reference in the book.
# ⚠ Keyed to the Book ENUM rather than to a title string, so a later rename of a
# display name upstream cannot quietly break the lookup.
_ALIASES = {
    "Psalm": bible.Book.PSALMS,
    "Song of Solomon": bible.Book.SONG_OF_SONGS,
}

_REF = re.compile(r"^(.+?)\s+(\d+):(\d+)$")

_text: dict | None = None          # {Book: {chapter: {verse: text}}}
_notice: str = ""                  # the publisher's own dedication


def _find() -> Path | None:
    """Where the text lives.

    ⭐ It is PACKAGED DATA, not user data, so it rides inside the bundle and
    `_bundle_root()` is the answer in a frozen app. Unlike the hymnal it is never
    something a church swaps out — it is one fixed public-domain text — so it
    needs no setting and no copy beside the executable.
    ⚠ `_app_root()` is checked too: it is the repo root on the dev server, and it
    would also catch a hand-placed copy beside a built app.
    """
    for root in (paths._bundle_root(), paths._app_root()):
        p = root / "bible-text" / "bsb.txt"
        if p.is_file():
            return p
    return None


def _load() -> dict:
    """Parse the file once, into {Book: {chapter: {verse: text}}}."""
    global _text, _notice
    if _text is not None:
        return _text
    _text = {}
    path = _find()
    if not path:
        return _text
    titles = {b.title: b for b in bible.Book}
    notice_lines = []
    with open(path, encoding="utf-8-sig") as f:
        for line in f:
            ref, _, body = line.partition("\t")
            m = _REF.match(ref.strip())
            if not m:
                # The first two lines are the publisher's dedication. Keep them:
                # they are the attribution we print, in their own words.
                if ref.strip() and "Verse" != ref.strip():
                    notice_lines.append(ref.strip())
                continue
            name, chapter, verse = m.group(1), int(m.group(2)), int(m.group(3))
            book = _ALIASES.get(name) or titles.get(name)
            if not book:
                continue
            _text.setdefault(book, {}).setdefault(chapter, {})[verse] = body.strip()
    _notice = " ".join(notice_lines)
    return _text


def available() -> bool:
    """True when the whole text is present — not merely that a file exists."""
    t = _load()
    return len(t) == 66


def notice() -> str:
    """The publisher's public-domain statement, for the bulletin and the slide."""
    _load()
    return _notice


def get_verses(book, chapter: int, start: int, end: int) -> list[dict]:
    """[{number, text}] for one reference, or [] if any of it is missing.

    ⛔ ALL OR NOTHING ON PURPOSE. A partial passage is worse than none: it would
    print as a complete reading that quietly stops early, and nothing on the page
    would say so. An empty result lets the caller fall back to the network.
    """
    t = _load()
    verses = t.get(book, {}).get(chapter, {})
    if not verses:
        return []
    out = []
    for n in range(start, end + 1):
        text = verses.get(n)
        if text is None:
            return []
        out.append({"number": n, "text": text})
    return out
