"""Licensed Bible translations from API.Bible, for churches that want one.

The texts OpenOrder ships with — BSB, KJV, ASV, WEB — are free or public domain
and need none of this. For anything copyrighted, a church gets its OWN free
API.Bible key (anyone can: three copyrighted translations, 5,000 calls a month)
and enters it in Settings. @decision:gold 2026-09-25

⭐⭐ THE KEY IS THE CHURCH'S, NOT OURS, AND THAT IS THE WHOLE DESIGN. Nothing is
shared, no quota is pooled, and the foundation pays nothing forever. It is also
what keeps this inside API.Bible's terms, which say the service is "solely for
your own noncommercial use" — a free tool given to churches satisfies that.

⚠ WHY THIS WAS DEAD UNTIL 2026-09-25, since someone will ask: the idea was
dropped on the belief that scripture living permanently inside a saved .docx
broke the caching rule. It does not. Their caching rule governs a CACHE — the
copy an app keeps to serve the text again — and requires under 500 consecutive
verses, refreshed at least every 30 days. A printed bulletin is a published
work, and what governs THAT is the individual translation's copyright policy,
which API.Bible's own terms defer to. Two different documents, two different
answers, and nobody had read either.

CALLED BY: app.scripture (routing a fetch), app.main (the settings endpoints).
"""

import json
import re
import time
import urllib.error
import urllib.parse
import urllib.request

from . import paths

API = "https://rest.api.bible/v1"

# ⚠ Read paths LATE, like everything else here — data_dir is user-configurable
# and can move while the app is running.
def _key_path():
    """⭐ Its own file, never settings.json: a key is a credential, and the
    settings blob gets handed around as a whole."""
    return paths.DATA_DIR / "api_bible.json"


# ⭐ CACHING IS NOT HANDLED HERE. `scripture.py` owns one cache for every
# translation, expiring at 30 days: one policy is a fact, two is a rule somebody
# maintains. API.Bible's terms are what set the 30. @decision:gold 2026-09-25

# Languages offered in the picker. A key also unlocks free Bibles in ~145
# languages; listing them all would bury the church's own three picks.
LANGUAGES = {"eng": "English", "spa": "Spanish", "kor": "Korean"}

# ⛔⛔ A KEY LISTS FAR MORE THAN IT BOUGHT. Measured 2026-09-25: a Starter key
# returns 251 Bibles, 40 of them English — the whole open-access library plus the
# three copyrighted ones the account picked. Offering all 40 is useless and ugly:
# it is mostly duplicates (WEB appears four times, KJV twice, BSB twice) that a
# dropdown cannot tell apart, and the free ones this app already ships.
# ⭐ SO SHOW ONLY WHAT THE KEY ACTUALLY BUYS — the licensed texts. Everything
# open says so in its own copyright line, and everything licensed makes a real
# rights claim. Derived from the data, never a hardcoded roster, because the
# picks differ per account.
# ⚠ `copyright` ONLY EXISTS WITH include-full-details. Without it the field is
# absent and every test against it silently passes as "not open".
_OPEN_LICENCE = re.compile(r"public domain|creative commons|\bCC[ -]BY\b|no copyright", re.I)
_RIGHTS_CLAIM = re.compile(r"©|\bcopyright\b", re.I)


def _is_licensed(b: dict) -> bool:
    c = (b.get("copyright") or "").strip()
    return bool(c) and not _OPEN_LICENCE.search(c) and bool(_RIGHTS_CLAIM.search(c))


_list_cache: dict = {}  # key -> (fetched_at, bibles). Metadata only, never scripture.


class BibleError(Exception):
    """Something the user should be told in words, not a 500."""


def get_key() -> str:
    try:
        return json.loads(_key_path().read_text(encoding="utf-8")).get("key", "")
    except (OSError, ValueError):
        return ""


def has_key() -> bool:
    return bool(get_key())


def set_key(key: str) -> list:
    """Validate by actually listing Bibles with it, then save.

    ⭐ A key that cannot list is not saved. Storing an unvalidated key means the
    failure surfaces on a Saturday night inside a generate, instead of here.
    """
    key = (key or "").strip()
    if not key:
        _key_path().unlink(missing_ok=True)
        _list_cache.clear()
        return []
    bibles = list_bibles(key)          # raises BibleError if the key is bad
    _key_path().parent.mkdir(parents=True, exist_ok=True)
    _key_path().write_text(json.dumps({"key": key}), encoding="utf-8")
    return bibles


def _get(path: str, key: str) -> dict:
    req = urllib.request.Request(f"{API}{path}", headers={"api-key": key})
    try:
        with urllib.request.urlopen(req, timeout=15) as r:
            return json.load(r)
    except urllib.error.HTTPError as e:
        if e.code in (401, 403):
            raise BibleError("API.Bible rejected that key.") from e
        raise BibleError(f"API.Bible returned HTTP {e.code}.") from e
    except (urllib.error.URLError, TimeoutError) as e:
        raise BibleError("Could not reach API.Bible. Check the connection.") from e


def list_bibles(key: str = "") -> list:
    """The Bibles this key can use: [{id, abbr, name, lang, copyright}]."""
    key = key or get_key()
    if not key:
        return []
    hit = _list_cache.get(key)
    if hit and time.time() - hit[0] < 3600:
        return hit[1]
    data = _get("/bibles?include-full-details=true", key)["data"]
    rows, seen = [], set()
    for b in data:
        lang = b.get("language", {}).get("id")
        if lang not in LANGUAGES or not _is_licensed(b):
            continue
        abbr = b.get("abbreviationLocal") or b.get("abbreviation") or ""
        # ⚠ Dedupe on the LOCAL abbreviation: the same translation comes back
        # several times under different ids, indistinguishable in a dropdown.
        if (lang, abbr) in seen:
            continue
        seen.add((lang, abbr))
        rows.append({
            "id": b["id"],
            "abbr": abbr,
            "name": b.get("nameLocal") or b.get("name") or "",
            "lang": lang,
            # ⭐ The publisher's own line. Not printed on a slide — the initials
            # are what non-salable media require — but it is what a salable use
            # would need, and it costs nothing to carry.
            "copyright": (b.get("copyright") or "").strip(),
        })
    bibles = sorted(rows, key=lambda b: (list(LANGUAGES).index(b["lang"]), b["abbr"].lower()))
    _list_cache[key] = (time.time(), bibles)
    return bibles


# USFM book codes, keyed by pythonbible's Book enum — the ids API.Bible wants.
def _usfm(book) -> str:
    from .scripture import BOOK_TO_USFM
    return BOOK_TO_USFM[book]


def _walk_text(node, out: list):
    """Collect text, keeping track of which verse each piece belongs to.

    ⚠ VERSE BOUNDARIES COME FROM `verse` MARKERS, NOT FROM A verseId ON EVERY
    NODE. OpenSign learned this the hard way: red-letter words of Jesus come back
    WITHOUT a verseId, so filtering on one drops them silently — Revelation 3:20
    came back empty. Track the CURRENT verse and attach everything after it.
    """
    if isinstance(node, list):
        for n in node:
            _walk_text(n, out)
        return
    if not isinstance(node, dict):
        if isinstance(node, str) and out:
            out[-1]["parts"].append(node)
        return
    # ⛔⛔ THE MARKER IS `name`, NOT `type`. Every tag node is
    # {"name": "verse", "type": "tag"} — testing type=="verse" matches nothing,
    # so no verse ever opened, text had nowhere to attach, and every passage came
    # back empty. It looked like a network or key problem and was neither.
    # ⭐ Its own items hold the verse NUMBER as text, so returning here drops it.
    if node.get("name") == "verse" and node.get("attrs", {}).get("sid"):
        out.append({"sid": node["attrs"]["sid"], "parts": []})
        return
    if node.get("type") == "text" and node.get("text"):
        if out:
            out[-1]["parts"].append(node["text"])
        return
    # Headings are not scripture: chapter labels, psalm titles, section heads.
    if node.get("name") in {"cl", "d", "ms", "ms1", "mr", "s", "s1", "s2", "r", "sp"}:
        return
    _walk_text(node.get("items") or node.get("content") or [], out)


def fetch_passage(bible_id: str, book, chapter: int, start: int, end: int) -> dict | None:
    """One passage from API.Bible: {verses, fums, copyright} or None.

    ⚠ UNVERIFIED AGAINST A LIVE KEY as of 2026-09-25 — written from their docs
    and from OpenSign's working implementation, but never run, because obtaining
    a key means creating an account and that is his to do. Treat the parsing as
    plausible rather than proven until a real fetch has been seen.
    """
    key = get_key()
    if not key:
        return None
    usfm = _usfm(book)
    pid = f"{usfm}.{chapter}.{start}"
    if end != start:
        pid += f"-{usfm}.{chapter}.{end}"
    q = urllib.parse.urlencode({
        "content-type": "json",
        "include-notes": "false",
        "include-titles": "false",
        "include-chapter-numbers": "false",
        # ⛔⛔ TRUE, DELIBERATELY, AND IT LOOKS BACKWARDS. Setting this false does
        # not merely hide the numbers — it REMOVES THE `verse` TAGS ALTOGETHER,
        # and those tags are the only thing marking where one verse ends and the
        # next begins. With false the payload has zero verse markers and every
        # passage parses to nothing. Measured 2026-09-25: false -> 0 verses,
        # true -> 2. ⭐ The number itself is dropped by `_walk_text`, which
        # returns at the marker without collecting its own items.
        "include-verse-numbers": "true",
        "include-verse-spans": "true",
    })
    data = _get(f"/bibles/{bible_id}/passages/{urllib.parse.quote(pid)}?{q}", key)
    payload = data.get("data", {})
    collected: list = []
    _walk_text(payload.get("content", []), collected)
    verses = []
    for item in collected:
        # sid is "JHN 3:16" — or "ROM 8:31-39" for a MERGED BLOCK.
        # ⛔⛔ A PARAPHRASE DOES NOT SUBDIVIDE. The Message returns Romans 8:38-39
        # as one span covering 31-39, because that is how it is published. An
        # earlier version required a plain digit here and silently dropped the
        # whole passage, so MSG returned nothing at all while NIV and NLT worked.
        # ⭐ Take the block, numbered from where it starts. ⚠ The reader gets MORE
        # than the reference asked for, which is correct and unavoidable — the
        # text cannot be split, and trimming it would misquote the translation.
        tail = item["sid"].rsplit(":", 1)[-1]
        start = tail.split("-")[0].strip()
        text = " ".join(p.strip() for p in item["parts"] if p.strip())
        text = " ".join(text.split())
        if text and start.isdigit():
            verses.append({"number": int(start), "text": text})
    if not verses:
        return None
    return {
        "verses": verses,
        # ⛔⛔ WHAT THEY GAVE US, WHICH MAY NOT BE WHAT WE ASKED FOR. A
        # paraphrase has no verse-level divisions in most places, so a request
        # for Romans 8:38-39 returns the block containing it — 8:31-39. Ignoring
        # this titles a slide "Romans 8:38-39" over verses 31 to 39, on a screen,
        # in front of a congregation. Measured 2026-09-25.
        "reference": (payload.get("reference") or "").strip(),
        # ⭐ Reported to their Fair Use Management System when the text is used.
        "fums": (data.get("meta") or {}).get("fumsToken", ""),
        "copyright": (payload.get("copyright") or "").strip(),
    }


def report_fums(token: str) -> None:
    """Tell API.Bible their text was used. Fire and forget, never raises.

    ⭐ Their Fair Use Management System: a token arrives with every passage, and
    reporting it is how publishers see usage — which is what API.Bible trades
    them instead of money, and therefore what keeps the free tier defensible.
    ⛔⛔ A FAILURE HERE MUST NEVER REACH THE USER. This runs inside a generate; a
    reporting call that raised would cost a church its slide deck over a
    courtesy. Short timeout, everything swallowed.
    ⚠ No device or session id: those identify a DISPLAY, and we are a document
    generator with no viewer to count. Sending an invented one would be a
    fabricated reading rather than a modest one.
    """
    if not token:
        return
    url = "https://fums.api.bible/f3?" + urllib.parse.urlencode({"t": token})
    try:
        urllib.request.urlopen(url, timeout=4).close()
    except Exception:
        pass
