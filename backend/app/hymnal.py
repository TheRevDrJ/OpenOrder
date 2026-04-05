"""Hymnal index and search functionality."""

import json
import re
from pathlib import Path

HYMNAL_DIR = Path(__file__).parent.parent.parent / "hymnal-json"

_index: list[dict] | None = None


def _load_index() -> list[dict]:
    global _index
    if _index is None:
        with open(HYMNAL_DIR / "index.json", "r", encoding="utf-8") as f:
            _index = json.load(f)
    return _index


def _strip_punctuation(s: str) -> str:
    """Remove punctuation for fuzzy matching."""
    return re.sub(r'[^\w\s]', '', s)


def search_hymns(query: str, limit: int = 20) -> list[dict]:
    """Search hymns by number or title. Case-insensitive, punctuation-tolerant."""
    index = _load_index()
    query_lower = query.lower().strip()

    if not query_lower:
        return []

    query_stripped = _strip_punctuation(query_lower)

    # Exact number match first
    exact = [h for h in index if h["number"] == query_lower.lstrip("0")]
    # Also try with leading zeros
    if not exact:
        exact = [h for h in index if h["number"].zfill(3) == query_lower.zfill(3)]

    # Then title contains (try exact first, then punctuation-stripped)
    title_matches = [
        h for h in index
        if (query_lower in h["title"].lower()
            or query_stripped in _strip_punctuation(h["title"].lower()))
        and h not in exact
    ]

    # Then number starts-with
    number_matches = [
        h for h in index
        if h["number"].startswith(query_lower) and h not in exact and h not in title_matches
    ]

    results = exact + title_matches + number_matches
    return results[:limit]


def get_hymn(source: str, filename: str) -> dict | None:
    """Load full hymn data from a JSON file."""
    path = HYMNAL_DIR / source / filename
    if not path.exists():
        return None
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def get_hymn_by_ref(source: str, number: str) -> dict | None:
    """Find and load a hymn by source and number."""
    index = _load_index()
    for entry in index:
        if entry["source"] == source and entry["number"] == number:
            return get_hymn(source, Path(entry["file"]).name)
    return None
