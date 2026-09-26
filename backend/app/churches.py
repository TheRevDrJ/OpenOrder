"""Church profiles — a church is DATA in the user's folder, never code.

A church is one FOLDER, `<data_dir>/churches/<id>/`, holding everything that is
that church's:

    church.json                 the profile below
    Template - Bulletin.docx    its letterhead, name and address
    images/                     optional slide art, overriding what ships

The folder name IS the id. @decision:gold 2026-09-24 — one folder per church, with
the template beside its profile, so a congregation is a single thing you could hand
to someone and the data folder does not fill up with parallel files.

The profile carries the three things that differ between congregations:

    template   the Word file, RELATIVE TO THE CHURCH'S OWN FOLDER. Optional —
               every church's template can simply be called the same thing.
    theme      the slide palette and typography
    images     per-slot art overrides, each a filename inside images/
    labels     what each section slide is titled
    positions  per-slot plate placement, overriding the theme's default
    washes     per-slot backdrop wash, overriding the theme's default
    defaults   who normally preaches there

WHY DATA AND NOT A PYTHON MODULE: a theme that ships inside the package is a theme
every person who installs or forks this has to delete, and it would carry one
church's details to all of them. Adding a congregation must never mean editing
source and cutting a release. @decision:gold 2026-09-24

WHERE THE CHURCH'S NAME ACTUALLY LIVES: in the Word template, not here. Nothing in
this file renders a church name onto a bulletin — the template does, because it is
the document that gets printed. `name` below is only a label for the picker.

FALLBACK IS BY DESIGN: with no `churches/` folder at all the app runs on GENERIC
below — a neutral profile with the shipped template and a plain palette. A fresh
install works with nothing configured, and it names no church.

CALLED BY: app.main (the API), app.bulletin (template), app.slides (theme).
"""

from __future__ import annotations

import json
import re
from pathlib import Path

from pptx.dml.color import RGBColor

from . import paths

# The file a church's template is called unless its profile says otherwise. The
# stored name carries the church id — `<id> - Template - Bulletin.docx`.
#
# ⭐⭐ THE NAME ON DISK IS THE NAME YOU ARE HANDED. @decision:gold 2026-09-25 ·
# BUG-019. The folder already says whose it is, so the id is duplicated — but a
# downloaded copy carries the id, and if that is not also what the app READS, then
# dropping an edited file straight into the church folder leaves it there being
# ignored while the old one keeps printing. His argument: *"What if someone more
# savvy wants to edit the file directly instead of going through the upload and
# download process?"* A silent ignored edit beats the cost, which is that the id
# now appears twice and a folder rename makes them disagree.
TEMPLATE_FILENAME = "Template - Bulletin.docx"

# The neutral profile. ⛔ No church name, no address, no real person — this ships
# in a public package and is what every new install starts from.
GENERIC: dict = {
    "name": "Default",
    "template": None,  # None -> the shared template, then the one that ships
    "defaults": {
        # The only staff value the app fills: it becomes {{SPEAKER}} in the
        # bulletin. Everything else about who serves is in the church's template.
        "speakerShort": "",
    },
    "images": {},   # nothing overridden -> whatever ships with the app
    "labels": dict(),  # filled from DEFAULT_LABELS at resolve time
    "positions": {},   # per-slot plate placement; empty -> the theme's default
    "washes": {},      # per-slot backdrop wash; empty -> the theme's default
    "theme": {
        "name": "Default",
        "description": "Neutral slide theme",
        "font": "Georgia",
        "title_color": "1F3864",
        "badge_bg": "1F3864",
        "badge_fg": "FFFFFF",
        "text_color": "000000",
        "speaker_label_color": "1F3864",
        "liturgy_uppercase": False,
        "shadow_enabled": True,
        # How a section slide's title is set over its photograph. The vocabulary
        # is slide_art.STYLES; the workshop bench renders every one of them.
        "plate_style": "frosted_plate",
        # Where that plate sits: center, or a corner/edge such as bottom-right.
        "plate_position": "bottom-right",
        # The two BACKDROPS (creed, hymn) washed toward white (+) or black (-),
        # -100..100. They sit under body text set in black, so positive is
        # normally what buys contrast.
        "backdrop_wash": 0,
    },
}

# ⭐ THE SLIDE ART, BY SLOT. Every one of these was a hardcoded path in slides.py
# carrying one congregation's identity — a tree-logo watermark and oak-leaf
# photographs — so a second church's deck came out wearing the first church's
# branding. A church overrides any subset by dropping files in its images/ folder.
#
# ⛔ THE FOUR SECTION SLOTS ARE NOW PHOTOGRAPHS ONLY. Their titles used to be
# pixels baked into the JPEG, which meant changing a word meant re-cutting art at
# the right size in the right font. The title is drawn by the app now (see
# `labels`), so any photograph can be dropped in and still read correctly.
IMAGE_SLOTS = {
    "hymn_background": "the square mark behind hymn lyrics",
    "creed_background": "full-bleed, behind the creed and liturgy",
    "prayer": "full-bleed section slide",
    "offering": "full-bleed section slide",
    "concerns": "full-bleed section slide",
    "announcements": "full-bleed section slide",
    "communion": "full-bleed section slide, only on communion Sundays",
}

# What each section slide is titled. Data, not pixels.
DEFAULT_LABELS = {
    "prayer": "Pastoral Prayer",
    "offering": "Offering",
    "concerns": "Joys and Concerns",
    "announcements": "Announcements",
    "communion": "Holy Communion",
}

# Theme keys whose value is a colour and must become an RGBColor for python-pptx.
_COLOR_KEYS = ("title_color", "badge_bg", "badge_fg", "text_color", "speaker_label_color")

_ID_OK = re.compile(r"^[A-Za-z0-9 ._-]+$")


def churches_dir() -> Path:
    """Where the church folders live. Read LATE from `paths` — never snapshot it at
    import, or a settings change keeps resolving to the old folder."""
    return paths.DATA_DIR / "churches"


def church_dir(church_id: str) -> Path:
    """One church's own folder."""
    return churches_dir() / church_id


def profile_path(church_id: str) -> Path:
    return church_dir(church_id) / "church.json"


# A communion Sunday is a different ORDER OF WORSHIP, not just an extra slide — the
# Table has its own liturgy on the printed page. So it is a second template, kept
# beside the ordinary one in the church's folder.
COMMUNION_TEMPLATE_FILENAME = "Template - Bulletin - Communion.docx"


def _stubbed(cid: str, filename: str) -> str:
    """`Template - Bulletin.docx` -> `<id> - Template - Bulletin.docx`."""
    return f"{cid} - {filename}"


def _candidates(cid: str, named: str | None, default: str) -> list[Path]:
    """What to look for, best first.

    ⭐ A profile that NAMES a template wins outright — it is an explicit choice and
    nothing here second-guesses it. Otherwise the stubbed name, then the bare one.
    ⚠⚠ THE BARE NAME IS A READ FALLBACK AND MUST OUTLIVE EVERY TEMPLATE THAT
    PREDATES THE STUB. Every church folder made before 2026-09-25 holds
    `Template - Bulletin.docx`, and dropping the fallback would not raise — it would
    silently fall through to the shipped default, so a congregation's own letterhead
    would just quietly stop appearing on its bulletin.
    """
    if named:
        return [church_dir(cid) / Path(named).name]
    return [church_dir(cid) / _stubbed(cid, default), church_dir(cid) / default]


def template_for(church: dict, communion: bool = False) -> Path | None:
    """The church's template to READ, or None if it has no folder of its own.

    ⛔ Resolved INSIDE the church's folder and never outside it: a `template` value
    is a filename, not a path, so a profile cannot reach into the rest of the disk.

    ⭐ ON A COMMUNION SUNDAY the communion template is used if the church has one.
    ⚠ FALLS BACK TO THE ORDINARY TEMPLATE WHEN IT DOES NOT — a congregation that has
    not made one still gets its bulletin, just without the Table's liturgy. Refusing
    would cost a Sunday over a file that was never created.
    @decision:gold 2026-09-25

    ⚠ When nothing exists this returns the CANONICAL path rather than None, so a
    caller's `is_file()` check behaves exactly as it did before the stub.
    """
    cid = church.get("id")
    if not cid or cid == "default":
        return None
    if communion:
        for path in _candidates(cid, church.get("templateCommunion"), COMMUNION_TEMPLATE_FILENAME):
            if path.is_file():
                return path
    ordinary = _candidates(cid, church.get("template"), TEMPLATE_FILENAME)
    for path in ordinary:
        if path.is_file():
            return path
    return ordinary[0]


# ⛔⛔ EXACTLY TWO, AND NOT A REGISTRY. @decision:gold 2026-09-25 — *"Right now there
# are 2 template types per church, and that's it, and we'll proceed as if that is all
# there will ever be."* ⭐ A free-typed third kind was considered and dropped because
# nothing on the service form could ever select one, and a template nothing selects is
# a file with no readers. ⚠ The id order is the display order.
TEMPLATE_KINDS: tuple[tuple[str, str], ...] = (
    ("normal", "Normal"),
    ("communion", "Communion"),
)


def _kind_parts(church: dict, kind: str) -> tuple[str | None, str]:
    """(the profile's own filename for this kind, the conventional one)."""
    if kind == "communion":
        return church.get("templateCommunion"), COMMUNION_TEMPLATE_FILENAME
    return church.get("template"), TEMPLATE_FILENAME


def template_of_kind(church: dict, kind: str) -> Path | None:
    """The file for THIS KIND ONLY — ⛔ never falling back to the other kind.

    ⭐⭐ THE DIFFERENCE FROM `template_for` IS THE WHOLE POINT. Generation falls back
    to the ordinary template on a communion Sunday on purpose, so a church that never
    made one still gets a bulletin. A MANAGER must not: falling back there would list
    the ordinary template twice and report a communion template that does not exist as
    present, which is the one thing the screen exists to tell you.
    """
    cid = church.get("id")
    if not cid or cid == "default":
        return None
    named, default = _kind_parts(church, kind)
    options = _candidates(cid, named, default)
    for path in options:
        if path.is_file():
            return path
    return options[0]


def template_write_of_kind(church: dict, kind: str) -> Path | None:
    """Where an upload of THIS KIND lands — always the stubbed name."""
    return template_write_for(church, kind == "communion")


def template_write_for(church: dict, communion: bool = False) -> Path | None:
    """Where an UPLOAD lands — always the stubbed name, never the legacy one.

    ⭐⭐ READS FALL BACK TO THE BARE NAME, WRITES NEVER DO — the same split as
    `main._service_path` / `_service_write_path`. So an upload MIGRATES the church
    onto the new name, and the old file stops being consulted the moment one exists.
    ⚠ Which leaves the superseded file sitting in the folder. That is deliberate: it
    is the user's own document and deleting it is not ours to do.
    """
    cid = church.get("id")
    if not cid or cid == "default":
        return None
    named = church.get("templateCommunion") if communion else church.get("template")
    if named:
        return church_dir(cid) / Path(named).name
    default = COMMUNION_TEMPLATE_FILENAME if communion else TEMPLATE_FILENAME
    return church_dir(cid) / _stubbed(cid, default)


def wash_for(church: dict, slot: str) -> int:
    """How far to wash this slot's backdrop. A per-slot value beats the theme's,
    so one congregation can pin a photograph and still tune a flat watermark."""
    per = (church.get("washes") or {}).get(slot)
    if per is not None:
        return max(-100, min(100, int(per)))
    return max(-100, min(100, int((church.get("theme") or {}).get("backdrop_wash", 0) or 0)))


def _hex_to_rgb(value) -> RGBColor:
    """Accept "5A6B2E" or "#5A6B2E". A bad value raises, and the caller turns that
    into a fallback rather than letting a half-built theme reach a slide."""
    if isinstance(value, RGBColor):
        return value
    s = str(value).lstrip("#").strip()
    if len(s) != 6:
        raise ValueError(f"colour must be 6 hex digits, got {value!r}")
    return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


def _resolve_theme(raw: dict | None) -> dict:
    """Merge a profile's theme over GENERIC's, so a profile may set only the two
    keys it cares about, then convert every colour to RGBColor."""
    merged = {**GENERIC["theme"], **(raw or {})}
    out = dict(merged)
    for key in _COLOR_KEYS:
        try:
            out[key] = _hex_to_rgb(merged[key])
        except (ValueError, TypeError):
            # One bad colour must not cost the whole deck — fall back to the
            # neutral value for THAT key and keep going.
            out[key] = _hex_to_rgb(GENERIC["theme"][key])
    return out


def images_dir(church_id: str) -> Path:
    """Where a church keeps its own slide art."""
    return church_dir(church_id) / "images"


def image_for(church: dict, slot: str) -> Path | None:
    """This church's file for `slot`, or None to use whatever ships.

    ⛔ Resolved INSIDE the church's images/ folder: the profile supplies a
    FILENAME, never a path, so it cannot reach elsewhere on disk.
    """
    cid = church.get("id")
    if not cid or cid == "default" or slot not in IMAGE_SLOTS:
        return None
    name = (church.get("images") or {}).get(slot)
    if not name:
        return None
    p = images_dir(cid) / Path(name).name
    return p if p.is_file() else None


def _resolve(raw: dict, church_id: str) -> dict:
    """A profile as the rest of the app wants it: defaults filled in, theme ready."""
    defaults = {**GENERIC["defaults"], **(raw.get("defaults") or {})}
    images = {k: v for k, v in (raw.get("images") or {}).items() if k in IMAGE_SLOTS}
    return {
        "id": church_id,
        "name": raw.get("name") or church_id,
        "template": raw.get("template") or None,
        "defaults": defaults,
        "images": images,
        "labels": {**DEFAULT_LABELS, **(raw.get("labels") or {})},
        "positions": {k: v for k, v in (raw.get("positions") or {}).items() if k in IMAGE_SLOTS},
        "washes": {k: int(v) for k, v in (raw.get("washes") or {}).items() if k in IMAGE_SLOTS},
        "theme": _resolve_theme(raw.get("theme")),
    }


def list_churches() -> list[dict]:
    """Every profile on disk, id and name only, sorted by name. Empty when none
    are configured — the caller decides whether that means "show GENERIC"."""
    d = churches_dir()
    if not d.is_dir():
        return []
    out = []
    for sub in sorted(p for p in d.iterdir() if p.is_dir()):
        f = sub / "church.json"
        if not f.is_file():
            continue
        try:
            raw = json.loads(f.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            # A malformed profile is skipped, not fatal: one bad file must not
            # make every church disappear from the picker.
            continue
        out.append({"id": sub.name, "name": raw.get("name") or sub.name})
    return sorted(out, key=lambda c: c["name"].lower())


def get_church(church_id: str | None = None) -> dict:
    """The profile to generate with.

    ⛔ Falls back to GENERIC rather than raising. A missing or misspelled id must
    still produce a bulletin — a neutral one — because failing a Sunday morning
    generation over a settings problem is the worse outcome.
    """
    if not church_id:
        return _resolve(GENERIC, "default")
    if not _ID_OK.match(church_id):
        # The id becomes a filename; reject separators rather than resolving them.
        return _resolve(GENERIC, "default")
    f = profile_path(church_id)
    if not f.is_file():
        return _resolve(GENERIC, "default")
    try:
        raw = json.loads(f.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return _resolve(GENERIC, "default")
    return _resolve(raw, church_id)
