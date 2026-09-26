"""Order of Worship — FastAPI backend."""

import json
import shutil
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

from fastapi import FastAPI, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, Response
from fastapi.staticfiles import StaticFiles

from .hymnal import search_hymns, get_hymn, get_hymn_by_ref
from .models import OrderOfWorship
from .bulletin import generate_bulletin
from .slides import generate_slides
from .scripture import fetch_scripture, get_available_translations, parse_reference
from . import bible_api, calendar_data, churches

app = FastAPI(title="Order of Worship")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

from . import paths
from .paths import FRONTEND_DIST_DIR, get_settings


def next_sunday() -> str:
    """Return the date of the upcoming Sunday (or today if it's Sunday)."""
    today = date.today()
    days_until_sunday = (6 - today.weekday()) % 7
    if days_until_sunday == 0 and today.weekday() != 6:
        days_until_sunday = 7
    return (today + timedelta(days=days_until_sunday)).isoformat()


# --- Health / Info ---

@app.get("/api/health")
def health():
    """⭐⭐ IT REPORTS WHETHER THE HYMNAL IS THERE, and that is not incidental.
    @decision:gold 2026-09-26 · BUG-024 — the hymnal is copyrighted, so it cannot ship;
    a church that has not pointed the app at one gets an EMPTY search with no error, and
    an empty search is indistinguishable from "no such hymn". The app looked broken when
    it was merely unconfigured. ⭐ Everything else — scripture, the calendar, the
    bulletin, the deck — works without it, which is exactly why it must SAY so rather
    than fail quietly.
    ⚠ The count comes off the loaded index, not off the folder: a folder that exists but
    holds no `index.json` is the same condition as no folder at all."""
    from .hymnal import _load_index
    return {
        "status": "ok",
        "nextSunday": next_sunday(),
        "hymnal": {"count": len(_load_index()), "dir": str(paths.HYMNAL_DIR)},
    }


# --- Hymnal ---

@app.get("/api/hymnal/search")
def hymnal_search(q: str = "", limit: int = 20):
    return search_hymns(q, limit)


@app.get("/api/hymnal/{source}/{number}")
def hymnal_get(source: str, number: str):
    hymn = get_hymn_by_ref(source, number)
    if not hymn:
        raise HTTPException(404, "Hymn not found")
    return hymn


# --- Services (save/load) ---
#
# ⭐⭐ A SERVICE IS KEYED ON DATE **AND CHURCH**: `{date} - {church} - Raw.json`.
# @decision:gold 2026-09-25 · FEATURE-008
#
# It used to be the date alone, which held only while both congregations shared one
# order of worship and differed in branding. It breaks the first Sunday they differ
# in CONTENT — a different sermon means different hymns — and that Sunday is real.
#
# ⭐ IT COSTS NOTHING ON AN ORDINARY WEEK, and that is why it is safe: generation
# saves first, and switching church does not clear the form. So filling it in once
# and generating for each congregation writes both files with identical content,
# through the flow that already exists.
#
# ⚠ THE DATE-ONLY NAME IS STILL READ, never written — see `_service_path`.

SNAPSHOT_KEEP = 20


def _service_path(service_date: str, church: str | None = None) -> Path:
    """Where this church's copy of a service lives.

    ⚠ FALLS BACK TO THE DATE-ONLY NAME WHEN NO CHURCH-SPECIFIC FILE EXISTS. Every
    service saved before 2026-09-25 is named `{date} - Raw.json`, and a read that
    demanded the new shape would make all of them vanish — silently, since a missing
    service just opens a blank form. ⛔ Reads fall back; WRITES always use the new
    name, so the old file is left exactly as it was.
    """
    if church and church != "default":
        own = paths.DATA_DIR / f"{service_date} - {church} - Raw.json"
        if own.exists():
            return own
        legacy = paths.DATA_DIR / f"{service_date} - Raw.json"
        return legacy if legacy.exists() else own
    return paths.DATA_DIR / f"{service_date} - Raw.json"


def _service_write_path(service_date: str, church: str | None = None) -> Path:
    """Where a save GOES — never the legacy name."""
    if church and church != "default":
        return paths.DATA_DIR / f"{service_date} - {church} - Raw.json"
    return paths.DATA_DIR / f"{service_date} - Raw.json"


def _snapshot_dir() -> Path:
    return paths.DATA_DIR / "snapshots"


def _snap_prefix(service_date: str, church: str | None) -> str:
    """`"2026-10-04 - <id> - "` — everything before the timestamp."""
    if church and church != "default":
        return f"{service_date} - {church} - "
    return f"{service_date} - "


def snapshot(service_date: str, church: str | None, reason: str) -> Path | None:
    """Copy the current saved service aside before something overwrites it.

    ⭐⭐ THIS IS WHAT AUTOSAVE TAKES AWAY AND HAS TO GIVE BACK. Closing without
    saving used to be the undo for "I opened the wrong service and typed into it".
    With the form saving continuously there is no such escape, so the escape is
    made explicit instead of being a side effect of not pressing a button.
    @decision:gold 2026-09-25

    ⛔ NOT ON EVERY KEYSTROKE — only at the boundaries where work can be lost:
    loading a service, switching church, and generating (which is a natural "this
    one was good" marker, since a document was just produced from it).
    """
    src = _service_path(service_date, church)
    if not src.exists():
        return None
    d = _snapshot_dir()
    d.mkdir(parents=True, exist_ok=True)
    # ⚠ EVERY FIELD IS SEPARATED BY " - " SO THE NAME CAN BE READ BACK. Gluing the
    # church to the stamp made `rsplit(" - ", 2)` recover "<id>-20260925-082557" as
    # the timestamp; it failed to parse and the revision was silently dropped from
    # the list — the snapshot existed and was invisible.
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    dest = d / f"{_snap_prefix(service_date, church)}{stamp} - {reason}.json"
    shutil.copy2(src, dest)
    _prune_snapshots(service_date, church)
    return dest


def _prune_snapshots(service_date: str, church: str | None) -> None:
    """Keep the newest SNAPSHOT_KEEP for this service. They are ~1.3KB each."""
    kept = sorted(_snapshot_dir().glob(f"{_snap_prefix(service_date, church)}*.json"),
                  reverse=True)
    for old in kept[SNAPSHOT_KEEP:]:
        old.unlink(missing_ok=True)


@app.get("/api/services")
def list_services(church: str | None = None):
    """Saved services for this church, newest first.

    ⚠ A legacy date-only file is listed only when this church has no file of its
    own for that date — otherwise migrating would make every date appear twice.
    """
    seen: dict[str, dict] = {}
    slug = f" - {church} - " if church and church != "default" else None
    for f in sorted(paths.DATA_DIR.glob("* - Raw.json"), reverse=True):
        name = f.name[: -len(" - Raw.json")]
        if slug and name.endswith(f" - {church}"):
            date_str = name[: -len(f" - {church}")]
            seen[date_str] = {"date": date_str, "filename": f.name, "own": True}
        elif " - " not in name:                      # a legacy date-only file
            seen.setdefault(name, {"date": name, "filename": f.name, "own": False})
    return sorted(seen.values(), key=lambda s: s["date"], reverse=True)


@app.get("/api/services/{service_date}")
def get_service(service_date: str, church: str | None = None):
    path = _service_path(service_date, church)
    if not path.exists():
        raise HTTPException(404, "Service not found")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _has_content(o: OrderOfWorship) -> bool:
    """Has anything actually been put in this service?

    ⚠ DEFAULTS ARE NOT CONTENT: `scriptureTranslation` starts at BSB,
    `liturgicalPrayer` at 895, and `speakerShortName` is prefilled from the church
    profile. A payload holding only those has had nothing typed into it.
    ⭐ Kept in step with `hasContent` in `frontend/src/types/index.ts` — the frontend
    copy stops the request, this one stops the WRITE, and the second is the one that
    protects the file from a client nobody is watching.
    """
    if any((getattr(o, f, "") or "").strip() for f in
           ("serviceTitle", "scripture", "sermonTitle", "sermonSubtitle", "offertoryNote")):
        return True
    if any(getattr(o, f, None) is not None for f in
           ("praiseHymn1", "praiseHymn2", "doxology", "creed", "prayerHymn", "closingHymn")):
        return True
    if o.heroImageFilename or o.communion:
        return True
    if o.liturgicalPrayer and getattr(o.liturgicalPrayer, "number", None) != "895":
        return True
    return False


@app.post("/api/services/{service_date}")
def save_service(service_date: str, data: OrderOfWorship, church: str | None = None,
                 snapshot_reason: str | None = None, allow_empty: bool = False):
    """Save this church's copy. ⭐ Called continuously by the form's autosave.

    ⛔⛔ AN EMPTY PAYLOAD NEVER OVERWRITES A SAVED SERVICE UNLESS IT SAYS IT MEANS TO.
    @decision:gold 2026-09-26 · BUG-022. A boundary save writes whatever is in the
    form, so opening a service or switching church before the boot load resolved wrote
    an empty order over a real one — and it cost his 2026-09-27 bulletin: seven hymn
    picks, the scripture, the sermon title and the hero image, gone with no error.
    ⭐ THE GUARD IS HERE AND NOT ONLY IN THE FORM, because the form is not the only
    writer. The packaged app and the dev server share this folder, and something wrote
    to it at 04:17 with nobody watching. A rule in one client protects one client.
    ⚠ Clearing the form IS a deliberate emptying and passes `allow_empty` — the confirm
    promises Revert brings it back, and the snapshot is what keeps that promise.
    """
    data.date = service_date
    path = _service_write_path(service_date, church)
    if not allow_empty and not _has_content(data) and path.exists():
        try:
            with open(path, "r", encoding="utf-8") as f:
                existing = OrderOfWorship(**json.load(f))
        except Exception:
            existing = None
        if existing is not None and _has_content(existing):
            raise HTTPException(
                409,
                f"Refusing to overwrite the saved service for {service_date} with an "
                f"empty form. Nothing was written.",
            )
    # ⛔ A REFUSED WRITE MUST SAY SO IN WORDS. This used to escape as a bare 500;
    # the UI then had nothing to show but a status code, over the one failure a
    # person can actually act on — the file is open somewhere, or read-only.
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data.model_dump(), f, indent=2, ensure_ascii=False)
    except PermissionError:
        raise HTTPException(
            409, f"Could not save — {path.name} is read-only or open in another program.")
    except OSError as e:
        raise HTTPException(409, f"Could not save {path.name}: {e.strerror or e}")
    # ⛔⛔ WRITE FIRST, THEN SNAPSHOT. Snapshotting before the write captured the
    # FILE rather than the FORM, so anything typed inside the autosave debounce was
    # one step behind — clear the form and Revert handed back the previous version,
    # silently losing the last edit while the dialog promised it was kept.
    # ⭐ The snapshot now holds exactly what was on screen at that moment, which is
    # what "before opening another service" / "before starting a new one" means.
    # @decision:gold 2026-09-25
    if snapshot_reason:
        snapshot(service_date, church, snapshot_reason)
    return {"saved": True, "path": str(path), "filename": path.name}


@app.get("/api/services/{service_date}/snapshots")
def list_snapshots(service_date: str, church: str | None = None):
    """Revision list for the Revert control, newest first."""
    prefix = _snap_prefix(service_date, church)
    out = []
    for f in sorted(_snapshot_dir().glob(f"{prefix}*.json"), reverse=True):
        rest = f.name[len(prefix):]
        if " - " not in rest:
            continue
        stem = f.stem.rsplit(" - ", 2)
        if len(stem) != 3:
            continue
        stamp, reason = stem[1], stem[2]
        try:
            when = datetime.strptime(stamp, "%Y%m%d-%H%M%S")
        except ValueError:
            continue
        out.append({"id": f.name, "at": when.isoformat(), "reason": reason})
    return out


@app.post("/api/services/{service_date}/revert")
def revert_service(service_date: str, body: dict, church: str | None = None):
    """Restore a snapshot over the live service.

    ⭐ The current state is snapshotted FIRST, so reverting is itself undoable —
    a revert to the wrong revision must not be the one move you cannot take back.
    """
    snap = _snapshot_dir() / str(body.get("id", ""))
    if snap.parent != _snapshot_dir() or not snap.is_file():
        raise HTTPException(404, "No such revision")
    snapshot(service_date, church, "before-revert")
    dest = _service_write_path(service_date, church)
    shutil.copy2(snap, dest)
    with open(dest, "r", encoding="utf-8") as f:
        return json.load(f)


# --- Hero image upload ---

@app.post("/api/services/{service_date}/hero-image")
async def upload_hero_image(service_date: str, file: UploadFile, church: str | None = None):
    ext = Path(file.filename).suffix.lower() or ".jpg"
    content = await file.read()

    # Convert unsupported formats (WebP, etc.) to PNG at upload time
    supported = {'.bmp', '.gif', '.jpg', '.jpeg', '.png', '.tiff', '.tif', '.wmf'}
    if ext not in supported:
        from PIL import Image
        import io
        img = Image.open(io.BytesIO(content))
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        content = buf.getvalue()
        ext = '.png'

    # ⭐ THE HERO IS PER CHURCH TOO (FEATURE-008, 2026-09-25). Different sermons
    # want different art, and it is keyed the same way the service is.
    # ⚠ The word "Theme" stays in the on-disk name deliberately — his existing files
    # use it, and changing it would leave two conventions in his own folder.
    slug = f"{church} - " if church and church != "default" else ""
    filename = f"{service_date} - {slug}Theme{ext}"
    # An uploaded hero image is INPUT — it's read back every time the service is
    # regenerated, so it belongs with the user's data.
    # ⚠ The on-disk name keeps "- Theme" on purpose: the 17 services saved before
    # 2026-09-24 already use it, and changing it here would leave two naming
    # conventions side by side in the user's own folder for no gain. The stored
    # filename is what matters, and the model reads either key.
    dest = paths.DATA_DIR / filename
    with open(dest, "wb") as f:
        f.write(content)
    return {"filename": filename}


@app.post("/api/services/{service_date}/hero-image/carry")
def carry_hero_image(service_date: str, filename: str, church: str | None = None):
    """Copy a hero image so the arriving church owns its own copy.

    ⭐⭐ CARRYING A SERVICE MUST NOT LEAVE TWO CHURCHES SHARING ONE FILE.
    @decision:gold 2026-09-26 · FEATURE-016. `heroImageFilename` resolves as
    `DATA_DIR / name`, so carrying the form across would point the arriving church's
    service at `{date} - {other church} - Theme.png`. It RENDERS, which is the whole
    problem — nobody notices until re-uploading one church's hero silently changes the
    other church's slide.

    ⚠ It overwrites the arriving church's hero for that date, deliberately: the person
    just chose to carry this service over, and the image is part of the service.
    ⛔ A missing source is not an error — the form simply has no hero to carry, and
    refusing here would block a church switch over a picture.
    """
    src = paths.DATA_DIR / Path(filename).name
    if not src.is_file():
        return {"filename": None}
    slug = f"{church} - " if church and church != "default" else ""
    dest = paths.DATA_DIR / f"{service_date} - {slug}Theme{src.suffix.lower()}"
    if dest.resolve() == src.resolve():
        return {"filename": src.name}
    try:
        shutil.copy2(src, dest)
    except OSError as e:
        raise HTTPException(409, f"Could not copy the hero image: {e.strerror or e}")
    return {"filename": dest.name}


def _default_template():
    """The template that ships with the app — the floor under every fallback."""
    from .bulletin import default_template_path
    return default_template_path()


# --- Churches ---
#
# ⭐ A church is DATA in the user's folder (app/churches.py): its Word template,
# its slide theme, and who normally preaches and leads worship there. Nothing here
# renders a church's NAME — the template does, because the template is the document
# that gets printed.


@app.get("/api/churches")
def list_churches():
    """Configured churches, for the picker. An empty list is a valid answer and
    means the app runs on the neutral default — it is not an error."""
    return {"churches": churches.list_churches(), "folder": str(churches.churches_dir())}


@app.get("/api/churches/{church_id}")
def get_church(church_id: str):
    """One profile, with its defaults, for prefilling a new service. ⛔ Colours are
    dropped here — they mean nothing to the UI and RGBColor will not serialise."""
    c = churches.get_church(church_id)
    return {
        "id": c["id"],
        "name": c["name"],
        "template": c["template"],
        "defaults": c["defaults"],
        # the bench seeds its per-slot sliders from these, so it opens showing
        # what is configured rather than an invented starting point
        "washes": {slot: churches.wash_for(c, slot)
                   for slot in ("creed_background", "hymn_background")},
    }


@app.get("/api/churches/{church_id}/preview/{slot}")
def church_art_preview(church_id: str, slot: str, style: str | None = None, label: str | None = None,
                       position: str | None = None, opacity: int | None = None,
                       blur: int | None = None):
    """Render a section slide EXACTLY as generation would, and return the PNG.

    ⭐ THE BENCH USES THIS RATHER THAN CSS. A workshop page that approximated the
    plate in CSS would be reviewing an impression; this runs the real compositor,
    so what is chosen on screen is what lands in the deck.
    """
    import io
    from .churches import IMAGE_SLOTS, image_for
    from . import slide_art
    if slot not in IMAGE_SLOTS:
        raise HTTPException(404, f"unknown slot: {slot}")
    c = churches.get_church(church_id)
    # ⛔ NO STYLE MEANS THE CHURCH'S OWN, never a hardcoded one. A default here
    # made the bench show a style the church was not configured for while
    # captioning it as the church's — the preview lied about the product.
    style = style or c["theme"].get("plate_style", "frosted_plate")
    if style not in slide_art.STYLES:
        raise HTTPException(404, f"unknown style: {style}")
    photo = image_for(c, slot) or _shipped_art(slot)
    if not photo or not photo.is_file():
        raise HTTPException(404, f"no art for {slot}")
    # ⭐ Overrides are for TRYING a value on the bench. None given means the
    # church's own, so the default view stays honest about the product.
    theme = dict(c["theme"])
    if opacity is not None:
        theme["plate_opacity"] = max(0, min(100, opacity))
    if blur is not None:
        theme["plate_blur"] = max(0, min(60, blur))
    pos = position or (c.get("positions") or {}).get(slot)
    img = slide_art.render(photo, label or c["labels"].get(slot, slot.title()), theme, style, pos)
    buf = io.BytesIO(); img.save(buf, "JPEG", quality=86); buf.seek(0)
    return Response(content=buf.read(), media_type="image/jpeg")


@app.get("/api/art/styles")
def art_styles():
    """The plate vocabulary, for the bench to enumerate rather than hardcode."""
    from . import slide_art
    return {"styles": [{"id": k, "description": v} for k, v in slide_art.STYLES.items()]}


@app.get("/api/churches/{church_id}/art/{slot}")
def church_art(church_id: str, slot: str, wash: int | None = None):
    """Serve the image this church would actually use for `slot`.

    ⭐ EXISTS SO A BENCH REVIEWS THE REAL THING. A bench that renders its own copy
    reviews the copy, and drifts from the product without anyone noticing. This runs
    the same resolver generation runs, so what it shows is what a deck would contain.
    """
    from .churches import IMAGE_SLOTS, image_for
    if slot not in IMAGE_SLOTS:
        raise HTTPException(404, f"unknown slot: {slot}")
    c = churches.get_church(church_id)
    path = image_for(c, slot) or _shipped_art(slot)
    if not path or not path.is_file():
        raise HTTPException(404, f"no art for {slot}")
    # ⭐ None means the church's own setting, so the bench's default view is what
    # a deck contains; a number is the bench TRYING a value.
    from .churches import wash_for
    amount = wash_for(c, slot) if wash is None else max(-100, min(100, wash))
    import io
    from . import slide_art
    # ⛔ EVEN AT WASH 0 THIS GOES THROUGH THE COMPOSITOR. Serving the file raw used
    # to be the cheap path, and it made the bench lie twice over: a transparent
    # emblem read BLACK against the bench's dark page, and the square was framed by
    # the browser rather than by the thing that builds the deck.
    img = slide_art.wash(path, int(amount), slide_art.canvas_for(slot))
    buf = io.BytesIO(); img.save(buf, "JPEG", quality=88); buf.seek(0)
    return Response(content=buf.read(), media_type="image/jpeg")


def _shipped_art(slot: str) -> Path | None:
    """The packaged fallback for a slot — the single place that mapping lives."""
    from .slides import (HYMN_BG, CREED_BG, PRAYER_SLIDE, OFFERING_SLIDE,
                         CONCERNS_SLIDE, ANNOUNCEMENTS_SLIDE)
    return {
        "hymn_background": HYMN_BG, "creed_background": CREED_BG,
        "prayer": PRAYER_SLIDE, "offering": OFFERING_SLIDE,
        "concerns": CONCERNS_SLIDE, "announcements": ANNOUNCEMENTS_SLIDE,
    }.get(slot)


# --- Generate ---

def _require_church(church_id: str | None, action: str = "generating") -> dict:
    """The profile to act on — or a refusal.

    ⛔⛔ @decision:gold 2026-09-25 · BUG-017 — TEMPLATES GO THROUGH HERE TOO. An
    upload with no church silently wrote the SHARED template at the data-folder
    root, which `bulletin.template_path` never reads again once a church has one
    of its own — so the file was written, the UI said it worked, and nothing on
    earth would ever use it. ⭐ `action` only names the verb in the message; the
    refusal is the same one, and there is deliberately only one of them.

    ⛔⛔ @decision:gold 2026-09-24 · BUG-010 — REFUSE, do not fall back. Once
    profiles exist, a missing id must STOP the generation rather than quietly
    resolve to the neutral default. It had been falling through, and the deck that
    came out looked almost right — the packaged art belongs to one congregation,
    so the only tell was a colour — which is exactly the kind of thing found on a
    Sunday morning instead of at the desk.

    ⭐ THE FALLBACK IN `churches.get_church` IS STILL CORRECT AND IS NOT TOUCHED.
    A fresh install with nothing configured SHOULD produce a neutral bulletin; the
    refusal only applies once the user has told us which congregations exist.
    """
    configured = churches.list_churches()
    if not configured:
        return churches.get_church(None)
    if not church_id:
        raise HTTPException(400, f"Choose a church before {action}.")
    c = churches.get_church(church_id)
    if c.get("id") != church_id:
        known = ", ".join(x["id"] for x in configured)
        raise HTTPException(400, f"Unknown church {church_id!r}. Configured: {known}")
    return c


@app.post("/api/generate/bulletin/{service_date}")
def gen_bulletin(service_date: str, church: str | None = None):
    # ⛔⛔ THE CHURCH GOES IN. A service is keyed on date AND church (FEATURE-008);
    # reading it without one looked for the pre-migration filename and reported
    # "Service not found — save first" over a service that was saved perfectly
    # well. ⚠ Every `_service_path` call takes a church — this pair was missed
    # because the tests exercised save/load/revert and not GENERATE.
    path = _service_path(service_date, church)
    if not path.exists():
        raise HTTPException(404, "Service not found — save first")
    with open(path, "r", encoding="utf-8") as f:
        data = OrderOfWorship(**json.load(f))
    # ⛔ OUTSIDE THE try: this raises HTTPException(400), and the broad
    # `except Exception` below would otherwise catch its own 400 and
    # re-raise it as a 500 — the refusal reached the user wearing the
    # wrong status and the words "generation failed", which is a lie.
    profile = _require_church(church)
    try:
        filepath = generate_bulletin(data, profile)
    except PermissionError:
        raise HTTPException(
            409, "The bulletin file is open in another program (probably Word). Close it and try again."
        )
    except Exception as e:
        import traceback, logging
        logging.getLogger("OpenOrder").error("Bulletin generation failed:\n%s", traceback.format_exc())
        raise HTTPException(500, f"Bulletin generation failed: {str(e)}")
    # The folder travels with the name so the UI can tell the user WHERE it
    # landed — there is no browser download bar in the desktop app.
    return {"filename": filepath.name, "folder": str(filepath.parent)}


@app.post("/api/generate/slides/{service_date}")
def gen_slides(service_date: str, church: str | None = None):
    # ⛔⛔ THE CHURCH GOES IN. A service is keyed on date AND church (FEATURE-008);
    # reading it without one looked for the pre-migration filename and reported
    # "Service not found — save first" over a service that was saved perfectly
    # well. ⚠ Every `_service_path` call takes a church — this pair was missed
    # because the tests exercised save/load/revert and not GENERATE.
    path = _service_path(service_date, church)
    if not path.exists():
        raise HTTPException(404, "Service not found — save first")
    with open(path, "r", encoding="utf-8") as f:
        data = OrderOfWorship(**json.load(f))
    # ⛔ OUTSIDE THE try: this raises HTTPException(400), and the broad
    # `except Exception` below would otherwise catch its own 400 and
    # re-raise it as a 500 — the refusal reached the user wearing the
    # wrong status and the words "generation failed", which is a lie.
    profile = _require_church(church)
    try:
        filepath = generate_slides(data, profile)
        from .slides import _LAST_FALLBACK
        fallback = dict(_LAST_FALLBACK)
    except PermissionError:
        raise HTTPException(
            409, "The slides file is open in another program (probably PowerPoint). Close it and try again."
        )
    except Exception as e:
        import traceback, logging
        logging.getLogger("OpenOrder").error("Slide generation failed:\n%s", traceback.format_exc())
        raise HTTPException(500, f"Slide generation failed: {str(e)}")
    # ⭐ Only present when the chosen translation had to be stood in for. The
    # UI says so; the alternative is a deck quietly in a translation nobody picked.
    return {"filename": filepath.name, "folder": str(filepath.parent),
            **({"fallback": fallback} if fallback else {})}


# --- Scripture ---

# --- Template management ---

EXPECTED_PLACEHOLDERS = [
    '{{DATE}}', '{{SERVICE_TITLE}}',
    '{{OPENING_HYMN_TITLE}}', '{{OFFERTORY_HYMN_TITLE}}', '{{OFFERTORY_HYMN}}',
    '{{DOX}}', '{{CREED}}', '{{CREED_TITLE}}',
    '{{PRAYER_HYMN_NUMBER}}', '{{PRAYER_HYMN_TITLE}}',
    '{{LITURGICAL_PRAYER}}',
    '{{SCRIPTURE}}', '{{SPEAKER}}',
    '{{SERMON_TITLE}}', '{{SERMON_SUBTITLE}}',
    '{{CLOSING_HYMN_NUMBER}}', '{{CLOSING_HYMN_TITLE}}',
    '{{CALENDAR_BLOCK}}',
]


def _placeholders_in(doc) -> tuple[list[str], list[str]]:
    """(found, missing) for one opened document — PARAGRAPHS AND TABLES.

    ⛔⛔ SCAN BOTH, BECAUSE SUBSTITUTION DOES. `generate_bulletin` replaces inside
    table cells as well (the QR block lives in one), so a scan of `doc.paragraphs`
    alone reports a placeholder that WORKS as missing. Harmless while it only fed a
    count; not harmless now that it feeds the tag diff on the replace confirm, where
    a false "removed" would talk somebody out of a good template.
    """
    parts = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                parts.extend(p.text for p in cell.paragraphs)
    full_text = "\n".join(parts)
    found = [p for p in EXPECTED_PLACEHOLDERS if p in full_text]
    missing = [p for p in EXPECTED_PLACEHOLDERS if p not in full_text]
    return found, missing


@app.get("/api/templates")
def list_templates(church: str | None = None):
    """Both of a church's templates and the state of each.

    ⭐⭐ THERE ARE EXACTLY TWO KINDS AND THERE IS NO REGISTRY. @decision:gold
    2026-09-25 — *"Right now there are 2 template types per church, and that's it,
    and we'll proceed as if that is all there will ever be."* A user-defined kind was
    considered and dropped: nothing on the service form could select one, so it would
    be a file with no readers (the shape of BUG-017).

    ⚠ STRICT PER KIND — `template_of_kind` does NOT fall back to the other kind, so a
    church with no communion template reports it as absent instead of showing the
    ordinary one twice. Generation still falls back, deliberately (`template_for`).
    """
    from docx import Document

    profile = _require_church(church, "managing templates")

    # ⭐ A FRESH INSTALL WITH NO CHURCHES STILL GETS A ROW. Its template is the
    # shared one at the data-folder root (or the shipped default), and it is the file
    # that install actually reads — so the screen must be able to replace it. One row,
    # no picker, and the UI keeps a single code path.
    if not profile.get("id") or profile.get("id") == "default":
        from .bulletin import template_path
        active = template_path(None)
        found, missing = ([], EXPECTED_PLACEHOLDERS)
        if active.is_file():
            found, missing = _placeholders_in(Document(str(active)))
        return {"church": None, "templates": [{
            "kind": "normal", "label": "Bulletin", "name": active.name,
            "exists": active.is_file(), "found": found, "missing": missing,
            "total_expected": len(EXPECTED_PLACEHOLDERS),
        }]}

    out = []
    for kind, label in churches.TEMPLATE_KINDS:
        path = churches.template_of_kind(profile, kind)
        row = {
            "kind": kind,
            "label": label,
            "name": path.name if path else None,
            "exists": bool(path and path.is_file()),
            "found": [],
            "missing": EXPECTED_PLACEHOLDERS,
            "total_expected": len(EXPECTED_PLACEHOLDERS),
        }
        if row["exists"]:
            found, missing = _placeholders_in(Document(str(path)))
            row["found"], row["missing"] = found, missing
        out.append(row)
    return {"church": profile.get("id"), "templates": out}


@app.get("/api/template/info")
def template_info(church: str | None = None):
    """Info about the bulletin template THIS CHURCH will use.

    ⭐ Templates live per church now (`churches/<id>/`), so "the template" is only
    meaningful once you say whose. With no church it reports the shared/shipped one.
    """
    from .bulletin import template_path, user_template_path
    from docx import Document

    c = churches.get_church(church) if church else None
    active = template_path(c)
    if not active.exists():
        return {"exists": False, "name": None, "placeholders": []}

    found, missing = _placeholders_in(Document(str(active)))

    return {
        "exists": True,
        "name": active.name,
        "custom": active != _default_template(),
        "church": (c or {}).get("id"),
        "found": found,
        "missing": missing,
        "total_expected": len(EXPECTED_PLACEHOLDERS),
    }


@app.post("/api/template/export")
def export_template(church: str | None = None, kind: str = "normal"):
    """
    Put a copy of the current bulletin template in the output folder.

    Deliberately a copy-to-output rather than a browser download: the desktop
    app has no download bar, so a FileResponse would land nowhere the user can
    see. This behaves identically in the app and in a browser, and the UI can
    say which folder it went to.
    """
    import shutil
    from .bulletin import _slug

    profile = _require_church(church, "downloading the template")
    active = churches.template_of_kind(profile, kind)
    if not active or not active.is_file():
        label = dict(churches.TEMPLATE_KINDS).get(kind, kind)
        raise HTTPException(404, f"There is no {label} template to download yet.")

    # ⛔⛔ THE COPY CARRIES THE CHURCH IN ITS NAME — @decision:gold 2026-09-25 ·
    # BUG-018. The copy OVERWRITES, so before this, downloading a second
    # congregation's template replaced the first one in the output folder with a
    # file that looked identical. Nothing in the name said whose it was, and
    # uploading the wrong one would have put one church's letterhead on the other's.
    # ⚠ The STORED name carries the id too now (BUG-019), so this only adds the stub
    # to a church still on the legacy filename — otherwise it would double it.
    # ⛔⛔ A DOWNLOAD DESTROYS NOTHING — @decision:gold 2026-09-25, his catch: it had
    # been a plain copy onto a fixed name, so editing the copy in place and downloading
    # again ate the edit. `unique_path` is the rule the rest of the app already follows.
    paths.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    slug = _slug(profile)
    name = active.name if active.name.startswith(slug) else f"{slug}{active.name}"
    dest = paths.unique_path(paths.OUTPUT_DIR / name)
    try:
        shutil.copy2(active, dest)
    except PermissionError:
        raise HTTPException(
            409, f"{dest.name} is open in another program (probably Word). Close it and try again."
        )
    return {"filename": dest.name, "folder": str(dest.parent)}


@app.post("/api/template/upload")
async def upload_template(
    file: UploadFile,
    church: str | None = None,
    kind: str = "normal",
    confirm: bool = False,
):
    """Replace one of a church's templates — in TWO PHASES.

    ⛔⛔ NOTHING IS WRITTEN WITHOUT `confirm=true`. @decision:gold 2026-09-25 — his
    ruling: *"Upload is destructive, so that needs to be a modal that says which
    church and template will be overwritten. And it should be checked for tags before
    it replaces, not after."* It used to write first and report the placeholder count
    afterwards, so a template missing half its tags had already replaced a good one by
    the time anybody could read about it.
    ⭐ The gate is in the ENDPOINT, not only in the UI, so the destructive path cannot
    be reached by accident — not by a stray click and not by a direct call.

    ⛔ IT NEVER REFUSES OVER TAGS, and that is deliberate: a template may leave an
    element out on purpose (FEATURE-002), so a missing placeholder is a legitimate
    choice. The answer is to SAY SO and let him decide — *"You told them about it,
    their choice."*

    ⚠ The file is sent twice, once to check and once to write. A quarter of a megabyte
    over loopback is cheaper than holding an unconfirmed upload in server state and
    having to expire it.
    """
    profile = _require_church(church, "uploading a template")
    if kind not in dict(churches.TEMPLATE_KINDS):
        raise HTTPException(400, f"Unknown template kind {kind!r}")

    if not file.filename.endswith('.docx'):
        raise HTTPException(400, "Template must be a .docx file")

    content = await file.read()

    # Validate it's a real docx by trying to open it
    try:
        import io
        from docx import Document
        doc = Document(io.BytesIO(content))
    except Exception:
        raise HTTPException(400, "Invalid .docx file — could not parse")

    found, missing = _placeholders_in(doc)

    # What is about to be replaced, and what its tags are — read BEFORE any write.
    from .bulletin import user_template_path
    dest = churches.template_write_of_kind(profile, kind) or user_template_path()
    current = churches.template_of_kind(profile, kind)
    replacing = bool(current and current.is_file())
    outgoing_found: list[str] | None = None
    if replacing:
        try:
            outgoing_found, _ = _placeholders_in(Document(str(current)))
        except Exception:
            # An unreadable file on disk is still being replaced; just do not
            # claim to know what was in it.
            outgoing_found = None

    # ⭐ THE DIFF IS AGAINST THE FILE BEING REPLACED, not against the full expected
    # list — *"previous template had x tags, this one has y."* Equal sets produce two
    # empty lists, and the UI then says nothing at all.
    added = sorted(set(found) - set(outgoing_found)) if outgoing_found is not None else []
    removed = sorted(set(outgoing_found) - set(found)) if outgoing_found is not None else []

    report = {
        "church": profile.get("id"),
        "kind": kind,
        "label": dict(churches.TEMPLATE_KINDS)[kind],
        "target": dest.name,
        "replacing": replacing,
        "found": found,
        "missing": missing,
        "outgoing_count": len(outgoing_found) if outgoing_found is not None else None,
        "added": added,
        "removed": removed,
        "total_expected": len(EXPECTED_PLACEHOLDERS),
    }

    if not confirm:
        return {"saved": False, **report}

    dest.parent.mkdir(parents=True, exist_ok=True)
    with open(dest, "wb") as f:
        f.write(content)

    return {"saved": True, "exists": True, "custom": True, "name": dest.name, **report}


# --- Scripture ---

@app.get("/api/scripture/translations")
def scripture_translations():
    return get_available_translations()


@app.get("/api/scripture/fetch")
def scripture_fetch(ref: str = "", translation: str = "BSB"):
    if not ref.strip():
        raise HTTPException(400, "Scripture reference is required")
    parsed = parse_reference(ref)
    if not parsed:
        raise HTTPException(400, f"Could not parse scripture reference: {ref}")
    data = fetch_scripture(ref, translation)
    if not data:
        raise HTTPException(404, f"Could not fetch scripture for {ref} ({translation})")
    return data


# --- API.Bible (a church's own key, for licensed translations) ---

@app.get("/api/bible/key")
def api_bible_key_status():
    """⛔ Never returns the key itself — only whether one is stored, and what it
    unlocks. A credential that can be read back out of an endpoint is a
    credential that ends up in a screenshot."""
    try:
        bibles = bible_api.list_bibles() if bible_api.has_key() else []
    except bible_api.BibleError as e:
        # ⛔⛔ A DEAD KEY AND A DEAD CONNECTION ARE DIFFERENT PROBLEMS WITH
        # DIFFERENT ANSWERS, and this used to report both as "could not reach".
        # One is fixed by getting a new key, the other by waiting. The fetch
        # layer already distinguishes them (401 vs unreachable); the reason was
        # simply discarded here. @decision:gold 2026-09-25
        # ⚠ In the PICKER both look identical, and that is correct — there is
        # nothing to offer either way. Settings is where the why belongs.
        return {"configured": True, "reachable": False, "bibles": [], "reason": str(e)}
    return {"configured": bible_api.has_key(), "reachable": True, "bibles": bibles}


@app.post("/api/bible/key")
def api_bible_set_key(body: dict):
    """Save a key, or clear it with an empty string. Validated before it lands."""
    try:
        bibles = bible_api.set_key(body.get("key", ""))
    except bible_api.BibleError as e:
        # ⭐ 400 in words, not a 500 — this is a user typing a key wrong.
        raise HTTPException(400, str(e))
    return {"configured": bool(body.get("key", "").strip()), "bibles": bibles}


# --- Settings ---

@app.get("/api/settings")
def api_get_settings():
    settings = get_settings()
    settings.update(paths.current_dirs())
    # Legacy key — the old UI read this one.
    settings["data_dir_current"] = str(paths.DATA_DIR)
    return settings


@app.post("/api/settings/dir")
def api_set_dir(body: dict):
    """
    Point one location at a new folder.

    key: "output_dir" | "data_dir" | "hymnal_dir"
    """
    key = (body.get("key") or "").strip()
    new_dir = (body.get("path") or "").strip()
    if key not in ("output_dir", "data_dir", "hymnal_dir"):
        raise HTTPException(400, f"Unknown setting: {key}")
    if not new_dir:
        raise HTTPException(400, "path is required")

    new_dir = str(Path(new_dir).expanduser().resolve())
    p = Path(new_dir)
    # The hymnal is source material the user supplies, so it must already
    # exist. The other two are ours to create.
    if key == "hymnal_dir":
        if not p.is_dir():
            raise HTTPException(400, f"Folder does not exist: {new_dir}")
    else:
        try:
            p.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            raise HTTPException(400, f"Could not create folder: {e}")

    paths.set_dir(key, new_dir)

    # The hymnal index is cached in memory — reload it when its folder moves.
    if key == "hymnal_dir":
        from .hymnal import _load_index
        _load_index(force=True)

    return {"status": "ok", **paths.current_dirs()}


@app.post("/api/settings/data-dir")
def api_set_data_dir(body: dict):
    """Legacy endpoint — sets the calendar/data folder."""
    return api_set_dir({"key": "data_dir", "path": body.get("data_dir", "")})


# --- Calendar ---

# --- Calendar ---
#
# ⛔⛔ EVERY CALENDAR ROUTE CARRIES `?church=` AND THE FILES ARE PER CHURCH.
# They used to be four shared files in data_dir, so a second congregation printed
# the first one's calendar on its own bulletin. There is deliberately NO fallback
# to the shared file (calendar_data._calendar_file) — a church with no calendar
# has an empty one, which is visibly true, rather than quietly borrowing.
# @decision:gold 2026-09-24

@app.get("/api/calendar/templates")
def get_calendar_templates(church: str | None = None):
    return calendar_data.list_templates(church)


@app.post("/api/calendar/templates")
def create_calendar_template(template: dict, church: str | None = None):
    return calendar_data.add_template(template, church)


@app.put("/api/calendar/templates/{template_id}")
def update_calendar_template(template_id: str, updates: dict, church: str | None = None):
    result = calendar_data.update_template(template_id, updates, church)
    if not result:
        raise HTTPException(404, "Template not found")
    return result


@app.delete("/api/calendar/templates/{template_id}")
def delete_calendar_template(template_id: str, church: str | None = None):
    if not calendar_data.delete_template(template_id, church):
        raise HTTPException(404, "Template not found")
    return {"deleted": True}


@app.get("/api/calendar/events")
def get_calendar_events(church: str | None = None):
    return calendar_data.list_events(church)


@app.post("/api/calendar/events")
def create_calendar_event(event: dict, church: str | None = None):
    return calendar_data.add_event(event, church)


@app.put("/api/calendar/events/{event_id}")
def update_calendar_event(event_id: str, updates: dict, church: str | None = None):
    result = calendar_data.update_event(event_id, updates, church)
    if not result:
        raise HTTPException(404, "Event not found")
    return result


@app.delete("/api/calendar/events/{event_id}")
def delete_calendar_event(event_id: str, church: str | None = None):
    if not calendar_data.delete_event(event_id, church):
        raise HTTPException(404, "Event not found")
    return {"deleted": True}


@app.get("/api/calendar/overrides")
def get_calendar_overrides(church: str | None = None):
    return calendar_data.list_overrides(church)


@app.post("/api/calendar/overrides/skip")
def set_calendar_skip(body: dict, church: str | None = None):
    """Toggle skip for a recurring event instance.
    Body: {templateId, date, skip: bool}"""
    return calendar_data.toggle_skip(
        body["templateId"], body["date"], bool(body.get("skip", False)), church
    )


@app.get("/api/calendar/for-service/{service_date}")
def get_calendar_for_service(service_date: str, church: str | None = None):
    """Get all events for the 4 weeks following the service date."""
    return calendar_data.get_calendar_for_service(service_date, church)


@app.get("/api/calendar/note/{service_date}")
def get_calendar_note(service_date: str, church: str | None = None):
    return calendar_data.get_note(service_date, church)


@app.post("/api/calendar/note/{service_date}")
def save_calendar_note(service_date: str, note: dict, church: str | None = None):
    return calendar_data.save_note(service_date, note, church)


# --- File downloads ---

@app.get("/api/download/{filename}")
def download_file(filename: str):
    # Serves both halves: generated documents live in the output folder, while
    # theme images the UI previews live with the user's data. Reject any path
    # separator first — the name is used to build a path.
    if "/" in filename or "\\" in filename or ".." in filename:
        raise HTTPException(400, "Invalid filename")
    for folder in (paths.OUTPUT_DIR, paths.DATA_DIR):
        path = folder / filename
        if path.exists():
            return FileResponse(path, filename=filename)
    raise HTTPException(404, "File not found")


# --- Serve frontend in production ---
FRONTEND_DIST = FRONTEND_DIST_DIR
if FRONTEND_DIST.exists():
    # Serve static assets (JS, CSS, images)
    app.mount("/assets", StaticFiles(directory=FRONTEND_DIST / "assets"), name="static-assets")
    # Serve other static files (favicon, etc.)
    for static_file in FRONTEND_DIST.iterdir():
        if static_file.is_file() and static_file.name != "index.html":
            @app.get(f"/{static_file.name}")
            def serve_static(f=static_file):
                return FileResponse(f)

    # SPA fallback — serve index.html for any non-API route
    @app.get("/{full_path:path}")
    def serve_spa(full_path: str):
        # Don't intercept API routes
        if full_path.startswith("api/"):
            raise HTTPException(404)
        return FileResponse(FRONTEND_DIST / "index.html")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=6801)
