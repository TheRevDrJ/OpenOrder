"""
Calendar data management.

Three JSON files in data/:
- calendar-templates.json   Recurring event definitions
- calendar-overrides.json   Per-date skips and overrides
- calendar-events.json      One-off events

Generation: given a service date, returns sorted list of events for the
next 4 weeks (28 days), with recurring events expanded and overrides applied.
"""

import json
import uuid
from datetime import date, timedelta
from pathlib import Path
from typing import Optional

from . import paths

# Resolved on every call, never snapshotted at import — the calendar folder is
# user-configurable and can change while the app is running.

def _calendar_file(name: str, church_id: str | None = None) -> Path:
    """Where one calendar file lives, for this church.

    ⛔⛔ A CHURCH READS ONLY ITS OWN, AND THERE IS NO FALLBACK TO THE SHARED FILE.
    Falling back would hand a second congregation the FIRST one's calendar, which
    is the exact defect this split exists to remove — and it would do it silently,
    on a printed bulletin. A church with no calendar yet has an EMPTY calendar,
    which is true and obvious.
    @decision:gold 2026-09-24

    ⭐ The flat file is still the answer when no church is given: that is a fresh
    install with nothing configured, the same case `churches.GENERIC` exists for.
    """
    if church_id and church_id != "default":
        from .churches import church_dir
        return church_dir(church_id) / name
    return paths.DATA_DIR / name


def _templates_file(church_id: str | None = None) -> Path:
    return _calendar_file("calendar-templates.json", church_id)


def _overrides_file(church_id: str | None = None) -> Path:
    return _calendar_file("calendar-overrides.json", church_id)


def _events_file(church_id: str | None = None) -> Path:
    return _calendar_file("calendar-events.json", church_id)


def _notes_file(church_id: str | None = None) -> Path:
    return _calendar_file("calendar-notes.json", church_id)

# How many weeks ahead to show in the bulletin
WEEKS_AHEAD = 4


# --- File I/O ---

def _load(path: Path) -> list:
    if not path.exists():
        return []
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def _save(path: Path, data: list):
    paths.DATA_DIR.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


# --- Templates (recurring events) ---

def list_templates(church_id: str | None = None) -> list[dict]:
    return _load(_templates_file(church_id))


def save_templates(templates: list[dict], church_id: str | None = None):
    _save(_templates_file(church_id), templates)


def add_template(template: dict, church_id: str | None = None) -> dict:
    """Add a new recurring event template. Returns the saved template."""
    templates = list_templates(church_id)
    if not template.get("id"):
        template["id"] = str(uuid.uuid4())
    templates.append(template)
    save_templates(templates, church_id)
    return template


def update_template(template_id: str, updates: dict, church_id: str | None = None) -> Optional[dict]:
    templates = list_templates(church_id)
    for t in templates:
        if t["id"] == template_id:
            t.update(updates)
            save_templates(templates, church_id)
            return t
    return None


def delete_template(template_id: str, church_id: str | None = None) -> bool:
    templates = list_templates(church_id)
    new_templates = [t for t in templates if t["id"] != template_id]
    if len(new_templates) < len(templates):
        save_templates(new_templates, church_id)
        # Also clean up any overrides for this template
        overrides = list_overrides(church_id)
        new_overrides = [o for o in overrides if o.get("templateId") != template_id]
        if len(new_overrides) < len(overrides):
            _save(_overrides_file(church_id), new_overrides)
        return True
    return False


# --- Overrides (per-instance skips and modifications) ---

def list_overrides(church_id: str | None = None) -> list[dict]:
    return _load(_overrides_file(church_id))


def set_override(template_id: str, date_str: str, skip: bool = False,
                  time: str = None, location: str = None,
                  church_id: str | None = None) -> dict:
    """Set or remove an override for a specific recurring event instance."""
    overrides = list_overrides(church_id)
    # Remove any existing override for this template/date combo
    overrides = [o for o in overrides
                  if not (o.get("templateId") == template_id and o.get("date") == date_str)]

    # If anything is non-default, add a new override
    if skip or time or location:
        override = {"templateId": template_id, "date": date_str}
        if skip:
            override["skip"] = True
        if time:
            override["time"] = time
        if location:
            override["location"] = location
        overrides.append(override)

    _save(_overrides_file(church_id), overrides)
    return {"success": True}


def toggle_skip(template_id: str, date_str: str, skip: bool,
                church_id: str | None = None) -> dict:
    """Toggle whether a specific instance is skipped."""
    overrides = list_overrides(church_id)

    # Find existing override
    existing = None
    for o in overrides:
        if o.get("templateId") == template_id and o.get("date") == date_str:
            existing = o
            break

    if skip:
        if existing:
            existing["skip"] = True
        else:
            overrides.append({"templateId": template_id, "date": date_str, "skip": True})
    else:
        # Remove skip flag
        if existing:
            existing.pop("skip", None)
            # If override is now empty (no time/location), remove it entirely
            if "time" not in existing and "location" not in existing:
                overrides = [o for o in overrides if o is not existing]

    _save(_overrides_file(church_id), overrides)
    return {"success": True}


# --- One-off events ---

def list_events(church_id: str | None = None) -> list[dict]:
    return _load(_events_file(church_id))


def add_event(event: dict, church_id: str | None = None) -> dict:
    events = list_events(church_id)
    if not event.get("id"):
        event["id"] = str(uuid.uuid4())
    events.append(event)
    _save(_events_file(church_id), events)
    return event


def update_event(event_id: str, updates: dict, church_id: str | None = None) -> Optional[dict]:
    events = list_events(church_id)
    for e in events:
        if e["id"] == event_id:
            e.update(updates)
            _save(_events_file(church_id), events)
            return e
    return None


def delete_event(event_id: str, church_id: str | None = None) -> bool:
    events = list_events(church_id)
    new_events = [e for e in events if e["id"] != event_id]
    if len(new_events) < len(events):
        _save(_events_file(church_id), new_events)
        return True
    return False


# --- Notes (per-service-date) ---

def _load_notes(church_id: str | None = None) -> dict:
    f_ = _notes_file(church_id)
    if not f_.exists():
        return {}
    try:
        with open(f_, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_notes(notes: dict, church_id: str | None = None):
    f_ = _notes_file(church_id)
    f_.parent.mkdir(parents=True, exist_ok=True)
    with open(f_, "w", encoding="utf-8") as f:
        json.dump(notes, f, indent=2, ensure_ascii=False)


def get_note(service_date: str, church_id: str | None = None) -> dict:
    notes = _load_notes(church_id)
    return notes.get(service_date, {
        "text": "",
        "enabled": False,
        "bold": False,
        "italic": False,
        "align": "left",
    })


def save_note(service_date: str, note: dict, church_id: str | None = None) -> dict:
    notes = _load_notes(church_id)
    align = note.get("align", "left")
    if align not in ("left", "center", "right"):
        align = "left"
    notes[service_date] = {
        "text": note.get("text", ""),
        "enabled": bool(note.get("enabled", False)),
        "bold": bool(note.get("bold", False)),
        "italic": bool(note.get("italic", False)),
        "align": align,
    }
    _save_notes(notes, church_id)
    return notes[service_date]


# --- Generation ---

def _compute_instances(template: dict, start: date, end: date) -> list[dict]:
    """Compute all instances of a recurring template within [start, end] inclusive."""
    instances = []
    day_of_week = template.get("dayOfWeek", 0)  # 0=Sunday in our convention
    # Python's weekday(): Monday=0, Sunday=6
    # Our convention: Sunday=0, Monday=1, ... Saturday=6
    py_dow = (day_of_week - 1) % 7  # convert to Python's Mon=0 convention

    # Find first occurrence of this day-of-week on or after start
    days_ahead = (py_dow - start.weekday()) % 7
    first = start + timedelta(days=days_ahead)

    cur = first
    while cur <= end:
        instances.append({
            "date": cur.isoformat(),
            "time": template.get("time", ""),
            "title": template.get("name", ""),
            "location": template.get("location", ""),
            "templateId": template["id"],
            "isRecurring": True,
        })
        cur += timedelta(days=7)

    return instances


def get_calendar_for_service(service_date_str: str, church_id: str | None = None) -> dict:
    """
    Get all events for the 4 weeks following the service date.

    Returns:
        {
          "start_date": "2026-04-12",
          "end_date": "2026-05-10",
          "events": [
            {"date": "2026-04-12", "time": "9:30 AM", "title": "Sunday School", ...},
            ...
          ]
        }
    """
    service_date = date.fromisoformat(service_date_str)
    end_date = service_date + timedelta(days=WEEKS_AHEAD * 7 - 1)

    templates = list_templates(church_id)
    overrides = list_overrides(church_id)
    one_offs = list_events(church_id)

    # Build override lookup: {(templateId, date): override}
    override_map = {}
    for o in overrides:
        key = (o.get("templateId"), o.get("date"))
        override_map[key] = o

    all_events = []

    # Expand recurring templates
    for template in templates:
        instances = _compute_instances(template, service_date, end_date)
        for inst in instances:
            override = override_map.get((template["id"], inst["date"]))
            if override:
                if override.get("skip"):
                    continue  # skip this instance
                if "time" in override:
                    inst["time"] = override["time"]
                if "location" in override:
                    inst["location"] = override["location"]
            all_events.append(inst)

    # Add one-offs in date range
    for event in one_offs:
        try:
            ev_date = date.fromisoformat(event["date"])
        except (ValueError, KeyError):
            continue
        if service_date <= ev_date <= end_date:
            all_events.append({
                "date": event["date"],
                "time": event.get("time", ""),
                "title": event.get("title", ""),
                "location": event.get("location", ""),
                "description": event.get("description", ""),
                "isRecurring": False,
                "id": event.get("id"),
            })

    # Sort by date, then time (string sort works for HH:MM AM/PM if times are well-formed)
    def time_sort_key(t: str) -> str:
        """Convert '9:30 AM' to a sortable key like '09:30 AM' -> 'AM 09:30'."""
        if not t:
            return "ZZZ"
        parts = t.strip().split()
        if len(parts) == 2:
            time_part, ampm = parts
            hh, mm = (time_part.split(":") + ["00"])[:2]
            return f"{ampm} {hh.zfill(2)}:{mm.zfill(2)}"
        return t

    all_events.sort(key=lambda e: (e["date"], time_sort_key(e["time"])))

    return {
        "start_date": service_date.isoformat(),
        "end_date": end_date.isoformat(),
        "weeks": WEEKS_AHEAD,
        "events": all_events,
    }


def get_upcoming_dates(service_date_str: str, day_of_week: int) -> list[str]:
    """Get the next 4 dates that fall on the given day-of-week starting from service_date.

    day_of_week: 0=Sunday, 1=Monday, ..., 6=Saturday (our convention)
    """
    service_date = date.fromisoformat(service_date_str)
    py_dow = (day_of_week - 1) % 7  # convert to Python's Mon=0
    days_ahead = (py_dow - service_date.weekday()) % 7
    first = service_date + timedelta(days=days_ahead)

    return [(first + timedelta(days=7 * i)).isoformat() for i in range(WEEKS_AHEAD)]
