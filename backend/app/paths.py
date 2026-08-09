"""
Centralized path resolution for OpenOrder.

Handles both normal Python execution and PyInstaller frozen bundles.

Three INDEPENDENTLY configurable locations, each with a sensible default so a
fresh install works with no setup at all:

    output_dir   Generated bulletins/slides   default: ~/Downloads
    data_dir     Calendar files (read+write)  default: ~/Documents/OpenOrder
    hymnal_dir   Hymnal JSON (read-only)      default: beside the app

Why three and not one: they have genuinely different lifetimes. Output is
disposable and belongs where downloads go. Calendar data is the user's own
work — it must live somewhere visible and backup-able, and must survive a
reinstall. The hymnal is copyrighted source material the user supplies, often
already sitting in a shared folder.

Settings live in the OS's standard per-user config location, NOT beside the
app. Anything stored inside the app bundle is destroyed by the next install,
which silently takes the user's folder choices (and any data still in there)
with it.

Bundled read-only resources (frontend/dist, resources/) always resolve from
the bundle itself via _MEIPASS.
"""

import json
import os
import sys
from pathlib import Path


def _app_root() -> Path:
    """Where the executable or the project lives."""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    else:
        return Path(__file__).parent.parent.parent


def _bundle_root() -> Path:
    """Where packaged read-only resources live."""
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS)
    else:
        return Path(__file__).parent.parent.parent


def _config_dir() -> Path:
    """
    The OS's standard per-user config directory for this app.

    macOS:   ~/Library/Application Support/OpenOrder
    Windows: %APPDATA%\\OpenOrder
    Linux:   $XDG_CONFIG_HOME/OpenOrder  (or ~/.config/OpenOrder)
    """
    if sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    elif os.name == "nt":
        base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
    return base / "OpenOrder"


SETTINGS_FILE = _config_dir() / "settings.json"

# Where settings used to live (beside the app). Read once, for migration.
_LEGACY_SETTINGS_FILE = _app_root() / "settings.json"


# --- Defaults ---------------------------------------------------------------

def _default_output_dir() -> Path:
    """Generated files go where a browser would put them."""
    return Path.home() / "Downloads"


def _default_data_dir() -> Path:
    """The user's own calendar data — visible, and theirs to back up."""
    return Path.home() / "Documents" / "OpenOrder"


def _default_hymnal_dir() -> Path:
    """Ships beside the app; the user can point this at their own copy."""
    return _app_root() / "hymnal-json"


# --- Settings ---------------------------------------------------------------

def _load_settings() -> dict:
    """
    Load settings from the per-user config location, falling back once to the
    legacy in-app location so an existing install keeps its folder choices.
    """
    for path in (SETTINGS_FILE, _LEGACY_SETTINGS_FILE):
        if path.exists():
            try:
                with open(path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                continue
    return {}


def save_settings(settings: dict):
    """Write settings to the per-user config location."""
    SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2, ensure_ascii=False)


def get_settings() -> dict:
    """Current settings (for API use)."""
    return _load_settings()


def _resolve(key: str, default: Path, creatable: bool = True) -> Path:
    """
    A configured directory, or its default.

    Falls back to the default when the configured folder is unusable — a drive
    that isn't mounted, or a path carried over from another machine. (A real
    case: the Windows-era settings named `D:/Dropbox/...`, which on the Mac is
    simply nowhere, and without this check every save would fail.)
    """
    value = _load_settings().get(key)
    if value:
        p = Path(value).expanduser()
        if p.is_dir():
            return p
        if creatable:
            try:
                p.mkdir(parents=True, exist_ok=True)
                return p
            except Exception:
                pass
    return default


# --- The resolved locations -------------------------------------------------
#
# NOTE: these are module attributes that the setters below REBIND, so a change
# takes effect without restarting the app. Consumers must therefore read them
# late — `from . import paths` then `paths.OUTPUT_DIR`, never
# `from .paths import OUTPUT_DIR` at module import, which snapshots the old
# value and silently keeps writing to the previous folder.

RESOURCES_DIR = _bundle_root() / "resources"
FRONTEND_DIST_DIR = _bundle_root() / "frontend" / "dist"

OUTPUT_DIR = _resolve("output_dir", _default_output_dir())
DATA_DIR = _resolve("data_dir", _default_data_dir())
HYMNAL_DIR = _resolve("hymnal_dir", _default_hymnal_dir(), creatable=False)
SCRIPTURE_CACHE_DIR = DATA_DIR / "scripture-cache"


def _ensure_dirs():
    """Create the writable directories. The hymnal is read-only — never created."""
    for d in (OUTPUT_DIR, DATA_DIR, SCRIPTURE_CACHE_DIR):
        try:
            d.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass


_ensure_dirs()


# --- Changing a location ----------------------------------------------------

_KEYS = {
    "output_dir": lambda: _default_output_dir(),
    "data_dir": lambda: _default_data_dir(),
    "hymnal_dir": lambda: _default_hymnal_dir(),
}


def _recompute():
    """Re-resolve every location from current settings."""
    global OUTPUT_DIR, DATA_DIR, HYMNAL_DIR, SCRIPTURE_CACHE_DIR
    OUTPUT_DIR = _resolve("output_dir", _default_output_dir())
    DATA_DIR = _resolve("data_dir", _default_data_dir())
    HYMNAL_DIR = _resolve("hymnal_dir", _default_hymnal_dir(), creatable=False)
    SCRIPTURE_CACHE_DIR = DATA_DIR / "scripture-cache"
    _ensure_dirs()


def set_dir(key: str, new_dir: str):
    """Point one location at a new folder and persist it."""
    if key not in _KEYS:
        raise ValueError(f"Unknown directory key: {key}")
    settings = _load_settings()
    settings[key] = str(Path(new_dir).expanduser())
    save_settings(settings)
    _recompute()


def current_dirs() -> dict:
    """The resolved locations, for the settings UI."""
    return {
        "output_dir": str(OUTPUT_DIR),
        "data_dir": str(DATA_DIR),
        "hymnal_dir": str(HYMNAL_DIR),
        "settings_file": str(SETTINGS_FILE),
    }


# --- Back-compat ------------------------------------------------------------

def _data_root() -> Path:
    """Legacy accessor — the calendar/data location."""
    return DATA_DIR


def update_data_dir(new_dir: str):
    """Legacy entry point — now sets only the calendar/data location."""
    set_dir("data_dir", new_dir)
