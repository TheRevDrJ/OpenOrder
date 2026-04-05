"""
Centralized path resolution for OpenOrder.

Handles both normal Python execution and PyInstaller frozen bundles.
- Bundled resources (frontend/dist, resources/) resolve from _MEIPASS
- User data (output/, hymnal-json/, scripture-cache/) resolve from:
  1. settings.json "data_dir" if configured (for Dropbox/shared setups)
  2. The exe's directory (frozen) or project root (development)
"""

import json
import sys
from pathlib import Path


def _app_root() -> Path:
    """Get the app root — where the exe or project lives."""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    else:
        return Path(__file__).parent.parent.parent


def _bundle_root() -> Path:
    """Get the bundle root (where packaged resources live)."""
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS)
    else:
        return Path(__file__).parent.parent.parent


def _load_settings() -> dict:
    """Load settings.json from app root if it exists."""
    settings_path = _app_root() / "settings.json"
    if settings_path.exists():
        try:
            with open(settings_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_settings(settings: dict):
    """Save settings to settings.json in app root."""
    settings_path = _app_root() / "settings.json"
    with open(settings_path, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2, ensure_ascii=False)


def _data_root() -> Path:
    """Get the data directory — from settings or default to app root."""
    settings = _load_settings()
    data_dir = settings.get("data_dir")
    if data_dir:
        p = Path(data_dir)
        if p.exists():
            return p
    return _app_root()


def get_settings() -> dict:
    """Get current settings (for API use)."""
    return _load_settings()


# Bundled resources (read-only, packaged with the app)
RESOURCES_DIR = _bundle_root() / "resources"
FRONTEND_DIST_DIR = _bundle_root() / "frontend" / "dist"

# User data (read-write, from data_dir setting or app root)
OUTPUT_DIR = _data_root() / "output"
HYMNAL_DIR = _data_root() / "hymnal-json"
SCRIPTURE_CACHE_DIR = _data_root() / "scripture-cache"

# Ensure writable directories exist
OUTPUT_DIR.mkdir(exist_ok=True)
SCRIPTURE_CACHE_DIR.mkdir(exist_ok=True)


def update_data_dir(new_dir: str):
    """Update the data directory and recalculate all paths."""
    global OUTPUT_DIR, HYMNAL_DIR, SCRIPTURE_CACHE_DIR

    settings = _load_settings()
    settings["data_dir"] = new_dir
    save_settings(settings)

    root = Path(new_dir)
    OUTPUT_DIR = root / "output"
    HYMNAL_DIR = root / "hymnal-json"
    SCRIPTURE_CACHE_DIR = root / "scripture-cache"

    OUTPUT_DIR.mkdir(exist_ok=True)
    SCRIPTURE_CACHE_DIR.mkdir(exist_ok=True)
