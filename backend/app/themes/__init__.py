"""Theme system for OpenOrder slide generation."""

from .oakpark import THEME as OAKPARK_THEME

# Registry of available themes
THEMES = {
    "oakpark": OAKPARK_THEME,
}

# Default theme
DEFAULT_THEME = "oakpark"


def get_theme(name: str = None) -> dict:
    """Get a theme by name. Returns default if name is None or not found."""
    if name is None:
        name = DEFAULT_THEME
    return THEMES.get(name.lower(), THEMES[DEFAULT_THEME])


def list_themes() -> list[dict]:
    """Return list of available themes with metadata."""
    return [
        {"id": tid, "name": t["name"], "description": t["description"]}
        for tid, t in THEMES.items()
    ]
