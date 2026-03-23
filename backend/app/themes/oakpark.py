"""Oak Park UMC slide theme."""

from pptx.dml.color import RGBColor

# Olive green palette
_OLIVE = RGBColor(0x5A, 0x6B, 0x2E)
_WHITE = RGBColor(0xFF, 0xFF, 0xFF)

THEME = {
    # Metadata
    "name": "Oak Park",
    "description": "Olive green theme for Oak Park UMC",

    # Typography
    "font": "Georgia",

    # Title pill
    "title_color": _OLIVE,

    # Source badge (bottom-right hymn number)
    "badge_bg": _OLIVE,
    "badge_fg": _WHITE,

    # Main text (lyrics, scripture) — use _WHITE for dark backgrounds
    "text_color": RGBColor(0x00, 0x00, 0x00),  # Black

    # Speaker labels in liturgy (Pastor:, People:, etc.)
    "speaker_label_color": _OLIVE,

    # Liturgy title casing
    "liturgy_uppercase": False,

    # Effects
    "shadow_enabled": True,
}
