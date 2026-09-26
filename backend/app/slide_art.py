"""Composite a section slide's title onto its photograph, outside PowerPoint.

WHY OUTSIDE: the titles used to be pixels baked into each JPEG, so changing a word
meant re-cutting artwork at the right size in the right font, per church. Drawing
them here means a photograph is just a photograph — drop in a new one and the
title still renders, in that church's font and colour.

⭐ AND WHY NOT IN POWERPOINT: a shape over a picture in pptx can be semi-transparent
but cannot BLUR what is behind it, so there is no real frosted glass there. Doing it
with Pillow gets a true gaussian blur, proper rounded corners and honest typography,
and the slide still receives exactly one full-bleed picture — the pptx side gets
simpler, not harder.

⛔ THE STYLES BELOW ARE THE WHOLE VOCABULARY. A bench renders every one of them
through THIS module, so what is chosen on screen is what a deck contains — never a
CSS impression of it.

CALLED BY: app.slides (generation), app.main (the workshop's preview endpoint).
"""

from __future__ import annotations

import glob
import os
from functools import lru_cache
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont

# Slides are 16:9 at 1920x1080; art is composited at that size and inserted whole.
CANVAS = (1920, 1080)

# ⛔ SLOTS WHOSE ART IS NOT 16:9. `hymn_background` is placed as a 7.5in SQUARE
# (slides.BG_IMG_SIZE), so it must never be reframed to the canvas. ⭐ It lives here
# rather than at each call site because BOTH generation and the bench have to agree
# — a bench that framed it differently would be reviewing something a deck cannot
# contain, which is the one thing a bench may never do.
SQUARE_SLOTS = {"hymn_background"}


def canvas_for(slot: str | None) -> tuple[int, int] | None:
    """The canvas a slot's art is washed against — None to keep its own shape."""
    return None if slot in SQUARE_SLOTS else CANVAS

# Where a font NAME becomes a font FILE. macOS first, then Windows, then Linux.
_FONT_DIRS = (
    "/System/Library/Fonts/Supplemental", "/Library/Fonts",
    os.path.expanduser("~/Library/Fonts"), "/System/Library/Fonts",
    "C:/Windows/Fonts",
    "/usr/share/fonts", "/usr/local/share/fonts",
)


@lru_cache(maxsize=32)
def _font_file(name: str, bold: bool) -> str | None:
    """Find a usable file for a font name. ⛔ Returns None rather than raising —
    a missing font must not cost a Sunday's slides; the caller falls back."""
    wanted = name.lower().replace(" ", "")
    best = None
    for d in _FONT_DIRS:
        for path in glob.glob(os.path.join(d, "*")):
            stem = os.path.basename(path).lower()
            if not stem.endswith((".ttf", ".otf", ".ttc")):
                continue
            flat = stem.rsplit(".", 1)[0].replace(" ", "").replace("-", "")
            if not flat.startswith(wanted):
                continue
            is_bold = "bold" in flat
            if is_bold == bold:
                return path
            best = best or path
    return best


def _load_font(name: str, size: int, bold: bool = True) -> ImageFont.FreeTypeFont:
    f = _font_file(name, bold) or _font_file("Georgia", bold) or _font_file("Arial", bold)
    if f:
        try:
            return ImageFont.truetype(f, size)
        except OSError:
            pass
    return ImageFont.load_default(size)


def _flatten(base: Image.Image) -> Image.Image:
    """RGB, with any transparency laid over WHITE.

    ⛔ `.convert("RGB")` ON ITS OWN DROPS ALPHA ONTO BLACK, and black is the wrong
    answer everywhere in this product: the slides are white and a backdrop is a
    watermark sitting under black text. A transparent emblem arrived as a BLACK
    field, so the wash had to be run to 100 just to climb back up to the white it
    should have started from — at which point it is not a knob, it is a repair.
    """
    if base.mode in ("RGBA", "LA", "P"):
        base = base.convert("RGBA")
        return Image.alpha_composite(
            Image.new("RGBA", base.size, (255, 255, 255, 255)), base).convert("RGB")
    return base.convert("RGB")


def _fit(base: Image.Image, canvas: tuple[int, int] = CANVAS) -> Image.Image:
    """Cover the canvas without distorting — crop the overflow, never squash."""
    base = _flatten(base)
    cw, ch = canvas
    scale = max(cw / base.width, ch / base.height)
    resized = base.resize((round(base.width * scale), round(base.height * scale)), Image.LANCZOS)
    left, top = (resized.width - cw) // 2, (resized.height - ch) // 2
    return resized.crop((left, top, left + cw, top + ch))


def _rounded_mask(size, radius: int) -> Image.Image:
    m = Image.new("L", size, 0)
    ImageDraw.Draw(m).rounded_rectangle([0, 0, size[0] - 1, size[1] - 1], radius=radius, fill=255)
    return m


# Each style is (label, description). The renderer branches on the key.
STYLES: dict[str, str] = {
    "frosted_bar":    "Full-width frosted band, blurred photo behind",
    "frosted_plate":  "Centred frosted plate, rounded",
    "dark_plate":     "Solid dark plate, rounded",
    "light_plate":    "Solid light plate, dark text",
    "themed_plate":   "Plate in the church's title colour",
    "bottom_band":    "Frosted band across the lower third",
    "scrim":          "Soft dark gradient, no plate edge",
    "plain":          "No plate — text with a drop shadow only",
    # Three attacks on "clearer". The original frosted plate darkens AND blurs
    # heavily, which reads as a smoked panel; these each give the picture back.
    "glass_light":    "Light frosted plate, dark text — glass rather than smoke",
    "glass_sharp":    "Barely blurred, light veil — the picture stays legible",
    "glass_edge":     "No veil at all — blur only, with a crisp outlined title",
}


def wash(photo: Path, amount: int,
         canvas: tuple[int, int] | None = CANVAS) -> Image.Image:
    """A backdrop washed toward white (positive) or black (negative), -100..100.

    ⭐ BACKDROPS CARRY BODY TEXT OVER THEM — the creed, the Lord's Prayer — so what
    matters is not how the picture looks alone but whether words read on it. The
    liturgy is set in BLACK, so washing the picture OUT is what buys contrast;
    darkening fights it. Both directions exist because a congregation could set
    light type one day, and the wrong half of the range would then be missing.

    ⛔⛔ `canvas=None` KEEPS THE SOURCE'S OWN SHAPE, AND THE SQUARE SLOT NEEDS IT.
    A wash is a COLOUR operation; reframing it to 16:9 silently wrecks any slot
    that is not 16:9. `hymn_background` is a 7.5in SQUARE, so covering the
    1920x1080 canvas cropped the top and bottom off the emblem and pptx then
    stretched the leftover strip back out into a square box — cropped AND
    distorted, and only ever when a wash was applied, so an unwashed church
    never saw it.
    """
    img = (_fit(Image.open(photo), canvas) if canvas
           else _flatten(Image.open(photo))).convert("RGBA")
    amount = max(-100, min(100, int(amount)))
    a = round(255 * abs(amount) / 100)
    if a:
        veil = (255, 255, 255, a) if amount > 0 else (0, 0, 0, a)
        img = Image.alpha_composite(img, Image.new("RGBA", img.size, veil))
    return img.convert("RGB")


def render(photo: Path, text: str, theme: dict, style: str = "frosted_bar",
           position: str | None = None) -> Image.Image:
    """Composite `text` onto `photo` in `style`, using the church's theme.

    `position` overrides the theme's default placement for this one slide, so a
    church can move a single title off whatever its artwork happens to put there.
    """
    img = _fit(Image.open(photo)).convert("RGBA")
    W, H = CANVAS
    font_name = theme.get("font") or "Georgia"

    rgb = lambda c: (c[0], c[1], c[2]) if not isinstance(c, str) else (0, 0, 0)
    title_rgb = rgb(theme.get("title_color", (0, 0, 0)))

    size = 132
    font = _load_font(font_name, size, bold=True)
    d0 = ImageDraw.Draw(img)
    while d0.textlength(text, font=font) > W * 0.78 and size > 48:
        size -= 6
        font = _load_font(font_name, size, bold=True)
    tw = d0.textlength(text, font=font)
    asc, desc = font.getmetrics()
    th = asc + desc

    pad_x, pad_y = int(size * 0.62), int(size * 0.40)
    plate_w = int(tw + pad_x * 2)
    plate_h = int(th + pad_y * 2)
    full_w = style in ("frosted_bar", "bottom_band")

    # ⭐ WHERE the plate sits is independent of WHAT it looks like, so it is its
    # own setting rather than a second copy of every style. Margin is a share of
    # the canvas, so it holds if the canvas size ever changes.
    margin = int(W * 0.045)
    pos = position or theme.get("plate_position", "bottom-right")
    if full_w:
        cx = W // 2                       # a band spans the slide; only y matters
    elif "left" in pos:
        cx = margin + plate_w // 2
    elif "right" in pos:
        cx = W - margin - plate_w // 2
    else:
        cx = W // 2
    # ⭐ A TOP PLATE SITS TWICE THE SIDE MARGIN DOWN, not flush at it. Pinned at
    # the same inset it reads as stuck to the edge; dropping it the same distance
    # again lets the title float over the picture. @decision:gold 2026-09-24
    if "bottom" in pos:
        cy = H - margin * 2 - plate_h // 2
    elif "top" in pos:
        cy = margin * 2 + plate_h // 2
    else:
        cy = H // 2
    if style == "bottom_band":
        cy = int(H * 0.78)

    box = [0 if full_w else cx - plate_w // 2, cy - plate_h // 2,
           W if full_w else cx + plate_w // 2, cy + plate_h // 2]
    radius = 0 if full_w else 28

    fg = (255, 255, 255, 255)

    # How much the plate darkens what it covers, 0-100. Lower lets more picture
    # through; the blur still separates the title from the detail behind it.
    veil = round(255 * max(0, min(100, theme.get("plate_opacity", 47))) / 100)
    # ⭐ BLUR IS THE KNOB THAT DECIDES WHETHER THE PLATE READS AS GLASS OR AS MUSH.
    # Heavy blur destroys the picture under the title, which reads as "not clear"
    # no matter how the tint is set. None here means the style's own default.
    blur_override = theme.get("plate_blur")

    outline = None

    if style in ("frosted_bar", "frosted_plate", "bottom_band"):
        b = 28 if blur_override is None else max(0, min(60, int(blur_override)))
        region = img.crop(tuple(box))
        if b: region = region.filter(ImageFilter.GaussianBlur(b))
        tint = Image.new("RGBA", region.size, (10, 12, 18, veil))
        region = Image.alpha_composite(region.convert("RGBA"), tint)
        img.paste(region, (box[0], box[1]), _rounded_mask(region.size, radius) if radius else None)
    elif style in ("glass_light", "glass_sharp", "glass_edge"):
        # blur radius and veil colour are what separate these three
        blur = {"glass_light": 22, "glass_sharp": 6, "glass_edge": 14}[style]
        if blur_override is not None:
            blur = max(0, min(60, int(blur_override)))
        region = img.crop(tuple(box))
        if blur: region = region.filter(ImageFilter.GaussianBlur(blur))
        region = region.convert("RGBA")
        if style == "glass_light":
            region = Image.alpha_composite(region, Image.new("RGBA", region.size, (255, 255, 255, veil + 40)))
            fg = (20, 22, 28, 255)
        elif style == "glass_sharp":
            region = Image.alpha_composite(region, Image.new("RGBA", region.size, (255, 255, 255, max(30, veil - 12))))
            fg = (20, 22, 28, 255)
        else:  # glass_edge — no veil, the blur alone separates it
            outline = (255, 255, 255, 230)
            fg = (16, 18, 24, 255)
        img.paste(region, (box[0], box[1]), _rounded_mask(region.size, radius))
    elif style in ("dark_plate", "light_plate", "themed_plate"):
        solid = min(255, round(veil * 1.25))   # a solid plate needs a touch more
        fill = {"dark_plate": (10, 12, 18, solid),
                "light_plate": (255, 255, 255, min(255, solid + 60)),
                "themed_plate": title_rgb + (min(255, solid + 60),)}[style]
        plate = Image.new("RGBA", (box[2] - box[0], box[3] - box[1]), fill)
        img.paste(plate, (box[0], box[1]), _rounded_mask(plate.size, radius))
        if style == "light_plate":
            fg = title_rgb + (255,)
    elif style == "scrim":
        grad = Image.new("L", (1, H))
        for y in range(H):
            t = abs(y - cy) / (H * 0.42)
            grad.putpixel((0, y), int(max(0, 165 * (1 - min(t, 1) ** 2))))
        veil = Image.new("RGBA", (W, H), (8, 10, 16, 255))
        veil.putalpha(grad.resize((W, H)))
        img = Image.alpha_composite(img, veil)

    layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    tx, ty = cx - tw / 2, cy - th / 2
    if outline:
        # a crisp ring rather than a soft shadow — this is the "clearer" one
        d.text((tx, ty), text, font=font, fill=fg, stroke_width=max(2, size // 36), stroke_fill=outline)
    elif style in ("plain", "scrim") or theme.get("shadow_enabled", True):
        d.text((tx + 3, ty + 4), text, font=font, fill=(0, 0, 0, 120))
    d.text((tx, ty), text, font=font, fill=fg)
    return Image.alpha_composite(img, layer).convert("RGB")
