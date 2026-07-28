"""
nav_icons.py

Draws glossy, 3D-looking circular icon buttons (Back / Forward / Refresh)
with Pillow -- the "shiny glass sphere with an icon on it" look, instead of
a flat tkinter/CTk button. Each icon is rendered once at 4x resolution and
downsized with LANCZOS so the edges, gradient and highlight all stay smooth.

Usage from dashboard.py:

    from ui.nav_icons import build_nav_button_images
    images = build_nav_button_images(size=42)   # {"back": {...}, "forward": {...}, "refresh": {...}}
    normal_img  = images["back"]["normal"]       # a customtkinter.CTkImage
    hover_img   = images["back"]["hover"]
    disabled_img = images["back"]["disabled"]
"""

import math
from PIL import Image, ImageDraw, ImageFilter
import customtkinter as ctk

# Brand-blue glossy theme, tuned to sit nicely on the app's #1E3A8A header.
_PALETTES = {
    "normal":   {"top": (79, 124, 209), "bottom": (28, 62, 128), "rim": (14, 34, 74)},
    "hover":    {"top": (103, 148, 232), "bottom": (37, 82, 163), "rim": (18, 44, 92)},
    "disabled": {"top": (140, 148, 168), "bottom": (96, 104, 122), "rim": (70, 76, 90)},
}

_SUPERSAMPLE = 4


def _glossy_base(size, palette):
    """Renders one glossy sphere-style circle (gradient fill + shine + rim)."""
    big = size * _SUPERSAMPLE
    img = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    top, bottom = palette["top"], palette["bottom"]
    for y in range(big):
        t = y / max(1, big - 1)
        r = int(top[0] + (bottom[0] - top[0]) * t)
        g = int(top[1] + (bottom[1] - top[1]) * t)
        b = int(top[2] + (bottom[2] - top[2]) * t)
        draw.line([(0, y), (big, y)], fill=(r, g, b, 255))

    mask = Image.new("L", (big, big), 0)
    ImageDraw.Draw(mask).ellipse([big * 0.02, big * 0.02, big * 0.98, big * 0.98], fill=255)
    img.putalpha(mask)

    rim = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    ImageDraw.Draw(rim).ellipse(
        [big * 0.02, big * 0.02, big * 0.98, big * 0.98],
        outline=palette["rim"] + (255,), width=max(2, big // 40)
    )
    img = Image.alpha_composite(img, rim)

    shine = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    ImageDraw.Draw(shine).ellipse(
        [big * 0.16, big * 0.06, big * 0.84, big * 0.52], fill=(255, 255, 255, 95)
    )
    shine = shine.filter(ImageFilter.GaussianBlur(big * 0.045))
    img = Image.alpha_composite(img, shine)

    return img, big


def _draw_chevron(draw, cx, cy, r, direction):
    """A solid triangular arrow pointing left or right."""
    if direction == "left":
        pts = [(cx - r, cy), (cx + r * 0.55, cy - r * 0.85), (cx + r * 0.55, cy + r * 0.85)]
    else:
        pts = [(cx + r, cy), (cx - r * 0.55, cy - r * 0.85), (cx - r * 0.55, cy + r * 0.85)]
    draw.polygon(pts, fill=(255, 255, 255, 235))


def _draw_refresh(draw, cx, cy, r, width):
    """A circular 'reload' arrow -- an open ring with an arrowhead at one end."""
    bbox = [cx - r, cy - r, cx + r, cy + r]
    start_angle, end_angle = -35, 250
    draw.arc(bbox, start=start_angle, end=end_angle, fill=(255, 255, 255, 235), width=width)

    theta = math.radians(end_angle)
    ex, ey = cx + r * math.cos(theta), cy + r * math.sin(theta)
    tangent = theta + math.radians(100)
    head_len = width * 1.9
    perp = tangent + math.radians(90)
    p1 = (ex + head_len * math.cos(tangent), ey + head_len * math.sin(tangent))
    p2 = (ex + width * 1.15 * math.cos(perp), ey + width * 1.15 * math.sin(perp))
    p3 = (ex - width * 1.15 * math.cos(perp), ey - width * 1.15 * math.sin(perp))
    draw.polygon([p1, p2, p3], fill=(255, 255, 255, 235))


def _render_icon(size, palette, kind):
    img, big = _glossy_base(size, palette)
    draw = ImageDraw.Draw(img)
    cx = cy = big / 2
    r = big * 0.24

    if kind == "back":
        _draw_chevron(draw, cx, cy, r, "left")
    elif kind == "forward":
        _draw_chevron(draw, cx, cy, r, "right")
    elif kind == "refresh":
        _draw_refresh(draw, cx, cy, r, width=max(3, int(big * 0.075)))

    return img.resize((size, size), Image.LANCZOS)


def build_nav_button_images(size=42):
    """
    Returns {"back": {"normal": CTkImage, "hover": CTkImage, "disabled": CTkImage},
             "forward": {...}, "refresh": {...}}
    ready to hand straight to a CTkButton's `image=` option.
    """
    out = {}
    for kind in ("back", "forward", "refresh"):
        out[kind] = {}
        for state, palette in _PALETTES.items():
            pil_img = _render_icon(size, palette, kind)
            out[kind][state] = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(size, size))
    return out