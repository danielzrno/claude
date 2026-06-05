#!/usr/bin/env python3
"""
Generate the small embedded PNG assets used by the Office artifacts:
  - mchip.png        coral rounded-square + white Inter-800 "M"  (for docx + pptx)
  - mchip_white.png  white rounded-square + coral "M"            (reverse, for coral fields)
  - hero_coral.png   coral hero field with a faint darker C9362B diagonal wave (pptx/docx cover)
  - photo_duotone.png / photo_neutral.png   rounded "photo" placeholder blocks

High-res, then placed small — keeps edges crisp in PowerPoint/Word.
Pillow only. No network images.
"""
from PIL import Image, ImageDraw, ImageFont, ImageFilter
import os

HERE = os.path.dirname(os.path.abspath(__file__))
INTER = "/root/.fonts/Inter.ttf"

CORAL      = (234, 73, 63)     # #EA493F
CORAL_DEEP = (201, 54, 43)     # #C9362B
CORAL_DARK = (168, 44, 34)     # deeper shade for wave tail
BLACK      = (0, 0, 0)
WHITE      = (255, 255, 255)


def inter(size, weight="ExtraBold"):
    f = ImageFont.truetype(INTER, size)
    try:
        f.set_variation_by_name(weight)
    except Exception:
        pass
    return f


def rounded_chip(path, bg, fg, px=512, radius_ratio=0.205):
    """A rounded-square tile with a centred 'M'. Supersampled for clean edges."""
    SS = 2
    S = px * SS
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    r = int(S * radius_ratio)
    d.rounded_rectangle([0, 0, S - 1, S - 1], radius=r, fill=bg)
    # the M — optically centred, weight 800, tight
    f = inter(int(S * 0.60), "ExtraBold")
    txt = "M"
    bb = d.textbbox((0, 0), txt, font=f)
    tw, th = bb[2] - bb[0], bb[3] - bb[1]
    x = (S - tw) / 2 - bb[0]
    y = (S - th) / 2 - bb[1] - int(S * 0.012)
    d.text((x, y), txt, font=f, fill=fg)
    img = img.resize((px, px), Image.LANCZOS)
    img.save(path)


def hero_field(path, w=2600, h=1500, base=CORAL, deep=CORAL_DEEP, tail=CORAL_DARK,
               highlight=True):
    """Coral field with a faint darker diagonal wave (lower-left -> up), a soft
    top highlight, and a low contour sweep. Subtle and premium, never a hard stripe."""
    img = Image.new("RGB", (w, h), base)

    # --- diagonal darker sweep from lower-left, built as a big soft radial blob ---
    blob = Image.new("L", (w, h), 0)
    bd = ImageDraw.Draw(blob)
    # ellipse anchored off the bottom-left corner
    bd.ellipse([-int(w * 0.55), int(h * 0.30), int(w * 0.75), int(h * 1.85)], fill=255)
    blob = blob.filter(ImageFilter.GaussianBlur(int(w * 0.16)))
    deep_layer = Image.new("RGB", (w, h), deep)
    img = Image.composite(deep_layer, img, blob.point(lambda v: int(v * 0.92)))

    # --- a deeper contour tail at the very bottom-left for a 'wave' read ---
    tailmask = Image.new("L", (w, h), 0)
    td = ImageDraw.Draw(tailmask)
    td.ellipse([-int(w * 0.40), int(h * 0.78), int(w * 0.40), int(h * 1.7)], fill=255)
    tailmask = tailmask.filter(ImageFilter.GaussianBlur(int(w * 0.10)))
    tail_layer = Image.new("RGB", (w, h), tail)
    img = Image.composite(tail_layer, img, tailmask.point(lambda v: int(v * 0.55)))

    # --- faint top-right highlight for depth ---
    if highlight:
        hl = Image.new("L", (w, h), 0)
        hd = ImageDraw.Draw(hl)
        hd.ellipse([int(w * 0.55), -int(h * 0.55), int(w * 1.45), int(h * 0.65)], fill=255)
        hl = hl.filter(ImageFilter.GaussianBlur(int(w * 0.14)))
        white_layer = Image.new("RGB", (w, h), WHITE)
        img = Image.composite(white_layer, img, hl.point(lambda v: int(v * 0.10)))

    img.save(path, "PNG")


def duotone_block(path, w=1400, h=1700, base=CORAL, deep=CORAL_DEEP):
    """A coral-duotone 'photo' placeholder block (vertical)."""
    img = Image.new("RGB", (w, h), base)
    grad = Image.new("L", (w, h), 0)
    gd = ImageDraw.Draw(grad)
    for y in range(h):
        gd.line([(0, y), (w, y)], fill=int(255 * (y / h)))
    deep_layer = Image.new("RGB", (w, h), deep)
    img = Image.composite(deep_layer, img, grad.point(lambda v: int(v * 0.85)))
    # soft top highlight
    hl = Image.new("L", (w, h), 0)
    hd = ImageDraw.Draw(hl)
    hd.ellipse([int(w * 0.35), -int(h * 0.35), int(w * 1.3), int(h * 0.45)], fill=255)
    hl = hl.filter(ImageFilter.GaussianBlur(int(w * 0.16)))
    white_layer = Image.new("RGB", (w, h), WHITE)
    img = Image.composite(white_layer, img, hl.point(lambda v: int(v * 0.14)))
    img.save(path, "PNG")


def neutral_block(path, w=1400, h=1700):
    """A warm-neutral 'studio' photo placeholder block."""
    top = (237, 238, 232)
    bot = (208, 209, 201)
    img = Image.new("RGB", (w, h), top)
    grad = Image.new("L", (w, h), 0)
    gd = ImageDraw.Draw(grad)
    for y in range(h):
        gd.line([(0, y), (w, y)], fill=int(255 * (y / h)))
    bot_layer = Image.new("RGB", (w, h), bot)
    img = Image.composite(bot_layer, img, grad)
    hl = Image.new("L", (w, h), 0)
    hd = ImageDraw.Draw(hl)
    hd.ellipse([int(w * 0.4), -int(h * 0.4), int(w * 1.3), int(h * 0.4)], fill=255)
    hl = hl.filter(ImageFilter.GaussianBlur(int(w * 0.14)))
    white_layer = Image.new("RGB", (w, h), WHITE)
    img = Image.composite(white_layer, img, hl.point(lambda v: int(v * 0.5)))
    img.save(path, "PNG")


def build_all():
    rounded_chip(os.path.join(HERE, "mchip.png"), CORAL, WHITE)
    rounded_chip(os.path.join(HERE, "mchip_white.png"), WHITE, CORAL)
    hero_field(os.path.join(HERE, "hero_coral.png"))
    duotone_block(os.path.join(HERE, "photo_duotone.png"))
    neutral_block(os.path.join(HERE, "photo_neutral.png"))
    print("assets written to", HERE)


if __name__ == "__main__":
    build_all()
