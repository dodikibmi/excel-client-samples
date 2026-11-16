"""docs/before-after.png (the two dashboards side by side) and docs/thumbnail.png (1200x627, Upwork card): the old
dashboard on the left, the redesign on the right, split by a slanted line, proof numbers read from the files.

    python scripts/make_thumb.py [checks]        e.g. 37/37 from test/verify.py
"""
import sys
from pathlib import Path
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
F = "C:/Windows/Fonts/"
font = lambda n, s: ImageFont.truetype(F + n, s)
CHECKS = sys.argv[1] if len(sys.argv) > 1 else (ROOT / "test" / "last-result.txt").read_text().strip()
GREEN, CREAM, TEXT, MUTED, RED = (47, 107, 79), (247, 244, 236), (35, 48, 42), (107, 117, 110), (184, 58, 46)
before = Image.open(ROOT / "docs" / "before.png").convert("RGB")
after = Image.open(ROOT / "docs" / "after.png").convert("RGB")
br = load_workbook(ROOT / "Sales-Dashboard-Redesign.xlsx", data_only=True)["Bridge"]
old_total, new_total = br["C7"].value, br["C11"].value

# ---- side by side
h = 900
b = before.resize((int(before.width * h / before.height), h), Image.LANCZOS)
a = after.resize((int(after.width * h / after.height), h), Image.LANCZOS)
pad, top = 30, 90
img = Image.new("RGB", (b.width + a.width + pad * 3, h + top + pad), CREAM)
d = ImageDraw.Draw(img)
img.paste(b, (pad, top)); img.paste(a, (pad * 2 + b.width, top))
d.text((pad, 24), f"BEFORE: the client's file · total shown ${old_total:,.0f} (wrong)", font=font("segoeuib.ttf", 30), fill=RED)
d.text((pad * 2 + b.width, 24), f"AFTER: redesign on cleaned data · ${new_total:,.0f}", font=font("segoeuib.ttf", 30), fill=GREEN)
img.save(ROOT / "docs" / "before-after.png")
print("docs/before-after.png")

# ---- thumbnail
W, H = 1200, 627
t = Image.new("RGB", (W, H), CREAM)
bl = before.resize((int(before.width * 470 / before.height), 470), Image.LANCZOS)
al = after.resize((640, int(after.height * 640 / after.width)), Image.LANCZOS)
t.paste(bl.crop((0, 0, 700, 470)), (0, 157))
mask = Image.new("L", (W, H), 0)
ImageDraw.Draw(mask).polygon([(560, H), (700, 120), (W, 120), (W, H)], fill=255)
layer = Image.new("RGB", (W, H), CREAM); layer.paste(al, (580, 122))
t.paste(layer, (0, 0), mask)
d = ImageDraw.Draw(t)
d.line([(560, H), (700, 120)], fill=(255, 255, 255), width=10)
d.rectangle((0, 0, W, 120), fill=CREAM)
d.text((36, 20), "Excel dashboard redesign: before", font=font("georgia.ttf", 40), fill=TEXT)
x = 36 + d.textlength("Excel dashboard redesign: before ", font=font("georgia.ttf", 40))
d.text((x, 22), "→", font=font("segoeui.ttf", 40), fill=GREEN)
d.text((x + 52, 20), "after", font=font("georgia.ttf", 40), fill=TEXT)
d.text((38, 74), "A cleaner look, and 9 wrong numbers found and fixed on the way", font=font("segoeui.ttf", 22), fill=MUTED)


def chip(x, y, text, col):
    f = font("segoeuib.ttf", 21)
    w = d.textlength(text, font=f)
    d.rounded_rectangle((x, y, x + w + 28, y + 42), 21, fill=col)
    d.text((x + 14, y + 8), text, font=f, fill=(255, 255, 255))


chip(24, 520, f"BEFORE ${old_total:,.0f}", RED)
chip(24, 572, "9 number errors", RED)
chip(880, 520, f"AFTER ${new_total:,.0f}", GREEN)
chip(880, 572, f"{CHECKS} checks pass", GREEN)
t.save(ROOT / "docs" / "thumbnail.png")
print("docs/thumbnail.png")
