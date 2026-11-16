"""docs/thumbnail.png (1200x627, Upwork card): title, Home, the refused booking form, proof chips.

    python scripts/make_thumb.py [checks]        e.g. 38/38 from test/verify.py
"""
import sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
F = "C:/Windows/Fonts/"
font = lambda n, s: ImageFont.truetype(F + n, s)
CHECKS = sys.argv[1] if len(sys.argv) > 1 else (ROOT / "test" / "last-result.txt").read_text().strip()
PLUM, PAGE, TEXT, MUTED, TERRA, GREEN = (91, 42, 78), (250, 247, 242), (43, 34, 48), (122, 110, 118), (200, 105, 63), (47, 138, 78)

W, H = 1200, 627
t = Image.new("RGB", (W, H), PAGE)
d = ImageDraw.Draw(t)
home = Image.open(ROOT / "docs" / "home.png").convert("RGB")
hw = 560
hl = home.resize((hw, int(home.height * hw / home.width)), Image.LANCZOS).crop((0, 0, hw, 470))
x0 = W - hw - 30
t.paste(hl, (x0, 140)); d.rectangle((x0 - 1, 139, x0 + hw, 140 + hl.height), outline=(220, 210, 198))
form = Image.open(ROOT / "docs" / "form-blocked.png").convert("RGB")
fw = 400
fl = form.resize((fw, int(form.height * fw / form.width)), Image.LANCZOS)
fx, fy = x0 - 90, 282
d.rectangle((fx - 4, fy - 4, fx + fw + 4, fy + fl.height + 4), fill=(200, 190, 180))
t.paste(fl, (fx, fy))
d.rectangle((0, 0, W, 118), fill=PLUM)
d.text((36, 18), "Event rental desk in Excel VBA", font=font("georgia.ttf", 40), fill=(255, 255, 255))
d.text((38, 72), "Bookings that cannot double-book · sign-in by role · late fees · audit trail", font=font("segoeui.ttf", 22), fill=(226, 205, 220))
y = 150
for line in ("Live stock per date", "Check out / check in", "Damage at replacement cost", "Invoice PDF, locked months"):
    d.text((36, y), line, font=font("segoeuib.ttf", 22), fill=TEXT); y += 34


def chip(x, y, text, col):
    f = font("segoeuib.ttf", 19)
    w = d.textlength(text, font=f)
    d.rounded_rectangle((x, y, x + w + 26, y + 38), 19, fill=col)
    d.text((x + 13, y + 7), text, font=f, fill=(255, 255, 255))


chip(36, 520, "0 items over-booked", PLUM)
chip(36, 568, f"{CHECKS} automated checks pass", GREEN)
chip(300, 520, "every attempt logged", TERRA)
t.save(ROOT / "docs" / "thumbnail.png")
print("docs/thumbnail.png")
