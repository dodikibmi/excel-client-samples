"""docs/thumbnail.png (1200x627) for the Upwork portfolio card: claim on the left, the real Summary screenshot on the right
(dark-gold, same palette as the workbook). Also docs/before-raw-csv.png: the checking export as the bank delivers it.

    python scripts/make_thumb.py [checks passed, e.g. 120/120]
"""
import sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
W, H = 1200, 627
BG, GOLD, GOLD_LT, IVORY, MUTED = (11, 15, 19), (201, 163, 91), (230, 207, 154), (236, 230, 218), (142, 152, 163)
F = "C:/Windows/Fonts/"
font = lambda name, size: ImageFont.truetype(F + name, size)
checks = sys.argv[1] if len(sys.argv) > 1 else "all"

img = Image.new("RGB", (W, H), BG)
d = ImageDraw.Draw(img)

shot = Image.open(ROOT / "docs" / "summary.png").convert("RGB")
cw = 700
crop = shot.resize((cw, int(shot.height * cw / shot.width)), Image.LANCZOS)
if crop.height > H - 168 - 30:
    crop = crop.crop((0, 0, cw, H - 168 - 30))
x, y = W - cw - 30, 168
glow = Image.new("RGBA", (cw + 60, crop.height + 60), (0, 0, 0, 0))
ImageDraw.Draw(glow).rectangle((30, 30, cw + 30, crop.height + 30), fill=(201, 163, 91, 70))
glow = glow.filter(ImageFilter.GaussianBlur(14))
img.paste(glow, (x - 30, y - 30), glow)
img.paste(crop, (x, y))
d.rectangle((x - 1, y - 1, x + cw, y + crop.height), outline=GOLD, width=1)

d.text((60, 40), "C O P P E R L E A F", font=font("segoeuib.ttf", 15), fill=GOLD)
d.text((58, 62), "Small business bookkeeping in Excel", font=font("georgia.ttf", 38), fill=IVORY)
d.text((60, 114), "Bank + card exports in: ledger, trial balance, P&L, cash flow, 13-week forecast and month-end close out",
       font=font("segoeui.ttf", 18), fill=MUTED)
d.line((60, 150, W - 30, 150), fill=GOLD, width=2)

points = [("186", "bank & card lines, categorised by rules"),
          ("0.00", "trial balance + 9 statements, every month"),
          ("13 wks", "cash forecast, 4.3 weeks of cash"),
          ("10 steps", "month-end close, ticked by formulas"),
          (checks, "automated checks pass")]
yy = 170
for big, small in points:
    d.text((60, yy), big, font=font("georgia.ttf", 30), fill=GOLD_LT)
    d.text((62, yy + 38), small, font=font("segoeui.ttf", 15), fill=IVORY)
    yy += 80
d.text((60, H - 40), "Excel + Power Query  \u00B7  slicers  \u00B7  no macros", font=font("segoeui.ttf", 16), fill=MUTED)
img.save(ROOT / "docs" / "thumbnail.png")
print("docs/thumbnail.png")

lines = (ROOT / "data" / "checking-2026-08-10.csv").read_text(encoding="utf-8").splitlines()[:11]
mono = font("consola.ttf", 20)
pad, lh = 28, 30
bw = max(max(int(d.textlength(l, font=mono)) for l in lines) + 2 * pad, 1080)
before = Image.new("RGB", (bw, len(lines) * lh + 2 * pad + 36), (250, 250, 250))
bd = ImageDraw.Draw(before)
bd.text((pad, 14), "checking-2026-08-10.csv  (as exported: title lines above the header, $ amounts stored as text, stray spaces)",
        font=font("segoeuib.ttf", 18), fill=(156, 0, 6))
for i, l in enumerate(lines):
    bd.text((pad, pad + 30 + i * lh), l, font=mono, fill=(0, 90, 160) if i == 3 else (150, 150, 150) if i < 3 else (40, 40, 40))
before.save(ROOT / "docs" / "before-raw-csv.png")
print("docs/before-raw-csv.png")
