"""docs/thumbnail.png (1200x627) for the Upwork portfolio card: claim on the left, the real Dashboard screenshot on the
right, dark-gold like the workbook. Also docs/before-raw-csv.png: the HR export as delivered (the "before" picture).

    python scripts/make_thumb.py 80/80        (the check count printed by test/verify.py)
"""
import sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
CHECKS = sys.argv[1] if len(sys.argv) > 1 else "?/?"
W, H = 1200, 627
BG, GOLD, GOLD_LT, IVORY, MUTED = (11, 15, 19), (201, 163, 91), (230, 207, 154), (236, 230, 218), (142, 152, 163)
F = "C:/Windows/Fonts/"
font = lambda name, size: ImageFont.truetype(F + name, size)

img = Image.new("RGB", (W, H), BG)
d = ImageDraw.Draw(img)

shot = Image.open(ROOT / "docs" / "dashboard.png").convert("RGB")
cw = 700
crop = shot.resize((cw, int(shot.height * cw / shot.width)), Image.LANCZOS)
x, y = W - cw - 30, 172
glow = Image.new("RGBA", (cw + 60, crop.height + 60), (0, 0, 0, 0))
ImageDraw.Draw(glow).rectangle((30, 30, cw + 30, crop.height + 30), fill=(201, 163, 91, 70))
glow = glow.filter(ImageFilter.GaussianBlur(14))
img.paste(glow, (x - 30, y - 30), glow)
img.paste(crop, (x, y))
d.rectangle((x - 1, y - 1, x + cw, y + crop.height), outline=GOLD, width=1)

d.text((60, 40), "R I V E R S I D E   C O M M U N I T Y   H E A L T H", font=font("segoeuib.ttf", 14), fill=GOLD)
d.text((58, 62), "FTE dashboard in Excel", font=font("georgia.ttf", 40), fill=IVORY)
d.text((60, 116), "HR roster vs Finance budget, reconciled person by person: pivots, slicers, a bridge and every gap explained",
       font=font("segoeui.ttf", 18), fill=MUTED)
d.line((60, 150, W - 30, 150), fill=GOLD, width=2)

points = [("+2.80 FTE", "budget vs HR gap, bridged by type"),
          ("23 lines", "each with both sides + what to do"),
          ("4 slicers", "drive 3 PivotTables + the dashboard"),
          ("3 months", "trend: gap +7.80 → +2.80"),
          (CHECKS, "automated checks pass")]
yy = 172
for big, small in points:
    d.text((60, yy), big, font=font("georgia.ttf", 30), fill=GOLD_LT)
    d.text((62, yy + 38), small, font=font("segoeui.ttf", 15), fill=IVORY)
    yy += 80
d.text((60, H - 40), "Power Query  ·  PivotTables + slicers  ·  no macros",
       font=font("segoeui.ttf", 16), fill=MUTED)
img.save(ROOT / "docs" / "thumbnail.png")
print("docs/thumbnail.png")

# docs/before-raw-csv.png: what the client actually exports (the "before" picture)
lines = (ROOT / "data" / "hr-roster-2026-10.csv").read_text(encoding="utf-8").splitlines()[:9]
mono = font("consola.ttf", 20)
pad, lh = 28, 30
bw = max(int(d.textlength(l, font=mono)) for l in lines) + 2 * pad
bw = max(bw, 1060)
before = Image.new("RGB", (bw, len(lines) * lh + 2 * pad + 36), (250, 250, 250))
bd = ImageDraw.Draw(before)
bd.text((pad, 14), "hr-roster-2026-10.csv  (as exported: FTE as %, fund centers typed 4 ways, terminated still listed)",
        font=font("segoeuib.ttf", 18), fill=(156, 0, 6))
for i, l in enumerate(lines):
    bd.text((pad, pad + 30 + i * lh), l, font=mono, fill=(40, 40, 40) if i else (0, 90, 160))
before.save(ROOT / "docs" / "before-raw-csv.png")
print("docs/before-raw-csv.png")
