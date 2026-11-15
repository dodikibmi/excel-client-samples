"""docs/thumbnail.png (1200x627) for the Upwork portfolio card: claim on the left, the real Overview page on the right.
Also docs/before-raw-csv.png: the raw sales export the client starts from.
Usage: python scripts/make_thumb.py <checks passed>/<checks total>   (numbers from test/verify.py)"""
import csv, sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
W, H = 1200, 627
BG, GOLD, GOLD_LT, IVORY, MUTED = (11, 15, 19), (201, 163, 91), (230, 207, 154), (236, 230, 218), (142, 152, 163)
F = "C:/Windows/Fonts/"
font = lambda name, size: ImageFont.truetype(F + name, size)
checks = sys.argv[1] if len(sys.argv) > 1 else "?/?"

img = Image.new("RGB", (W, H), BG)
d = ImageDraw.Draw(img)

shot = Image.open(ROOT / "docs" / "dashboard-west.png").convert("RGB")
cw = 700
crop = shot.resize((cw, int(shot.height * cw / shot.width)), Image.LANCZOS)
x, y = W - cw - 34, 168
crop = crop.crop((0, 0, cw, min(crop.height, H - y - 22)))
glow = Image.new("RGBA", (cw + 60, crop.height + 60), (0, 0, 0, 0))
ImageDraw.Draw(glow).rectangle((30, 30, cw + 30, crop.height + 30), fill=(201, 163, 91, 70))
glow = glow.filter(ImageFilter.GaussianBlur(14))
img.paste(glow, (x - 30, y - 30), glow)
img.paste(crop, (x, y))
d.rectangle((x - 1, y - 1, x + cw, y + crop.height), outline=GOLD, width=1)

d.text((60, 40), "B R I G H T L I N E", font=font("segoeuib.ttf", 15), fill=GOLD)
d.text((58, 62), "Excel sales dashboard", font=font("georgia.ttf", 40), fill=IVORY)
d.text((60, 116), "Budget vs actual vs prior year: slicers, timeline, auto insights, forecast, 6 linked pages",
       font=font("segoeui.ttf", 19), fill=MUTED)
d.line((60, 150, W - 34, 150), fill=GOLD, width=2)

points = [("6 pages", "slicers + timeline drive every number"),
          ("6 insights", "sentences written by formulas"),
          ("-61%", "Garden in West: the miss, found in 1 click"),
          ("80/20", "Pareto, rep gauges, year-end forecast"),
          (checks, "automated checks pass")]
yy = 172
for big, small in points:
    d.text((60, yy), big, font=font("georgia.ttf", 32), fill=GOLD_LT)
    d.text((62, yy + 40), small, font=font("segoeui.ttf", 16), fill=IVORY)
    yy += 82
d.text((60, H - 40), "Power Query · traceable SUMIFS · no macros", font=font("segoeui.ttf", 16), fill=MUTED)
img.save(ROOT / "docs" / "thumbnail.png")
print("docs/thumbnail.png")

lines = (ROOT / "data" / "sales-lines.csv").read_text(encoding="utf-8").splitlines()
n = len(lines) - 1
lines = lines[:9]
mono = font("consola.ttf", 20)
pad, lh = 28, 30
bw = max(int(d.textlength(l, font=mono)) for l in lines) + 2 * pad
bw = max(bw, 1000)
before = Image.new("RGB", (bw, len(lines) * lh + 2 * pad + 36), (250, 250, 250))
bd = ImageDraw.Draw(before)
bd.text((pad, 14), f"sales-lines.csv  ({n:,} invoice lines, 22 months: what the accounting system exports)",
        font=font("segoeuib.ttf", 18), fill=(156, 0, 6))
for i, l in enumerate(lines):
    bd.text((pad, pad + 30 + i * lh), l, font=mono, fill=(40, 40, 40) if i else (0, 90, 160))
before.save(ROOT / "docs" / "before-raw-csv.png")
print("docs/before-raw-csv.png")
