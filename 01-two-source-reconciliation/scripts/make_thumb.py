"""docs/thumbnail.png (1200x627) for the Upwork portfolio card: claim on the left, real Summary screenshot on the right."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
W, H = 1200, 627
NAVY, TEAL, WHITE, MUTED = (24, 45, 72), (0, 150, 136), (255, 255, 255), (184, 200, 218)
F = "C:/Windows/Fonts/"
font = lambda name, size: ImageFont.truetype(F + name, size)

img = Image.new("RGB", (W, H), NAVY)
d = ImageDraw.Draw(img)

# Screenshot card on the right, slightly overflowing the bottom edge
shot = Image.open(ROOT / "docs" / "summary.png").convert("RGB")
crop = shot.crop((0, 0, shot.width, int(shot.height * 0.53)))  # title band, KPI cards, status table
cw = 720
crop = crop.resize((cw, int(crop.height * cw / crop.width)), Image.LANCZOS)
x, y = W - cw - 24, 150
shadow = Image.new("RGBA", (cw + 40, crop.height + 40), (0, 0, 0, 0))
ImageDraw.Draw(shadow).rectangle((20, 20, cw + 20, crop.height + 20), fill=(0, 0, 0, 120))
img.paste(shadow.filter(ImageFilter.GaussianBlur(12)), (x - 14, y - 8), shadow.filter(ImageFilter.GaussianBlur(12)))
img.paste(crop, (x, y))

# Left column
d.rectangle((48, 52, 56, 128), fill=TEAL)
d.text((72, 46), "Bank reconciliation", font=font("segoeuib.ttf", 38), fill=WHITE)
d.text((72, 92), "in Excel + Power Query", font=font("segoeui.ttf", 28), fill=MUTED)

points = [("95%", "of bank lines matched automatically"),
          ("13", "exceptions, each with a next step"),
          ("0.00", "unexplained difference, proven"),
          ("31/31", "automated checks pass")]
yy = 180
for big, small in points:
    d.text((72, yy), big, font=font("segoeuib.ttf", 34), fill=(120, 220, 205))
    d.text((72, yy + 44), small, font=font("segoeui.ttf", 19), fill=WHITE)
    yy += 92

d.text((72, H - 52), "No macros  ·  one-click Refresh  ·  messy CSVs in", font=font("segoeui.ttf", 17), fill=MUTED)
img.save(ROOT / "docs" / "thumbnail.png")
print("docs/thumbnail.png")

# docs/before-raw-csv.png: what the client actually exports (the "before" picture)
lines = (ROOT / "data" / "bank-statement-2026-10.csv").read_text(encoding="utf-8").splitlines()[:9]
mono = font("consola.ttf", 20)
pad, lh = 28, 30
bw = max(int(d.textlength(l, font=mono)) for l in lines) + 2 * pad
before = Image.new("RGB", (bw, len(lines) * lh + 2 * pad + 36), (250, 250, 250))
bd = ImageDraw.Draw(before)
bd.text((pad, 14), "bank-statement-2026-10.csv  (as exported: text amounts, US dates, references typed 4 ways)",
        font=font("segoeuib.ttf", 18), fill=(156, 0, 6))
for i, l in enumerate(lines):
    bd.text((pad, pad + 30 + i * lh), l, font=mono, fill=(40, 40, 40) if i else (0, 90, 160))
before.save(ROOT / "docs" / "before-raw-csv.png")
print("docs/before-raw-csv.png")
