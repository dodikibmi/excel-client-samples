"""docs/thumbnail.png (1200x627) for the Upwork portfolio card: claim on the left, the real Power BI Overview page on
the right. Also docs/before-raw-csv.png: the checking export as the bank delivers it (the "before" picture)."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
W, H = 1200, 627
BG, GOLD, GOLD_LT, IVORY, MUTED = (11, 15, 19), (201, 163, 91), (230, 207, 154), (236, 230, 218), (142, 152, 163)  # report palette
F = "C:/Windows/Fonts/"
font = lambda name, size: ImageFont.truetype(F + name, size)

img = Image.new("RGB", (W, H), BG)
d = ImageDraw.Draw(img)

shot = Image.open(ROOT / "docs" / "overview.png").convert("RGB")
cw = 690
crop = shot.resize((cw, int(shot.height * cw / shot.width)), Image.LANCZOS)
x, y = W - cw - 34, 168
glow = Image.new("RGBA", (cw + 60, crop.height + 60), (0, 0, 0, 0))
ImageDraw.Draw(glow).rectangle((30, 30, cw + 30, crop.height + 30), fill=(201, 163, 91, 70))
glow = glow.filter(ImageFilter.GaussianBlur(14))
img.paste(glow, (x - 30, y - 30), glow)
img.paste(crop, (x, y))
d.rectangle((x - 1, y - 1, x + cw, y + crop.height), outline=GOLD, width=1)

d.text((60, 40), "C O P P E R L E A F", font=font("segoeuib.ttf", 15), fill=GOLD)
d.text((58, 62), "Power BI finance report", font=font("georgia.ttf", 40), fill=IVORY)
d.text((60, 116), "Bank + card exports in: P&L vs budget, 13-week cash forecast, ledger and checks out",
       font=font("segoeui.ttf", 19), fill=MUTED)
d.line((60, 150, W - 34, 150), fill=GOLD, width=2)

points = [("7", "pages, one dark-gold theme"),
          ("105", "DAX measures, insights written by DAX"),
          ("13 wks", "cash forecast with a what-if buffer"),
          ("0.00", "books vs 9 bank and card statements"),
          ("41/41", "automated checks pass")]
yy = 172
for big, small in points:
    d.text((60, yy), big, font=font("georgia.ttf", 32), fill=GOLD_LT)
    d.text((62, yy + 40), small, font=font("segoeui.ttf", 16), fill=IVORY)
    yy += 82
d.text((60, H - 40), "Power Query  ·  DAX  ·  PBIP (TMDL + PBIR)  ·  .pbix included", font=font("segoeui.ttf", 16), fill=MUTED)
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
