"""docs/thumbnail.png (1200x627) for the Upwork portfolio card: claim and proof numbers on the left, the real Dashboard
screenshot on the right (dark-gold palette, same as the workbook). Run after shots.ps1 and test/verify.py:
    python scripts/make_thumb.py <checks passed, e.g. 70/70>"""
import sys
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

shot = Image.open(ROOT / "docs" / "dashboard-top.png").convert("RGB")
cw = 720
crop = shot.resize((cw, int(shot.height * cw / shot.width)), Image.LANCZOS)
if crop.height > H - 190:
    crop = crop.crop((0, 0, cw, H - 190))
x, y = W - cw - 30, 168
glow = Image.new("RGBA", (cw + 60, crop.height + 60), (0, 0, 0, 0))
ImageDraw.Draw(glow).rectangle((30, 30, cw + 30, crop.height + 30), fill=(201, 163, 91, 70))
glow = glow.filter(ImageFilter.GaussianBlur(14))
img.paste(glow, (x - 30, y - 30), glow)
img.paste(crop, (x, y))
d.rectangle((x - 1, y - 1, x + cw, y + crop.height), outline=GOLD, width=1)

d.text((60, 40), "B R I G H T W I R E", font=font("segoeuib.ttf", 15), fill=GOLD)
d.text((58, 62), "Project cost tracker in Excel", font=font("georgia.ttf", 40), fill=IVORY)
d.text((60, 116), "PM sheets roll up to a VP dashboard: earned value, S-curve, change orders, risks, WIP and cash",
       font=font("segoeui.ttf", 19), fill=MUTED)
d.line((60, 150, W - 30, 150), fill=GOLD, width=2)

points = [("6", "report sheets, one dark-gold theme"),
          ("CPI · SPI", "TCPI, 3 forecast methods, S-curve"),
          ("5×5", "risk matrix + change orders in the forecast"),
          ("0 macros", "weekly history saved by Refresh All"),
          (checks, "automated checks pass")]
yy = 172
for big, small in points:
    d.text((60, yy), big, font=font("georgia.ttf", 32), fill=GOLD_LT)
    d.text((62, yy + 40), small, font=font("segoeui.ttf", 16), fill=IVORY)
    yy += 82
d.text((60, H - 40), "Excel 2016+  ·  Power Query  ·  earned value management  ·  fictional data",
       font=font("segoeui.ttf", 16), fill=MUTED)
img.save(ROOT / "docs" / "thumbnail.png")
print("docs/thumbnail.png")
