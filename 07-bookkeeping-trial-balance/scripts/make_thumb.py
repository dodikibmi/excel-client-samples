"""docs/thumbnail.png (1200x627) for the Upwork portfolio card: claim on the left, real Summary screenshot on the right.
Also docs/before-raw-csv.png: the checking export as the bank delivers it (the "before" picture)."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
W, H = 1200, 627
NAVY, TEAL, WHITE, MUTED = (24, 45, 72), (0, 150, 136), (255, 255, 255), (184, 200, 218)
F = "C:/Windows/Fonts/"
font = lambda name, size: ImageFont.truetype(F + name, size)

img = Image.new("RGB", (W, H), NAVY)
d = ImageDraw.Draw(img)

shot = Image.open(ROOT / "docs" / "trial-balance.png").convert("RGB")
crop = shot.crop((0, int(shot.height * 0.13), shot.width, shot.height * 0.94))  # the trial balance itself, totals and PASS row
cw = 680
crop = crop.resize((cw, int(crop.height * cw / crop.width)), Image.LANCZOS)
x, y = W - cw - 36, 150
shadow = Image.new("RGBA", (cw + 40, crop.height + 40), (0, 0, 0, 0))
ImageDraw.Draw(shadow).rectangle((20, 20, cw + 20, crop.height + 20), fill=(0, 0, 0, 120))
img.paste(shadow.filter(ImageFilter.GaussianBlur(12)), (x - 14, y - 8), shadow.filter(ImageFilter.GaussianBlur(12)))
img.paste(crop, (x, y))

d.rectangle((48, 52, 56, 128), fill=TEAL)
d.text((72, 46), "Small business bookkeeping", font=font("segoeuib.ttf", 36), fill=WHITE)
d.text((72, 92), "bank + card exports to a balanced trial balance", font=font("segoeui.ttf", 24), fill=MUTED)

points = [("186", "bank & card lines, categorised by rules"),
          ("0.00", "trial balance and bank rec difference"),
          ("P&L", "monthly vs budget, with variance"),
          ("65/65", "automated checks pass")]
yy = 170
for big, small in points:
    d.text((72, yy), big, font=font("segoeuib.ttf", 34), fill=(120, 220, 205))
    d.text((72, yy + 44), small, font=font("segoeui.ttf", 18), fill=WHITE)
    yy += 92
d.text((72, H - 52), "No macros  ·  one-click Refresh  ·  Excel + Power Query", font=font("segoeui.ttf", 17), fill=MUTED)
img.save(ROOT / "docs" / "thumbnail.png")
print("docs/thumbnail.png")

lines = (ROOT / "data" / "checking-2026-08-10.csv").read_text(encoding="utf-8").splitlines()[:11]
mono = font("consola.ttf", 20)
pad, lh = 28, 30
bw = max(int(d.textlength(l, font=mono)) for l in lines) + 2 * pad
bw = max(bw, 1080)
before = Image.new("RGB", (bw, len(lines) * lh + 2 * pad + 36), (250, 250, 250))
bd = ImageDraw.Draw(before)
bd.text((pad, 14), "checking-2026-08-10.csv  (as exported: title lines above the header, $ amounts stored as text, stray spaces)",
        font=font("segoeuib.ttf", 18), fill=(156, 0, 6))
for i, l in enumerate(lines):
    bd.text((pad, pad + 30 + i * lh), l, font=mono, fill=(0, 90, 160) if i == 3 else (150, 150, 150) if i < 3 else (40, 40, 40))
before.save(ROOT / "docs" / "before-raw-csv.png")
print("docs/before-raw-csv.png")
