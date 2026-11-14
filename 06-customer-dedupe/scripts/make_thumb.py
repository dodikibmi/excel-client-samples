"""docs/thumbnail.png (1200x627) for the Upwork portfolio card, and docs/before-raw-csv.png (the exports as they arrive)."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
W, H = 1200, 627
NAVY, TEAL, WHITE, MUTED = (24, 45, 72), (0, 150, 136), (255, 255, 255), (184, 200, 218)
F = "C:/Windows/Fonts/"
font = lambda name, size: ImageFont.truetype(F + name, size)

img = Image.new("RGB", (W, H), NAVY)
d = ImageDraw.Draw(img)

# Screenshot card on the right: the Merge log (which record went into which customer, and why)
shot = Image.open(ROOT / "docs" / "merge-log.png").convert("RGB")
crop = shot.crop((int(shot.width * 0.105), 0, int(shot.width * 0.80), int(shot.height * 0.80)))
cw = 700
crop = crop.resize((cw, int(crop.height * cw / crop.width)), Image.LANCZOS)
x, y = W - cw - 30, 175
shadow = Image.new("RGBA", (cw + 40, crop.height + 40), (0, 0, 0, 0))
ImageDraw.Draw(shadow).rectangle((20, 20, cw + 20, crop.height + 20), fill=(0, 0, 0, 120))
img.paste(shadow.filter(ImageFilter.GaussianBlur(12)), (x - 14, y - 8), shadow.filter(ImageFilter.GaussianBlur(12)))
img.paste(crop, (x, y))

d.rectangle((48, 52, 56, 128), fill=TEAL)
d.text((72, 46), "Customer list cleanup", font=font("segoeuib.ttf", 38), fill=WHITE)
d.text((72, 92), "3 sources, duplicates merged, CRM-ready file", font=font("segoeui.ttf", 28), fill=MUTED)

points = [("1,962 → 1,387", "rows in, unique customers out"),
          ("561", "duplicates merged, each with its rule"),
          ("11", "look-alikes left for a person to decide"),
          ("33/33", "automated checks pass")]
yy = 180
for big, small in points:
    d.text((72, yy), big, font=font("segoeuib.ttf", 34), fill=(120, 220, 205))
    d.text((72, yy + 44), small, font=font("segoeui.ttf", 19), fill=WHITE)
    yy += 92
d.text((72, H - 52), "No macros  ·  one-click Refresh  ·  email, phone, name + ZIP, nicknames", font=font("segoeui.ttf", 17), fill=MUTED)
img.save(ROOT / "docs" / "thumbnail.png")
print("docs/thumbnail.png")

# docs/before-raw-csv.png: the webshop and POS exports as they arrive
mono = font("consola.ttf", 19)
blocks = [("webshop-customers-2026-10.csv  (names in any case, emails with capitals, phones 4 ways, ZIP+4, 'Colorado')",
           (ROOT / "data" / "webshop-customers-2026-10.csv").read_text(encoding="utf-8").splitlines()[:7]),
          ("pos-customers-2026-10.csv  (\"Smith, Anna\", email often missing, ZIP lost its leading zero, M/D/YYYY)",
           (ROOT / "data" / "pos-customers-2026-10.csv").read_text(encoding="utf-8").splitlines()[:7])]
pad, lh = 28, 28
probe = ImageDraw.Draw(Image.new("RGB", (10, 10)))
bw = max(int(probe.textlength(l, font=mono)) for _, ls in blocks for l in ls) + 2 * pad
bh = sum(36 + len(ls) * lh + 24 for _, ls in blocks) + pad
before = Image.new("RGB", (bw, bh), (250, 250, 250))
bd = ImageDraw.Draw(before)
yy = pad - 10
for title, ls in blocks:
    bd.text((pad, yy), title, font=font("segoeuib.ttf", 18), fill=(156, 0, 6))
    yy += 36
    for i, l in enumerate(ls):
        bd.text((pad, yy), l, font=mono, fill=(40, 40, 40) if i else (0, 90, 160))
        yy += lh
    yy += 24
before.save(ROOT / "docs" / "before-raw-csv.png")
print("docs/before-raw-csv.png")
