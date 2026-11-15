"""docs/thumbnail.png (1200x627) for the Upwork portfolio card, and docs/before-raw-csv.png (the exports as they arrive)."""
from pathlib import Path
import openpyxl
from PIL import Image, ImageDraw, ImageFont, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
W, H = 1200, 627
BG, GOLD, GOLD_LT, IVORY, MUTED = (11, 15, 19), (201, 163, 91), (230, 207, 154), (236, 230, 218), (142, 152, 163)  # E9 palette
F = "C:/Windows/Fonts/"
font = lambda name, size: ImageFont.truetype(F + name, size)
CHECKS = "74/74"  # automated checks in test/verify.py (update after a test run)

# every number on the card comes from the saved workbook (cached values of the cockpit)
s = openpyxl.load_workbook(ROOT / "Customer-Cleanup.xlsx", data_only=True)["Summary"]
rows_in, aside, merged, unique, review = (int(s[c].value) for c in ("D8", "I8", "N8", "S8", "X8"))
q = {s[f"AD{r}"].value: (s[f"AE{r}"].value, s[f"AG{r}"].value) for r in range(21, 25)}

img = Image.new("RGB", (W, H), BG)
d = ImageDraw.Draw(img)
shot = Image.open(ROOT / "docs" / "summary.png").convert("RGB")
cw = 700
crop = shot.resize((cw, int(shot.height * cw / shot.width)), Image.LANCZOS)
x, y = W - cw - 30, 168
glow = Image.new("RGBA", (cw + 60, crop.height + 60), (0, 0, 0, 0))
ImageDraw.Draw(glow).rectangle((30, 30, cw + 30, crop.height + 30), fill=(201, 163, 91, 70))
glow = glow.filter(ImageFilter.GaussianBlur(14))
img.paste(glow, (x - 30, y - 30), glow)
img.paste(crop, (x, y))
d.rectangle((x - 1, y - 1, x + cw, y + crop.height), outline=GOLD, width=1)

d.text((60, 40), "S U M M I T   R I D G E", font=font("segoeuib.ttf", 15), fill=GOLD)
d.text((58, 62), "Customer list cleanup in Excel", font=font("georgia.ttf", 40), fill=IVORY)
d.text((60, 116), "Webshop + POS + sales sheet in: duplicates merged, data quality cockpit, CRM-ready file out",
       font=font("segoeui.ttf", 19), fill=MUTED)
d.line((60, 150, W - 30, 150), fill=GOLD, width=2)
e0, e1 = q["Valid email"]
points = [(f"{rows_in:,} → {unique:,}", "rows in, unique customers out"),
          (f"{merged}", "duplicates merged, each with its rule"),
          (f"{e0:.0%} → {e1:.0%}", "valid emails, before → after"),
          (f"{review}", "look-alikes: you decide Merge / Keep separate"),
          (CHECKS, "automated checks pass")]
yy = 172
for big, small in points:
    x0, big_f, arrow_f = 60, font("georgia.ttf", 32), font("segoeui.ttf", 28)  # Georgia has no arrow glyph
    for i, part in enumerate(big.split("→")):
        if i:
            d.text((x0 + 4, yy + 2), "→", font=arrow_f, fill=GOLD)
            x0 += int(d.textlength("→", font=arrow_f)) + 10
        d.text((x0, yy), part.strip(), font=big_f, fill=GOLD_LT)
        x0 += int(d.textlength(part.strip(), font=big_f))
    d.text((62, yy + 40), small, font=font("segoeui.ttf", 16), fill=IVORY)
    yy += 82
d.text((60, H - 40), "Power Query  ·  FILTER + XLOOKUP  ·  no macros", font=font("segoeui.ttf", 16), fill=MUTED)
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
