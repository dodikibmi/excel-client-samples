"""docs/thumbnail.png (1200x627) for the Upwork portfolio card, dark-gold like the workbook: a before/after story with
real screenshots (the raw orders CSV -> the Dashboard sheet) and proof numbers read from the saved workbook.
Also docs/before-raw-csv.png: the order lines as the sales team typed them (the "before" picture).

    python scripts/make_thumb.py <automated checks passed>     e.g. 74/74 from test/verify.py
"""
import sys
from pathlib import Path
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
W, H = 1200, 627
BG, CARD, GOLD, GOLD_LT, IVORY, MUTED, LINE = (11, 15, 19), (26, 33, 41), (201, 163, 91), (230, 207, 154), (236, 230, 218), (142, 152, 163), (42, 51, 61)
F = "C:/Windows/Fonts/"
font = lambda name, size: ImageFont.truetype(F + name, size)
CHECKS = sys.argv[1] if len(sys.argv) > 1 else "all"

# ---- numbers from the workbook as Excel saved it (cached values)
wb = load_workbook(ROOT / "Sales-Commission-Calculator.xlsx", data_only=True)
names = {n: wb.defined_names[n].attr_text for n in ("TestsTotal", "TestsFailed")}
cell = lambda ref: (lambda sh, a: wb[sh.strip("'")][a.replace("$", "")].value)(*ref.split("!"))
tests_total, tests_failed = int(cell(names["TestsTotal"])), int(cell(names["TestsFailed"]))
n_lines = sum(1 for _ in open(ROOT / "data" / "orders.csv", encoding="utf-8")) - 1

# ---- before picture: the raw CSV
lines = (ROOT / "data" / "orders.csv").read_text(encoding="utf-8").splitlines()[:10]
mono = font("consola.ttf", 20)
pad, lh = 28, 30
tmp = ImageDraw.Draw(Image.new("RGB", (10, 10)))
bw = max(int(tmp.textlength(l, font=mono)) for l in lines) + 2 * pad
before = Image.new("RGB", (bw, len(lines) * lh + 2 * pad + 36), (250, 250, 250))
bd = ImageDraw.Draw(before)
bd.text((pad, 14), f"orders.csv  ({n_lines:,} order lines typed by the sales team, Nov 2025 - Oct 2026)",
        font=font("segoeuib.ttf", 18), fill=(156, 0, 6))
for i, l in enumerate(lines):
    bd.text((pad, pad + 30 + i * lh), l, font=mono, fill=(40, 40, 40) if i else (0, 90, 160))
before.save(ROOT / "docs" / "before-raw-csv.png")
print("docs/before-raw-csv.png")

# ---- thumbnail
img = Image.new("RGB", (W, H), BG)
d = ImageDraw.Draw(img)


def framed(src, box_w, x, y, glow=True, outline=GOLD):
    im = src.resize((box_w, int(src.height * box_w / src.width)), Image.LANCZOS)
    if glow:
        g = Image.new("RGBA", (box_w + 60, im.height + 60), (0, 0, 0, 0))
        ImageDraw.Draw(g).rectangle((30, 30, box_w + 30, im.height + 30), fill=(201, 163, 91, 70))
        g = g.filter(ImageFilter.GaussianBlur(14))
        img.paste(g, (x - 30, y - 30), g)
    img.paste(im, (x, y))
    d.rectangle((x - 1, y - 1, x + box_w, y + im.height), outline=outline, width=1)
    return im.height


dash = Image.open(ROOT / "docs" / "dashboard.png").convert("RGB")
dash = dash.crop((0, 0, dash.width, int(dash.height * 0.66)))   # header, cards, trend and donut
aw = min(700, int(400 * dash.width / dash.height))
ax, ay = W - aw - 30, 186
ah = framed(dash, aw, ax, ay)
chip = lambda x, y, text, fg, bg: (d.rounded_rectangle((x, y, x + 12 + int(d.textlength(text, font=font("segoeuib.ttf", 13))) + 12, y + 24),
                                                       6, fill=bg, outline=fg), d.text((x + 12, y + 3), text, font=font("segoeuib.ttf", 13), fill=fg))
chip(ax, ay - 32, "AFTER  ·  one workbook, every number checked", GOLD_LT, (40, 34, 22))

# before: a small, dimmed copy of the raw CSV, bottom-left
bsrc = before.crop((0, 0, before.width, before.height))
bw2 = 340
bx, by = 60, 468
bh = framed(bsrc.point(lambda v: int(v * 0.85)), bw2, bx, by, glow=False, outline=MUTED)
chip(bx, by - 32, "BEFORE  ·  a CSV and a calculator", MUTED, CARD)
d.text((bx + bw2 + 14, by + bh // 2 - 14), "→", font=font("segoeuib.ttf", 26), fill=GOLD)

d.text((60, 32), "H A R B O R V I E W", font=font("segoeuib.ttf", 15), fill=GOLD)
d.text((58, 54), "Sales commission calculator", font=font("georgia.ttf", 40), fill=IVORY)
d.text((60, 108), "Excel: tiered bands, accelerator, manager override, split deals, clawback, payout status, what-if",
       font=font("segoeui.ttf", 18), fill=MUTED)
d.line((60, 142, W - 30, 142), fill=GOLD, width=2)

points = [("6 rules", "bands, accelerator, bonus, override, splits, clawback"),
          (f"{tests_total - tests_failed}/{tests_total}", "hand-worked test cases pass"),
          (CHECKS, "automated checks: every rep-month recomputed"),
          ("1 click", "statements for every rep, bulk PDF")]
yy = 158
for big, small in points:
    d.text((60, yy), big, font=font("georgia.ttf", 28), fill=GOLD_LT)
    d.text((62, yy + 36), small, font=font("segoeui.ttf", 15), fill=IVORY)
    yy += 64
d.text((ax, ay + ah + 10), f"{n_lines:,} order lines · 12 months · no macros · Excel 2016+", font=font("segoeui.ttf", 15), fill=MUTED)
img.save(ROOT / "docs" / "thumbnail.png")
print("docs/thumbnail.png", tests_total - tests_failed, tests_total)
