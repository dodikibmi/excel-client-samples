"""docs/before-raw-exports.png: the two exports as the store gets them (the "before" picture), and, once
docs/overview.png exists (scripts/shots.py), docs/thumbnail.png (1200x627) for the Upwork portfolio card.
Run:  python scripts/make_thumb.py
"""
import csv, json
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
F = "C:/Windows/Fonts/"
font = lambda name, size: ImageFont.truetype(F + name, size)
PAPER, ACCENT, ACCENT_LT, INK, MUTED, LINE = (250, 246, 240), (194, 96, 61), (227, 165, 135), (59, 47, 42), (125, 110, 99), (232, 223, 210)
RED, BLUE, GREY = (176, 40, 30), (0, 90, 160), (150, 150, 150)

# ------------------------------------------------------------------ before: the two exports
with open(ROOT / "data" / "shop-orders-export.csv", encoding="utf-8") as f:
    shop = list(csv.reader(f))
cols = ["Name", "Financial Status", "Currency", "Subtotal", "Total", "Discount Code", "Created at", "Lineitem quantity",
        "Lineitem name", "Lineitem price", "Lineitem sku", "Refunded Amount"]
idx = [shop[0].index(c) for c in cols]
start = next(i for i, r in enumerate(shop) if i > 1 and r[2] == "partially_refunded" and r[0] == shop[i + 1][0] == shop[i + 2][0])
rows = [shop[0]] + shop[start - 2:start + 6]
with open(ROOT / "data" / "marketplace-orders-eu.csv", encoding="utf-8") as f:
    mp = list(csv.reader(f, delimiter=";"))
seen, dup_at = set(), None
for i, r in enumerate(mp):
    if r[0] != "Order ID" and tuple(r) in seen:
        dup_at = i
        break
    seen.add(tuple(r))
mrows = mp[dup_at - 3:dup_at + 3]

mono = font("consola.ttf", 17)
W = 1430
img = Image.new("RGB", (W, 640), (252, 252, 252))
d = ImageDraw.Draw(img)
d.text((24, 14), "shop-orders-export.csv  (Shopify layout: status, currency, totals and refund only on the FIRST line of each order; SKUs typed by hand)",
       font=font("segoeuib.ttf", 17), fill=RED)
widths = [70, 150, 70, 80, 80, 100, 230, 40, 230, 70, 110, 90]
y = 46
for n, r in enumerate(rows):
    x = 24
    for c, w in zip(idx, widths):
        v = r[c] if c < len(r) else ""
        if n == 0:
            v = {"Financial Status": "Status", "Lineitem quantity": "Qty", "Lineitem name": "Lineitem", "Lineitem price": "Price",
                 "Lineitem sku": "SKU", "Refunded Amount": "Refunded", "Discount Code": "Code"}.get(v, v)
        shown = v if d.textlength(v, font=mono) < w - 6 else v[: max(1, int(len(v) * (w - 14) / d.textlength(v, font=mono)))] + "…"
        colour = BLUE if n == 0 else GREY if v == "" else RED if (c == shop[0].index("Lineitem sku") and v != v.strip().upper()) else INK
        d.text((x, y), shown if shown else "·", font=mono, fill=colour)
        x += w
    y += 26
d.text((24, y + 18), "marketplace-orders-eu.csv  (semicolons, dd.mm.yyyy, decimal commas; monthly exports pasted together: the header repeats, a row appears twice)",
       font=font("segoeuib.ttf", 17), fill=RED)
y += 50
for r in mrows:
    line = ";".join(r)
    colour = BLUE if r[0] == "Order ID" else RED if mrows.count(r) > 1 else INK
    d.text((24, y), line, font=mono, fill=colour)
    y += 26
img = img.crop((0, 0, W, y + 20))
img.save(ROOT / "docs" / "before-raw-exports.png")
print("docs/before-raw-exports.png")

# ------------------------------------------------------------------ thumbnail for the portfolio card
shot_file = ROOT / "docs" / "overview.png"
if shot_file.exists():
    exp = json.loads((ROOT / "test" / "expected.json").read_text(encoding="utf-8"))
    s = exp["summary"]
    W, H = 1200, 627
    img = Image.new("RGB", (W, H), PAPER)
    d = ImageDraw.Draw(img)
    shot = Image.open(shot_file).convert("RGB")
    cw = 690
    crop = shot.resize((cw, int(shot.height * cw / shot.width)), Image.LANCZOS)
    x, y = W - cw - 34, 168
    glow = Image.new("RGBA", (cw + 60, crop.height + 60), (0, 0, 0, 0))
    ImageDraw.Draw(glow).rectangle((30, 30, cw + 30, crop.height + 30), fill=(194, 96, 61, 60))
    glow = glow.filter(ImageFilter.GaussianBlur(14))
    img.paste(glow, (x - 30, y - 30), glow)
    img.paste(crop, (x, y))
    d.rectangle((x - 1, y - 1, x + cw, y + crop.height), outline=LINE, width=1)
    d.text((60, 40), "L A R K S P U R", font=font("segoeuib.ttf", 15), fill=ACCENT)
    d.text((58, 62), "Power BI e-commerce sales", font=font("georgia.ttf", 40), fill=INK)
    d.text((60, 116), "Webshop + marketplace exports in 5 currencies, live exchange rates from an API",
           font=font("segoeui.ttf", 19), fill=MUTED)
    d.line((60, 150, W - 34, 150), fill=ACCENT, width=2)
    points = [(f"{s['net_2026_ytd'] / s['net_2025_ytd'] - 1:+.1%}", "net sales vs last year, in USD"),
              (f"{s['net_cc_2026'] / s['net_2025_ytd'] - 1:+.1%}", "at last year's exchange rates"),
              ("API", "daily ECB rates, saved copy as fallback"),
              ("cohorts", "new vs returning, MER, cost per customer"),
              ("31/31", "automated checks pass")]
    yy = 172
    for big, small in points:
        d.text((60, yy), big, font=font("georgia.ttf", 32), fill=ACCENT)
        d.text((62, yy + 40), small, font=font("segoeui.ttf", 16), fill=INK)
        yy += 82
    d.text((60, H - 40), "Power Query  ·  Web API  ·  DAX  ·  PBIP (TMDL + PBIR)  ·  .pbix included", font=font("segoeui.ttf", 16), fill=MUTED)
    img.save(ROOT / "docs" / "thumbnail.png")
    print("docs/thumbnail.png")
else:
    print("docs/overview.png not there yet: thumbnail skipped")
