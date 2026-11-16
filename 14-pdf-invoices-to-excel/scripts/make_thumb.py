"""docs/before-invoices.png (the 4 invoice layouts side by side) and docs/thumbnail.png (1200x627, Upwork card).

    python scripts/make_thumb.py [checks]        e.g. 27/27 from test/verify.py
"""
import json
import sys
from pathlib import Path

import pdfplumber
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
F = "C:/Windows/Fonts/"
font = lambda n, s: ImageFont.truetype(F + n, s)
CHECKS = sys.argv[1] if len(sys.argv) > 1 else (ROOT / "test" / "last-result.txt").read_text().strip()
BG, CARD, GOLD, GOLD_LT, IVORY, MUTED = (11, 15, 19), (26, 33, 41), (201, 163, 91), (230, 207, 154), (236, 230, 218), (142, 152, 163)
JADE, BAD = (79, 163, 140), (229, 115, 95)

# ---- the four layouts as they arrive
exp = json.loads((ROOT / "test" / "expected.json").read_text(encoding="utf-8"))
inv = {i["number"]: i for i in exp["invoices"]}
pick = {}
for f in exp["files"]:
    i = inv[f["number"]]
    if f["status"] == "Read" and f["role"] == "original" and i["kind"] == "Invoice":
        key = i["layout"]
        if key == "C" and f["pages"] < 2:
            continue
        pick.setdefault(key, f["file"])
pages = []
for key in "ABCD":
    with pdfplumber.open(ROOT / "invoices" / pick[key]) as p:
        im = p.pages[0].crop((30, 30, p.pages[0].width - 30, p.pages[0].height * 0.62)).to_image(resolution=110).original
        pages.append(im.convert("RGB"))
w = 520
scaled = [im.resize((w, int(im.height * w / im.width)), Image.LANCZOS) for im in pages]
h = max(s.height for s in scaled)
pad = 18
col = Image.new("RGB", (pad + 4 * (w + pad), h + 2 * pad + 40), (232, 232, 232))
dc = ImageDraw.Draw(col)
for k, (key, s) in enumerate(zip("ABCD", scaled)):
    x = pad + k * (w + pad)
    col.paste(s, (x, pad + 40))
    dc.rectangle((x - 1, pad + 39, x + w, pad + 40 + s.height), outline=(170, 170, 170))
    dc.text((x, pad + 6), pick[key], font=font("segoeuib.ttf", 20), fill=(40, 40, 40))
col.save(ROOT / "docs" / "before-invoices.png")
print("docs/before-invoices.png")

# ---- thumbnail
wb = load_workbook(ROOT / "PDF-Invoices-to-Excel.xlsx", data_only=True)
ds = wb["Dashboard"]
files, auto = ds["B6"].value, ds["D6"].value
W, H = 1200, 627
t = Image.new("RGB", (W, H), BG)
d = ImageDraw.Draw(t)
dash = Image.open(ROOT / "docs" / "dashboard.png").convert("RGB")
dw = 690
dl = dash.resize((dw, int(dash.height * dw / dash.width)), Image.LANCZOS)
x0 = W - dw - 28
t.paste(dl, (x0, 150))
d.rectangle((x0 - 1, 149, x0 + dl.width, 150 + dl.height), outline=(60, 70, 80), width=1)
d.text((34, 22), "300 PDF invoices to one Excel register", font=font("georgia.ttf", 40), fill=GOLD_LT)
d.text((36, 80), "Power Query reads every supplier's layout, checks the sums, flags what a person must see",
       font=font("segoeui.ttf", 20), fill=MUTED)
y = 160
for line, f, c in (("8 suppliers, 4 layouts,", "segoeuib.ttf", IVORY), ("multi-page invoices", "segoeuib.ttf", IVORY),
                   ("dates read per supplier", "segoeui.ttf", MUTED), ("duplicates & credit notes", "segoeui.ttf", MUTED),
                   ("bills import for Xero / QBO", "segoeui.ttf", MUTED)):
    d.text((34, y), line, font=font(f, 25 if f.endswith("b.ttf") else 22), fill=c)
    y += 36 if f.endswith("b.ttf") else 32


def chip(x, y, text, colr):
    f = font("segoeuib.ttf", 19)
    w = d.textlength(text, font=f)
    d.rounded_rectangle((x, y, x + w + 28, y + 40), 20, fill=colr)
    d.text((x + 14, y + 8), text, font=f, fill=(15, 15, 15))


chip(34, 400, f"{auto} of {files} read with no typing", GOLD)
chip(34, 452, f"{CHECKS} automated checks pass", JADE)
chip(34, 504, "3 supplier sum errors held", BAD)
d.text((36, 575), "Sample project · fictional company · no macros", font=font("segoeui.ttf", 17), fill=MUTED)
t.save(ROOT / "docs" / "thumbnail.png")
print("docs/thumbnail.png")
