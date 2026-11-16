"""docs/thumbnail.png (1200x627, Upwork card): title, the dashboard, and proof chips read from the files.

    python scripts/make_thumb.py [checks]        e.g. 29/29 from test/verify.py
"""
import sys
from pathlib import Path
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
F = "C:/Windows/Fonts/"
font = lambda n, s: ImageFont.truetype(F + n, s)
CHECKS = sys.argv[1] if len(sys.argv) > 1 else (ROOT / "test" / "last-result.txt").read_text().strip()
NAVY, PAGE, TEXT, MUTED, TEAL, AMBER = (31, 58, 95), (244, 246, 249), (30, 40, 51), (102, 113, 126), (42, 157, 143), (233, 162, 59)

wb = load_workbook(ROOT / "Three-Statement-Model.xlsx", data_only=True)
sc = wb["Scenarios"]
rows = {sc[f"B{r}"].value: r for r in range(15, 27)}
fund_p = sc[f"F{rows['New funding needed']}"].value

W, H = 1200, 627
t = Image.new("RGB", (W, H), PAGE)
d = ImageDraw.Draw(t)
dash = Image.open(ROOT / "docs" / "dashboard.png").convert("RGB")
dw = 700
dl = dash.resize((dw, int(dash.height * dw / dash.width)), Image.LANCZOS)
x0 = W - dw - 30
t.paste(dl, (x0, 140)); d.rectangle((x0 - 1, 139, x0 + dl.width, 140 + dl.height), outline=(205, 212, 222), width=1)
d.rectangle((0, 0, W, 118), fill=NAVY)
d.text((36, 18), "3-statement financial model in Excel", font=font("georgia.ttf", 40), fill=(255, 255, 255))
d.text((38, 72), "P&L, balance sheet and cash flow, 36 months, from one assumptions sheet", font=font("segoeui.ttf", 22), fill=(200, 212, 228))
y = 150
for line, f, col in (("Base, optimistic and", "segoeuib.ttf", TEXT), ("pessimistic scenarios", "segoeuib.ttf", TEXT),
                     ("price x volume sensitivity", "segoeui.ttf", MUTED), ("loan, capex, working capital", "segoeui.ttf", MUTED),
                     ("investor dashboard", "segoeui.ttf", MUTED)):
    d.text((36, y), line, font=font(f, 25 if "b" in f[-6:-4] else 22), fill=col); y += 36 if "b" in f[-6:-4] else 32


def chip(x, y, text, col):
    f = font("segoeuib.ttf", 20)
    w = d.textlength(text, font=f)
    d.rounded_rectangle((x, y, x + w + 28, y + 40), 20, fill=col)
    d.text((x + 14, y + 8), text, font=f, fill=(255, 255, 255))


chip(36, 410, "Balances in 36 / 36 months", NAVY)
chip(36, 462, f"{CHECKS} automated checks pass", TEAL)
chip(36, 514, f"Pessimistic: ${fund_p:,.0f} funding gap found", AMBER)
d.text((38, 575), "Sample project · fictional company · no macros", font=font("segoeui.ttf", 17), fill=MUTED)
t.save(ROOT / "docs" / "thumbnail.png")
print("docs/thumbnail.png")
