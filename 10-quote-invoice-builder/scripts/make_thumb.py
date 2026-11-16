"""docs/thumbnail.png (1200x627) for the Upwork portfolio card, dark-gold: the real UserForm (from the demo recording)
in front of a real quote PDF, with proof numbers read from the files.

    python scripts/make_thumb.py <automated checks passed>     e.g. 78/78 from test/verify.py
"""
import json, sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
W, H = 1200, 627
BG, CARD, GOLD, GOLD_LT, IVORY, MUTED, LINE = (11, 15, 19), (26, 33, 41), (201, 163, 91), (230, 207, 154), (236, 230, 218), (142, 152, 163), (42, 51, 61)
F = "C:/Windows/Fonts/"
font = lambda name, size: ImageFont.truetype(F + name, size)
CHECKS = sys.argv[1] if len(sys.argv) > 1 else (ROOT / "test" / "last-result.txt").read_text().strip()
NUM = json.loads((ROOT / "data" / "numbers.json").read_text())
n_docs = len(NUM["quotes"]) + len(NUM["invoices"])

img = Image.new("RGB", (W, H), BG)
d = ImageDraw.Draw(img)


def shadow_paste(src, xy):
    sh = Image.new("RGBA", (src.width + 40, src.height + 40), (0, 0, 0, 0))
    ImageDraw.Draw(sh).rectangle((20, 20, src.width + 20, src.height + 20), fill=(0, 0, 0, 170))
    sh = sh.filter(ImageFilter.GaussianBlur(12))
    img.paste(sh, (xy[0] - 12, xy[1] - 8), sh)
    img.paste(src, xy)


# the quote PDF behind, the form in front
pdf = Image.open(ROOT / "docs" / "quote-pdf.png").convert("RGB")
pdf = pdf.crop((0, 0, pdf.width, int(pdf.height * 0.62)))
pdf = pdf.resize((330, int(330 * pdf.height / pdf.width)), Image.LANCZOS)
shadow_paste(pdf, (850, 40))
form = Image.open(ROOT / "docs" / "form-ready.png").convert("RGB")
form = form.resize((560, int(560 * form.height / form.width)), Image.LANCZOS)
shadow_paste(form, (590, 150))
d = ImageDraw.Draw(img)
d.rectangle((838, 28, 1192, 34), fill=GOLD)

d.text((48, 52), "E X C E L   V B A   ·   U S E R F O R M", font=font("segoeuib.ttf", 15), fill=GOLD)
d.text((44, 82), "Quote & invoice", font=font("georgia.ttf", 52), fill=IVORY)
d.text((44, 142), "builder", font=font("georgia.ttf", 52), fill=IVORY)
d.text((48, 214), "A form in; a numbered PDF and a ready", font=font("segoeui.ttf", 22), fill=MUTED)
d.text((48, 242), "email out, checked before it goes.", font=font("segoeui.ttf", 22), fill=MUTED)
chips = [(f"{n_docs}", "documents made by the macros"), ("6", "rules block a bad quote"), (CHECKS, "automated checks pass")]
y = 300
for big, small in chips:
    d.rounded_rectangle((44, y, 520, y + 70), 10, fill=CARD, outline=LINE)
    d.text((64, y + 12), big, font=font("georgiab.ttf", 34), fill=GOLD_LT)
    bw = d.textlength(big, font=font("georgiab.ttf", 34))
    d.text((64 + bw + 16, y + 24), small, font=font("segoeui.ttf", 21), fill=IVORY)
    y += 84
d.text((48, 566), "Sample project · fictional data · Excel 2016 and later", font=font("segoeui.ttf", 16), fill=MUTED)
img.save(ROOT / "docs" / "thumbnail.png")
print("docs/thumbnail.png")
