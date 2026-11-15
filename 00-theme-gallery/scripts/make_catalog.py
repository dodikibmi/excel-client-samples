"""From the screenshots in docs/ (made by shots.py):
  docs/theme-catalog.png  all 8 themed dashboards on one image (4 x 2) with number and name: attach it to a proposal
  docs/thumbnail.png      1200x627 Upwork card: claim on the left, 4 themes on the right
Usage: python scripts/make_catalog.py <checks passed>/<checks total>   (numbers from test/verify.py)"""
import os, sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageFilter

sys.path.insert(0, os.path.dirname(__file__))
from themes import THEMES

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"
F = "C:/Windows/Fonts/"
font = lambda name, size: ImageFont.truetype(F + name, size)
hexrgb = lambda h: tuple(int(h.lstrip("#")[k:k + 2], 16) for k in (0, 2, 4))
checks = sys.argv[1] if len(sys.argv) > 1 else "?/?"
BG, INK, MUTED, GOLD, IVORY = (236, 234, 230), (30, 34, 40), (107, 111, 118), (201, 163, 91), (236, 230, 218)
shots = [Image.open(DOCS / f"theme-{i + 1}-{t['key']}.png").convert("RGB") for i, t in enumerate(THEMES)]

# ---------------------------------------------------------------- catalogue
TW, GAP, M = 760, 36, 60
th = int(shots[0].height * TW / shots[0].width)
W = M * 2 + TW * 4 + GAP * 3
H = 230 + (th + 92) * 2 + 70
cat = Image.new("RGB", (W, H), BG)
d = ImageDraw.Draw(cat)
d.rectangle((0, 0, W, 170), fill=(21, 24, 28))
d.rectangle((0, 170, W, 175), fill=GOLD)
d.text((M, 34), "C O L O U R   T H E M E S", font=font("segoeuib.ttf", 20), fill=GOLD)
d.text((M - 2, 64), "Pick the look for your dashboard", font=font("georgia.ttf", 52), fill=IVORY)
d.text((M, 132), "Same data, formulas and charts in 8 themes.  Reply with a number, mix two, or send your brand colours.",
       font=font("segoeui.ttf", 22), fill=(169, 174, 181))
for i, (t, shot) in enumerate(zip(THEMES, shots)):
    x = M + (i % 4) * (TW + GAP)
    y = 215 + (i // 4) * (th + 92)
    im = shot.resize((TW, th), Image.LANCZOS)
    sh = Image.new("RGBA", (TW + 40, th + 40), (0, 0, 0, 0))
    ImageDraw.Draw(sh).rectangle((20, 24, TW + 20, th + 24), fill=(0, 0, 0, 60))
    sh = sh.filter(ImageFilter.GaussianBlur(9))
    cat.paste(sh, (x - 20, y - 20), sh)
    cat.paste(im, (x, y))
    acc = hexrgb(t["accent"])
    d.rounded_rectangle((x, y + th + 16, x + 44, y + th + 60), 8, fill=acc)
    on = (255, 255, 255) if sum(acc) < 420 else (17, 17, 17)
    num = str(i + 1)
    d.text((x + 22 - d.textlength(num, font=font("segoeuib.ttf", 24)) / 2, y + th + 21), num,
           font=font("segoeuib.ttf", 24), fill=on)
    d.text((x + 58, y + th + 14), t["name"], font=font("georgia.ttf", 26), fill=INK)
    d.text((x + 58, y + th + 46), f"{t['mood']}  ·  {t['best_for'].lower()}", font=font("segoeui.ttf", 17), fill=MUTED)
    for k, role in enumerate(("page", "card", "accent", "s2", "s3", "s4")):
        sx = x + TW - (6 - k) * 30
        d.rectangle((sx, y + th + 22, sx + 24, y + th + 46), fill=hexrgb(t[role]), outline=(200, 196, 188))
d.text((M, H - 52), "Every theme passes WCAG AA contrast  ·  hex codes for Excel, Power BI and Google Sheets included  ·  "
       "fictional data", font=font("segoeui.ttf", 19), fill=MUTED)
cat.save(DOCS / "theme-catalog.png")
small = cat.resize((1800, int(H * 1800 / W)), Image.LANCZOS)
small.save(DOCS / "theme-catalog-small.png")
print("docs/theme-catalog.png", cat.size, "+ docs/theme-catalog-small.png", small.size)

# ---------------------------------------------------------------- thumbnail
W, H = 1200, 627
BGD, GOLD_LT = (11, 15, 19), (230, 207, 154)
img = Image.new("RGB", (W, H), BGD)
d = ImageDraw.Draw(img)
d.text((56, 40), "E X C E L   D A S H B O A R D   T H E M E S", font=font("segoeuib.ttf", 14), fill=GOLD)
d.text((54, 62), "Pick your colours", font=font("georgia.ttf", 42), fill=IVORY)
d.text((56, 118), "One sales dashboard in 8 ready-made colour themes", font=font("segoeui.ttf", 19), fill=(142, 152, 163))
d.line((56, 152, W - 34, 152), fill=GOLD, width=2)
points = [("8 themes", "5 dark, 3 light, same data"),
          ("14 colours", "per theme, hex codes included"),
          ("WCAG AA", "text contrast at least 4.5:1"),
          (checks, "automated checks pass")]
yy = 178
for big, sm in points:
    d.text((56, yy), big, font=font("georgia.ttf", 30), fill=GOLD_LT)
    d.text((58, yy + 38), sm, font=font("segoeui.ttf", 16), fill=IVORY)
    yy += 84
d.text((56, H - 44), "Excel · no macros · your brand = theme 9", font=font("segoeui.ttf", 16), fill=(142, 152, 163))
cw = 262
ch = int(shots[0].height * cw / shots[0].width)
x0, y0, g = W - 2 * cw - 18 - 34, 172, 18
for k, idx in enumerate((0, 3, 5, 6)):
    x, y = x0 + (k % 2) * (cw + g), y0 + (k // 2) * (ch + 44)
    im = shots[idx].resize((cw, ch), Image.LANCZOS)
    img.paste(im, (x, y))
    d.rectangle((x - 1, y - 1, x + cw, y + ch), outline=(60, 66, 74), width=1)
    t = THEMES[idx]
    d.rounded_rectangle((x, y + ch + 8, x + 26, y + ch + 34), 5, fill=hexrgb(t["accent"]))
    on = (255, 255, 255) if sum(hexrgb(t["accent"])) < 420 else (17, 17, 17)
    d.text((x + 8, y + ch + 9), str(idx + 1), font=font("segoeuib.ttf", 17), fill=on)
    d.text((x + 36, y + ch + 9), t["name"], font=font("segoeui.ttf", 17), fill=IVORY)
img.save(DOCS / "thumbnail.png")
print("docs/thumbnail.png")
