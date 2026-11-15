"""docs/thumbnail.png (1200x627) for the Upwork portfolio card, in the dashboard's dark-gold style: claim and proof
numbers on the left, the real Dashboard screenshot (docs/dashboard.png) on the right.
Numbers come from test/expected.json (worked out in Python) and the check counts below (from the last test runs)."""
import json
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
GOOGLE_CHECKS, OFFLINE_CHECKS = "58/58", "58/58"   # Checks sheet (Run self-test) and node test/run-tests.js
W, H = 1200, 627
BG, GOLD, GOLD_LT, IVORY, MUTED = (11, 15, 19), (201, 163, 91), (230, 207, 154), (236, 230, 218), (142, 152, 163)
F = "C:/Windows/Fonts/"
font = lambda name, size: ImageFont.truetype(F + name, size)
exp = json.loads((ROOT / "test" / "expected.json").read_text(encoding="utf-8"))

img = Image.new("RGB", (W, H), BG)
d = ImageDraw.Draw(img)

shot = Image.open(ROOT / "docs" / "dashboard.png").convert("RGB")
crop = shot.crop((70, 245, 2065, 1110))   # title, navigation, KPI cards and both charts
cw = 700
crop = crop.resize((cw, int(crop.height * cw / crop.width)), Image.LANCZOS)
if crop.height > H - 190:
    crop = crop.crop((0, 0, cw, H - 190))
x, y = W - cw - 30, 168
glow = Image.new("RGBA", (cw + 60, crop.height + 60), (0, 0, 0, 0))
ImageDraw.Draw(glow).rectangle((30, 30, cw + 30, crop.height + 30), fill=(201, 163, 91, 70))
glow = glow.filter(ImageFilter.GaussianBlur(14))
img.paste(glow, (x - 30, y - 30), glow)
img.paste(crop, (x, y))
d.rectangle((x - 1, y - 1, x + cw, y + crop.height), outline=GOLD, width=1)

d.text((60, 40), "N O R T H B E A M   S T U D I O", font=font("segoeuib.ttf", 15), fill=GOLD)
d.text((58, 62), "Invoice tracker & collections desk", font=font("georgia.ttf", 38), fill=IVORY)
d.text((60, 114), "Google Sheets + Apps Script: payment reminders that never repeat, an action queue, DSO and risk",
       font=font("segoeui.ttf", 18), fill=MUTED)
d.line((60, 148, W - 30, 148), fill=GOLD, width=2)

k = exp["kpis"]
points = [("4 stages", "reminders, never the same one twice"),
          (f"{len(exp['queue'])} actions", "ranked for today: call, remind, record promise"),
          (f"{k['dso']:.1f} days", "DSO, with 6-month trends and aging heat"),
          ("Paused", "when the customer replies or promises to pay"),
          (f"{GOOGLE_CHECKS} + {OFFLINE_CHECKS}", "checks pass in Google and offline")]
yy = 168
for big, small in points:
    d.text((60, yy), big, font=font("georgia.ttf", 30), fill=GOLD_LT)
    d.text((62, yy + 38), small, font=font("segoeui.ttf", 15), fill=IVORY)
    yy += 80
d.text((60, H - 38), "ARRAYFORMULA · MAP/LAMBDA · SPARKLINE · HtmlService sidebar · triggers · PDF statement",
       font=font("segoeui.ttf", 15), fill=MUTED)
img.save(ROOT / "docs" / "thumbnail.png")
print("docs/thumbnail.png")
