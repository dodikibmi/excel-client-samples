"""docs/thumbnail.png (1200x627) for the Upwork portfolio card: claim on the left, real Dashboard screenshot on the right."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
W, H = 1200, 627
NAVY, TEAL, WHITE, MUTED = (24, 45, 72), (0, 150, 136), (255, 255, 255), (184, 200, 218)
F = "C:/Windows/Fonts/"
font = lambda name, size: ImageFont.truetype(F + name, size)

img = Image.new("RGB", (W, H), NAVY)
d = ImageDraw.Draw(img)

shot = Image.open(ROOT / "docs" / "dashboard.png").convert("RGB")
crop = shot.crop((80, int(shot.height * 0.142), int(shot.width * 0.82), int(shot.height * 0.455)))  # title, tiles, aging, who owes most
cw = 720
crop = crop.resize((cw, int(crop.height * cw / crop.width)), Image.LANCZOS)
x, y = W - cw - 24, 150
shadow = Image.new("RGBA", (cw + 40, crop.height + 40), (0, 0, 0, 0))
ImageDraw.Draw(shadow).rectangle((20, 20, cw + 20, crop.height + 20), fill=(0, 0, 0, 120))
blur = shadow.filter(ImageFilter.GaussianBlur(12))
img.paste(blur, (x - 14, y - 8), blur)
img.paste(crop, (x, y))

d.rectangle((48, 52, 56, 128), fill=TEAL)
d.text((72, 46), "Invoice tracker", font=font("segoeuib.ttf", 38), fill=WHITE)
d.text((72, 92), "payment reminders, Google Sheets", font=font("segoeui.ttf", 28), fill=MUTED)

points = [("4 stages", "heads-up, friendly, second, final"),
          ("0 repeats", "a stage is never emailed twice"),
          ("Preview", "see every email before anything sends"),
          ("31 + 30", "checks pass in Google and offline")]
yy = 180
for big, small in points:
    d.text((72, yy), big, font=font("segoeuib.ttf", 34), fill=(120, 220, 205))
    d.text((72, yy + 44), small, font=font("segoeui.ttf", 19), fill=WHITE)
    yy += 92

d.text((72, H - 52), "Apps Script  ·  daily trigger  ·  weekly summary email", font=font("segoeui.ttf", 17), fill=MUTED)
img.save(ROOT / "docs" / "thumbnail.png")
print("docs/thumbnail.png")

