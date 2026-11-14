"""docs/thumbnail.png (1200x627) for the Upwork portfolio card: claim on the left, real CPI trend chart on the right."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
W, H = 1200, 627
NAVY, TEAL, WHITE, MUTED = (24, 45, 72), (0, 150, 136), (255, 255, 255), (184, 200, 218)
F = "C:/Windows/Fonts/"
font = lambda name, size: ImageFont.truetype(F + name, size)

img = Image.new("RGB", (W, H), NAVY)
d = ImageDraw.Draw(img)

shot = Image.open(ROOT / "docs" / "weekly-trend.png").convert("RGB")
crop = shot.crop((int(shot.width * 0.02), int(shot.height * 0.47), int(shot.width * 0.50), int(shot.height * 0.94)))  # CPI trend chart
cw = 640
crop = crop.resize((cw, int(crop.height * cw / crop.width)), Image.LANCZOS)
x, y = W - cw - 40, 120
shadow = Image.new("RGBA", (cw + 40, crop.height + 40), (0, 0, 0, 0))
ImageDraw.Draw(shadow).rectangle((20, 20, cw + 20, crop.height + 20), fill=(0, 0, 0, 120))
img.paste(shadow.filter(ImageFilter.GaussianBlur(12)), (x - 14, y - 8), shadow.filter(ImageFilter.GaussianBlur(12)))
img.paste(crop, (x, y))

d.rectangle((48, 52, 56, 128), fill=TEAL)
d.text((72, 46), "Project cost tracker", font=font("segoeuib.ttf", 36), fill=WHITE)
d.text((72, 92), "PM sheets roll up to a VP dashboard", font=font("segoeui.ttf", 24), fill=MUTED)

points = [("3 PM sheets", "one VP dashboard, budget vs actual"),
          ("8 weeks", "of history, saved by Refresh All"),
          ("CPI 0.90", "warehouse flagged At risk early"),
          ("44/44", "automated checks pass")]
yy = 170
for big, small in points:
    d.text((72, yy), big, font=font("segoeuib.ttf", 32), fill=(120, 220, 205))
    d.text((72, yy + 42), small, font=font("segoeui.ttf", 18), fill=WHITE)
    yy += 90

d.text((72, H - 52), "No macros  ·  earned value  ·  Excel 2016+", font=font("segoeui.ttf", 17), fill=MUTED)
img.save(ROOT / "docs" / "thumbnail.png")
print("docs/thumbnail.png")
