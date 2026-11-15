"""docs/thumbnail.png (1200x627, Upwork portfolio card) in the dark-gold style: claim, CSV -> Power Query -> 0.00 flow and
proof numbers on the left, the real Summary screenshot on the right. Numbers are read from the saved workbook
(cached values) and from test/verify_result.json, so they always match the last build and test run.
Also docs/before-raw-csv.png: what the client actually exports (the "before" picture)."""
import json
from pathlib import Path
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
W, H = 1200, 627
INK, PAGE, CARD, LINE = (11, 15, 19), (18, 23, 29), (26, 33, 41), (42, 51, 61)
GOLD, GOLD2, IVORY, MUTED, GOOD, JADE = (201, 163, 91), (230, 207, 154), (236, 230, 218), (142, 152, 163), (91, 191, 138), (79, 163, 140)
F = "C:/Windows/Fonts/"
font = lambda name, size: ImageFont.truetype(F + name, size)
SERIF, SERIFB, SANS, SANSB = "georgia.ttf", "georgiab.ttf", "segoeui.ttf", "segoeuib.ttf"


def numbers():
    wb = load_workbook(ROOT / "Two-Source-Reconciliation.xlsx", data_only=True, read_only=False)
    names = wb.defined_names

    def val(nm):
        sheet, ref = list(names[nm].destinations)[0]
        return wb[sheet][ref.replace("$", "")].value
    res = json.loads((ROOT / "test" / "verify_result.json").read_text(encoding="utf-8"))
    return dict(rate=val("MatchRate"), unexplained=val("StUnexplained"), periods=val("PeriodsChecked"),
                failed=val("PeriodsFailed"), checks=f"{res['passed']}/{res['total']}", account=wb["Summary"]["C7"].value)


def main():
    n = numbers()
    assert n["unexplained"] == 0 and n["failed"] == 0, n
    img = Image.new("RGB", (W, H), INK)
    d = ImageDraw.Draw(img)

    # brand, title, subtitle, gold rule
    x = 56
    d.text((x, 40), "  ".join("LAKESIDE"), font=font(SANSB, 13), fill=GOLD)
    d.text((x, 62), "Bank reconciliation", font=font(SERIF, 44), fill=IVORY)
    d.text((x, 118), "in Excel + Power Query, balanced to 0.00", font=font(SANS, 20), fill=MUTED)
    d.line((x, 156, W - 56, 156), fill=GOLD, width=2)

    # flow chips: CSV -> Power Query -> 0.00
    y = 178
    cx = x
    for i, (t, col) in enumerate([("24 CSV files", IVORY), ("Power Query", IVORY), ("0.00", GOOD)]):
        f = font(SANSB if i == 2 else SANS, 15)
        w = int(d.textlength(t, font=f)) + 26
        d.rounded_rectangle((cx, y, cx + w, y + 32), 7, fill=CARD, outline=GOLD if i == 2 else LINE, width=1)
        d.text((cx + 13, y + 6), t, font=f, fill=col)
        cx += w
        if i < 2:
            d.text((cx + 7, y + 3), "→", font=font("seguisym.ttf", 18), fill=GOLD)
            cx += 30

    # proof numbers
    points = [(f"{n['rate']:.1%}".replace(".0%", "%"), "of October bank lines matched by 3 rules"),
              (f"{n['periods']}/{n['periods']}", "account-months balance to 0.00"),
              ("Aging", "0-7 / 8-30 / 30+ days, owners + comments"),
              (n["checks"], "automated checks pass (recomputed in Python)")]
    yy = 238
    for big, small in points:
        d.text((x, yy), big, font=font(SERIF, 34), fill=GOLD2)
        d.text((x, yy + 44), small, font=font(SANS, 15), fill=IVORY)
        yy += 80
    d.text((x, H - 44), "Workbench  ·  tolerance rules  ·  2 accounts x 6 months  ·  no macros", font=font(SANS, 14), fill=MUTED)

    # real screenshot on the right: header, KPI cards and the three charts
    shot = Image.open(ROOT / "docs" / "summary.png").convert("RGB")
    crop = shot.crop((0, 0, shot.width, int(shot.height * 0.505)))
    cw = 650
    crop = crop.resize((cw, int(crop.height * cw / crop.width)), Image.LANCZOS)
    sx, sy = W - cw - 40, 178
    shadow = Image.new("RGBA", (cw + 60, crop.height + 60), (0, 0, 0, 0))
    ImageDraw.Draw(shadow).rectangle((30, 30, cw + 30, crop.height + 30), fill=(0, 0, 0, 170))
    shadow = shadow.filter(ImageFilter.GaussianBlur(14))
    img.paste(shadow, (sx - 30, sy - 22), shadow)
    img.paste(crop, (sx, sy))
    d.rectangle((sx - 1, sy - 1, sx + cw, sy + crop.height), outline=GOLD, width=1)
    img.save(ROOT / "docs" / "thumbnail.png")
    print("docs/thumbnail.png", n)


def before():
    lines = (ROOT / "data" / "bank_Operating_2026-10.csv").read_text(encoding="utf-8").splitlines()[:9]
    mono = font("consola.ttf", 20)
    pad, lh = 30, 30
    probe = ImageDraw.Draw(Image.new("RGB", (10, 10)))
    bw = max(int(probe.textlength(l, font=mono)) for l in lines) + 2 * pad
    img = Image.new("RGB", (max(bw, 1060), len(lines) * lh + 2 * pad + 44), PAGE)
    d = ImageDraw.Draw(img)
    d.text((pad, 16), "bank_Operating_2026-10.csv  as exported: text amounts, US dates, references typed 4 ways",
           font=font(SANSB, 18), fill=(229, 115, 95))
    d.line((pad, 48, img.width - pad, 48), fill=GOLD, width=1)
    for i, l in enumerate(lines):
        d.text((pad, pad + 34 + i * lh), l, font=mono, fill=IVORY if i else GOLD)
    img.save(ROOT / "docs" / "before-raw-csv.png")
    print("docs/before-raw-csv.png")


if __name__ == "__main__":
    before()
    main()
