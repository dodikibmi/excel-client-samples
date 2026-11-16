"""docs/*.png: sheet area -> PDF (Excel's own renderer) -> PNG cropped to content. Opened read-only, never saved.
    python ../../tools/with_lock.py e14 -- python scripts/shots.py
"""
import sys
import tempfile
from pathlib import Path

import pdfplumber
import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
TMP = Path(tempfile.gettempdir()) / "e14-shots"
TMP.mkdir(exist_ok=True)
for f in TMP.glob("*.pdf"):
    f.unlink()
SHOTS = [("dashboard", "Dashboard", "A1:R45"), ("needs-attention", "Needs attention", "A1:H18"),
         ("register", "Register", "A1:Q30"), ("lines", "Lines", "A1:O30"), ("checks", "Checks", "A1:E18"),
         ("labels", "Labels", "A1:E36"), ("suppliers", "Suppliers", "A1:F15"), ("manual-entry", "Manual entry", "A1:M16"),
         ("accounting-import", "Accounting import", "A1:M30")]
only = sys.argv[1:]
xl = win32.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(str(ROOT / "PDF-Invoices-to-Excel.xlsx"), 0, True)
    for name, sheet, area in SHOTS:
        if only and name not in only:
            continue
        ws = wb.Worksheets(sheet)
        ps = ws.PageSetup
        ps.Orientation = 2
        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = 1
        ps.PrintArea = area
        ps.PrintGridlines = False
        for m in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
            setattr(ps, m, 10)
        ws.ExportAsFixedFormat(0, str(TMP / f"{name}.pdf"))
    wb.Close(False)
finally:
    xl.Quit()

(ROOT / "docs").mkdir(exist_ok=True)
for pdf in sorted(TMP.glob("*.pdf")):
    with pdfplumber.open(pdf) as p:
        page = p.pages[0]
        objs = page.chars + page.rects + page.lines + page.images + page.curves
        x0 = max(0, min(o["x0"] for o in objs) - 6); x1 = min(page.width, max(o["x1"] for o in objs) + 6)
        t0 = max(0, min(o["top"] for o in objs) - 6); b1 = min(page.height, max(o["bottom"] for o in objs) + 6)
        crop = page.crop((x0, t0, x1, b1))
        res = min(216, int(1800 / max(1, (x1 - x0)) * 72))
        crop.to_image(resolution=res).save(ROOT / "docs" / f"{pdf.stem}.png")
        print("docs/" + pdf.stem + ".png")
