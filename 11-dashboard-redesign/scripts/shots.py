"""docs/*.png from both workbooks: sheet area -> PDF (Excel's own renderer) -> PNG cropped to content.
Both files are opened read-only and never saved.
    python scripts/shots.py [before]      (run under tools/with_lock.py; PDF -> PNG uses PyMuPDF from the system Python)
"""
import subprocess, sys, tempfile
from pathlib import Path
import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
TMP = Path(tempfile.gettempdir()) / "e11-shots"
TMP.mkdir(exist_ok=True)
for f in TMP.glob("*.pdf"):
    f.unlink()
ONLY_BEFORE = "before" in sys.argv[1:]
SHOTS = [("Before-Client-Dashboard.xlsx", "before", "Dashboard", "A1:Q41")]
if not ONLY_BEFORE:
    SHOTS += [("Sales-Dashboard-Redesign.xlsx", n, s, a) for n, s, a in (
        ("after", "Dashboard", "A1:T52"), ("issues", "Issues found", "A1:H18"), ("bridge", "Bridge", "A1:E37"),
        ("data", "Data", "A1:Q30"), ("checks", "Checks", "A1:E20"))]

xl = win32.DispatchEx("Excel.Application")
xl.Visible = False; xl.DisplayAlerts = False
try:
    books = {}
    for book, name, sheet, area in SHOTS:
        wb = books.get(book) or books.setdefault(book, xl.Workbooks.Open(str(ROOT / book), 0, True))
        ws = wb.Worksheets(sheet)
        ws.Unprotect()
        ps = ws.PageSetup
        ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1; ps.PrintArea = area
        ps.PrintGridlines = (name == "before")          # the old file is looked at with its gridlines
        for m in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
            setattr(ps, m, 10)
        ws.ExportAsFixedFormat(0, str(TMP / f"{name}.pdf"))
    for wb in books.values():
        wb.Close(False)
finally:
    xl.Quit()

code = r"""
import fitz, pathlib, sys
tmp, docs = pathlib.Path(sys.argv[1]), pathlib.Path(sys.argv[2])
for pdf in sorted(tmp.glob('*.pdf')):
    page = fitz.open(pdf)[0]
    r = fitz.Rect(page.rect); bb = None
    for b in page.get_drawings(): bb = b['rect'] if bb is None else bb | b['rect']
    for b in page.get_text('blocks'): bb = fitz.Rect(b[:4]) if bb is None else bb | fitz.Rect(b[:4])
    for img in page.get_image_info(): bb = fitz.Rect(img['bbox']) if bb is None else bb | fitz.Rect(img['bbox'])
    bb = (bb or r) + (-6, -6, 6, 6)
    zoom = 1800 / max(1, (bb & r).width)
    page.get_pixmap(matrix=fitz.Matrix(min(zoom, 3), min(zoom, 3)), clip=bb & r).save(docs / (pdf.stem + '.png'))
    print('docs/' + pdf.stem + '.png')
"""
(ROOT / "docs").mkdir(exist_ok=True)
subprocess.run(["python", "-c", code, str(TMP), str(ROOT / "docs")], check=True)
