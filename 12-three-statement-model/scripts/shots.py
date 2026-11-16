"""docs/*.png from both workbooks: sheet area -> PDF (Excel's own renderer) -> PNG cropped to content.
The workbook is opened read-only and never saved.
    python scripts/shots.py      (run under tools/with_lock.py; PDF -> PNG uses PyMuPDF from the system Python)
"""
import subprocess, sys, tempfile
from pathlib import Path
import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
TMP = Path(tempfile.gettempdir()) / "e12-shots"
TMP.mkdir(exist_ok=True)
for f in TMP.glob("*.pdf"):
    f.unlink()
BOOK = "Three-Statement-Model.xlsx"
SHOTS = [(BOOK, n, s, a) for n, s, a in (
    ("dashboard", "Dashboard", "A1:Y54"), ("assumptions", "Assumptions", "A1:H76"), ("scenarios", "Scenarios", "A1:H46"),
    ("pl", "P&L", "A1:Q32"), ("balance", "Balance Sheet", "A1:Q31"), ("cashflow", "Cash Flow", "A1:Q29"),
    ("schedules", "Schedules", "A1:Q50"), ("annual", "Annual", "A1:F60"), ("checks", "Checks", "A1:D23"))]

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
        ps.PrintGridlines = False
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
