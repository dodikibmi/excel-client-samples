"""docs/*.png: sheet areas -> PDF (Excel's own renderer) -> PNG cropped to content; quote / invoice pictures are the
real PDFs the macros wrote. The workbook is opened read-only and never saved.
Run under the shared lock:  python tools/with_lock.py E10 -- <python> scripts/shots.py
PDF -> PNG uses PyMuPDF from the system Python (python on PATH)."""
import json, subprocess, tempfile
from pathlib import Path
import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
TMP = Path(tempfile.gettempdir()) / "e10-shots"
TMP.mkdir(exist_ok=True)
for f in TMP.glob("*.pdf"):
    f.unlink()
NUM = json.loads((ROOT / "data" / "numbers.json").read_text())

SHOTS = [("home", "Home", "A1:T50", 2), ("register", "Register", "A1:Z34", 2), ("lines", "Lines", "A1:N30", 2),
         ("log", "Log", "A1:E40", 2), ("checks", "Checks", "A1:E19", 2), ("products", "Products", "A1:G30", 2),
         ("settings", "Settings", "A1:C25", 2), ("how-to-use", "How to use", "A1:A21", 2)]

xl = win32.DispatchEx("Excel.Application")
xl.Visible = False; xl.DisplayAlerts = False; xl.EnableEvents = False
try:
    wb = xl.Workbooks.Open(str(ROOT / "Quote-Invoice-Builder.xlsm"), 0, True)
    for name, sheet, area, orient in SHOTS:
        ws = wb.Worksheets(sheet)
        ws.Unprotect()
        ps = ws.PageSetup
        ps.Orientation = orient; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1; ps.PrintArea = area
        ps.PrintTitleRows = ""
        for m in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
            setattr(ps, m, 10)
        ws.ExportAsFixedFormat(0, str(TMP / f"{name}.pdf"))
    # a draft on the page: the DRAFT mark and the blocked reasons live in the register
    doc = wb.Worksheets("Document"); doc.Unprotect()
    doc.Range("M2").Value = NUM["quotes"]["d2"]; xl.Calculate()
    doc.ExportAsFixedFormat(0, str(TMP / "document-draft.pdf"))
    wb.Close(False)
finally:
    xl.Quit()

from openpyxl import load_workbook
wbv = load_workbook(ROOT / "Quote-Invoice-Builder.xlsm", data_only=True, read_only=True)
rows = [r for r in wbv["Register"].iter_rows(min_row=6, values_only=True) if r[0]]
paid_inv = max((r for r in rows if r[1] == "Invoice" and r[13] == "Paid" and r[10] and r[8]) or
               (r for r in rows if r[1] == "Invoice" and r[13] == "Paid" and r[10]), key=lambda r: (r[6], r[12]))
quote = max((r for r in rows if r[1] == "Quote" and r[13] == "Accepted" and r[8]), key=lambda r: (r[6], r[12]))
for nm, r in (("invoice-pdf", paid_inv), ("quote-pdf", quote)):
    (TMP / f"{nm}.pdf").write_bytes((ROOT / "PDF" / r[18]).read_bytes())
print("invoice", paid_inv[0], "quote", quote[0])

code = r"""
import fitz, pathlib, sys
tmp, docs = pathlib.Path(sys.argv[1]), pathlib.Path(sys.argv[2])
for pdf in sorted(tmp.glob('*.pdf')):
    page = fitz.open(pdf)[0]
    r = fitz.Rect(page.rect)
    if pdf.stem.endswith('-pdf') or pdf.stem == 'document-draft':
        bb = r                                   # a real document: keep the whole A4 page
    else:
        bb = None
        for b in page.get_drawings(): bb = b['rect'] if bb is None else bb | b['rect']
        for b in page.get_text('blocks'): bb = fitz.Rect(b[:4]) if bb is None else bb | fitz.Rect(b[:4])
        for img in page.get_image_info(): bb = fitz.Rect(img['bbox']) if bb is None else bb | fitz.Rect(img['bbox'])
        bb = (bb or r) + (-6, -6, 6, 6)
    zoom = (1400 if bb is r else 1800) / max(1, (bb & r).width)
    page.get_pixmap(matrix=fitz.Matrix(min(zoom, 3), min(zoom, 3)), clip=bb & r).save(docs / (pdf.stem + '.png'))
    print('docs/' + pdf.stem + '.png')
"""
(ROOT / "docs").mkdir(exist_ok=True)
subprocess.run(["python", "-c", code, str(TMP), str(ROOT / "docs")], check=True)
