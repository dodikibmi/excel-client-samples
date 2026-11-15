"""Renders docs/*.png from Sales-Dashboard.xlsx: each page -> PDF (Excel's own renderer) -> PNG cropped to content.
The workbook is opened read-only and never saved. Slicers are set through their caches, exactly like a click.
Run under the shared lock:  python tools/with_lock.py E2 -- <python> scripts/shots.py
PDF -> PNG uses PyMuPDF from the system Python (python on PATH).
"""
import os, subprocess, sys, tempfile
from pathlib import Path
import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
TMP = Path(tempfile.gettempdir()) / "e2-shots"
TMP.mkdir(exist_ok=True)
for f in TMP.glob("*.pdf"):
    f.unlink()

REGIONS = ["North", "South", "East", "West"]


def set_items(wb, cache, wanted=None):
    sc = wb.SlicerCaches(cache)
    if wanted is None:
        sc.ClearManualFilter(); return
    for it in sc.SlicerItems:
        if it.Name in wanted:
            it.Selected = True
    for it in sc.SlicerItems:
        if it.Name not in wanted:
            it.Selected = False


SHOTS = [  # file, sheet, view, regions
    ("dashboard-all", "Dashboard", "MTD", None),
    ("dashboard-west", "Dashboard", "MTD", ["West"]),
    ("dashboard-ytd", "Dashboard", "YTD", None),
    ("products", "Products", "MTD", None),
    ("customers", "Customers", "MTD", None),
    ("reps", "Reps", "MTD", None),
    ("forecast", "Forecast", "MTD", None),
    ("checks", "Checks", "MTD", None),
]
LIGHT = {"calc": ("Calc", "A1:I38"), "sales-data": ("Sales data", "A1:R25"), "how-to-use": ("How to use", "A1:A20")}

xl = win32.DispatchEx("Excel.Application")
xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(str(ROOT / "Sales-Dashboard.xlsx"), 0, True)
    wb.SlicerCaches("Timeline_Month").TimelineState.SetFilterDateRange("2026-10-01", "2026-10-31")
    for name, sheet, view, regions in SHOTS:
        set_items(wb, "Slicer_View", [view])
        set_items(wb, "Slicer_Region", regions)
        xl.Calculate()
        wb.Worksheets(sheet).ExportAsFixedFormat(0, str(TMP / f"{name}.pdf"))
    set_items(wb, "Slicer_View", ["MTD"]); set_items(wb, "Slicer_Region", None); xl.Calculate()
    for name, (sheet, area) in LIGHT.items():
        ws = wb.Worksheets(sheet)
        ps = ws.PageSetup
        ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1; ps.PrintArea = area
        for m in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
            setattr(ps, m, 18)
        ws.ExportAsFixedFormat(0, str(TMP / f"{name}.pdf"))
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
    pad = 0 if pdf.stem not in ('calc', 'sales-data', 'how-to-use') else 8
    bb = (bb or r) + (-pad, -pad, pad, pad)
    zoom = 1800 / max(1, (bb & r).width)
    page.get_pixmap(matrix=fitz.Matrix(min(zoom, 3), min(zoom, 3)), clip=bb & r).save(docs / (pdf.stem + '.png'))
    print('docs/' + pdf.stem + '.png')
"""
subprocess.run(["python", "-c", code, str(TMP), str(ROOT / "docs")], check=True)
