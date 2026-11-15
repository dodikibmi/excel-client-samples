"""Renders docs/*.png from Theme-Gallery.xlsx: each sheet -> PDF (Excel's own renderer) -> PNG cropped to content.
The workbook is opened read-only and never saved.
Run under the shared lock:  python tools/with_lock.py E0 -- <python> scripts/shots.py
PDF -> PNG uses PyMuPDF from the system Python (python on PATH).
"""
import os, subprocess, sys, tempfile
from pathlib import Path
import win32com.client as win32

sys.path.insert(0, os.path.dirname(__file__))
from themes import THEMES, sheet_name

ROOT = Path(__file__).resolve().parent.parent
TMP = Path(tempfile.gettempdir()) / "e10-shots"
TMP.mkdir(exist_ok=True)
for f in TMP.glob("*.pdf"):
    f.unlink()

SHOTS = [("themes", "Themes"), ("palette", "Palette"), ("how-to-use", "How to use")]
SHOTS += [(f"theme-{i + 1}-{t['key']}", sheet_name(i)) for i, t in enumerate(THEMES)]

xl = win32.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(str(ROOT / "Theme-Gallery.xlsx"), 0, True)
    xl.Calculate()
    for name, sheet in SHOTS:
        wb.Worksheets(sheet).ExportAsFixedFormat(0, str(TMP / f"{name}.pdf"))
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
    pad = 8 if pdf.stem == 'how-to-use' else 0
    bb = (bb or r) + (-pad, -pad, pad, pad)
    zoom = 1800 / max(1, (bb & r).width)
    page.get_pixmap(matrix=fitz.Matrix(min(zoom, 3), min(zoom, 3)), clip=bb & r).save(docs / (pdf.stem + '.png'))
    print('docs/' + pdf.stem + '.png')
"""
subprocess.run(["python", "-c", code, str(TMP), str(ROOT / "docs")], check=True)
