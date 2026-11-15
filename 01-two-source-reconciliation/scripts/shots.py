"""Renders docs/*.png from Two-Source-Reconciliation.xlsx: each sheet's print area -> PDF page (real Excel) -> PNG
cropped to the content. Run under the shared Excel lock:
    python tools/with_lock.py E1 -- python scripts/shots.py
"""
import os, subprocess, sys, tempfile
from pathlib import Path
try:
    import fitz  # PyMuPDF
except ImportError:  # the Excel venv may lack PyMuPDF: the PNG step then runs in a Python that has it
    fitz = None
PYMUPDF_PY = r"C:\Python314\python.exe"

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"
AREAS = {"Summary": ("summary", "A1:Y63"), "Exceptions workbench": ("workbench", "A1:O33"),
         "Reconciliation": ("reconciliation", "A1:U30"), "Cleaning log": ("cleaning-log", "A1:O32"),
         "Settings": ("settings", "A1:E22")}
EXTRA = [("Summary", "summary-payroll", "A1:Y63", ("Payroll", None))]  # the same dashboard, other account picked


def export(ws, area, pdf):
    ps = ws.PageSetup
    ps.Orientation = 2
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1
    ps.PrintArea = area
    ps.LeftMargin = ps.RightMargin = ps.TopMargin = ps.BottomMargin = 10
    ps.CenterHorizontally = False
    ws.ExportAsFixedFormat(0, str(pdf))


def to_png(pdf, png):
    page = fitz.open(pdf)[0]
    bb = None
    for b in page.get_drawings():
        bb = b["rect"] if bb is None else bb | b["rect"]
    for b in page.get_text("blocks"):
        bb = fitz.Rect(b[:4]) if bb is None else bb | fitz.Rect(b[:4])
    for img in page.get_image_info():
        bb = fitz.Rect(img["bbox"]) if bb is None else bb | fitz.Rect(img["bbox"])
    bb = (bb or page.rect) & page.rect
    page.get_pixmap(dpi=170, clip=bb).save(png)
    print("docs/" + png.name)


def main():
    tmp = Path(tempfile.mkdtemp(prefix="e1-shots"))
    DOCS.mkdir(exist_ok=True)
    import win32com.client as win32
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(ROOT / "Two-Source-Reconciliation.xlsx"), 0, True)
        for sheet, (stem, area) in AREAS.items():
            export(wb.Worksheets(sheet), area, tmp / f"{stem}.pdf")
        for sheet, stem, area, (acct, _) in EXTRA:
            wb.Worksheets("Summary").Range("C7").Value = acct
            xl.Calculate()
            export(wb.Worksheets(sheet), area, tmp / f"{stem}.pdf")
        wb.Close(False)
    finally:
        xl.Quit()
    if fitz is None:
        subprocess.run([PYMUPDF_PY, __file__, "--png", str(tmp)], check=True)
    else:
        pngs(tmp)


def pngs(tmp):
    for pdf in sorted(Path(tmp).glob("*.pdf")):
        to_png(pdf, DOCS / f"{pdf.stem}.png")


if __name__ == "__main__":
    if sys.argv[1:2] == ["--png"]:
        import fitz
        pngs(sys.argv[2])
    else:
        main()
