"""Bulk PDF export: one commission statement per rep for a month, plus all of them in one PDF.

The workbook itself stays macro-free; this script drives Excel from outside (COM). It opens the workbook read-only,
picks each rep on the Statement sheet, trims the print area to the lines actually listed, exports a PDF per rep and
merges them.  (Inside Excel the same job can be done with the optional macro in scripts/ExportStatements.bas.)

    python scripts/export_statements.py [YYYY-MM]          default: the last month in the plan
Writes docs/statements/<YYYY-MM>-<rep>.pdf and docs/statements/all-reps-<YYYY-MM>.pdf
Run it under the shared Excel lock:  python tools/with_lock.py E3 -- <python> scripts/export_statements.py
"""
import json, sys
from datetime import date
from pathlib import Path
import win32com.client as win32
from pypdf import PdfReader, PdfWriter

ROOT = Path(__file__).resolve().parent.parent
PLAN = json.loads((ROOT / "data" / "plan.json").read_text())
FIRST_LINE = 47          # first order line row on the Statement sheet (see build_workbook.STMT_FIRST + 1)
OUT = ROOT / "docs" / "statements"

first = date.fromisoformat(PLAN["first_month"])
last = date(first.year + (first.month - 2 + PLAN["months"]) // 12, (first.month - 2 + PLAN["months"]) % 12 + 1, 1)
month = date.fromisoformat(sys.argv[1] + "-01") if len(sys.argv) > 1 else last
OUT.mkdir(parents=True, exist_ok=True)

xl = win32.DispatchEx("Excel.Application")
xl.Visible = False; xl.DisplayAlerts = False
files = []
try:
    wb = xl.Workbooks.Open(str(ROOT / "Sales-Commission-Calculator.xlsx"), 0, True)
    s = wb.Worksheets("Statement")
    s.Unprotect("")
    wb.Names("StmtMonth").RefersToRange.Value2 = (month - date(1899, 12, 30)).days
    for rep in [r[0] for r in PLAN["reps"]]:
        wb.Names("StmtRep").RefersToRange.Value = rep
        xl.Calculate()
        n = int(xl.WorksheetFunction.Count(s.Range(f"I{FIRST_LINE}:I{FIRST_LINE + 99}")))
        ps = s.PageSetup
        ps.PrintArea = f"$A$7:$J${FIRST_LINE + n + 1}"          # header, working, YTD, the listed lines and their total
        ps.PrintTitleRows = f"${FIRST_LINE - 1}:${FIRST_LINE - 1}"
        ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = False
        f = OUT / f"{month:%Y-%m}-{rep.replace(' ', '-')}.pdf"
        s.ExportAsFixedFormat(0, str(f))
        files.append(f)
        print(f"{rep}: {n} lines -> {f.relative_to(ROOT)}")
    wb.Close(False)
finally:
    xl.Quit()

merged = PdfWriter()
for f in files:
    for page in PdfReader(f).pages:
        merged.add_page(page)
all_pdf = OUT / f"all-reps-{month:%Y-%m}.pdf"
with open(all_pdf, "wb") as fh:
    merged.write(fh)
print(f"{len(files)} statements, {len(merged.pages)} pages -> {all_pdf.relative_to(ROOT)}")
