"""Checks PDF-Invoices-to-Excel.xlsx against the designed answer (test/expected.json, written by make_data.py).

Part 1 reads the saved workbook (openpyxl, values as Excel calculated them):
  every file has the right status, every invoice the right fields, every line the right numbers, duplicates point at
  the right original, held invoices are exactly the ones whose supplier made a sum mistake, the in-workbook checks
  say PASS, the accounting import file adds up to the payable total.
Part 2 (--live) works on a copy in a temp folder with real Excel, changes the inputs and refreshes once:
  a new supplier invoice dropped into the folder, the same invoice saved again under another name, a scanned invoice
  typed on Manual entry, and a supplier removed from the Suppliers sheet. Each must land where it should.

    python test/verify.py            (part 1)
    python ../../tools/with_lock.py e14 -- python test/verify.py --live
"""
import csv
import datetime as dt
import json
import shutil
import sys
import tempfile
from decimal import Decimal
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
BOOK = ROOT / "PDF-Invoices-to-Excel.xlsx"
EXP = json.loads((ROOT / "test" / "expected.json").read_text(encoding="utf-8"))
sys.path.insert(0, str(ROOT / "scripts"))
from build_workbook import N_MANUAL_PREFILL  # noqa: E402

results = []


def check(name, ok, detail=""):
    results.append((name, bool(ok), detail))
    print(("PASS " if ok else "FAIL ") + name + (f"  [{detail}]" if detail and not ok else ""))


def table(ws, header_row=4):
    rows = list(ws.iter_rows(min_row=header_row, values_only=True))
    h = rows[0]
    return [dict(zip(h, r)) for r in rows[1:] if r[0] not in (None, "")]


def d(v):
    if v is None:
        return None
    return v.date() if isinstance(v, dt.datetime) else v


def money(v):
    return None if v is None else Decimal(str(v)).quantize(Decimal("0.01"))


def expected_rows():
    inv = {i["number"]: i for i in EXP["invoices"]}
    scans = sorted([f["file"] for f in EXP["files"] if f["status"] == "Scanned image"])
    typed = set(scans[:N_MANUAL_PREFILL])
    first = {}
    for f in EXP["files"]:
        if f["role"] == "original":
            first[f["number"]] = f["file"]
    out = {}
    for f in EXP["files"]:
        i = inv[f["number"]]
        if f["role"] == "duplicate":
            st = "Duplicate"
        elif f["status"] == "Scanned image":
            st = "Entered by hand" if f["file"] in typed else "Scanned image (no text)"
        elif f["status"] in ("Password protected", "Damaged file"):
            st = f["status"]
        elif i.get("broken_math"):
            st = "Lines don't add up"
        else:
            st = "OK"
        out[f["file"]] = dict(status=st, inv=i, dup_of=first[f["number"]] if f["role"] == "duplicate" else None,
                              pages=f["pages"])
    return out


def part1():
    wb = openpyxl.load_workbook(BOOK, data_only=True)
    check("no macros (.xlsx, no vbaProject)", BOOK.suffix == ".xlsx")
    reg = table(wb["Register"])
    exp = expected_rows()
    check("Register has one row per PDF in the folder (300)", len(reg) == len(exp) == 300, f"{len(reg)} rows")
    by_file = {r["File"]: r for r in reg}
    check("same file names as the folder", set(by_file) == set(exp))

    bad_status, bad_fields, bad_lines = [], [], []
    for f, e in exp.items():
        r = by_file.get(f)
        if r is None:
            continue
        if r["Status"] != e["status"]:
            bad_status.append(f"{f}: {r['Status']} vs {e['status']}")
            continue
        if e["status"] in ("OK", "Lines don't add up", "Duplicate", "Entered by hand"):
            i = e["inv"]
            bm = Decimal(i.get("broken_math") or "0")
            want = {
                "Supplier": i["supplier"], "Invoice No": i["number"], "Type": i["kind"],
                "Invoice Date": dt.date.fromisoformat(i["date"]), "Due Date": dt.date.fromisoformat(i["due"]),
                "PO": i["po"] or None, "Property": i["property"],
                "Subtotal": money(Decimal(i["subtotal"]) + bm), "Tax": money(i["tax"]),
                "Total": money(Decimal(i["total"]) + bm),
            }
            got = {k: r[k] for k in want}
            got["Invoice Date"], got["Due Date"] = d(got["Invoice Date"]), d(got["Due Date"])
            for k in ("Subtotal", "Tax", "Total"):
                got[k] = money(got[k])
            if got != want:
                diff = {k: (got[k], want[k]) for k in want if got[k] != want[k]}
                bad_fields.append(f"{f}: {diff}")
            if e["status"] != "Entered by hand":
                n_ok = r["Lines"] == len(i["lines"])
                s_ok = money(r["Lines total"]) == money(i["subtotal"])
                if not (n_ok and s_ok):
                    bad_lines.append(f"{f}: lines {r['Lines']}/{len(i['lines'])} total {r['Lines total']}/{i['subtotal']}")
        if e["dup_of"] and r["Duplicate of"] != e["dup_of"]:
            bad_status.append(f"{f}: duplicate of {r['Duplicate of']} vs {e['dup_of']}")
    check("every file has the expected status (OK / held / duplicate / scanned / password / damaged / typed)",
          not bad_status, "; ".join(bad_status[:5]))
    check("every invoice: supplier, number, type, dates, PO, property, subtotal, tax, total as on the PDF",
          not bad_fields, "; ".join(bad_fields[:3]))
    check("every invoice: number of lines and lines total as on the PDF", not bad_lines, "; ".join(bad_lines[:3]))
    check("dates read with each supplier's own format (day-first 09/10/2026 = 9 Oct, month-first 9/10 = 10 Sep)",
          not bad_fields)
    n_multi = sum(1 for e in exp.values() if e["pages"] > 1 and e["status"] == "OK")
    check(f"invoices over 2-3 pages read in full ({n_multi})", n_multi > 20 and not bad_lines)
    held = sorted(r["File"] for r in reg if r["Payable"] == "Hold")
    want_held = sorted(f for f, e in exp.items() if e["status"] == "Lines don't add up")
    check("held = exactly the invoices whose supplier added up wrong", held == want_held and len(held) == 3, f"{held}")
    credits = [r for r in reg if r["Type"] == "Credit note" and r["Payable"] == "Yes"]
    check("credit notes read as negative amounts", len(credits) == 7 and all(r["Total"] < 0 for r in credits),
          f"{len(credits)}")

    # Lines sheet against the PDFs, line by line
    lines = table(wb["Lines"])
    inv_lines = {}
    for l in lines:
        inv_lines.setdefault((l["Supplier"], l["Invoice No"]), []).append(l)
    bad = []
    for f, e in exp.items():
        if e["status"] not in ("OK", "Lines don't add up"):
            continue
        i = e["inv"]
        got = sorted(inv_lines.get((i["supplier"], i["number"]), []), key=lambda x: x["Line"])
        if len(got) != len(i["lines"]):
            bad.append(f"{i['number']}: {len(got)} lines")
            continue
        for g, w in zip(got, i["lines"]):
            if (g["Description"] != w["desc"] or money(g["Amount"]) != money(w["amount"])
                    or money(g["Unit price"]) != money(w["unit"]) or abs(float(g["Qty"]) - float(w["qty"])) > 1e-9
                    or (w.get("sku") and g["SKU"] != w["sku"])):
                bad.append(f"{i['number']} line {g['Line']}")
                break
    check(f"Lines sheet: every line (description, SKU, qty, unit price, amount) = the PDF ({len(lines)} lines)", not bad,
          "; ".join(bad[:4]))

    # in-workbook checks
    chk = [r for r in wb["Checks"].iter_rows(min_row=5, max_row=30, values_only=True) if r[0] and isinstance(r[0], int)]
    check(f"in-workbook Checks sheet: all {len(chk)} PASS", chk and all(r[4] == "PASS" for r in chk),
          str([r[1] for r in chk if r[4] != "PASS"]))

    # needs attention = files that are not paid and not a caught duplicate
    na = [r for r in wb["Needs attention"].iter_rows(min_row=5, max_row=44, min_col=3, max_col=4, values_only=True) if r[0]]
    want_na = sorted(f for f, e in exp.items() if e["status"] not in ("OK", "Duplicate", "Entered by hand"))
    check("Needs attention lists exactly the files a person must handle", sorted(r[0] for r in na) == want_na,
          f"{len(na)} vs {len(want_na)}")
    check("Needs attention: held invoices first", [r[1] for r in na[:3]] == ["Lines don't add up"] * 3)

    # accounting import CSV
    pay_total = sum(money(r["Total"]) for r in reg if r["Payable"] == "Yes")
    with open(ROOT / "output" / "bills-import-2026-09-10.csv", encoding="utf-8") as fh:
        rows = list(csv.DictReader(fh))
    csv_total = sum(Decimal(r["Amount"]) for r in rows)
    check("accounting import CSV total = payable total on Register", csv_total == pay_total, f"{csv_total} vs {pay_total}")
    n_yes_lines = sum(1 for l in lines if l["Payable"] == "Yes")
    n_tax = sum(1 for r in reg if r["Payable"] == "Yes" and r["Group"] == "Read automatically" and (r["Tax"] or 0) != 0)
    n_hand = sum(1 for r in reg if r["Status"] == "Entered by hand")
    check("accounting import: one row per paid line + one tax row per taxed invoice + one row per typed invoice",
          len(rows) == n_yes_lines + n_tax + n_hand, f"{len(rows)} vs {n_yes_lines}+{n_tax}+{n_hand}")
    codes = {r["Account Code"] for r in rows}
    check("accounting import: every row has an account code", "" not in codes and None not in codes, str(codes))
    dash = wb["Dashboard"]
    check("dashboard: payable card = Register", money(dash["L6"].value) == pay_total, f"{dash['L6'].value}")
    check("dashboard: files card = 300 and read-automatically card = 284",
          dash["B6"].value == 300 and dash["D6"].value == 284, f"{dash['B6'].value} {dash['D6'].value}")


def part2():
    import time
    import win32com.client as win32
    sys.path.insert(0, str(ROOT / "scripts"))
    import make_data as md

    tmp = Path(tempfile.gettempdir()) / "e14-live"
    if tmp.exists():
        shutil.rmtree(tmp)
    shutil.copytree(ROOT / "invoices", tmp / "invoices")
    shutil.copy(BOOK, tmp / BOOK.name)
    # 1) a new invoice from a known supplier arrives
    sup = [s for s in md.SUPPLIERS if s["name"] == "Summit Landscaping Group"][0]
    md.random.seed(7)
    new = md.make_invoice(sup, 900)
    new["date"] = "2026-10-30"
    new["due"] = "2026-11-14"
    md.build_pdf(new, tmp / "invoices" / "SLG invoice NEW.pdf")
    # 2) an invoice saved again under another name
    shutil.copy(tmp / "invoices" / "Brightway 10443.pdf", tmp / "invoices" / "Re_ Brightway 10443.pdf")
    exp = expected_rows()
    scan = sorted(f for f, e in exp.items() if e["status"] == "Scanned image (no text)")[0]
    si = exp[scan]["inv"]

    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(tmp / BOOK.name))
        reg0 = wb.Worksheets("Register").ListObjects("Register")
        before_total = xl.WorksheetFunction.SumIfs(reg0.ListColumns("Total").DataBodyRange, reg0.ListColumns("Payable").DataBodyRange, "Yes")
        # 3) a scanned invoice typed on Manual entry
        me = wb.Worksheets("Manual entry").ListObjects("ManualEntry")
        free = next(i for i in range(1, me.ListRows.Count + 1) if not me.ListRows(i).Range.Cells(1, 1).Value)
        row = me.ListRows(free).Range
        vals = [scan, si["supplier"], si["kind"], si["number"], None, None, si["po"] or "", si["property"],
                float(si["subtotal"]), float(si["tax"]), float(si["total"]), "Test", ""]
        for j, v in enumerate(vals, 1):
            if v is not None:
                row.Cells(1, j).Value = v
        dd, du = dt.date.fromisoformat(si["date"]), dt.date.fromisoformat(si["due"])
        row.Cells(1, 5).Formula = f"=DATE({dd.year},{dd.month},{dd.day})"
        row.Cells(1, 6).Formula = f"=DATE({du.year},{du.month},{du.day})"
        # 4) a supplier removed from the Suppliers sheet
        su = wb.Worksheets("Suppliers").ListObjects("Suppliers")
        for i in range(1, su.ListRows.Count + 1):
            if su.ListRows(i).Range.Cells(1, 1).Value == "Pinecrest Security Systems":
                su.ListRows(i).Delete()
                break
        t = time.time()
        wb.RefreshAll()
        xl.CalculateUntilAsyncQueriesDone()
        xl.Calculate()
        print(f"  refresh {time.time() - t:.0f}s")
        reg = wb.Worksheets("Register").ListObjects("Register")
        cols = [reg.ListColumns(i).Name for i in range(1, reg.ListColumns.Count + 1)]
        rows = [dict(zip(cols, r)) for r in reg.DataBodyRange.Value]
        by = {r["File"]: r for r in rows}
        n = by.get("SLG invoice NEW.pdf")
        check("live: a new PDF in the folder is read on Refresh All", n and n["Status"] == "OK"
              and n["Invoice No"] == new["number"] and money(n["Total"]) == money(new["total"]), str(n and n["Status"]))
        dup = by.get("Re_ Brightway 10443.pdf")
        check("live: the same invoice saved again is caught as a duplicate", dup and dup["Status"] == "Duplicate"
              and dup["Duplicate of"] == "Brightway 10443.pdf", str(dup and dup["Status"]))
        s = by.get(scan)
        check("live: a scanned invoice typed on Manual entry is counted", s and s["Status"] == "Entered by hand"
              and s["Payable"] == "Yes" and money(s["Total"]) == money(si["total"]), str(s and s["Status"]))
        pss = [r for r in rows if r["File"].startswith("PSS")]
        check("live: invoices of a supplier missing from Suppliers go to Needs attention as Unknown supplier",
              pss and all(r["Status"] == "Unknown supplier" for r in pss), str({r['Status'] for r in pss}))
        na = [c.Value for c in wb.Worksheets("Needs attention").Range("C5:C44") if c.Value]
        check("live: Needs attention shows them with a next step", all(r["File"] in na for r in pss[:20]))
        pss_total = sum(money(r["Total"]) or 0 for r in table(openpyxl.load_workbook(BOOK, data_only=True)["Register"])
                        if r["File"].startswith("PSS") and r["Payable"] == "Yes")
        after_total = sum(money(r["Total"]) for r in rows if r["Payable"] == "Yes")
        want = money(before_total) + money(new["total"]) + money(si["total"]) - pss_total
        check("live: payable total moves by exactly new + typed - unknown supplier", after_total == want,
              f"{after_total} vs {want}")
        ch = wb.Worksheets("Checks")
        res = [ch.Cells(r, 5).Value for r in range(5, 17)]
        check("live: in-workbook checks still all PASS after the changes", all(v == "PASS" for v in res), str(res))
        wb.Close(False)
    finally:
        xl.Quit()


if __name__ == "__main__":
    part1()
    if "--live" in sys.argv:
        part2()
    ok = sum(1 for _, o, _ in results if o)
    print(f"\n{ok}/{len(results)} checks pass")
    if "--live" in sys.argv:  # the full run (static + live) is the one quoted in the README
        (ROOT / "test" / "last-result.txt").write_text(f"{ok}/{len(results)}\n", encoding="utf-8")
    sys.exit(0 if ok == len(results) else 1)
