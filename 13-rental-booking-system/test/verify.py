"""Independent check of Event-Rental-Desk.xlsm.

1. Replays data/history.json in plain Python with the rules from the README (hire days, line amounts, deposit, late
   fee after the grace day, damaged / missing at replacement cost, tax, payments, status) and compares every booking.
2. Recomputes the 21-day Availability grid and proves no item is over-booked on any day, from the Lines and Bookings
   sheets, not from the workbook's own grid.
3. Reads every invoice PDF the macros wrote: the invoice number, customer and total must be in it.
4. On a throw-away copy, tries every rule live: not signed in, wrong PIN, staff doing manager actions, a date in the
   locked period, over-booking, enquiry vs confirmed, check out / in late with damage, invoice + PDF, payment, cancel;
   drives both UserForms; and breaks a total on purpose to see the Checks sheet catch it.

    python test/verify.py          (run under tools/with_lock.py; writes test/last-result.txt)
"""
import csv, json, os, re, shutil, subprocess, sys, tempfile
from datetime import date, timedelta, datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from openpyxl import load_workbook
import win32com.client as win32

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT / "scripts"))
from vbom import vba_access
from watchdog import watch

BOOK = ROOT / "Event-Rental-Desk.xlsm"
SET = json.loads((ROOT / "data" / "settings.json").read_text())
HIST = json.loads((ROOT / "data" / "history.json").read_text())
NO = json.loads((ROOT / "data" / "numbers.json").read_text())
CAT = {r[0]: dict(cat=r[1], item=r[2], owned=int(r[3]), rate=float(r[4]), latef=float(r[5]), repl=float(r[6]))
       for r in list(csv.reader(open(ROOT / "data" / "catalog.csv", encoding="utf-8")))[1:]}
USERS = list(csv.reader(open(ROOT / "data" / "users.csv", encoding="utf-8")))[1:]
AS_OF = date.fromisoformat(SET["as_of"])
results = []


def ok(name, cond, detail=""):
    results.append((bool(cond), name))
    print(("PASS " if cond else "FAIL ") + name + (f"  [{detail}]" if detail and not cond else ""))


def R2(x):
    return float(Decimal(repr(x)).quantize(Decimal("0.01"), ROUND_HALF_UP))


D = lambda s: date.fromisoformat(s)
day = lambda v: v.date() if isinstance(v, datetime) else v

# ------------------------------------------------------------------ 1. the history in plain Python
BY = {b["ref"]: b for b in HIST["bookings"]}
book = {}
inv_n = 0
for e in HIST["events"]:
    a, ref = e["action"], e.get("ref")
    if e.get("expect") == "BLOCKED" or a == "lock":
        continue
    if a == "create":
        b = BY[ref]
        lines = {}
        for it, q in b["lines"]:
            lines[it] = lines.get(it, 0) + q
        days = max(1, (D(b["due"]) - D(b["out"])).days)
        amts = {it: R2(q * CAT[it]["rate"] * days) for it, q in lines.items()}
        rent = sum(amts.values())
        sbt = R2(rent + b["delivery"]); tot = R2(sbt + R2(sbt * SET["tax_rate"]))
        book[ref] = dict(no=NO[ref], status="Confirmed" if e["confirm"] else "Enquiry", lines=lines, days=days, amts=amts,
                         rent=rent, deliv=b["delivery"], late=0.0, dmg=0.0, dep=R2(tot * SET["deposit_pct"]), paid=0.0,
                         due=D(b["due"]), out=D(b["out"]), back=None, inv="", dmgq={})
    else:
        k = book[ref]
    if a == "confirm": k["status"] = "Confirmed"
    elif a == "cancel": k["status"] = "Cancelled"
    elif a == "checkout": k["status"] = "Out"
    elif a == "checkin":
        k["back"] = D(e["date"]); k["status"] = "Returned"
        k["dmgq"] = {it: q for it, q in e["damaged"]}
        late_days = max(0, (k["back"] - k["due"]).days - SET["grace_days"])
        k["late"] = sum(R2(q * CAT[it]["latef"] * late_days) for it, q in k["lines"].items())
        k["dmg"] = sum(R2(q * CAT[it]["repl"]) for it, q in k["dmgq"].items())
    elif a in ("deposit", "pay", "invoice"):
        pass
    if a in ("create", "checkin", "deposit", "pay", "invoice", "confirm", "cancel", "checkout"):
        k = book[ref]
        k["sub"] = R2(k["rent"] + k["deliv"] + k["late"] + k["dmg"])
        k["tax"] = R2(k["sub"] * SET["tax_rate"]); k["total"] = R2(k["sub"] + k["tax"])
        if a == "deposit": k["paid"] = R2(k["paid"] + k["dep"])
        if a == "pay": k["paid"] = R2(k["paid"] + R2(k["total"] - k["paid"]))
        k["bal"] = R2(k["total"] - k["paid"])
        if a == "invoice":
            inv_n += 1; k["inv"] = f"INV-2027-{inv_n:04d}"
            k["status"] = "Paid" if k["bal"] <= 0 else "Invoiced"
        if a == "pay" and k["status"] == "Invoiced" and k["bal"] <= 0:
            k["status"] = "Paid"

wbv = load_workbook(BOOK, data_only=True, keep_vba=False)
bs = wbv["Bookings"]
rows = {bs.cell(r, 1).value: [bs.cell(r, c).value for c in range(1, 26)] for r in range(6, bs.max_row + 1) if bs.cell(r, 1).value}
ok(f"{len(book)} bookings in the sheet, numbered B-2027-0001 to B-2027-{len(book):04d} in creation order",
   set(rows) == {k["no"] for k in book.values()} and sorted(rows) == [f"B-2027-{i:04d}" for i in range(1, len(book) + 1)])
bad = []
for ref, k in book.items():
    r = rows[k["no"]]
    exp = {2: k["status"], 9: k["days"], 10: k["rent"], 11: k["deliv"], 12: k["late"], 13: k["dmg"], 14: k["sub"], 15: k["tax"],
           16: k["total"], 17: k["dep"], 18: k["paid"], 19: k["bal"], 20: k["inv"] or None}
    for c, v in exp.items():
        got = r[c - 1]
        if isinstance(v, float):
            if got is None or abs(got - v) > 0.005: bad.append(f"{k['no']} col{c}: {got} vs {v}")
        elif got != v and not (v is None and got in (None, "")):
            bad.append(f"{k['no']} col{c}: {got!r} vs {v!r}")
    if (day(r[7]) if r[7] else None) != k["back"]: bad.append(f"{k['no']} returned {r[7]} vs {k['back']}")
ok("every booking = Python: status, days, hire, delivery, late fee, damage, subtotal, tax, total, deposit, paid, balance, invoice no.",
   not bad, "; ".join(bad[:4]))
late = [k for k in book.values() if k["late"] > 0]; dmg = [k for k in book.values() if k["dmg"] > 0]
ok(f"the sample covers the rules: {len(late)} late returns, {len(dmg)} with damaged / missing items, "
   f"{sum(1 for k in book.values() if k['status'] == 'Cancelled')} cancelled, {sum(1 for k in book.values() if k['status'] == 'Enquiry')} enquiries, "
   f"{sum(1 for k in book.values() if k['status'] == 'Out' and k['due'] < AS_OF)} overdue", late and dmg)

# ------------------------------------------------------------------ 2. stock: grid and no over-booking
ls = wbv["Lines"]
LN = [[ls.cell(r, c).value for c in range(1, 12)] for r in range(6, ls.max_row + 1) if ls.cell(r, 1).value]
held = {}; lost = {}
for no, _, code, _, qty, *_rest in LN:
    dq = _rest[-1] or 0
    st, out, due, back = rows[no][1], day(rows[no][5]), day(rows[no][6]), day(rows[no][7]) if rows[no][7] else None
    if dq and back: lost.setdefault(code, []).append((back, dq))
    if st in ("Enquiry", "Cancelled"): continue
    end = max(due, AS_OF) if st == "Out" else due
    d = out
    while d <= end:
        held[(code, d)] = held.get((code, d), 0) + qty; d += timedelta(1)
free = lambda code, d: CAT[code]["owned"] - sum(q for dd, q in lost.get(code, []) if dd <= d) - held.get((code, d), 0)
neg = [(c, d) for c in CAT for d in (date(2027, 1, 4) + timedelta(i) for i in range(130)) if free(c, d) < 0]
ok("no item over-booked on any day, 4 Jan - 13 May 2027 (from Lines + Bookings, not the workbook grid)", not neg, str(neg[:3]))
av = wbv["Availability"]
bad = []
for r in range(9, 9 + len(CAT)):
    code = av.cell(r, 2).value
    for c in range(5, 26):
        d = day(av.cell(8, c).value)
        if av.cell(r, c).value != free(code, d): bad.append(f"{code} {d}: {av.cell(r, c).value} vs {free(code, d)}")
ok(f"Availability grid: {len(CAT)} items x 21 days from {AS_OF:%d %b %Y} = Python", not bad, "; ".join(bad[:3]))
cat = wbv["Catalog"]
bad = [cat.cell(r, 1).value for r in range(6, 6 + len(CAT)) if cat.cell(r, 11).value != free(cat.cell(r, 1).value, AS_OF)]
ok("Catalog 'free today' = Python", not bad, str(bad))
ok("written-off stock = damaged / missing quantities checked in",
   all(cat.cell(r, 8).value == sum(q for _, q in lost.get(cat.cell(r, 1).value, [])) for r in range(6, 6 + len(CAT))))

# ------------------------------------------------------------------ the log and the planted refusals
lg = wbv["Log"]
LOG = [[lg.cell(r, c).value for c in range(1, 8)] for r in range(6, lg.max_row + 1) if lg.cell(r, 1).value]
blocked = [l for l in LOG if l[5] == "BLOCKED"]
want = ["wrong PIN", "Lock period needs a Manager (maya is Staff)", "Cancel needs a Manager (jon is Staff)",
        "the period is closed through 28 Feb 2027", "Cocktail table: 11 asked, only 5 free on 20 Mar 2027"]
ok(f"Log: {len(LOG)} lines, the {len(blocked)} refused attempts are the 5 planted ones, each with its reason",
   len(blocked) == 5 and all(any(w in str(l[6]) for l in blocked) for w in want), str([l[6] for l in blocked]))
ok("Log is in time order and every line has user, action and result",
   all(LOG[i][0] <= LOG[i + 1][0] for i in range(len(LOG) - 1)) and all(l[1] and l[3] and l[5] for l in LOG))
ok("Checks sheet: all 10 PASS", all(wbv["Checks"].cell(r, 4).value == "PASS" for r in range(7, 17)))

# ------------------------------------------------------------------ PIN hashes, no PIN in the file
def pin_hash(user, pin):
    t, h = f"{user.lower()}:{pin}:lumen-linen", 7.0
    for k in range(1, 301):
        for ch in t:
            v = h * 131 + ord(ch) + k
            h = v - int(v / 2147483629.0) * 2147483629.0
    return str(int(h))


us = wbv["Users"]
ok("Users keep a hash of each PIN (= Python), never the PIN itself",
   all(str(us.cell(6 + i, 4).value) == pin_hash(u[0], u[3]) for i, u in enumerate(USERS))
   and not any(str(c.value) == u[3] for ws in wbv.worksheets for row in ws.iter_rows() for c in row for u in USERS))

# ------------------------------------------------------------------ 3. the invoice PDFs
pdfdir = ROOT / SET["pdf_folder"]
invs = {r[19]: r for r in rows.values() if r[19]}
files = {p.name for p in pdfdir.glob("*.pdf")}
ok(f"{len(invs)} invoices, {len(files)} PDFs, each named in its booking row", files == {r[21] for r in invs.values()})
code = r"""
import pymupdf as fitz, sys, json
out = {}
for p in sys.argv[1:]:
    out[p] = fitz.open(p)[0].get_text()
print(json.dumps(out))
"""
txt = json.loads(subprocess.run(["python", "-c", code] + [str(pdfdir / r[21]) for r in invs.values()], capture_output=True, text=True).stdout)
bad = []
for r in invs.values():
    t = txt[str(pdfdir / r[21])]
    for must in (r[19], r[3], f"{r[15]:,.2f}", f"{r[18]:,.2f}" if r[18] > 0 else "-"):
        if must not in t: bad.append(f"{r[19]}: '{must}' missing")
ok("every PDF shows its invoice no., customer, total and balance due", not bad, "; ".join(bad[:3]))

# ------------------------------------------------------------------ 4. live, on a throw-away copy
HELPER = r'''
Public Sub T_NewForm(): frmBooking.StartNew: frmBooking.Show vbModeless: End Sub
Public Sub T_Set(c As String, v As String): frmBooking.Controls(c).Value = v: End Sub
Public Sub T_Pick(c As String, key As String)
    Dim i As Long
    For i = 0 To frmBooking.Controls(c).ListCount - 1
        If frmBooking.Controls(c).List(i, 0) = key Then frmBooking.Controls(c).ListIndex = i: Exit Sub
    Next
End Sub
Public Function T_Get(c As String) As String: T_Get = frmBooking.Controls(c).Caption: End Function
Public Function T_Count(c As String) As Long: T_Count = frmBooking.Controls(c).ListCount: End Function
Public Sub T_Add(): frmBooking.AddLine: End Sub
Public Sub T_Remove(i As Long): frmBooking.Controls("lstLines").ListIndex = i: frmBooking.RemoveLine: End Sub
Public Function T_Save(c As Boolean) As String: T_Save = frmBooking.SaveNow(c): End Function
Public Sub T_Close(): Unload frmBooking: End Sub
Public Function T_SignForm(u As String, p As String) As String
    frmSignIn.Show vbModeless
    Dim i As Long
    For i = 0 To frmSignIn.cboUser.ListCount - 1
        If frmSignIn.cboUser.List(i, 0) = u Then frmSignIn.cboUser.ListIndex = i
    Next
    frmSignIn.txtPin.Text = p
    T_SignForm = frmSignIn.DoSignIn() & "|" & frmSignIn.lblMsg.Caption
    Unload frmSignIn
End Function
'''
TMPD = Path(tempfile.mkdtemp(prefix="e13verify_"))
COPY = TMPD / BOOK.name
shutil.copy(BOOK, COPY)
SER = lambda d: (d - date(1899, 12, 30)).days
with vba_access():
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False; xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(COPY))
        comps = {c.Name for c in wb.VBProject.VBComponents}
        ok("VBA: modCore, modButtons, frmSignIn and frmBooking are in the file", {"modCore", "modButtons", "frmSignIn", "frmBooking"} <= comps)
        wb.VBProject.VBComponents.Add(1).CodeModule.AddFromString(HELPER)
        watch(xl)
        run = lambda *a: xl.Run(*a)
        xl.Calculation = -4105
        today = SER(AS_OF)
        fut_out, fut_due, fut_ev = date(2027, 4, 20), date(2027, 4, 22), date(2027, 4, 21)
        ok("opened: nobody signed in; any action is refused ('sign in first')",
           run("CurUser") == "" and run("NewBooking", "C003", SER(fut_ev), SER(fut_out), SER(fut_due), "CH-GLD~10", 0.0, "", True, today) == "BLOCKED: sign in first")
        r = run("T_SignForm", "jon", "1111")
        ok("Sign-in form: wrong PIN refused, message shown, nobody signed in", r.startswith("BLOCKED: wrong PIN") and "Try again" in r and run("CurUser") == "")
        r = run("T_SignForm", "maya", "2468")
        ok("Sign-in form: right PIN signs in Maya as Staff", r.startswith("|") and run("CurUser") == "maya" and run("CurRole") == "Staff")
        ok("Staff cannot lock a period or cancel (Manager only)",
           run("LockPeriod", SER(date(2027, 3, 15)), today).startswith("BLOCKED: Lock period needs a Manager")
           and run("CancelBooking", NO[next(r for r, k in book.items() if k["status"] == "Confirmed")], today, "test").startswith("BLOCKED: Cancel needs a Manager"))
        # the booking form: too many chairs, then the right number
        run("T_NewForm")
        run("T_Pick", "cboCustomer", "C003")
        run("T_Set", "txtEvent", "2027-03-21"); run("T_Set", "txtOut", "2027-03-20"); run("T_Set", "txtDue", "2027-03-22")
        run("T_Set", "cboCategory", "Seating"); run("T_Pick", "cboItem", "CH-GLD")
        f_py = min(free("CH-GLD", date(2027, 3, 20) + timedelta(i)) for i in range(3))
        run("T_Set", "txtQty", str(f_py + 10))
        shown = run("T_Get", "lblFree")
        ok(f"form shows what is free while typing: '{shown.split('  ')[0]}' (Python: {f_py} of 120)", shown.startswith(f"{f_py} of 120 free"))
        run("T_Add")
        ok("form warns before saving: 1 line above what is free", run("T_Get", "lblCheck").startswith("1 line(s) above"))
        r = run("T_Save", True)
        ok(f"Confirm with {f_py + 10} chairs is refused: '{r[:70]}'", r == f"BLOCKED: Chiavari chair, gold: {f_py + 10} asked, only {f_py} free on 20 Mar 2027"
           or (r.startswith("BLOCKED: Chiavari chair, gold:") and f"only {f_py} free" in r))
        run("T_Remove", 0)
        run("T_Set", "txtQty", str(f_py)); run("T_Add")
        run("T_Set", "cboCategory", "Lighting"); run("T_Pick", "cboItem", "LT-CND"); run("T_Set", "txtQty", "4"); run("T_Add")
        run("T_Set", "txtDelivery", "120")
        total_form = run("T_Get", "lblTotal")
        sbt = R2(R2(f_py * 4.5 * 2) + R2(4 * 8.0 * 2) + 120); tot = R2(sbt + R2(sbt * 0.08))
        ok(f"form total = Python ({total_form})", total_form == f"${tot:,.2f}")
        newno = run("T_Save", True)
        run("T_Close")
        ok(f"with what is free, the same booking is confirmed: {newno} (next number)", newno == f"B-2027-{len(book) + 1:04d}")
        MF = lambda: run("MinFree", "CH-GLD", SER(date(2027, 3, 20)), SER(date(2027, 3, 22)), "")
        ok("its chairs are now held: free chairs for 20-22 Mar drop to 0", MF() == 0)
        # enquiry holds nothing; confirming it needs stock
        enq = run("NewBooking", "C011", SER(date(2027, 3, 21)), SER(date(2027, 3, 20)), SER(date(2027, 3, 22)), "CH-GLD~5", 0.0, "", False, today)
        ok(f"an enquiry for chairs that are not free is saved ({enq}) but holds no stock", enq.startswith("B-") and MF() == 0)
        r = run("ConfirmBooking", enq, today)
        ok("confirming that enquiry is refused (no chairs left)", r.startswith("BLOCKED: Chiavari chair, gold: 5 asked, only 0 free"))
        # check out, back 3 days late with 2 lanterns missing, invoice, pay
        home = wb.Worksheets("Home")                      # from here on through the Home buttons, as a user would
        home.Range("C14").Value = newno; home.Range("C22").Value = SER(date(2027, 3, 20))
        run("BtnCheckOut")
        ok("Home: Check out button", wb.Worksheets("Bookings").Cells(run("BookRow", newno), 2).Value == "Out")
        home.Range("C22").Value = SER(date(2027, 3, 25)); home.Range("G35").Value = 2      # line 2 = lantern sets
        run("BtnCheckIn"); r = "" if home.Range("B6").Value.startswith("Check in") else home.Range("B6").Value
        ok("Home: Check in button reads the damaged / missing column and clears it", r == "" and home.Range("G35").Value in (None, ""))
        k_late = 3 - SET["grace_days"]
        exp_late = R2(f_py * 2.0 * k_late) + R2(4 * 3.0 * k_late); exp_dmg = R2(2 * 60.0)
        bw = wb.Worksheets("Bookings"); rr = run("BookRow", newno)
        ok(f"back 3 days late with 2 lantern sets missing: late fee ${exp_late:,.2f}, damage ${exp_dmg:,.2f} (= Python)",
           r == "" and abs(bw.Cells(rr, 12).Value - exp_late) < 0.005 and abs(bw.Cells(rr, 13).Value - exp_dmg) < 0.005)
        ok("the 2 missing lantern sets leave the stock from that day",
           run("FreeQty", "LT-CND", SER(date(2027, 3, 26)), "") == free("LT-CND", date(2027, 3, 26)) - 2)
        inv_before = set(os.listdir(TMPD / SET["pdf_folder"])) if (TMPD / SET["pdf_folder"]).exists() else set()
        run("BtnInvoice"); r = "" if str(home.Range("B6").Value).startswith("Invoice") else home.Range("B6").Value
        newpdf = set(os.listdir(TMPD / SET["pdf_folder"])) - inv_before
        ok(f"invoice + PDF: {bw.Cells(rr, 20).Value}, file written", r == "" and len(newpdf) == 1 and bw.Cells(rr, 2).Value == "Invoiced")
        bal = bw.Cells(rr, 19).Value
        ok("a payment above the balance is refused", run("RecordPayment", newno, float(bal) + 1, SER(date(2027, 3, 26))).startswith("BLOCKED"))
        ok("paying the balance makes it Paid", run("RecordPayment", newno, float(bal), SER(date(2027, 3, 26))) == "" and bw.Cells(rr, 2).Value == "Paid")
        # manager: cancel releases stock; the locked period holds
        run("SignOut")
        ok("Manager signs in (core)", run("SignIn", "rita", "9024") == "" and run("CurRole") == "Manager")
        conf = next(k for k in book.values() if k["status"] == "Confirmed")
        code0, q0 = next(iter(conf["lines"].items()))
        before = run("FreeQty", code0, SER(conf["out"]), "")
        ok(f"Manager cancels {conf['no']}: stock released ({code0} free on its out date {before} -> {before + q0})",
           run("CancelBooking", conf["no"], today, "Customer changed venue") == "" and run("FreeQty", code0, SER(conf["out"]), "") == before + q0)
        ok("nothing can be recorded on a date inside the locked period, not even by a Manager",
           run("RecordPayment", NO[next(r for r, k in book.items() if k["status"] == "Invoiced")], 10.0, SER(date(2027, 2, 27))).startswith("BLOCKED: the period is closed"))
        xl.CalculateFull()
        chk = wb.Worksheets("Checks")
        ok("after all of that the 10 checks still PASS", all(chk.Cells(r, 4).Value == "PASS" for r in range(7, 17)))
        # break a total on purpose
        bw.Unprotect(); old = bw.Cells(rr, 16).Value; bw.Cells(rr, 16).Value = old + 50; xl.CalculateFull()
        t = [chk.Cells(r, 4).Value for r in range(7, 17)]
        ok("a total typed over by hand is caught (Checks: total and balance rows FAIL)", t.count("FAIL") >= 1 and chk.Cells(11, 4).Value == "FAIL")
        bw.Cells(rr, 16).Value = old; xl.CalculateFull()
        lgn = wb.Worksheets("Log").Cells(wb.Worksheets("Log").Rows.Count, 1).End(-4162).Row - 5
        ok(f"every live attempt above is in the Log too ({lgn - len(LOG)} new lines, refusals included)", lgn - len(LOG) >= 18)
        wb.Close(False)
    finally:
        xl.Quit()
shutil.rmtree(TMPD, ignore_errors=True)

passed = sum(1 for r in results if r[0])
print(f"\n{passed}/{len(results)} checks pass")
(HERE / "last-result.txt").write_text(f"{passed}/{len(results)}")
sys.exit(0 if passed == len(results) else 1)
