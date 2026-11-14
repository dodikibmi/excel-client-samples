"""Independent check of Customer-Cleanup.xlsx.

1. Re-does the whole cleanup in plain Python straight from the three source files (no Excel involved):
   clean every field, drop exact duplicates and test rows, link records that share an email, a phone or
   first name + last name + ZIP (nicknames: Bill = William), keep the most recent record as master.
2. Compares it with the answer that was planted when the data was generated (test/expected.json): every group of
   records must be exactly one real person, and the look-alike pairs must be flagged for review, not merged.
3. Opens the workbook with Excel and compares every customer, merge, review pair and data check with Python.
4. Copies workbook + data to another folder, refreshes there, and checks the answer is the same.
Exit code 1 if anything differs.   Run:  python test/verify.py   (needs openpyxl; pywin32 for steps 3-4)
"""
import csv, datetime as dt, json, re, sys
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
results = []


def check(name, ok, detail=""):
    results.append(bool(ok))
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  ({detail})" if detail and not ok else ""))


# ---------------------------------------------------------------- the cleaning rules (same as the Power Query steps)
NICK = {"JON": "JOHN", "JOHNNY": "JOHN", "BILL": "WILLIAM", "WILL": "WILLIAM", "BOB": "ROBERT", "ROB": "ROBERT",
        "LIZ": "ELIZABETH", "BETH": "ELIZABETH", "KATE": "KATHERINE", "KATIE": "KATHERINE", "MIKE": "MICHAEL",
        "JIM": "JAMES", "JIMMY": "JAMES", "TOM": "THOMAS", "DAN": "DANIEL", "DANNY": "DANIEL", "CHRIS": "CHRISTOPHER",
        "MATT": "MATTHEW", "NICK": "NICHOLAS", "SAM": "SAMUEL", "ALEX": "ALEXANDER", "TONY": "ANTHONY",
        "JEN": "JENNIFER", "JENNY": "JENNIFER", "MEG": "MARGARET", "PEGGY": "MARGARET", "STEVE": "STEVEN",
        "DAVE": "DAVID", "ANDY": "ANDREW", "BEN": "BENJAMIN", "JOE": "JOSEPH", "ABBY": "ABIGAIL", "GABE": "GABRIEL",
        "TED": "EDWARD", "ED": "EDWARD", "RICK": "RICHARD", "SUE": "SUSAN", "BECKY": "REBECCA", "NATE": "NATHAN"}
STATES = {"COLORADO": "CO", "UTAH": "UT", "WASHINGTON": "WA", "OREGON": "OR", "CALIFORNIA": "CA", "ARIZONA": "AZ",
          "NEW MEXICO": "NM", "TEXAS": "TX", "ILLINOIS": "IL", "MINNESOTA": "MN", "MASSACHUSETTS": "MA",
          "CONNECTICUT": "CT", "NEW JERSEY": "NJ", "VERMONT": "VT", "MAINE": "ME", "NEW YORK": "NY"}
ABBR = {"STREET": "St", "ST": "St", "AVENUE": "Ave", "AVE": "Ave", "ROAD": "Rd", "RD": "Rd", "DRIVE": "Dr", "DR": "Dr",
        "LANE": "Ln", "LN": "Ln", "COURT": "Ct", "CT": "Ct", "BOULEVARD": "Blvd", "BLVD": "Blvd"}
NO_COMPANY = {"N/A", "NA", "NONE", "-", "SELF-EMPLOYED"}
PRIORITY = {"Webshop": 1, "POS": 2, "Sales sheet": 3}

T = lambda x: "" if x is None else str(x).strip()
squash = lambda s: " ".join(p for p in s.strip().split(" ") if p)
digits = lambda s: re.sub(r"\D", "", s)
proper = lambda s: s.lower().title()


def clean(rec):
    """rec: Source, Source row, Name typed, First typed, Last typed, raw fields, Last activity, Spend."""
    if rec["first_t"] or rec["last_t"]:
        first, last = squash(T(rec["first_t"])), squash(T(rec["last_t"]))
    else:
        s = squash(T(rec["name"]))
        if "," in s:
            p = s.split(",")
            last, first = p[0].strip(), ",".join(p[1:]).strip()
        else:
            p = s.split(" ")
            first, last = p[0], " ".join(p[1:])
    first, last = proper(first), proper(last)
    e = T(rec["email"]).replace(" ", "").lower() or None
    d = digits(T(rec["phone"]))
    d = d[1:] if len(d) == 11 and d.startswith("1") else d
    phone = f"({d[:3]}) {d[3:6]}-{d[6:]}" if len(d) == 10 else None
    z = digits(T(rec["zip"]).split("-")[0])
    zip_ = z if len(z) == 5 else "0" + z if len(z) == 4 else None
    st = T(rec["state"]).upper()
    st = (st if len(st) == 2 else STATES.get(st, st)) or None
    a = squash(T(rec["addr"]).replace(".", ""))
    addr = " ".join(ABBR.get(t.upper()) or proper(t) for t in a.split(" ")) if a else None
    city = proper(squash(T(rec["city"]))) or None
    co = squash(T(rec["company"]))
    co = None if not co or co.upper() in NO_COMPANY else co
    o = T(rec["optin"]).upper()
    optin = "Yes" if o in ("YES", "Y") else "No" if o in ("NO", "N") else None
    canon = NICK.get(first.upper(), first.upper())
    typed = T(rec["name"]) if rec["name"] is not None else f"{T(rec['first_t'])} {T(rec['last_t'])}"
    return dict(rec, first=first, last=last, email=e, phone=phone, zip=zip_, st=st, address=addr, cityc=city,
                co=co, opt=optin, canon=canon, typed=typed,
                kname=f"{canon}|{last.upper()}|{zip_}" if first and last and zip_ else None,
                filled=sum(v is not None for v in (e, phone, co, addr)))


# ---------------------------------------------------------------- 1. read the three sources
recs = []
with open(ROOT / "data" / "webshop-customers-2026-10.csv", encoding="utf-8") as f:
    for i, r in enumerate(csv.DictReader(f), 1):
        recs.append(dict(id=f"W-{i:04d}", src="Webshop", raw=tuple(r.values()), name=r["Full Name"], first_t="", last_t="",
                         email=r["Email"], phone=r["Phone"], company=r["Company"], addr=r["Street"], city=r["City"],
                         state=r["State"], zip=r["ZIP"], optin=r["Marketing Opt-in"],
                         date=dt.date.fromisoformat(r["Last Order"].strip()),
                         spend=float(r["Total Spent"].replace("$", "").replace(",", ""))))
with open(ROOT / "data" / "pos-customers-2026-10.csv", encoding="utf-8") as f:
    for i, r in enumerate(csv.DictReader(f), 1):
        m, d_, y = map(int, r["Last Visit"].split("/"))
        recs.append(dict(id=f"P-{i:04d}", src="POS", raw=tuple(r.values()), name=r["Customer"], first_t="", last_t="",
                         email=r["Email"], phone=r["Phone"], company="", addr=r["Address"], city=r["City"],
                         state=r["St"], zip=r["Zip"], optin="", date=dt.date(y, m, d_), spend=float(r["Spend"] or 0)))
ws = openpyxl.load_workbook(ROOT / "data" / "sales-team-accounts.xlsx").active
cols = [c.value for c in ws[3]]
company, n = None, 0
for row in ws.iter_rows(min_row=4, values_only=True):
    if all(v is None or str(v).strip() == "" for v in row):
        continue
    r = dict(zip(cols, row))
    company = r["Company"] if r["Company"] not in (None, "") else company  # merged cells: fill down
    n += 1
    v = r["Last contact"]
    date = v.date() if isinstance(v, dt.datetime) else dt.datetime.strptime(str(v).strip(), "%d %b %Y").date()
    recs.append(dict(id=f"S-{n:04d}", src="Sales sheet", raw=tuple(row), name=None, first_t=T(r["First"]),
                     last_t=T(r["Last"]), email=r["Email"], phone=r["Phone"], company=company, addr=r["Address"],
                     city=r["City"], state=r["State"], zip=r["ZIP"], optin=r["Newsletter"], date=date, spend=0.0))
for i, r in enumerate(recs):
    r["order"] = i
recs = [clean(r) for r in recs]

seen = set()
for r in recs:
    k = (r["src"], r["raw"])
    up = r["typed"].upper()
    if k in seen:
        r["status"] = "Discarded: exact duplicate row"
    elif "TEST" in up or (r["email"] or "").startswith("test@") or "WALK IN" in up.replace("-", " "):
        r["status"] = "Discarded: test or walk-in record"
    elif not r["last"] and not r["email"] and not r["phone"]:
        r["status"] = "Discarded: no name or contact"
    else:
        r["status"] = "Kept"
    seen.add(k)
kept = [r for r in recs if r["status"] == "Kept"]

# link records that share a key (union-find)
parent = {r["id"]: r["id"] for r in kept}


def find(x):
    while parent[x] != x:
        parent[x] = parent[parent[x]]
        x = parent[x]
    return x


for key in ("email", "phone", "kname"):
    first_of = {}
    for r in kept:
        if r[key]:
            if r[key] in first_of:
                a, b = find(first_of[r[key]]), find(r["id"])
                if a != b:
                    parent[b] = a
            else:
                first_of[r[key]] = r["id"]
groups = defaultdict(list)
for r in kept:
    groups[find(r["id"])].append(r)
customers = {}  # master record id -> customer
for g in groups.values():
    g.sort(key=lambda r: (-r["date"].toordinal(), -r["filled"], PRIORITY[r["src"]], r["order"]))
    m = g[0]
    first_val = lambda k: next((r[k] for r in g if r[k]), None)
    nm = next((r for r in g if r["last"]), m)
    ad = next((r for r in g if r["address"]), m)
    email = first_val("email")
    opts = {r["opt"] for r in g}
    srcs = sorted({r["src"] for r in g}, key=PRIORITY.get)
    customers[m["id"]] = dict(first=nm["first"], last=nm["last"], email=email, phone=first_val("phone"),
                              other=sorted({r["email"] for r in g if r["email"] and r["email"] != email}),
                              company=first_val("co"), address=ad["address"], city=ad["cityc"], st=ad["st"], zip=ad["zip"],
                              opt="No" if "No" in opts else "Yes" if "Yes" in opts else "Unknown",
                              spend=round(sum(r["spend"] for r in g), 2), n=len(g), sources=" + ".join(srcs),
                              lastact=m["date"], members=g)
    for r in g[1:]:
        cnt = lambda k: sum(1 for x in g if x[k] and x[k] == r[k])
        r["why"] = ("Same email" if r["email"] and cnt("email") > 1 else "Same phone" if r["phone"] and cnt("phone") > 1
                    else "Same name and ZIP")
        r["into"] = m["id"]
gid = {r["id"]: find(r["id"]) for r in kept}
masters = {find(r["id"]): g[0]["id"] for g in groups.values() for r in g}
pairs = {}
for keyf, why in ((lambda r: f"{r['canon'][:1]}|{r['last'].upper()}|{r['zip']}" if r["first"] and r["last"] and r["zip"] else None,
                   "Similar first name, same last name and ZIP"),
                  (lambda r: f"{r['canon']}|{r['last'].upper()}" if r["first"] and r["last"] else None, "Same name, different ZIP")):
    bucket = defaultdict(set)
    for r in kept:
        k = keyf(r)
        if k:
            bucket[k].add(masters[gid[r["id"]]])
    for ms in bucket.values():
        ms = sorted(ms)
        for i, a in enumerate(ms):
            for b in ms[i + 1:]:
                pairs.setdefault((a, b), why)

received = Counter(r["src"] for r in recs)
status = Counter(r["status"] for r in recs)
merged = len(kept) - len(customers)
print(f"Python: received {len(recs)}, discarded {len(recs) - len(kept)}, customers {len(customers)}, merged {merged}, review {len(pairs)}")

# ---------------------------------------------------------------- 2. against the planted answer
exp = json.loads((ROOT / "test" / "expected.json").read_text(encoding="utf-8"))
truth = exp["truth"]
check("Rows read per source = rows written", all(received[s] == exp["rows"][k] for s, k in
                                                  (("Webshop", "W"), ("POS", "P"), ("Sales sheet", "S"))), received)
check("Exact duplicate rows found = planted (3 webshop + 2 POS)", status["Discarded: exact duplicate row"] == 5)
check("Test and walk-in rows found = planted", sum(v == "junk" for v in truth.values()) ==
      status["Discarded: test or walk-in record"] + status["Discarded: no name or contact"])
check("Every discarded row is a planted duplicate or test row, every kept row a real person",
      all((r["status"] == "Kept") == isinstance(truth[r["id"]], int) for r in recs))
pure = all(len({truth[r["id"]] for r in g}) == 1 for g in groups.values())
check("Every merged group is one real person (no false merge)", pure)
check("Every real person ends up as exactly one customer (no missed duplicate)",
      len(customers) == exp["customers"] == len({truth[r["id"]] for r in kept}))
check("Records merged away = planted", merged == exp["merged_away"], f"{merged} vs {exp['merged_away']}")
check("Lifetime spend of all customers = planted spend", round(sum(c["spend"] for c in customers.values()), 2) == exp["lifetime_spend"])
pid_master = {truth[m]: m for m in customers}
planted = {tuple(sorted((pid_master[a], pid_master[b]))): why for a, b, why in exp["review_pairs"]}
check(f"Look-alike pairs flagged for review = the {len(planted)} planted pairs", pairs == planted,
      f"extra {set(pairs) - set(planted)}, missing {set(planted) - set(pairs)}")
check("Look-alikes are NOT merged (they stay two customers)", all(a != b for a, b in planted))
check("No email belongs to two customers", len({c["email"] for c in customers.values() if c["email"]}) ==
      sum(1 for c in customers.values() if c["email"]))
check("Opt-out wins: a customer is opted in only if no record says No",
      all(c["opt"] != "Yes" or all(r["opt"] != "No" for r in c["members"]) for c in customers.values()))

# ---------------------------------------------------------------- 3. against the workbook
try:
    import win32com.client as win32
except ImportError:
    print("SKIP  workbook checks (pywin32 not installed)")
    win32 = None


def table(wb, sheet, name):
    lo = wb.Worksheets(sheet).ListObjects(name)
    cols = [lo.ListColumns(i).Name for i in range(1, lo.ListColumns.Count + 1)]
    if lo.DataBodyRange is None:
        return []
    return [dict(zip(cols, r)) for r in lo.DataBodyRange.Value]


def xdate(v):
    return dt.date(v.year, v.month, v.day) if v else None


def check_workbook(wb, tag=""):
    cust = table(wb, "Customers", "Customers")
    check(f"{tag}Workbook customers = Python ({len(customers)})", len(cust) == len(customers), f"{len(cust)}")
    bad = []
    for c in cust:
        p = customers.get(c["Master record"])
        if p is None:
            bad.append(f"{c['Master record']} not a Python master"); continue
        got = (c["First name"], c["Last name"], c["Email"], c["Phone"], c["Company"], c["Address"], c["City"], c["State"],
               c["ZIP"], c["Marketing opt-in"], round(c["Lifetime spend"], 2), int(c["Records merged"]), c["Sources"],
               xdate(c["Last activity"]), "; ".join(p["other"]) or None)
        want = (p["first"], p["last"], p["email"], p["phone"], p["company"], p["address"], p["city"], p["st"], p["zip"],
                p["opt"], p["spend"], p["n"], p["sources"], p["lastact"], "; ".join(p["other"]) or None)
        got = tuple(None if g == "" else g for g in got)
        if got != want:
            bad.append(f"{c['Master record']}: " + ", ".join(f"{g!r}<>{w!r}" for g, w in zip(got, want) if g != w))
    check(f"{tag}Every customer field = Python (survivorship, other emails, opt-in, spend, sources)", not bad, "; ".join(bad[:3]))
    ids = {c["Master record"]: c["Customer ID"] for c in cust}
    log = table(wb, "Merge log", "MergeLog")
    py_log = {r["id"]: (ids.get(r["into"]), r["why"]) for g in groups.values() for r in g[1:]}
    x_log = {r["Record"]: (r["Merged into"], r["Matched by"]) for r in log}
    check(f"{tag}Merge log = Python: every merged record, its customer and the rule ({len(py_log)})", x_log == py_log,
          f"{len(x_log)} rows; first diff {next(((k, x_log.get(k), v) for k, v in py_log.items() if x_log.get(k) != v), None)}")
    rev = table(wb, "Needs review", "Review")
    x_pairs = {tuple(sorted((r["Master A"], r["Master B"]))): r["Reason"] for r in rev}
    check(f"{tag}Needs review = Python ({len(pairs)} pairs, with the reason)", x_pairs == pairs, f"{len(x_pairs)}")
    allr = table(wb, "All records", "AllRecords")
    xs = Counter(("Discarded" if r["Status"].startswith("Discarded") else r["Status"]) for r in allr)
    check(f"{tag}All records: every source row once, with the right status",
          len(allr) == len(recs) and {r["Record"]: r["Status"] for r in allr if r["Status"].startswith("Discarded")} ==
          {r["id"]: r["status"] for r in recs if r["status"] != "Kept"}, f"{xs}")
    crm = table(wb, "CRM import", "CRMImport")
    check(f"{tag}CRM import: one row per customer, same emails", len(crm) == len(cust) and
          sorted(str(r["Email"] or "") for r in crm) == sorted(str(c["Email"] or "") for c in cust))
    return cust


if win32:
    xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(ROOT / "Customer-Cleanup.xlsx"), 0, True)
        check_workbook(wb)
        s = wb.Worksheets("Summary")
        for r in range(23, 30):
            check(f"Workbook data check: {s.Range(f'B{r}').Value}", s.Range(f"M{r}").Value == "PASS")
        wb.Close(False)

        import shutil, tempfile
        with tempfile.TemporaryDirectory() as tmp:
            shutil.copy(ROOT / "Customer-Cleanup.xlsx", tmp)
            shutil.copytree(ROOT / "data", Path(tmp) / "data")
            wb = xl.Workbooks.Open(str(Path(tmp) / "Customer-Cleanup.xlsx"))
            wb.RefreshAll(); xl.CalculateUntilAsyncQueriesDone(); xl.Calculate()
            folder = wb.Worksheets("Settings").Range("B4").Value
            check("Refresh in another folder reads that folder's data", str(folder).lower().startswith(str(tmp).lower()[:12]), folder)
            check_workbook(wb, "After refresh: ")
            s = wb.Worksheets("Summary")
            check("After refresh: all seven data checks PASS", all(s.Range(f"M{r}").Value == "PASS" for r in range(23, 30)))
            wb.Close(False)
    finally:
        xl.Quit()

print(f"\n{sum(results)}/{len(results)} checks passed")
sys.exit(0 if all(results) else 1)
