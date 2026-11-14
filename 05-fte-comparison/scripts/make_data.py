"""Generates the fictional October 2026 data for "Riverside Community Health" (made-up organisation).

Writes two deliberately messy CSV exports into ../data/ and the designed answer into ../test/expected.json:
  hr-roster-2026-10.csv       - HR system export: IDs like E01042, FTE as "80%", fund centers typed
                                "FC-210" / "fc210" / "FC 210", a Status column (terminated people still listed),
                                2 rows exported twice
  finance-budget-2026-10.csv  - Finance position budget: Emp No without the E and the leading zeros, FTE as
                                decimals, cost center as "210", functional area typed in any case, vacant positions
Every discrepancy is planted on purpose so the workbook can be checked against a known answer.
"""
import csv, json, random
from pathlib import Path

random.seed(20261115)
ROOT = Path(__file__).resolve().parent.parent

FUND_CENTERS = ["110", "120", "210", "220", "310", "410"]
# each fund center mostly pays for one or two functional areas
AREAS = ["Clinical Care", "Community Health", "Administration", "Facilities & IT"]
FC_AREAS = {"110": ["Administration"], "120": ["Administration", "Facilities & IT"],
            "210": ["Clinical Care"], "220": ["Clinical Care", "Community Health"],
            "310": ["Community Health"], "410": ["Facilities & IT", "Administration"]}
FIRST = ["Aaliyah", "Ben", "Carmen", "Daniel", "Elena", "Farid", "Grace", "Hiro", "Isabel", "James", "Kofi", "Lena",
         "Marco", "Nadia", "Omar", "Priya", "Quinn", "Rosa", "Samuel", "Tara", "Uma", "Victor", "Wendy", "Xavier",
         "Yara", "Zane", "Alicia", "Brandon", "Chloe", "Derek", "Esther", "Felix", "Gloria", "Hector", "Ivy", "Jonah"]
LAST = ["Abbott", "Barnes", "Castillo", "Dawson", "Ellis", "Fischer", "Gomez", "Hayes", "Iqbal", "Jensen", "Khan",
        "Lopez", "Morgan", "Nguyen", "Okafor", "Patel", "Quinn", "Reyes", "Silva", "Turner", "Usman", "Vance",
        "Walsh", "Young", "Zimmer", "Brooks", "Carter", "Diaz", "Evans", "Foster"]

used_names = set()


def name():
    while True:
        n = f"{random.choice(FIRST)} {random.choice(LAST)}"
        if n not in used_names:
            used_names.add(n)
            return n


def hr_fc(fc):
    """How HR clerks type a fund center."""
    return random.choice([f"FC-{fc}", f"FC{fc}", f"fc{fc}", f"FC {fc}", f"FC-{fc} ", f"fc-{fc}"])


def hr_fte(f):
    p = round(f * 100)
    return random.choice([f"{p}%", f"{p}%", f"{p} %", f" {p}%"])


def fin_area(a):
    return random.choice([a, a, a.upper(), a.lower(), a + " ", " " + a])


def fin_fte(f):
    return random.choice([f"{f:.2f}", f"{f:.1f}" if round(f, 1) == f else f"{f:.2f}", f"{f:g}"])


hr, fin, expected = [], [], []
emp_no = 1000


def new_emp():
    global emp_no
    emp_no += random.randint(1, 9)
    return emp_no


def fte():
    return random.choice([1.0] * 7 + [0.8, 0.6, 0.5])


def assign():
    fc = random.choice(FUND_CENTERS)
    return fc, random.choice(FC_AREAS[fc])


def add(rows, e, nm, fc, area, f, status="Active"):
    rows.append(dict(id=e, name=nm, fc=fc, area=area, fte=f, status=status))


# 1) Matched: one fund center each
for _ in range(118):
    e, nm, (fc, area), f = new_emp(), name(), assign(), fte()
    st = random.choice(["Active"] * 12 + ["On leave"])
    add(hr, e, nm, fc, area, f, st); add(fin, e, nm, fc, area, f)
    expected.append(dict(emp=e, fc=fc, status="Matched"))

# 2) Matched: split between two fund centers in both sources
for _ in range(6):
    e, nm = new_emp(), name()
    fcs = random.sample(FUND_CENTERS, 2)
    for fc in fcs:
        area = random.choice(FC_AREAS[fc])
        add(hr, e, nm, fc, area, 0.5); add(fin, e, nm, fc, area, 0.5)
        expected.append(dict(emp=e, fc=fc, status="Matched"))

# 3) FTE differs
for hr_f, fin_f in [(1.0, 0.8), (0.8, 1.0), (0.5, 0.6), (1.0, 0.5), (0.6, 0.8)]:
    e, nm, (fc, area) = new_emp(), name(), assign()
    add(hr, e, nm, fc, area, hr_f); add(fin, e, nm, fc, area, fin_f)
    expected.append(dict(emp=e, fc=fc, status="FTE differs", diff=round(fin_f - hr_f, 2)))

# 4) Functional area differs (same fund center, same FTE)
for fc, a_hr, a_fin in [("120", "Administration", "Facilities & IT"), ("220", "Community Health", "Clinical Care"),
                        ("410", "Facilities & IT", "Administration")]:
    e, nm, f = new_emp(), name(), fte()
    add(hr, e, nm, fc, a_hr, f); add(fin, e, nm, fc, a_fin, f)
    expected.append(dict(emp=e, fc=fc, status="Functional area differs"))

# 5) Fund center differs (one row each side, same person)
for fc_hr, fc_fin in [("210", "220"), ("110", "120"), ("310", "220"), ("410", "120")]:
    e, nm, f = new_emp(), name(), fte()
    area = FC_AREAS[fc_hr][0] if FC_AREAS[fc_hr][0] in FC_AREAS[fc_fin] else FC_AREAS[fc_fin][0]
    add(hr, e, nm, fc_hr, area, f); add(fin, e, nm, fc_fin, area, f)
    expected.append(dict(emp=e, fc=fc_hr, fc_budget=fc_fin, status="Fund center differs"))

# 6) Only in HR: hired but not in the budget
for _ in range(4):
    e, nm, (fc, area), f = new_emp(), name(), assign(), fte()
    add(hr, e, nm, fc, area, f)
    expected.append(dict(emp=e, fc=fc, status="Only in HR roster"))

# 7) Terminated in HR but still funded in the budget
for _ in range(3):
    e, nm, (fc, area), f = new_emp(), name(), assign(), 1.0
    add(hr, e, nm, fc, area, f, "Terminated"); add(fin, e, nm, fc, area, f)
    expected.append(dict(emp=e, fc=fc, status="Only in budget", why="terminated"))

# 8) Vacant positions in the budget
for fc, area in [("210", "Clinical Care"), ("310", "Community Health")]:
    fin.append(dict(id=None, name="VACANT", fc=fc, area=area, fte=1.0, status=""))
    expected.append(dict(emp=None, fc=fc, status="Only in budget", why="vacant"))

# 9) In the budget, not on the HR roster at all
for _ in range(2):
    e, nm, (fc, area), f = new_emp(), name(), assign(), fte()
    add(fin, e, nm, fc, area, f)
    expected.append(dict(emp=e, fc=fc, status="Only in budget", why="unknown"))

random.shuffle(hr)
random.shuffle(fin)
dups = random.sample(range(len(hr)), 2)
(ROOT / "data").mkdir(exist_ok=True)
with open(ROOT / "data" / "hr-roster-2026-10.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Employee ID", "Employee Name", "Fund Center", "Functional Area", "FTE", "Status"])
    for i, r in enumerate(hr):
        row = [f"E{r['id']:05d}", r["name"] + random.choice(["", "", "", " "]), hr_fc(r["fc"]), r["area"],
               hr_fte(r["fte"]), r["status"]]
        w.writerow(row)
        if i in dups:
            w.writerow(row)  # the export repeats a row now and then

with open(ROOT / "data" / "finance-budget-2026-10.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Position", "Emp No", "Employee", "Cost Center", "Function", "Budgeted FTE"])
    for i, r in enumerate(fin, 1):
        w.writerow([f"P-{3000 + i}", "" if r["id"] is None else str(r["id"]), r["name"], r["fc"],
                    fin_area(r["area"]), fin_fte(r["fte"])])

active = [r for r in hr if r["status"] != "Terminated"]
exp = dict(hr_rows_in_file=len(hr) + 2, hr_duplicates_in_file=2, hr_terminated=len(hr) - len(active),
           hr_fte=round(sum(r["fte"] for r in active), 2), budget_rows=len(fin),
           budget_fte=round(sum(r["fte"] for r in fin), 2), cases=expected)
(ROOT / "test").mkdir(exist_ok=True)
(ROOT / "test" / "expected.json").write_text(json.dumps(exp, indent=1), encoding="utf-8")
print(f"HR rows {exp['hr_rows_in_file']} (FTE {exp['hr_fte']}), budget rows {exp['budget_rows']} (FTE {exp['budget_fte']})")
