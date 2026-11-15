"""Screenshots of the E4 Google Sheet for docs/ (Browser Kerja tab "Invoice Tracker", CDP).
Run under the shared lock:  python tools/with_lock.py E4 -- <python> scripts/shots.py [name ...]
Names: dashboard invoices customers replies statement settings log checks preview sidebar (default: all)."""
import json, sys, time
from pathlib import Path

sys.path.insert(0, r"D:\DATA\Dodik\PROJECT\2026\BELAJAR\tools")
sys.path.insert(0, r"D:\DATA\Dodik\PROJECT\2026\BELAJAR\tools\video")
import e4_shots as E  # noqa: E402

DOCS = Path(__file__).resolve().parent.parent / "docs"
import threading  # noqa: E402
s = E.S.__new__(E.S)  # the sheet tab by URL (the Apps Script editor tab's title also says "Invoice tracker")
s.ws, s.i, s.lock = E.bk._WS(next(t for t in E.bk.tabs() if "1L1VOTy3AUF289PfG9lxm0bMtTGEC7bCOxzYyuiCXTio" in t["url"])["ws"], 120), 0, threading.Lock()
s("Page.bringToFront")
s("Emulation.setFocusEmulationEnabled", enabled=True)


def view(w, h):
    s("Emulation.setDeviceMetricsOverride", width=w, height=h, deviceScaleFactor=1.5, mobile=False)
    time.sleep(2)


def tab(name):
    for _ in range(4):
        p = s.center(f"Array.from(document.querySelectorAll('.docs-sheet-tab')).find(e=>e.innerText.trim()==={json.dumps(name)})")
        s.click(p[0] - 20, p[1])
        time.sleep(3)
        if s.js("document.querySelector('.docs-sheet-active-tab').innerText.trim()") == name:
            break
    for t in ("keyDown", "keyUp"):  # Ctrl+Home: back to A1
        s("Input.dispatchKeyEvent", type=t, modifiers=2, windowsVirtualKeyCode=36, code="Home", key="Home")
    time.sleep(1)
    s.js("document.querySelectorAll('.native-scrollbar-y,.native-scrollbar-x').forEach(e=>{e.scrollTop=0;e.scrollLeft=0})")
    time.sleep(3)


def menu(top, item, wait):
    p = s.center(f"Array.from(document.querySelectorAll('.menu-button')).find(e=>e.innerText.trim()==={json.dumps(top)})")
    s.click(*p)
    time.sleep(1.5)
    p = s.center(f"Array.from(document.querySelectorAll('.goog-menuitem')).find(e=>e.offsetParent&&e.innerText.trim()==={json.dumps(item)})")
    s.click(*p)
    time.sleep(wait)


def close_dialog():
    p = s.center("document.querySelector('[aria-label=\"Close dialogue\"],[aria-label=\"Close dialog\"]')")
    if p:
        s.click(*p)
        time.sleep(1.5)


SHOTS = {
    "dashboard": lambda: (view(1440, 1470), tab("Dashboard")),
    "invoices": lambda: (view(1500, 900), tab("Invoices")),
    "customers": lambda: (view(1500, 640), tab("Customers")),
    "replies": lambda: (view(1300, 480), tab("Replies")),
    "statement": lambda: (view(1300, 1080), tab("Statement")),
    "settings": lambda: (view(1500, 900), tab("Settings")),
    "reminder-log": lambda: (view(1300, 640), tab("Reminder log")),
    "checks": lambda: (view(1500, 1560), tab("Checks")),
    "preview": lambda: (view(1500, 980), tab("Dashboard"), menu("Invoice tracker", "Preview reminders", 22)),
    "sidebar": lambda: (view(1500, 1100), tab("Dashboard"), menu("Invoice tracker", "Open collections desk", 22)),
}

def close_side_panels():
    """Google sometimes opens its own side panel (e.g. "Tables" tips); close it so it does not cover the sheet."""
    for _ in range(3):
        p = s.center("[...document.querySelectorAll('[aria-label=\"Close\"],[aria-label=\"Close panel\"],[aria-label=\"Close sidebar\"]')]"
                     ".find(e=>e.offsetParent&&e.getBoundingClientRect().x>innerWidth*0.6)")
        if not p:
            break
        s.click(*p)
        time.sleep(1.5)


view(1440, 900)
close_side_panels()
names = sys.argv[1:] or list(SHOTS)
for n in names:
    SHOTS[n]()
    out = DOCS / f"{n}.png"
    s.shot(str(out))
    if n == "preview":
        close_dialog()
view(1280, 720)
