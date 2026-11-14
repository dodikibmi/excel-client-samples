"""Screenshots for docs/ taken from Power BI Desktop itself (Store version), plus the .pbix copy.

Opens Copperleaf-Finance.pbip, closes the optional sign-in window (Desktop works without an account), presses
Refresh now, folds the side panes away and grabs each report page and the model view. With --pbix it also saves a
.pbix copy (File > Save as). Window positions assume a 1920x1200 screen at 125% scaling with Desktop maximised.
Run:  python scripts/shots.py [--pbix] [--keep-open]
"""
import argparse, ctypes, os, subprocess, sys, time
from pathlib import Path

ctypes.windll.shcore.SetProcessDpiAwareness(2)
import win32api, win32con, win32gui, win32process, win32ui  # noqa: E402
import win32com.client  # noqa: E402
from PIL import Image  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
PBIP = ROOT / "Copperleaf-Finance.pbip"
DOCS = ROOT / "docs"
PAGES = {"overview": 313, "pnl-vs-budget": 463, "cash-and-statements": 647, "checks": 803}  # page tab x, y = 1089
CANVAS = (181, 228, 1653, 1052)  # report page inside the window, panes folded


def exe_name(h):
    _, pid = win32process.GetWindowThreadProcessId(h)
    try:
        return win32process.GetModuleFileNameEx(win32api.OpenProcess(0x1000, False, pid), 0).split("\\")[-1].lower()
    except Exception:
        return ""


def windows():
    out = []
    win32gui.EnumWindows(lambda h, _: out.append(h) if win32gui.IsWindowVisible(h) else None, None)
    return out


def main_window():
    best = None
    for h in windows():
        if exe_name(h) == "pbidesktop.exe" and win32gui.GetWindowText(h):
            r = win32gui.GetWindowRect(h)
            if best is None or (r[2] - r[0]) > (win32gui.GetWindowRect(best)[2] - win32gui.GetWindowRect(best)[0]):
                best = h
    return best


def close_sign_in():
    """The Microsoft account prompt is a separate system window: close it (no login needed)."""
    for h in windows():
        r = win32gui.GetWindowRect(h)
        if win32gui.GetClassName(h) == "ApplicationFrameWindow" and exe_name(h) == "explorer.exe" and r[2] - r[0] < 1500:
            win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
            return True
    return False


def front(h=None):
    """Bring Desktop to the front before every click or key, and stop if it is not there: input must never land on
    another window (the screen may be in use)."""
    h = h or main_window()
    if exe_name(win32gui.GetForegroundWindow()) != "pbidesktop.exe":   # a Desktop dialog in front is fine
        keys_raw("%")                 # an Alt tap lets this process move the foreground
        try:
            win32gui.SetForegroundWindow(h)
        except Exception:
            pass
        time.sleep(0.8)
    fg = win32gui.GetForegroundWindow()
    if fg != h and exe_name(fg) != "pbidesktop.exe":
        raise SystemExit("Power BI Desktop is not the front window: stopped so no click lands elsewhere")


def keys_raw(k):
    win32com.client.Dispatch("WScript.Shell").SendKeys(k)


def click(x, y, wait=2.0, right=False):
    h = main_window()
    front(h)
    l, t, _, _ = win32gui.GetWindowRect(h)
    win32api.SetCursorPos((l + x, t + y))
    time.sleep(0.2)
    down, up = (win32con.MOUSEEVENTF_RIGHTDOWN, win32con.MOUSEEVENTF_RIGHTUP) if right else (win32con.MOUSEEVENTF_LEFTDOWN, win32con.MOUSEEVENTF_LEFTUP)
    win32api.mouse_event(down, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(up, 0, 0)
    time.sleep(wait)


def park_mouse():
    h = main_window()
    l, t, _, _ = win32gui.GetWindowRect(h)
    win32api.SetCursorPos((l + 1799, t + 199))  # empty ribbon area: no tooltips on the page
    time.sleep(1.5)


def keys(k, wait=1.0):
    front()
    win32com.client.Dispatch("WScript.Shell").SendKeys(k)
    time.sleep(wait)


def grab(h=None):
    h = h or main_window()
    l, t, r, b = win32gui.GetWindowRect(h)
    w, hh = r - l, b - t
    hdc = win32gui.GetWindowDC(h)
    src = win32ui.CreateDCFromHandle(hdc)
    mem = src.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(src, w, hh)
    mem.SelectObject(bmp)
    ctypes.windll.user32.PrintWindow(h, mem.GetSafeHdc(), 2)
    info, bits = bmp.GetInfo(), bmp.GetBitmapBits(True)
    img = Image.frombuffer("RGB", (info["bmWidth"], info["bmHeight"]), bits, "raw", "BGRX", 0, 1)
    win32gui.DeleteObject(bmp.GetHandle()); mem.DeleteDC(); src.DeleteDC(); win32gui.ReleaseDC(h, hdc)
    return img


def screen_to_window(box):
    l, t, _, _ = win32gui.GetWindowRect(main_window())
    return (box[0] - l, box[1] - t, box[2] - l, box[3] - t)


def open_report(target=PBIP, wait=70):
    subprocess.run(["taskkill", "/IM", "PBIDesktop.exe", "/F"], capture_output=True)
    time.sleep(3)
    subprocess.Popen([os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\WindowsApps\PBIDesktopStore.exe"), str(target)])
    t0 = time.time()
    while time.time() - t0 < wait:
        time.sleep(3)
        close_sign_in()
        h = main_window()
        if h and PBIP.stem in win32gui.GetWindowText(h) and time.time() - t0 > 35:
            break
    win32gui.ShowWindow(main_window(), win32con.SW_MAXIMIZE)
    time.sleep(2)


def refresh_and_fold(refresh=True):
    if refresh:
        click(1305, 259, 45)        # "Refresh now" on the yellow bar: a PBIP keeps no data
    click(1909, 263)                # fold the Data pane
    click(1859, 263)                # fold the Visualizations pane (the filter pane starts folded)


def goto(page):
    click(PAGES[page], 1089, 4)
    if page == "pnl-vs-budget":     # right-click the matrix > Expand > All: show every account under its category
        click(300, 503, 2, right=True)
        click(300 + 95, 503 + 51, 2)
        click(300 + 355, 503 + 119, 3)
        click(1490, 377, 1.5)       # click the empty page so no visual stays selected
    park_mouse()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pbix", action="store_true")
    ap.add_argument("--keep-open", action="store_true")
    a = ap.parse_args()
    DOCS.mkdir(exist_ok=True)
    open_report()
    refresh_and_fold()
    for page in PAGES:
        goto(page)
        grab().crop(screen_to_window(CANVAS)).save(DOCS / f"{page}.png")
        print(f"docs/{page}.png")
    click(39, 378, 6)                # model view: the star schema
    click(1909, 263)                 # fold the Data pane
    click(1509, 263)                 # fold the Properties pane
    click(1887, 1131, 3)             # fit to screen
    win32api.SetCursorPos((1000, 1125))  # status bar: no tooltip on the diagram
    time.sleep(1.5)
    grab().crop(screen_to_window((60, 228, 1190, 1045))).save(DOCS / "model-view.png")
    print("docs/model-view.png")
    click(39, 263, 3)                # back to the report view
    goto("overview")                 # the .pbix opens on the first page
    if a.pbix:
        save_pbix()
    if not a.keep_open:
        subprocess.run(["taskkill", "/IM", "PBIDesktop.exe", "/F"], capture_output=True)


def control_click(h, wait=1.5):
    front()
    l, t, r, b = win32gui.GetWindowRect(h)
    win32api.SetCursorPos(((l + r) // 2, (t + b) // 2))
    time.sleep(0.2)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(wait)


def save_pbix():
    """File > Save as (F12), "Save as type" = Power BI file (*.pbix), Save. The data loaded by Refresh goes inside the
    .pbix, so it opens with numbers on any PC. The dialog's controls are found by their window class, then clicked."""
    out = ROOT / "Copperleaf-Finance.pbix"
    if out.exists():
        out.unlink()
    keys("{F12}", 5)
    dlg = next(h for h in windows() if win32gui.GetClassName(h) == "#32770" and win32gui.GetWindowText(h) == "Save As")
    kids = []
    win32gui.EnumChildWindows(dlg, lambda h, _: kids.append(h), None)
    type_box = next(h for h in kids if win32gui.GetClassName(h) == "ComboBox" and win32gui.SendMessage(h, win32con.CB_GETCOUNT, 0, 0) == 3)
    control_click(type_box)          # open the list; "Power BI file (*.pbix)" is the first entry
    keys("{HOME}", 0.5)
    keys("{ENTER}", 1.5)
    control_click(win32gui.GetDlgItem(dlg, 1), 25)  # Save
    print("Copperleaf-Finance.pbix" if out.exists() else "pbix NOT saved")


if __name__ == "__main__":
    main()
