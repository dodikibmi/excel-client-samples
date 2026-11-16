"""Screenshots for docs/ taken from Power BI Desktop itself (Store version), plus the .pbix copy (E15).

Opens Larkspur-Sales.pbip, closes the optional sign-in window (Desktop works without an account), presses
Refresh now, folds the side panes away and grabs each report page and the model view. With --pbix it also saves a
.pbix copy (File > Save as). Window positions assume a 1920x1200 screen at 125% scaling with Desktop maximised.
Run:  python scripts/shots.py [--pbix] [--keep-open]
"""
import argparse, ctypes, os, subprocess, sys, time
from pathlib import Path

ctypes.windll.shcore.SetProcessDpiAwareness(2)
import win32api, win32con, win32gui, win32process, win32ui  # noqa: E402
import win32com.client  # noqa: E402
from PIL import Image, ImageChops  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
PBIP = ROOT / "Larkspur-Sales.pbip"
DOCS = ROOT / "docs"
PAGES = ["overview", "products", "customers", "markets-and-fx", "checks"]
NAV_X, NAV_W, NAV_Y = 560, 700, 30   # the page buttons in the report header (page pixels, see build_pbip.py)
RULE_RGB = (0xC2, 0x60, 0x3D)  # the terracotta line under the header of every report page (page y = 60..62)


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


def page_box(img):
    """The report page inside the window: found from the terracotta line under the header (full page width, page
    y = 60), so the crop works at any screen size or zoom. Height follows the 1280x720 page."""
    bands = [b.point(lambda v, t=t: 255 if abs(v - t) < 6 else 0) for b, t in zip(img.convert("RGB").split(), RULE_RGB)]
    mask = ImageChops.multiply(ImageChops.multiply(bands[0], bands[1]), bands[2])
    w, h = mask.size
    for y in range(200, h):          # below the ribbon: the first row that is mostly the line colour
        row = mask.crop((0, y, w, y + 1))
        if row.histogram()[255] > w * 0.4:
            x0, _, x1, _ = row.getbbox()
            k = (x1 - x0) / 1280
            top = y - round(60 * k)
            return (x0, top, x1, top + round((x1 - x0) * 720 / 1280))
    raise SystemExit("report page not found in the window")


def deselect():
    """Click the grey area right of the page so no visual stays selected."""
    x0, y0, x1, y1 = page_box(grab())
    click(x1 + 30, (y0 + y1) // 2, 1.5)


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


def privacy_dialog():
    """A fresh PBIP may ask for privacy levels before it combines the CSV files (all local, fictional): tick
    "Ignore Privacy Levels checks for this file" and Save."""
    for h in windows():
        l, t, r, b = win32gui.GetWindowRect(h)
        if exe_name(h) == "pbidesktop.exe" and not win32gui.GetWindowText(h) and abs(r - l - 914) < 30 and 400 < b - t < 800:
            for x, y in ((48, 201), (722, b - t - 56)):   # the tick box, then Save (the list grows with the files)
                keys_raw("%")
                win32gui.SetForegroundWindow(h)
                time.sleep(0.8)
                win32api.SetCursorPos((l + x, t + y))
                time.sleep(0.2)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
                time.sleep(0.05)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
                time.sleep(1.5)
            return True
    return False


def viz_pane_open():
    """The Visualizations pane is open when its blue visual icons show right of the page."""
    img = grab().convert("RGB")
    return any(img.getpixel((x, y))[2] > 170 and img.getpixel((x, y))[0] < 110
               for x in range(1650, 1850, 3) for y in range(390, 470, 3))


def data_bar_shown():
    """The light "incomplete or no data" bar: a dark info icon at its left end, light grey behind the text."""
    img = grab().convert("L")
    icon = min(img.crop((86, 250, 108, 272)).getdata())
    text_bg = list(img.crop((300, 250, 600, 270)).getdata())
    return icon < 120 and sum(text_bg) / len(text_bg) > 200


def page_width():
    try:
        x0, _, x1, _ = page_box(grab())
        return x1 - x0
    except SystemExit:
        return 0


def leave_focus_mode():
    """Desktop's auto-recovery can reopen a visual in focus mode ("Back to report"): leave it."""
    if page_width() == 0:
        click(166, 293, 3)


def refresh_now(wait=45):
    """Home > Refresh (the ribbon icon, same place whatever the panes do) until the data is in."""
    for _ in range(3):
        if not data_bar_shown():
            return
        click(872, 138, 8)
        privacy_dialog()
        time.sleep(wait)
    if data_bar_shown():
        raise SystemExit("Refresh did not load the data")


def fold_panes():
    """Fold the Visualizations and Data panes: each chevron is clicked only if it makes the page wider."""
    for x in (1651, 1859, 1909):
        before = page_width()
        if before >= 1460:          # folded: the page fills the window
            return
        click(x, 264, 2)
        if page_width() < before:
            click(x, 264, 2)         # that opened a pane instead: close it again


def refresh_and_fold(refresh=True):
    """Window positions for a 1938x1158 maximised window (Power BI Desktop, November 2026)."""
    leave_focus_mode()
    if refresh:
        refresh_now()
    fold_panes()


def goto(page):
    """Ctrl+click the page button in the report header (in Desktop a plain click only selects the button)."""
    x0, y0, x1, _ = page_box(grab())
    k = (x1 - x0) / 1280
    i = PAGES.index(page)
    front()
    win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
    try:
        click(x0 + round((NAV_X + (i + 0.5) * NAV_W / len(PAGES)) * k), y0 + round(NAV_Y * k), 0.5)
    finally:
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(4)
    if page == "products":     # right-click the matrix > Expand > All: show every account under its category
        x0, y0, x1, _ = page_box(grab())
        k = (x1 - x0) / 1280             # page pixels -> window pixels at the current zoom
        x, y = x0 + round(76 * k), y0 + round(240 * k)   # the first category row of the matrix
        click(x, y, 2, right=True)
        click(x + 97, y + 53, 2)         # Expand >
        click(x + 356, y + 120, 3)       # All
        click(x0 + round(640 * k), y0 + round(711 * k), 1.5)  # the empty strip under the visuals: nothing stays selected
    park_mouse()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pbix", action="store_true")
    ap.add_argument("--keep-open", action="store_true")
    a = ap.parse_args()
    DOCS.mkdir(exist_ok=True)
    open_report()
    time.sleep(5)
    privacy_dialog()                 # first refresh on a new PC asks for privacy levels
    refresh_and_fold()
    for page in PAGES:
        goto(page)
        img = grab()
        img.crop(page_box(img)).save(DOCS / f"{page}.png")
        print(f"docs/{page}.png")
    click(39, 378, 6)                # model view: the star schema
    click(1909, 263)                 # fold the Data pane
    click(1509, 263)                 # fold the Properties pane
    click(1887, 1131, 3)             # fit to screen
    win32api.SetCursorPos((1000, 1125))  # status bar: no tooltip on the diagram
    time.sleep(1.5)
    grab().crop((125, 270, 1107, 1024)).save(DOCS / "model-view.png")  # window pixels, 1938x1158 window
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
    out = ROOT / "Larkspur-Sales.pbix"
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
    print("Larkspur-Sales.pbix" if out.exists() else "pbix NOT saved")


if __name__ == "__main__":
    main()
