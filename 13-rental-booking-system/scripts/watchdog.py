"""A hidden Excel that hits a VBA error or a MsgBox waits forever for a click nobody can see.
watch(xl) starts a thread that looks for such a dialog every 2 s; if one appears it prints the dialog text and kills
that Excel, so the build or the test fails loudly instead of hanging."""
import os, threading, time
import win32gui, win32process


def _dialogs(pid):
    found = []

    def cb(h, _):
        if win32process.GetWindowThreadProcessId(h)[1] == pid and win32gui.IsWindowVisible(h) \
                and win32gui.GetClassName(h) == "#32770":
            texts = []
            win32gui.EnumChildWindows(h, lambda c, _: texts.append(win32gui.GetWindowText(c)) or True, None)
            found.append(win32gui.GetWindowText(h) + ": " + " | ".join(t for t in texts if t and t not in ("OK", "Help")))
        return True
    win32gui.EnumWindows(cb, None)
    return found


def watch(xl, allow=()):
    pid = win32process.GetWindowThreadProcessId(xl.Hwnd)[1]
    stop = threading.Event()

    def loop():
        while not stop.wait(2):
            d = [x for x in _dialogs(pid) if not any(a in x for a in allow)]
            if d:
                print("\n!! Excel dialog, killing Excel:", *d, sep="\n   ", flush=True)
                os.system(f"taskkill /PID {pid} /F >nul 2>&1")
                return
    threading.Thread(target=loop, daemon=True).start()
    return stop
