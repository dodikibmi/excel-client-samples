"""Excel only lets a script write VBA code (modules, the UserForm) when "Trust access to the VBA project object model"
is on. vba_access() switches it on for the current user while the build/test runs and puts the old value back after."""
import contextlib, winreg

KEY = r"Software\Microsoft\Office\16.0\Excel\Security"


@contextlib.contextmanager
def vba_access():
    k = winreg.CreateKey(winreg.HKEY_CURRENT_USER, KEY)
    try:
        old = winreg.QueryValueEx(k, "AccessVBOM")[0]
    except FileNotFoundError:
        old = None
    winreg.SetValueEx(k, "AccessVBOM", 0, winreg.REG_DWORD, 1)
    try:
        yield
    finally:
        if old is None:
            winreg.DeleteValue(k, "AccessVBOM")
        else:
            winreg.SetValueEx(k, "AccessVBOM", 0, winreg.REG_DWORD, old)
        winreg.CloseKey(k)
