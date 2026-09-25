"""Click and screenshot the KNIME window."""
import ctypes
import sys
import time
from ctypes import wintypes
from pathlib import Path

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    user32.SetProcessDPIAware()

class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

def find_knime():
    found = []
    @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    def cb(hwnd, _):
        if user32.IsWindowVisible(hwnd):
            n = user32.GetWindowTextLengthW(hwnd)
            buf = ctypes.create_unicode_buffer(n + 1)
            user32.GetWindowTextW(hwnd, buf, n + 1)
            t = buf.value
            if t == "KNIME Analytics Platform":
                found.append((hwnd, t))
        return True
    user32.EnumWindows(cb, 0)
    return found

def rect(hwnd):
    r = RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(r))
    return r

ULONG_PTR = ctypes.c_ulonglong if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_ulong

class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", ctypes.c_long), ("dy", ctypes.c_long),
        ("mouseData", ctypes.c_ulong), ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong), ("dwExtraInfo", ULONG_PTR),
    ]

class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.c_ushort), ("wScan", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong), ("time", ctypes.c_ulong),
        ("dwExtraInfo", ULONG_PTR),
    ]

class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ("uMsg", ctypes.c_ulong), ("wParamL", ctypes.c_ushort), ("wParamH", ctypes.c_ushort),
    ]

class INPUTUNION(ctypes.Union):
    _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT), ("hi", HARDWAREINPUT)]

class INPUT(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong), ("union", INPUTUNION)]

SendInput = user32.SendInput
SendInput.argtypes = [ctypes.c_uint, ctypes.POINTER(INPUT), ctypes.c_int]
SendInput.restype = ctypes.c_uint

kernel32 = ctypes.windll.kernel32


def _send(inputs):
    arr = (INPUT * len(inputs))(*inputs)
    sent = SendInput(len(inputs), arr, ctypes.sizeof(INPUT))
    if sent != len(inputs):
        print("SendInput sent", sent, "expected", len(inputs), "err", kernel32.GetLastError())


def foreground(hwnd):
    user32.ShowWindow(hwnd, 9)  # SW_RESTORE
    user32.ShowWindow(hwnd, 3)  # SW_MAXIMIZE
    fg = user32.GetForegroundWindow()
    pid = wintypes.DWORD()
    fg_tid = user32.GetWindowThreadProcessId(fg, ctypes.byref(pid))
    cur_tid = kernel32.GetCurrentThreadId()
    user32.AttachThreadInput(cur_tid, fg_tid, True)
    # Alt down/up to bypass foreground lock
    user32.keybd_event(0x12, 0, 0, 0)
    user32.SetForegroundWindow(hwnd)
    user32.BringWindowToTop(hwnd)
    user32.keybd_event(0x12, 0, 2, 0)
    user32.AttachThreadInput(cur_tid, fg_tid, False)
    time.sleep(0.4)
    print("foreground now", user32.GetForegroundWindow(), "target", hwnd)


def click(x, y):
    x, y = int(x), int(y)
    sw = user32.GetSystemMetrics(0)
    sh = user32.GetSystemMetrics(1)
    ax = int(x * 65535 / max(sw - 1, 1))
    ay = int(y * 65535 / max(sh - 1, 1))
    MOVE = 0x0001 | 0x8000
    DOWN = 0x0002
    UP = 0x0004
    seq = []
    for flags in (MOVE, DOWN, UP):
        inp = INPUT()
        inp.type = 0
        if flags == MOVE:
            inp.union.mi = MOUSEINPUT(ax, ay, 0, flags, 0, 0)
        else:
            inp.union.mi = MOUSEINPUT(0, 0, 0, flags, 0, 0)
        seq.append(inp)
    _send(seq)
    time.sleep(0.15)


def hotkey(*vks):
    seq = []
    for vk in vks:
        inp = INPUT()
        inp.type = 1
        inp.union.ki = KEYBDINPUT(vk, 0, 0, 0, 0)
        seq.append(inp)
    for vk in reversed(vks):
        inp = INPUT()
        inp.type = 1
        inp.union.ki = KEYBDINPUT(vk, 0, 0x0002, 0, 0)
        seq.append(inp)
    _send(seq)
    time.sleep(0.2)


def type_text(text: str):
    KEYEVENTF_UNICODE = 0x0004
    KEYEVENTF_KEYUP = 0x0002
    seq = []
    for ch in text:
        code = ord(ch)
        down = INPUT(); down.type = 1
        down.union.ki = KEYBDINPUT(0, code, KEYEVENTF_UNICODE, 0, 0)
        up = INPUT(); up.type = 1
        up.union.ki = KEYBDINPUT(0, code, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP, 0, 0)
        seq.extend([down, up])
    _send(seq)
    time.sleep(0.1)

def screenshot(hwnd, path: Path):
    r = rect(hwnd)
    w, h = r.right - r.left, r.bottom - r.top
    hwnd_dc = user32.GetWindowDC(hwnd)
    mem_dc = gdi32.CreateCompatibleDC(hwnd_dc)
    bmp = gdi32.CreateCompatibleBitmap(hwnd_dc, w, h)
    gdi32.SelectObject(mem_dc, bmp)
    # PrintWindow
    user32.PrintWindow(hwnd, mem_dc, 2)
    # BITMAPINFO
    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [
            ("biSize", ctypes.c_uint32), ("biWidth", ctypes.c_int32),
            ("biHeight", ctypes.c_int32), ("biPlanes", ctypes.c_uint16),
            ("biBitCount", ctypes.c_uint16), ("biCompression", ctypes.c_uint32),
            ("biSizeImage", ctypes.c_uint32), ("biXPelsPerMeter", ctypes.c_int32),
            ("biYPelsPerMeter", ctypes.c_int32), ("biClrUsed", ctypes.c_uint32),
            ("biClrImportant", ctypes.c_uint32),
        ]
    class BITMAPINFO(ctypes.Structure):
        _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", ctypes.c_uint32 * 3)]
    bmi = BITMAPINFO()
    bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bmi.bmiHeader.biWidth = w
    bmi.bmiHeader.biHeight = -h
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 32
    buf = ctypes.create_string_buffer(w * h * 4)
    gdi32.GetDIBits(mem_dc, bmp, 0, h, buf, ctypes.byref(bmi), 0)
    gdi32.DeleteObject(bmp)
    gdi32.DeleteDC(mem_dc)
    user32.ReleaseDC(hwnd, hwnd_dc)
    from PIL import Image
    img = Image.frombuffer("RGBA", (w, h), buf, "raw", "BGRA", 0, 1)
    path.parent.mkdir(parents=True, exist_ok=True)
    img.convert("RGB").save(path)
    print("saved", path, img.size)

def main():
    wins = find_knime()
    print("windows", wins)
    if not wins:
        sys.exit(2)
    hwnd, title = wins[0]
    r = rect(hwnd)
    print("rect", r.left, r.top, r.right, r.bottom, title)
    cmd = sys.argv[1] if len(sys.argv) > 1 else "shot"
    out = Path(sys.argv[-1]) if len(sys.argv) > 2 else Path("knime.png")
    foreground(hwnd)
    if cmd == "click":
        rx, ry = float(sys.argv[2]), float(sys.argv[3])
        x = r.left + rx * (r.right - r.left)
        y = r.top + ry * (r.bottom - r.top)
        print("click", x, y)
        click(x, y)
        time.sleep(0.8)
        screenshot(hwnd, Path(sys.argv[4]) if len(sys.argv) > 4 else out)
    elif cmd == "clickxy":
        x, y = int(sys.argv[2]), int(sys.argv[3])
        click(x, y)
        time.sleep(0.8)
        screenshot(hwnd, Path(sys.argv[4]))
    elif cmd == "keys":
        # keys ctrl+n / alt+f / enter ...
        mapping = {
            "ctrl": 0x11, "alt": 0x12, "shift": 0x10, "enter": 0x0D,
            "tab": 0x09, "esc": 0x1B, "f8": 0x77, "del": 0x2E,
            "back": 0x08, "space": 0x20, "down": 0x28, "up": 0x26,
            "left": 0x25, "right": 0x27, "home": 0x24, "end": 0x23,
        }
        vks = []
        for tok in sys.argv[2].lower().split("+"):
            if tok in mapping:
                vks.append(mapping[tok])
            elif len(tok) == 1:
                vks.append(ord(tok.upper()))
            else:
                raise SystemExit("unknown key " + tok)
        print("hotkey", sys.argv[2], vks)
        hotkey(*vks)
        time.sleep(0.8)
        if len(sys.argv) > 3:
            screenshot(hwnd, Path(sys.argv[3]))
    elif cmd == "type":
        type_text(sys.argv[2])
        time.sleep(0.3)
        if len(sys.argv) > 3:
            screenshot(hwnd, Path(sys.argv[3]))
    else:
        screenshot(hwnd, out)

if __name__ == "__main__":
    main()
