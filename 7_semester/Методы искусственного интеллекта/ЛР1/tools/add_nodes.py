"""Add KNIME nodes: click/type in one session without Alt-stealing focus."""
from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import knime_ui as k

SHOT = Path(__file__).resolve().parents[1] / "screenshots"


def cursor():
    pt = k.wintypes.POINT()
    k.user32.GetCursorPos(pt)
    return pt.x, pt.y


def click_abs(x, y, *, double=False):
    x, y = int(x), int(y)
    k.user32.SetCursorPos(x, y)
    time.sleep(0.12)
    print("cursor", cursor(), "target", x, y)
    k.click(x, y)
    if double:
        time.sleep(0.08)
        k.click(x, y)


def img_to_screen(hwnd, ix, iy):
    r = k.rect(hwnd)
    return r.left + ix, r.top + iy


def main():
    wins = k.find_knime()
    print("windows", wins)
    hwnd = wins[0][0]
    # focus once, then ESC to clear Alt menu
    k.foreground(hwnd)
    k.hotkey(0x1B)  # ESC
    time.sleep(0.2)
    r = k.rect(hwnd)
    print("rect", r.left, r.top, r.right, r.bottom, "size", r.right - r.left, r.bottom - r.top)

    mode = sys.argv[1] if len(sys.argv) > 1 else "search"

    if mode == "search":
        # Search box center in screenshot: ~ (185, 114)
        sx, sy = img_to_screen(hwnd, 185, 114)
        click_abs(sx, sy)
        time.sleep(0.25)
        k.type_text("CSV Reader")
        time.sleep(0.8)
        k.screenshot(hwnd, SHOT / "knime_search_csv.png")
    elif mode == "csvclick":
        # CSV Reader icon ~ (120, 450) in full screenshot
        sx, sy = img_to_screen(hwnd, 120, 450)
        click_abs(sx, sy, double=True)
        time.sleep(0.8)
        k.screenshot(hwnd, SHOT / "knime_csv_clicked.png")
    elif mode == "dragcsv":
        x0, y0 = img_to_screen(hwnd, 120, 450)
        x1, y1 = img_to_screen(hwnd, 900, 420)
        k.user32.SetCursorPos(int(x0), int(y0))
        time.sleep(0.15)
        k.click(x0, y0)  # pick
        time.sleep(0.1)
        # press down, move, up
        sw = k.user32.GetSystemMetrics(0)
        sh = k.user32.GetSystemMetrics(1)

        def abs_xy(x, y):
            return int(x * 65535 / max(sw - 1, 1)), int(y * 65535 / max(sh - 1, 1))

        ax0, ay0 = abs_xy(x0, y0)
        ax1, ay1 = abs_xy(x1, y1)
        MOVE, DOWN, UP = 0x0001 | 0x8000, 0x0002 | 0x8000, 0x0004 | 0x8000
        seq = []
        for ax, ay, flags in ((ax0, ay0, MOVE), (ax0, ay0, DOWN), (ax1, ay1, MOVE), (ax1, ay1, UP)):
            inp = k.INPUT()
            inp.type = 0
            inp.union.mi = k.MOUSEINPUT(ax, ay, 0, flags, 0, 0)
            seq.append(inp)
        k._send(seq)
        time.sleep(0.8)
        k.screenshot(hwnd, SHOT / "knime_csv_drag.png")
    elif mode == "kai":
        # K-AI rail button ~ (22, 268)
        sx, sy = img_to_screen(hwnd, 22, 270)
        click_abs(sx, sy)
        time.sleep(0.8)
        k.screenshot(hwnd, SHOT / "knime_kai.png")
    else:
        raise SystemExit(mode)


if __name__ == "__main__":
    main()
