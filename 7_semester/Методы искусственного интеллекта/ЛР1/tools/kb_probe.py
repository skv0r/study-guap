"""Keyboard-only exploration of KNIME Modern UI."""
from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import knime_ui as k

SHOT = Path(__file__).resolve().parents[1] / "screenshots"


def shot(hwnd, name):
    k.screenshot(hwnd, SHOT / name)
    print("shot", name)


def main():
    hwnd = k.find_knime()[0][0]
    k.foreground(hwnd)
    time.sleep(0.3)
    k.hotkey(0x1B)  # esc
    time.sleep(0.2)

    seq = sys.argv[1] if len(sys.argv) > 1 else "slash"
    if seq == "slash":
        k.type_text("/")
        time.sleep(0.5)
        shot(hwnd, "knime_slash.png")
    elif seq == "ctrlf":
        k.hotkey(0x11, ord("F"))
        time.sleep(0.5)
        shot(hwnd, "knime_ctrlf.png")
    elif seq == "ctrlk":
        k.hotkey(0x11, ord("K"))
        time.sleep(0.5)
        shot(hwnd, "knime_ctrlk.png")
    elif seq == "tabn":
        n = int(sys.argv[2]) if len(sys.argv) > 2 else 5
        for i in range(n):
            k.hotkey(0x09)
            time.sleep(0.25)
        shot(hwnd, f"knime_tab{n}.png")
    elif seq == "typec":
        k.type_text("csv")
        time.sleep(0.6)
        shot(hwnd, "knime_typec.png")
    elif seq == "space":
        k.hotkey(0x20)
        time.sleep(0.4)
        shot(hwnd, "knime_space.png")
    elif seq == "f1":
        k.hotkey(0x70)
        time.sleep(0.5)
        shot(hwnd, "knime_f1.png")
    else:
        raise SystemExit(seq)


if __name__ == "__main__":
    main()
