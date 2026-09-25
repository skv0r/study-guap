"""Return to workflow tab and try Ctrl+Space node insert."""
from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import knime_ui as k

SHOT = Path(__file__).resolve().parents[1] / "screenshots"


def main():
    hwnd = k.find_knime()[0][0]
    k.foreground(hwnd)
    time.sleep(0.25)
    k.hotkey(0x1B)
    time.sleep(0.2)
    k.hotkey(0x11, 0x09)  # Ctrl+Tab
    time.sleep(0.6)
    k.screenshot(hwnd, SHOT / "knime_ctrl_tab.png")
    k.hotkey(0x11, 0x20)  # Ctrl+Space
    time.sleep(0.6)
    k.screenshot(hwnd, SHOT / "knime_ctrl_space.png")


if __name__ == "__main__":
    main()
