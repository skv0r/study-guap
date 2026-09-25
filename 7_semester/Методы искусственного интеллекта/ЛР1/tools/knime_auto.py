"""Drive KNIME Modern UI via Windows UI Automation."""
from __future__ import annotations

import sys
import time
from pathlib import Path

import uiautomation as auto

SHOT = Path(__file__).resolve().parents[1] / "screenshots"


def find_knime():
    wins = []
    for w in auto.GetRootControl().GetChildren():
        name = w.Name or ""
        if "KNIME Analytics Platform" in name and "Setup" not in name and "Word" not in name:
            wins.append(w)
        elif name == "KNIME Analytics Platform":
            wins.append(w)
    return wins


def dump(ctrl, depth=0, max_depth=6, lines=None):
    if lines is None:
        lines = []
    if depth > max_depth:
        return lines
    try:
        rect = ctrl.BoundingRectangle
        r = f"{rect.left},{rect.top},{rect.right},{rect.bottom}"
    except Exception:
        r = "?"
    try:
        lines.append(
            f"{'  '*depth}{ctrl.ControlTypeName} name={ctrl.Name!r} auto={ctrl.AutomationId!r} class={ctrl.ClassName!r} rect={r}"
        )
    except Exception as e:
        lines.append(f"{'  '*depth}ERR {e}")
        return lines
    try:
        for c in ctrl.GetChildren():
            dump(c, depth + 1, max_depth, lines)
    except Exception:
        pass
    return lines


def main():
    auto.SetGlobalSearchTimeout(3.0)
    cmd = sys.argv[1] if len(sys.argv) > 1 else "dump"
    wins = find_knime()
    print("windows", [(w.Name, w.ClassName, w.NativeWindowHandle) for w in wins])
    if not wins:
        sys.exit(2)
    w = wins[0]
    w.SetActive()
    w.SetFocus()
    time.sleep(0.3)
    if cmd == "dump":
        depth = int(sys.argv[2]) if len(sys.argv) > 2 else 5
        lines = dump(w, max_depth=depth)
        out = SHOT / "uia_dump.txt"
        out.write_text("\n".join(lines), encoding="utf-8")
        print("dumped", len(lines), "lines to", out)
        for line in lines[:80]:
            print(line)
    elif cmd == "clickname":
        name = sys.argv[2]
        ctrl = w.Control(searchDepth=12, Name=name)
        print("found", ctrl.Name, ctrl.ControlTypeName, ctrl.BoundingRectangle)
        ctrl.Click()
        time.sleep(0.8)
    elif cmd == "listnames":
        names = []
        def walk(c, d=0):
            if d > 10:
                return
            if c.Name:
                names.append((c.ControlTypeName, c.Name, str(c.BoundingRectangle)))
            for ch in c.GetChildren():
                walk(ch, d + 1)
        walk(w)
        for t, n, r in names:
            print(f"{t}\t{n}\t{r}")
    else:
        print("unknown", cmd)


if __name__ == "__main__":
    main()
