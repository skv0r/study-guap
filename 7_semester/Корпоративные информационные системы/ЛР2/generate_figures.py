# -*- coding: utf-8 -*-
"""Рисует диаграммы IDEF0. В текст отчёта не входит."""

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, Rectangle

from idef_model import DIAGRAMS, check_balance

OUT = Path(__file__).resolve().parent / "figures"
FONT = "Times New Roman"
W = 486.0
BH = 48.0


def wrap(text, limit):
    if "\n" in text:
        return text
    words = text.split()
    lines, cur = [], ""
    for word in words:
        trial = (cur + " " + word).strip()
        if len(trial) <= limit:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = word
    if cur:
        lines.append(cur)
    return "\n".join(lines)


def T(ax, x, y, text, size=16, ha="center", va="center"):
    ax.text(
        x, y, text, ha=ha, va=va, fontsize=size, fontname=FONT, color="black",
        linespacing=1.02,
        bbox=dict(facecolor="white", edgecolor="none", pad=0.08),
        zorder=6, clip_on=False,
    )


def arrow(ax, p, q):
    ax.add_patch(FancyArrowPatch(
        p, q, arrowstyle="-|>", mutation_scale=13, lw=1.2, color="black",
        shrinkA=0, shrinkB=0, zorder=3, clip_on=False,
    ))


def poly(ax, pts):
    for a, b in zip(pts, pts[1:-1]):
        ax.plot([a[0], b[0]], [a[1], b[1]], color="black", lw=1.2, zorder=3, solid_capstyle="butt")
    arrow(ax, pts[-2], pts[-1])


def hline(ax, x1, x2, y):
    ax.plot([x1, x2], [y, y], color="black", lw=1.2, zorder=3, solid_capstyle="butt")


def title_of(spec):
    if spec["id"] == "А-0":
        return "А-0    Ремонтировать велосипеды"
    if "«" in spec["section"]:
        name = spec["section"].split("«", 1)[1].rstrip("»")
        return f"{spec['id']}    {name}"
    return spec["id"] + "    " + spec["section"]


def render(spec):
    boxes = spec["boxes"]
    n = len(boxes)
    flows = list(spec.get("flow", []))
    ups = list(spec.get("up", []))
    ins = spec.get("in", [])
    outs = spec.get("out", [])
    ctrls = spec.get("ctrl", [])
    mechs = spec.get("mech", [])
    purpose = spec.get("purpose")

    lanes = []
    for item in ins:
        for target in item["to"]:
            if n > 1 and target != 0:
                lanes.append(("in", item["label"], target))
    for flow_item in flows:
        if flow_item["to"] != flow_item["frm"] + 1:
            lanes.append(("skip", flow_item["label"], flow_item))
    for item in outs:
        if item["frm"] != n - 1:
            lanes.append(("out", item["label"], item["frm"]))

    n_lane = len(lanes)
    n_mech = max(len(mechs), 1)
    n_ctrl = max(len(ctrls), 1)

    y = 12 + (18 if purpose else 0)
    lane_y = []
    for _ in range(max(n_lane, 0)):
        lane_y.append(y + 4)
        y += 16
    if n_lane:
        y += 8
    else:
        y += 4
    mech_y = []
    for _ in range(len(mechs)):
        mech_y.append(y + 6)
        y += 16
    y += 8
    yb = y
    top = yb + BH
    y = top + 10
    up_y = y if ups else None
    if ups:
        y += 14
    ctrl_y = []
    for _ in range(len(ctrls)):
        ctrl_y.append(y)
        y += 16
    y += 16
    height = y

    left, right = 70, 78
    if n == 1:
        bw = 120
        xs = [(W - bw) / 2.0]
    else:
        bw = 62 if n <= 3 else 50
        usable = W - left - right
        gap = (usable - n * bw) / (n - 1)
        xs = [left + i * (bw + gap) for i in range(n)]

    fig, ax = plt.subplots(figsize=(W / 25.4, height / 25.4), dpi=140)
    ax.set_xlim(0, W)
    ax.set_ylim(0, height)
    ax.set_aspect("equal")
    ax.axis("off")
    fig.subplots_adjust(0, 0, 1, 1)
    ax.add_patch(Rectangle((1.5, 1.5), W - 3, height - 3, fill=False, ec="black", lw=1.1, zorder=1))
    T(ax, W / 2, height - 9, title_of(spec), size=18)

    side_levels = (0.80, 0.28, 0.54)

    def sy(i):
        return yb + BH * side_levels[min(i, 2)]

    straight = [item for item in ins if 0 in item["to"]]
    for i, item in enumerate(straight):
        yy = sy(i)
        arrow(ax, (4, yy), (xs[0], yy))
        T(ax, 6, yy + 1.6, wrap(item["label"], 20), size=16, ha="left", va="bottom")

    for i, (kind, label, payload) in enumerate(lanes):
        ly = lane_y[i]
        if kind == "in":
            target = payload
            entry = yb + 7
            poly(ax, [(4, ly), (xs[target] - 6, ly), (xs[target] - 6, entry), (xs[target], entry)])
            T(ax, 6, ly + 1.8, wrap(label, 22), size=16, ha="left", va="bottom")
        elif kind == "skip":
            flow_item = payload
            a, b = flow_item["frm"], flow_item["to"]
            leave = yb + BH * 0.30
            poly(ax, [
                (xs[a] + bw, leave),
                (xs[a] + bw + 5, leave),
                (xs[a] + bw + 5, ly),
                (xs[b] - 7, ly),
                (xs[b] - 7, leave),
                (xs[b], leave),
            ])
            T(ax, (xs[a] + xs[b] + bw) / 2, ly + 1.8, wrap(label, 18), size=16, va="bottom")
        else:
            src = payload
            leave = yb + BH * 0.5
            poly(ax, [
                (xs[src] + bw, leave),
                (xs[src] + bw + 5, leave),
                (xs[src] + bw + 5, ly),
                (W - 6, ly),
            ])
            T(ax, (xs[src] + bw + W) / 2, ly + 1.8, wrap(label, 22), size=16, va="bottom")

    pair = {}
    for flow_item in flows:
        if flow_item["to"] == flow_item["frm"] + 1:
            pair.setdefault((flow_item["frm"], flow_item["to"]), []).append(flow_item)
    for (a, b), group in pair.items():
        for i, flow_item in enumerate(group[:3]):
            yy = sy(i)
            arrow(ax, (xs[a] + bw, yy), (xs[b], yy))
            if i == 0:
                T(ax, (xs[a] + bw + xs[b]) / 2, yy + 1.5, wrap(flow_item["label"], 16), size=16, va="bottom")
            else:
                T(ax, (xs[a] + bw + xs[b]) / 2, yy - 1.5, wrap(flow_item["label"], 16), size=16, va="top")

    last_outs = [item for item in outs if item["frm"] == n - 1]
    for i, item in enumerate(last_outs):
        yy = sy(i)
        arrow(ax, (xs[n - 1] + bw, yy), (W - 4, yy))
        if i == 0:
            T(ax, (xs[n - 1] + bw + W - 4) / 2, yy + 1.5, wrap(item["label"], 18), size=16, va="bottom")
        else:
            T(ax, (xs[n - 1] + bw + W - 4) / 2, yy - 1.5, wrap(item["label"], 18), size=16, va="top")

    # входы сверху: управления и стрелки «выход → управление»
    top_used = [0] * n

    def take_top(index):
        k = top_used[index]
        top_used[index] = k + 1
        fracs = (0.28, 0.50, 0.74, 0.86)
        return xs[index] + bw * fracs[min(k, 3)]

    for up in ups:
        x1 = take_top(up["frm"])
        x2 = take_top(up["to"])
        poly(ax, [(x1, top), (x1, up_y), (x2, up_y), (x2, top)])
        T(ax, (x1 + x2) / 2, up_y + 1.6, wrap(up["label"], 18), size=16, va="bottom")

    for ci, ctrl in enumerate(ctrls):
        bus = ctrl_y[ci]
        drops = [take_top(i) for i in ctrl["to"]]
        if len(drops) == 1:
            arrow(ax, (drops[0], bus + 9), (drops[0], top))
            T(ax, drops[0], bus + 10, wrap(ctrl["label"], 20), size=16, va="bottom")
        else:
            hline(ax, min(drops), max(drops), bus)
            for x in drops:
                arrow(ax, (x, bus), (x, top))
            T(ax, sum(drops) / len(drops), bus + 1.6, wrap(ctrl["label"], 20), size=16, va="bottom")

    for mi, mech in enumerate(mechs):
        bus = mech_y[mi]
        frac = 0.36 + mi * 0.22
        drops = [xs[i] + bw * min(frac, 0.78) for i in mech["to"]]
        if len(drops) == 1:
            arrow(ax, (drops[0], bus), (drops[0], yb))
            T(ax, drops[0], bus - 1.6, wrap(mech["label"], 16), size=16, va="top")
        else:
            hline(ax, min(drops) - 2, max(drops), bus)
            for x in drops:
                arrow(ax, (x, bus), (x, yb))
            T(ax, min(drops) - 3, bus, wrap(mech["label"], 16), size=16, ha="right", va="center")

    for i, box in enumerate(boxes):
        x = xs[i]
        ax.add_patch(Rectangle((x, yb), bw, BH, fill=False, ec="black", lw=1.5, zorder=4))
        T(ax, x + bw / 2, yb + BH * 0.62, box["title"], size=16)
        num = box.get("num", str(i + 1))
        T(ax, x + bw - 2, yb + 2, num, size=16, ha="right", va="bottom")
        if box.get("child"):
            T(ax, x + bw - 2, yb + 8, box["child"], size=16, ha="right", va="bottom")

    if purpose:
        T(ax, W / 2, 14, purpose, size=16)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / spec["file"]
    fig.savefig(path, dpi=140, facecolor="white")
    plt.close(fig)
    spec["_height"] = height
    return path


def main():
    boxes, errors = check_balance()
    if errors:
        raise SystemExit("ICOM:\n" + "\n".join(errors))
    for spec in DIAGRAMS:
        path = render(spec)
        print(path.name, "h", round(spec["_height"], 1))
    print("diagrams", len(DIAGRAMS), "boxes", boxes)


if __name__ == "__main__":
    main()
