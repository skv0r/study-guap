# -*- coding: utf-8 -*-
"""Рисунки ЛР3. В текст отчёта не входит."""

from __future__ import annotations

import math
from pathlib import Path

import matplotlib.pyplot as plt
from PIL import Image, PngImagePlugin
from matplotlib.patches import Ellipse, FancyArrowPatch, FancyBboxPatch, Rectangle

OUT = Path(__file__).resolve().parent / "figures"
FONT = "Times New Roman"


def cm_fig(w, h):
    fig, ax = plt.subplots(figsize=(w / 2.54, h / 2.54), dpi=160)
    fig.subplots_adjust(0, 0, 1, 1)
    ax.set_xlim(0, w)
    ax.set_ylim(0, h)
    ax.set_aspect("equal")
    ax.axis("off")
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    return fig, ax


def save(fig, name):
    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / name
    fig.savefig(path, dpi=160, facecolor="white")
    plt.close(fig)
    with Image.open(path) as image:
        clean = Image.new(image.mode, image.size)
        clean.paste(image)
        clean.save(path, format="PNG", pnginfo=PngImagePlugin.PngInfo())
    return path


def txt(ax, x, y, s, size=12, weight="normal", ha="center", va="center", style="normal", color="black"):
    return ax.text(
        x, y, s,
        fontsize=size, fontname=FONT, fontweight=weight, fontstyle=style,
        ha=ha, va=va, color=color, zorder=5,
    )


def line(ax, x1, y1, x2, y2, lw=1.05, dashed=False):
    ax.plot(
        [x1, x2], [y1, y2],
        color="black", lw=lw, solid_capstyle="round",
        linestyle=(0, (3.2, 2.2)) if dashed else "-",
        zorder=2,
    )


def arrow(ax, x1, y1, x2, y2, dashed=False, open_head=False, lw=1.05):
    style = "-|>" if not open_head else "->"
    patch = FancyArrowPatch(
        (x1, y1), (x2, y2),
        arrowstyle=style,
        mutation_scale=11,
        lw=lw,
        color="black",
        linestyle=(0, (3.2, 2.2)) if dashed else "-",
        shrinkA=0,
        shrinkB=0,
        zorder=3,
        fill=not open_head,
    )
    ax.add_patch(patch)


def data_span(fig, ax, artist):
    fig.canvas.draw()
    bb = artist.get_window_extent(renderer=fig.canvas.get_renderer())
    inv = ax.transData.inverted()
    x0, y0 = inv.transform((bb.x0, bb.y0))
    x1, y1 = inv.transform((bb.x1, bb.y1))
    return min(x0, x1), max(x0, x1), min(y0, y1), max(y0, y1)


def text_width(fig, ax, s, size):
    probe = txt(ax, 0, 0, s, size=size, ha="left", va="bottom")
    probe.set_alpha(0)
    x0, x1, _, _ = data_span(fig, ax, probe)
    probe.remove()
    return x1 - x0


def actor(ax, x, y, name, side):
    r = 0.32
    head = plt.Circle((x, y), r, fill=False, lw=1.15, color="black", zorder=3)
    ax.add_patch(head)
    y_neck = y - r
    y_hip = y_neck - 0.95
    line(ax, x, y_neck, x, y_hip, 1.15)
    line(ax, x - 0.62, y_neck - 0.32, x + 0.62, y_neck - 0.32, 1.15)
    line(ax, x, y_hip, x - 0.48, y_hip - 0.85, 1.15)
    line(ax, x, y_hip, x + 0.48, y_hip - 0.85, 1.15)
    txt(ax, x, y_hip - 1.05, name, size=12, va="top")
    return x + side * 0.78, y_neck - 0.32


def ellipse_point(cx, cy, rx, ry, px, py):
    dx, dy = px - cx, py - cy
    norm = math.hypot(dx / rx, dy / ry)
    if norm == 0:
        return cx + rx, cy
    t = 1 / norm
    return cx + dx * t, cy + dy * t


def use_case(ax, cx, cy, rx, ry, label):
    ax.add_patch(Ellipse((cx, cy), rx * 2, ry * 2, fill=True, facecolor="white",
                         edgecolor="black", lw=1.15, zorder=3))
    txt(ax, cx, cy, label, size=11.5)


def draw_usecase():
    w, h = 24.0, 12.2
    fig, ax = cm_fig(w, h)
    ax.add_patch(Rectangle((4.15, 0.55), 15.7, 11.15, fill=False, lw=1.25, zorder=1))
    txt(ax, 12.0, 11.25, "СпринтКонтур", size=14, weight="bold")

    cases = {
        "reg": (8.15, 9.55, "Зарегистрировать\nсрочную задачу"),
        "assign": (8.15, 7.15, "Назначить\nразработчика"),
        "term": (8.15, 4.75, "Проверить\nсрок сдачи"),
        "close": (8.15, 2.35, "Закрыть задачу"),
        "tell": (15.85, 9.55, "Сообщить\nтребование"),
        "status": (15.85, 7.15, "Получить\nстатус задачи"),
        "accept": (15.85, 4.75, "Принять задачу"),
        "protocol": (15.85, 2.35, "Оформить протокол\nпроверки"),
    }
    rx, ry = 2.55, 0.82
    for cx, cy, label in cases.values():
        use_case(ax, cx, cy, rx, ry, label)

    lead = actor(ax, 2.05, 8.55, "Руководитель\nкоманды", side=1)
    customer = actor(ax, 21.95, 8.55, "Заказчик", side=-1)
    developer = actor(ax, 21.95, 3.55, "Разработчик", side=-1)

    def link(anchor, key, from_left):
        cx, cy, _ = cases[key]
        x2 = cx - rx if from_left else cx + rx
        line(ax, anchor[0], anchor[1], x2, cy)

    for key in ("reg", "assign", "close"):
        link(lead, key, True)
    for key in ("tell", "status"):
        link(customer, key, False)
    for key in ("accept", "protocol"):
        link(developer, key, False)

    x = cases["assign"][0]
    y_from = cases["assign"][1] - ry
    y_to = cases["term"][1] + ry
    arrow(ax, x, y_from - 0.02, x, y_to + 0.02, dashed=True, open_head=True, lw=1.05)
    txt(ax, x + 1.15, (y_from + y_to) / 2, "«include»", size=11, style="italic")
    return save(fig, "fig01_usecase.png")


def class_box(ax, x, top, w, name, attrs, ops):
    head = 0.72
    row = 0.46
    attr_h = 0.20 + len(attrs) * row + 0.10
    op_h = (0.16 + len(ops) * row + 0.12) if ops else 0.0
    h = head + attr_h + op_h
    bottom = top - h
    ax.add_patch(Rectangle((x, bottom), w, h, fill=True, facecolor="white", edgecolor="black", lw=1.15, zorder=3))
    sep1 = top - head
    ax.plot([x, x + w], [sep1, sep1], color="black", lw=1.0, zorder=4)
    txt(ax, x + w / 2, top - head / 2, name, size=13, weight="bold")
    for i, attr in enumerate(attrs):
        txt(ax, x + 0.22, sep1 - 0.12 - i * row, attr, size=11, ha="left", va="top")
    if ops:
        sep2 = bottom + op_h
        ax.plot([x, x + w], [sep2, sep2], color="black", lw=1.0, zorder=4)
        for i, op in enumerate(ops):
            txt(ax, x + 0.22, sep2 - 0.10 - i * row, "+ " + op, size=11, ha="left", va="top")
    return bottom


def draw_classes():
    w, h = 24.0, 12.8
    fig, ax = cm_fig(w, h)
    top = 12.35
    x_mod, w_mod = 0.25, 5.2
    x_task, w_task = 8.55, 6.7
    x_team, w_team = 18.15, 5.55
    b_mod = class_box(ax, x_mod, top, w_mod, "Модуль", ["код", "название", "версия"], [])
    b_task = class_box(
        ax, x_task, top, w_task, "Задача",
        ["номер", "признакСрочности", "состояние", "срокСдачи", "описание"],
        ["зарегистрировать()", "назначитьРазработчика()", "закрыть()"],
    )
    b_team = class_box(
        ax, x_team, top, w_team, "Команда",
        ["код", "руководитель", "естьСвободный"],
        ["принятьЗадачу()"],
    )
    act_top = b_task - 1.55
    b_proto = class_box(
        ax, x_task, act_top, w_task, "ПротоколПроверки",
        ["номер", "результат", "замечание", "дата"],
        ["оформить()"],
    )

    y_link = top - 1.45
    line(ax, x_mod + w_mod, y_link, x_task, y_link)
    txt(ax, x_mod + w_mod + 0.12, y_link - 0.18, "1", size=12, ha="left", va="top")
    txt(ax, x_task - 0.12, y_link - 0.18, "0..*", size=12, ha="right", va="top")
    txt(ax, (x_mod + w_mod + x_task) / 2, y_link + 0.12, "относится к", size=12, style="italic", va="bottom")

    line(ax, x_task + w_task, y_link, x_team, y_link)
    txt(ax, x_task + w_task + 0.12, y_link - 0.18, "0..*", size=12, ha="left", va="top")
    txt(ax, x_team - 0.12, y_link - 0.18, "0..1", size=12, ha="right", va="top")
    txt(ax, (x_task + w_task + x_team) / 2, y_link + 0.12, "назначена", size=12, style="italic", va="bottom")

    x_v = x_task + w_task / 2
    line(ax, x_v, b_task, x_v, act_top)
    txt(ax, x_v - 0.18, b_task - 0.1, "1", size=12, ha="right", va="top")
    txt(ax, x_v - 0.18, act_top + 0.1, "0..1", size=12, ha="right", va="bottom")
    txt(ax, x_v + 0.18, (b_task + act_top) / 2, "имеет", size=12, style="italic", ha="left")

    x_side = x_team + w_team / 2
    y_side = (act_top + b_proto) / 2
    line(ax, x_side, b_team, x_side, y_side)
    line(ax, x_side, y_side, x_task + w_task, y_side)
    txt(ax, x_side - 0.15, b_team - 0.08, "1", size=11, ha="right", va="top")
    txt(ax, x_task + w_task + 0.12, y_side + 0.12, "0..*", size=11, ha="left", va="bottom")
    txt(ax, x_side - 0.18, (b_team + y_side) / 2, "составляет", size=11, style="italic", ha="right")
    return save(fig, "fig02_classes.png")


def head_box(ax, cx, top, text, bw=4.6, bh=0.85):
    ax.add_patch(Rectangle((cx - bw / 2, top - bh), bw, bh, fill=True, facecolor="white",
                           edgecolor="black", lw=1.1, zorder=4))
    txt(ax, cx, top - bh / 2, text, size=12, weight="bold")
    return top - bh


def draw_sequence():
    w, h = 24.0, 13.6
    fig, ax = cm_fig(w, h)
    cols = [3.7, 12.15, 20.35]
    names = ["Руководитель команды", "СпринтКонтур", "Разработчик"]
    widths = [6.5, 4.9, 4.5]
    top = 13.05
    bottoms = []
    for cx, name, bw in zip(cols, names, widths):
        bottoms.append(head_box(ax, cx, top, name, bw=bw))
    life_top = min(bottoms)
    life_bot = 0.35
    for cx in cols:
        line(ax, cx, life_top, cx, life_bot, lw=0.9, dashed=True)

    labels = []
    y = life_top - 0.95

    def call(x1, x2, y_pos, text, ret=False):
        if x2 >= x1:
            arrow(ax, x1, y_pos, x2 - 0.06, y_pos, dashed=ret, open_head=ret)
        else:
            arrow(ax, x1, y_pos, x2 + 0.06, y_pos, dashed=ret, open_head=ret)
        left = min(x1, x2) + 0.42
        right = max(x1, x2) - 0.42
        lines = text.split("\n")
        line_h = 0.38
        for i, part in enumerate(reversed(lines)):
            artist = txt(ax, (left + right) / 2, y_pos + 0.12 + i * line_h, part, size=10.5, va="bottom")
            labels.append((artist, left, right))

    def self_msg(x, y_pos, text, next_x):
        gap = 0.55
        tw = text_width(fig, ax, text, 10.5)
        reach = gap + tw + 0.45
        if x + reach > next_x - 0.5:
            raise RuntimeError(f"Самосообщение не помещается справа от линии: {text}")
        drop = 0.72
        line(ax, x, y_pos, x + reach, y_pos)
        line(ax, x + reach, y_pos, x + reach, y_pos - drop)
        arrow(ax, x + reach, y_pos - drop, x + 0.06, y_pos - drop)
        artist = txt(ax, x + gap, y_pos + 0.12, text, size=10.5, ha="left", va="bottom")
        labels.append((artist, x + 0.28, next_x - 0.35))
        return y_pos - drop

    call(cols[0], cols[1], y, "1: зарегистрироватьСрочную\n(модуль)")
    y -= 1.22
    y = self_msg(cols[1], y, "2: состояние := «Новая»", cols[2])
    y -= 0.95
    call(cols[1], cols[0], y, "3: номер и срок, «Новая»", ret=True)
    y -= 1.05
    call(cols[0], cols[1], y, "4: назначитьРазработчика\n(номер)")
    y -= 1.22
    y = self_msg(cols[1], y, "5: проверить срок сдачи", cols[2])
    y -= 0.95
    call(cols[1], cols[2], y, "6: задача, «Назначена»")
    y -= 1.05
    call(cols[2], cols[1], y, "7: принятьЗадачу(номер)")
    y -= 1.05
    y = self_msg(cols[1], y, "8: состояние := «В работе»", cols[2])
    y -= 0.95
    call(cols[1], cols[0], y, "9: состояние «В работе»", ret=True)

    fig.canvas.draw()
    for artist, left, right in labels:
        x0, x1, _, _ = data_span(fig, ax, artist)
        raw = artist.get_text().replace("\n", " ")
        if x0 < left - 0.02 or x1 > right + 0.02:
            raise RuntimeError(f"Подпись пересекает линию жизни: {raw}")
        for cx in cols:
            if x0 - 0.05 < cx < x1 + 0.05:
                raise RuntimeError(f"Подпись пересекает пунктир: {raw}")
    return save(fig, "fig03_sequence.png")


def state_box(ax, cx, cy, text, bw=2.85, bh=1.15):
    ax.add_patch(FancyBboxPatch(
        (cx - bw / 2, cy - bh / 2), bw, bh,
        boxstyle="round,pad=0.02,rounding_size=0.35",
        facecolor="white", edgecolor="black", lw=1.15, zorder=3,
    ))
    txt(ax, cx, cy, text, size=13, weight="bold")
    return cx - bw / 2, cx + bw / 2, cy - bh / 2, cy + bh / 2


def initial(ax, x, y, r=0.22):
    ax.add_patch(plt.Circle((x, y), r, facecolor="black", edgecolor="black", zorder=4))


def final(ax, x, y, r=0.26):
    ax.add_patch(plt.Circle((x, y), r, fill=False, lw=1.15, color="black", zorder=4))
    ax.add_patch(plt.Circle((x, y), r * 0.55, facecolor="black", edgecolor="black", zorder=4))


def draw_states():
    w, h = 24.0, 10.2
    fig, ax = cm_fig(w, h)
    y1, y2 = 6.55, 2.25
    xs = [5.15, 10.55, 15.95]
    initial(ax, 0.55, y1)
    labels = ["Новая", "Назначена", "В работе"]
    boxes = []
    for x, name in zip(xs, labels):
        boxes.append(state_box(ax, x, y1, name))
    checked = state_box(ax, xs[2], y2, "Проверена")
    closed = state_box(ax, 21.05, y2, "Закрыта", bw=2.85)
    final(ax, 23.45, y2)

    def hlabel(x1, x2, y, line1, line2=None):
        arrow(ax, x1, y, x2, y)
        if line2:
            txt(ax, (x1 + x2) / 2, y + 0.62, line1, size=10.5)
            txt(ax, (x1 + x2) / 2, y + 0.22, line2, size=10.5)
        else:
            txt(ax, (x1 + x2) / 2, y + 0.28, line1, size=10.5, va="bottom")

    hlabel(0.82, boxes[0][0], y1, "зарегистрировать", "срочную")
    hlabel(boxes[0][1], boxes[1][0], y1, "назначить", "разработчика")
    hlabel(boxes[1][1], boxes[2][0], y1, "принять", "задачу")

    top = boxes[2][3]
    enter_x = xs[2] - 1.05
    leave_x = xs[2] + 1.05
    loop_y = top + 0.82
    line(ax, leave_x, top, leave_x, loop_y)
    line(ax, leave_x, loop_y, enter_x, loop_y)
    arrow(ax, enter_x, loop_y, enter_x, top + 0.02)
    txt(ax, xs[2], loop_y + 0.46, "проверка", size=10.5, va="bottom")
    txt(ax, xs[2], loop_y + 0.08, "не пройдена", size=10.5, va="bottom")

    x = xs[2]
    arrow(ax, x, boxes[2][2], x, checked[3])
    mid = (boxes[2][2] + checked[3]) / 2
    txt(ax, x + 0.28, mid + 0.2, "проверка", size=10.5, ha="left")
    txt(ax, x + 0.28, mid - 0.22, "пройдена", size=10.5, ha="left")

    hlabel(checked[1], closed[0], y2, "закрыть", "задачу")
    arrow(ax, closed[1], y2, 23.15, y2)
    return save(fig, "fig04_states.png")


def action_box(ax, cx, cy, text, bw=6.7, bh=0.92):
    ax.add_patch(FancyBboxPatch(
        (cx - bw / 2, cy - bh / 2), bw, bh,
        boxstyle="round,pad=0.02,rounding_size=0.28",
        facecolor="white", edgecolor="black", lw=1.1, zorder=3,
    ))
    txt(ax, cx, cy, text, size=12)
    return cy - bh / 2, cy + bh / 2


def draw_activity():
    w, h = 24.0, 14.6
    fig, ax = cm_fig(w, h)
    lanes = [0.0, 8.0, 16.0]
    lane_w = 8.0
    ax.add_patch(Rectangle((0, 0.25), w, h - 0.5, fill=False, lw=1.15, zorder=1))
    for x in (8.0, 16.0):
        line(ax, x, 0.25, x, h - 0.25, lw=1.05)
    line(ax, 0, h - 1.35, w, h - 1.35, lw=1.05)
    for x, name in zip(lanes, ["Разработчик", "СпринтКонтур", "Руководитель команды"]):
        txt(ax, x + lane_w / 2, h - 0.8, name, size=13, weight="bold")

    centers = [lane + lane_w / 2 for lane in lanes]
    steps = [
        (0, "start", ""),
        (0, "act", "Зафиксировать результат\nи замечание"),
        (0, "act", "Оформить протокол\nпроверки"),
        (1, "act", "Перевести в «Проверена»"),
        (2, "act", "Сверить протокол\nсо сроком сдачи"),
        (2, "act", "Закрыть задачу"),
        (1, "act", "Перевести в «Закрыта»"),
        (1, "end", ""),
    ]
    top = h - 1.55
    bot = 0.45
    slot = (top - bot) / len(steps)
    nodes = []
    for i, (lane, kind, text) in enumerate(steps):
        cy = top - slot * (i + 0.5)
        cx = centers[lane]
        if kind == "start":
            txt(ax, cx, cy + 0.42, "состояние «В работе»", size=12, va="bottom")
            initial(ax, cx, cy - 0.08, r=0.22)
            nodes.append((cx, cy - 0.30, cy + 0.14, "start"))
        elif kind == "end":
            final(ax, cx, cy, r=0.28)
            nodes.append((cx, cy - 0.28, cy + 0.28, "end"))
        else:
            yb, yt = action_box(ax, cx, cy, text)
            nodes.append((cx, yb, yt, "act"))

    for a, b in zip(nodes, nodes[1:]):
        x1, yb1, yt1, _ = a
        x2, yb2, yt2, _ = b
        if abs(x1 - x2) < 0.2:
            arrow(ax, x1, yb1, x2, yt2)
        else:
            gap = yb1 - yt2
            if gap < 0.36:
                raise RuntimeError("Мало места для стрелки между дорожками")
            mid = yt2 + gap / 2
            line(ax, x1, yb1, x1, mid)
            line(ax, x1, mid, x2, mid)
            arrow(ax, x2, mid, x2, yt2)
    return save(fig, "fig05_activity.png")


def main():
    paths = [draw_usecase(), draw_classes(), draw_sequence(), draw_states(), draw_activity()]
    for p in paths:
        print(p.name, p.stat().st_size)


if __name__ == "__main__":
    main()
