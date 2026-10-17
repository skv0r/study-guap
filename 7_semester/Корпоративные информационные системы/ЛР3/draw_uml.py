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
    txt(ax, x, y_hip - 1.05, name, size=11.5, va="top")
    return x + side * 0.78, y_neck - 0.32


def use_case(ax, cx, cy, rx, ry, label):
    ax.add_patch(Ellipse((cx, cy), rx * 2, ry * 2, fill=True, facecolor="white",
                         edgecolor="black", lw=1.15, zorder=3))
    txt(ax, cx, cy, label, size=11)


def draw_usecase():
    w, h = 24.0, 13.0
    fig, ax = cm_fig(w, h)
    ax.add_patch(Rectangle((4.0, 0.45), 14.2, 12.15, fill=False, lw=1.25, zorder=1))
    txt(ax, 11.1, 12.15, "СпринтКонтур / СервисЗадач", size=13, weight="bold")

    cases = {
        "reg": (7.7, 10.35, "Зарегистрировать\nсрочную задачу"),
        "assign": (7.7, 7.85, "Назначить\nразработчика"),
        "persist": (7.7, 5.35, "Записать состояние\nв БазаДанных"),
        "close": (7.7, 2.85, "Закрыть задачу"),
        "tell": (14.5, 10.35, "Сообщить\nтребование"),
        "status": (14.5, 7.85, "Получить\nстатус задачи"),
        "accept": (14.5, 5.35, "Принять задачу"),
        "protocol": (14.5, 2.85, "Оформить протокол\nпроверки"),
    }
    rx, ry = 2.45, 0.78
    for cx, cy, label in cases.values():
        use_case(ax, cx, cy, rx, ry, label)

    lead = actor(ax, 1.85, 9.05, "Руководитель\nкоманды", side=1)
    customer = actor(ax, 21.85, 9.55, "Заказчик", side=-1)
    developer = actor(ax, 21.85, 5.15, "Разработчик", side=-1)
    db = actor(ax, 1.85, 3.55, "БазаДанных", side=1)

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
    link(db, "persist", True)

    # include: назначение → запись в БД
    x = cases["assign"][0]
    y_from = cases["assign"][1] - ry
    y_to = cases["persist"][1] + ry
    arrow(ax, x, y_from - 0.02, x, y_to + 0.02, dashed=True, open_head=True, lw=1.05)
    txt(ax, x + 1.05, (y_from + y_to) / 2, "«include»", size=10.5, style="italic")

    # include: закрытие → запись в БД (стрелка к включаемому)
    x2 = cases["close"][0] + 0.55
    y_close = cases["close"][1] + ry
    y_persist = cases["persist"][1] - ry
    arrow(ax, x2, y_close + 0.02, x2, y_persist - 0.02, dashed=True, open_head=True, lw=1.05)
    txt(ax, x2 + 1.0, (y_close + y_persist) / 2, "«include»", size=10.5, style="italic")

    # secondary: сервис обращается к БД при статусе (dashed association from status to persist conceptually via DB actor already)
    return save(fig, "fig01_usecase.png")


def class_box(ax, x, top, w, name, attrs, ops, stereo=None):
    head = 0.72 if stereo is None else 0.95
    row = 0.42
    attr_h = 0.16 + len(attrs) * row + 0.08 if attrs else 0.18
    op_h = (0.14 + len(ops) * row + 0.10) if ops else 0.0
    h = head + attr_h + op_h
    bottom = top - h
    ax.add_patch(Rectangle((x, bottom), w, h, fill=True, facecolor="white", edgecolor="black", lw=1.15, zorder=3))
    sep1 = top - head
    ax.plot([x, x + w], [sep1, sep1], color="black", lw=1.0, zorder=4)
    if stereo:
        txt(ax, x + w / 2, top - 0.28, stereo, size=10, style="italic")
        txt(ax, x + w / 2, top - 0.68, name, size=12, weight="bold")
    else:
        txt(ax, x + w / 2, top - head / 2, name, size=12.5, weight="bold")
    for i, attr in enumerate(attrs):
        txt(ax, x + 0.18, sep1 - 0.10 - i * row, attr, size=10.5, ha="left", va="top")
    if ops:
        sep2 = bottom + op_h
        ax.plot([x, x + w], [sep2, sep2], color="black", lw=1.0, zorder=4)
        for i, op in enumerate(ops):
            txt(ax, x + 0.18, sep2 - 0.08 - i * row, "+ " + op, size=10.5, ha="left", va="top")
    return bottom


def draw_classes():
    w, h = 26.0, 17.5
    fig, ax = cm_fig(w, h)

    # слой приложения (верх)
    txt(ax, 0.4, 17.15, "слой приложения и хранилища", size=10.5, ha="left", style="italic", color="#333333")
    top_app = 16.75
    b_svc = class_box(
        ax, 0.5, top_app, 11.0, "СервисЗадач",
        [],
        ["зарегистрироватьСрочную()", "назначитьРазработчика()",
         "принятьЗадачу()", "закрытьПослеПроверки()"],
        stereo="«приложение»",
    )
    b_db = class_box(
        ax, 14.5, top_app, 11.0, "БазаДанных",
        ["таблицы: задачи, модули,", "назначения, протоколы"],
        ["сохранитьЗадачу()", "обновитьСостояние()",
         "сохранитьНазначение()", "сохранитьПротокол()"],
        stereo="«хранилище»",
    )

    # связь сервис → БД (между блоками, не по центру текста)
    y_app = (top_app + min(b_svc, b_db)) / 2
    arrow(ax, 11.5, y_app, 14.5, y_app, open_head=True)
    txt(ax, 13.0, y_app + 0.28, "использует", size=11, style="italic", va="bottom")

    # промежуток между слоями
    gap = 1.35
    top_d = min(b_svc, b_db) - gap
    txt(ax, 0.4, top_d + 0.55, "доменный слой", size=10.5, ha="left", style="italic", color="#333333")

    # домен: три класса в ряд, протокол ниже задачи
    x_mod, w_mod = 0.5, 5.4
    x_task, w_task = 7.4, 6.6
    x_team, w_team = 16.0, 5.4

    b_mod = class_box(ax, x_mod, top_d, w_mod, "Модуль",
                      ["код", "название", "версия"], [])
    b_task = class_box(
        ax, x_task, top_d, w_task, "Задача",
        ["номер", "признакСрочности", "состояние", "срокСдачи"],
        ["изменитьСостояние()"],
    )
    b_team = class_box(
        ax, x_team, top_d, w_team, "Команда",
        ["код", "руководитель"],
        ["естьСвободный()"],
    )

    w_proto = 6.2
    x_proto = x_task + (w_task - w_proto) / 2
    top_proto = min(b_mod, b_task, b_team) - 1.65
    b_proto = class_box(
        ax, x_proto, top_proto, w_proto, "ПротоколПроверки",
        ["номер", "результат", "дата"],
        ["оформить()"],
    )

    # Сервис управляет Задачей (вертикаль слева от задачи, без пересечения блоков)
    x_ctrl = x_task + 0.55
    line(ax, x_ctrl, b_svc, x_ctrl, top_d)
    arrow(ax, x_ctrl, top_d + 0.02, x_ctrl, top_d, open_head=True)
    txt(ax, x_ctrl + 0.22, (b_svc + top_d) / 2, "управляет", size=10.5, style="italic", ha="left")

    # БД хранит домен (справа, к Команде)
    x_store = x_team + w_team / 2
    arrow(ax, x_store, b_db, x_store, top_d, open_head=True, dashed=True)
    txt(ax, x_store + 0.25, (b_db + top_d) / 2, "хранит", size=10.5, style="italic", ha="left")

    # горизонтальные связи домена — ниже нижнего края всех трёх коробок
    y_link = min(b_mod, b_task, b_team) - 0.45
    line(ax, x_mod + w_mod, y_link, x_task, y_link)
    txt(ax, x_mod + w_mod + 0.15, y_link - 0.22, "1", size=11, ha="left", va="top")
    txt(ax, x_task - 0.15, y_link - 0.22, "0..*", size=11, ha="right", va="top")
    txt(ax, (x_mod + w_mod + x_task) / 2, y_link + 0.18, "относится к", size=10.5, style="italic", va="bottom")

    line(ax, x_task + w_task, y_link, x_team, y_link)
    txt(ax, x_task + w_task + 0.15, y_link - 0.22, "0..*", size=11, ha="left", va="top")
    txt(ax, x_team - 0.15, y_link - 0.22, "0..1", size=11, ha="right", va="top")
    txt(ax, (x_task + w_task + x_team) / 2, y_link + 0.18, "назначена", size=10.5, style="italic", va="bottom")

    # Задача — Протокол (вертикаль вниз по центру)
    x_tp = x_task + w_task / 2
    line(ax, x_tp, b_task, x_tp, top_proto)
    txt(ax, x_tp - 0.25, b_task - 0.15, "1", size=11, ha="right", va="top")
    txt(ax, x_tp + 0.25, top_proto + 0.18, "0..1", size=11, ha="left", va="bottom")
    txt(ax, x_tp + 0.45, (b_task + top_proto) / 2, "имеет", size=10.5, style="italic", ha="left")

    return save(fig, "fig02_classes.png")


def head_box(ax, cx, top, text, bw=4.6, bh=0.85):
    ax.add_patch(Rectangle((cx - bw / 2, top - bh), bw, bh, fill=True, facecolor="white",
                           edgecolor="black", lw=1.1, zorder=4))
    txt(ax, cx, top - bh / 2, text, size=11.5, weight="bold")
    return top - bh


def draw_sequence():
    w, h = 24.0, 15.2
    fig, ax = cm_fig(w, h)
    cols = [2.55, 8.35, 14.35, 20.55]
    names = ["Руководитель\nкоманды", "СервисЗадач", "БазаДанных", "Разработчик"]
    widths = [4.4, 4.2, 4.0, 3.8]
    top = 14.75
    bottoms = []
    for cx, name, bw in zip(cols, names, widths):
        bottoms.append(head_box(ax, cx, top, name, bw=bw, bh=1.05))
    life_top = min(bottoms)
    life_bot = 0.30
    for cx in cols:
        line(ax, cx, life_top, cx, life_bot, lw=0.9, dashed=True)

    labels = []
    y = life_top - 0.75

    def call(x1, x2, y_pos, text, ret=False):
        if x2 >= x1:
            arrow(ax, x1, y_pos, x2 - 0.06, y_pos, dashed=ret, open_head=ret)
        else:
            arrow(ax, x1, y_pos, x2 + 0.06, y_pos, dashed=ret, open_head=ret)
        # подпись только в первом сегменте от источника, чтобы не пересекать чужие линии жизни
        lo, hi = min(x1, x2), max(x1, x2)
        mids = sorted(c for c in cols if lo + 0.15 < c < hi - 0.15)
        if x1 < x2:
            left = x1 + 0.38
            right = (mids[0] if mids else x2) - 0.38
        else:
            right = x1 - 0.38
            left = (mids[-1] if mids else x2) + 0.38
        lines = text.split("\n")
        line_h = 0.34
        for i, part in enumerate(reversed(lines)):
            artist = txt(ax, (left + right) / 2, y_pos + 0.10 + i * line_h, part, size=9.8, va="bottom")
            labels.append((artist, left, right))

    def self_msg(x, y_pos, text, next_x):
        gap = 0.42
        tw = text_width(fig, ax, text, 9.8)
        reach = gap + tw + 0.38
        if x + reach > next_x - 0.45:
            raise RuntimeError(f"Самосообщение не помещается справа от линии: {text}")
        drop = 0.62
        line(ax, x, y_pos, x + reach, y_pos)
        line(ax, x + reach, y_pos, x + reach, y_pos - drop)
        arrow(ax, x + reach, y_pos - drop, x + 0.06, y_pos - drop)
        artist = txt(ax, x + gap, y_pos + 0.10, text, size=9.8, ha="left", va="bottom")
        labels.append((artist, x + 0.22, next_x - 0.30))
        return y_pos - drop

    # 1 руководитель → сервис
    call(cols[0], cols[1], y, "1: зарегистрироватьСрочную\n(модуль)")
    y -= 1.05
    # 2 сервис → БД
    call(cols[1], cols[2], y, "2: сохранитьЗадачу\n(«Новая»)")
    y -= 1.00
    # 3 БД → сервис
    call(cols[2], cols[1], y, "3: номер, срок", ret=True)
    y -= 0.90
    # 4 сервис → руководитель
    call(cols[1], cols[0], y, "4: номер и срок, «Новая»", ret=True)
    y -= 0.95
    # 5 руководитель → сервис
    call(cols[0], cols[1], y, "5: назначитьРазработчика\n(номер)")
    y -= 1.05
    # 6 самосообщение сервиса
    y = self_msg(cols[1], y, "6: проверить срок сдачи", cols[2])
    y -= 0.78
    # 7 сервис → БД
    call(cols[1], cols[2], y, "7: обновитьСостояние\n(«Назначена»)")
    y -= 1.00
    # 8 сервис → разработчик
    call(cols[1], cols[3], y, "8: задача, «Назначена»")
    y -= 0.95
    # 9 разработчик → сервис
    call(cols[3], cols[1], y, "9: принятьЗадачу(номер)")
    y -= 0.95
    # 10 сервис → БД
    call(cols[1], cols[2], y, "10: обновитьСостояние\n(«В работе»)")
    y -= 1.00
    # 11 сервис → руководитель
    call(cols[1], cols[0], y, "11: состояние «В работе»", ret=True)

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
    w, h = 24.0, 11.0
    fig, ax = cm_fig(w, h)
    txt(ax, 12.0, 10.45, "Состояние задачи хранится в БазаДанных; переходы выполняет СервисЗадач",
        size=11, style="italic")
    y1, y2 = 6.85, 2.45
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


def action_box(ax, cx, cy, text, bw=5.4, bh=0.88):
    ax.add_patch(FancyBboxPatch(
        (cx - bw / 2, cy - bh / 2), bw, bh,
        boxstyle="round,pad=0.02,rounding_size=0.28",
        facecolor="white", edgecolor="black", lw=1.1, zorder=3,
    ))
    txt(ax, cx, cy, text, size=11)
    return cy - bh / 2, cy + bh / 2


def draw_activity():
    w, h = 24.0, 15.4
    fig, ax = cm_fig(w, h)
    lanes = [0.0, 6.0, 12.0, 18.0]
    lane_w = 6.0
    ax.add_patch(Rectangle((0, 0.25), w, h - 0.5, fill=False, lw=1.15, zorder=1))
    for x in (6.0, 12.0, 18.0):
        line(ax, x, 0.25, x, h - 0.25, lw=1.05)
    line(ax, 0, h - 1.35, w, h - 1.35, lw=1.05)
    names = ["Разработчик", "СервисЗадач", "БазаДанных", "Руководитель\nкоманды"]
    for x, name in zip(lanes, names):
        txt(ax, x + lane_w / 2, h - 0.75, name, size=12, weight="bold")

    centers = [lane + lane_w / 2 for lane in lanes]
    # lane indices: 0 dev, 1 svc, 2 db, 3 lead
    steps = [
        (0, "start", ""),
        (0, "act", "Зафиксировать результат\nи замечание"),
        (0, "act", "Передать протокол\nв СервисЗадач"),
        (1, "act", "Проверить результат\nпротокола"),
        (2, "act", "Сохранить протокол"),
        (1, "act", "Запросить обновление\nсостояния"),
        (2, "act", "Записать «Проверена»"),
        (3, "act", "Сверить протокол\nсо сроком сдачи"),
        (3, "act", "Закрыть задачу"),
        (1, "act", "Запросить закрытие"),
        (2, "act", "Записать «Закрыта»"),
        (1, "end", ""),
    ]
    top = h - 1.55
    bot = 0.40
    slot = (top - bot) / len(steps)
    nodes = []
    for i, (lane, kind, text) in enumerate(steps):
        cy = top - slot * (i + 0.5)
        cx = centers[lane]
        if kind == "start":
            txt(ax, cx, cy + 0.38, "состояние «В работе»", size=11, va="bottom")
            initial(ax, cx, cy - 0.08, r=0.20)
            nodes.append((cx, cy - 0.28, cy + 0.12, "start"))
        elif kind == "end":
            final(ax, cx, cy, r=0.26)
            nodes.append((cx, cy - 0.26, cy + 0.26, "end"))
        else:
            yb, yt = action_box(ax, cx, cy, text, bw=5.35, bh=0.82)
            nodes.append((cx, yb, yt, "act"))

    for a, b in zip(nodes, nodes[1:]):
        x1, yb1, yt1, _ = a
        x2, yb2, yt2, _ = b
        if abs(x1 - x2) < 0.2:
            arrow(ax, x1, yb1, x2, yt2)
        else:
            gap = yb1 - yt2
            if gap < 0.28:
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
