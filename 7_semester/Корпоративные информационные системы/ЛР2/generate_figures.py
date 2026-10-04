# -*- coding: utf-8 -*-
"""Схемы IDEF0 мастерской «Обод»."""

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, Rectangle

OUT = Path(__file__).resolve().parent / "figures"
FONT = "Times New Roman"


def _new(w, h):
    fig, ax = plt.subplots(figsize=(w / 25.4, h / 25.4), dpi=200)
    ax.set_xlim(0, w)
    ax.set_ylim(0, h)
    ax.set_aspect("equal")
    ax.axis("off")
    fig.subplots_adjust(0, 0, 1, 1)
    ax.add_patch(Rectangle((0.5, 0.5), w - 1, h - 1, fill=False, ec="black", lw=0.8))
    return fig, ax


def _save(fig, name):
    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / name
    fig.savefig(path, dpi=200, facecolor="white")
    plt.close(fig)
    return path


def T(ax, x, y, s, size=8, ha="center", va="center"):
    ax.text(
        x, y, s, ha=ha, va=va, fontsize=size, fontname=FONT, color="black",
        linespacing=1.12,
        bbox=dict(facecolor="white", edgecolor="none", pad=0.12),
        zorder=5, clip_on=False,
    )


def arrow(ax, p, q):
    ax.add_patch(FancyArrowPatch(
        p, q, arrowstyle="-|>", mutation_scale=8, lw=0.8, color="black",
        shrinkA=0, shrinkB=0, zorder=2, clip_on=False,
    ))


def poly(ax, pts):
    for a, b in zip(pts, pts[1:-1]):
        ax.plot([a[0], b[0]], [a[1], b[1]], color="black", lw=0.8, zorder=2, solid_capstyle="projecting")
    arrow(ax, pts[-2], pts[-1])


def box(ax, x, y, w, h, title, num, child=None):
    ax.add_patch(Rectangle((x, y), w, h, fill=False, ec="black", lw=1.2, zorder=3))
    T(ax, x + w / 2, y + h * 0.58, title, size=8)
    T(ax, x + w - 1.4, y + 1.2, str(num), size=7.5, ha="right", va="bottom")
    if child:
        T(ax, x + w - 0.2, y - 0.3, child, size=7.5, ha="right", va="top")


def harr(ax, x1, x2, y, text, place="above", size=7):
    arrow(ax, (x1, y), (x2, y))
    xm = (x1 + x2) / 2
    if place == "above":
        T(ax, xm, y + 0.45, text, size=size, va="bottom")
    else:
        T(ax, xm, y - 0.45, text, size=size, va="top")


def varr(ax, x, y1, y2, text=None, size=7):
    if y2 > y1:
        arrow(ax, (x, y1), (x, y2))
        if text:
            T(ax, x, y1 - 0.4, text, size=size, va="top")
    else:
        arrow(ax, (x, y1), (x, y2))
        if text:
            T(ax, x, y1 + 0.4, text, size=size, va="bottom")


# Геометрия ряда из четырёх блоков
BW, BH = 36, 34
XS = [30, 84, 138, 192]  # зазор 18 мм; правый край последнего = 228
YB = 52
TOP = YB + BH


def row_fig(title, names, children, filename, draw):
    W, H = 250, 154
    fig, ax = _new(W, H)
    T(ax, W / 2, H - 6, title, size=11)
    for i, (name, child) in enumerate(zip(names, children)):
        box(ax, XS[i], YB, BW, BH, name, str(i + 1), child)
    draw(ax)
    return _save(fig, filename)


def controls(ax, spec):
    """spec: list of (box_index, x_frac, label). Стрелка сверху в блок."""
    for i, frac, label in spec:
        x = XS[i] + BW * frac
        varr(ax, x, 116, TOP)
        T(ax, x, 117.5, label, size=6.6, va="bottom")


def mechs(ax, spec):
    """spec: list of (box_index, x_frac, label). Стрелка снизу в блок."""
    for i, frac, label in spec:
        x = XS[i] + BW * frac
        varr(ax, x, 40, YB)
        T(ax, x, 38.2, label, size=7, va="top")


def fig_context():
    W, H = 250, 124
    fig, ax = _new(W, H)
    T(ax, W / 2, H - 6, "А-0    Ремонтировать велосипеды", size=12)
    bw, bh = 78, 30
    x, y = (W - bw) / 2, 48
    box(ax, x, y, bw, bh, "Ремонтировать\nвелосипеды", "0", "А0")

    for yy, name in ((y + 22, "Велосипед клиента"), (y + 15, "Заявка клиента"), (y + 8, "Комплектующие")):
        arrow(ax, (6, yy), (x, yy))
        T(ax, 8, yy + 0.5, name, size=8.5, ha="left", va="bottom")

    for xx, name in ((x + 18, "Прайс-лист услуг"), (x + 60, "Регламент мастерской")):
        arrow(ax, (xx, y + bh + 24), (xx, y + bh))
        T(ax, xx, y + bh + 25.2, name, size=8.5, va="bottom")

    for yy, name in ((y + 19, "Отремонтированный велосипед"), (y + 9, "Акт выполненных работ")):
        arrow(ax, (x + bw, yy), (W - 8, yy))
        T(ax, (x + bw + W - 8) / 2, yy + 0.6, name, size=8, va="bottom")

    for xx, name in ((x + 16, "Приёмщик"), (x + 39, "Мастер"), (x + 62, "Инструмент")):
        arrow(ax, (xx, 28), (xx, y))
        T(ax, xx, 26, name, size=8.5, va="top")

    T(ax, W / 2, 12,
      "Цель: описать выполнение одного заказа на ремонт велосипеда.\nТочка зрения: владелец мастерской.",
      size=9)
    return _save(fig, "fig01_a-0.png")


def fig_a0():
    def draw(ax):
        controls(ax, [
            (0, 0.28, "Прайс-лист\nуслуг"),
            (3, 0.28, "Прайс-лист\nуслуг"),
            (0, 0.72, "Регламент\nмастерской"),
            (1, 0.50, "Регламент\nмастерской"),
            (2, 0.50, "Регламент\nмастерской"),
            (3, 0.72, "Регламент\nмастерской"),
        ])
        mechs(ax, [
            (0, 0.35, "Приёмщик"),
            (1, 0.32, "Мастер"),
            (1, 0.72, "Инструмент"),
            (2, 0.32, "Мастер"),
            (2, 0.72, "Инструмент"),
            (3, 0.35, "Приёмщик"),
        ])
        # входы в блок 1
        harr(ax, 2, XS[0], YB + 24, "Велосипед\nклиента", "above", 6.6)
        harr(ax, 2, XS[0], YB + 12, "Заявка\nклиента", "below", 6.6)
        # комплектующие в блок 3, нижний коридор
        poly(ax, [(2, 16), (XS[2] - 3, 16), (XS[2] - 3, YB + 8), (XS[2], YB + 8)])
        T(ax, 4, 17.2, "Комплектующие", size=7, ha="left", va="bottom")

        harr(ax, XS[0] + BW, XS[1], YB + 24, "Принятый\nвелосипед", "above", 6.4)
        harr(ax, XS[0] + BW, XS[1], YB + 11, "Карточка\nзаказа", "below", 6.4)
        # второй выход «Карточка заказа» уходит сверху и входит в блок 4 как управление выдачей
        x_out = XS[0] + BW
        poly(ax, [
            (x_out, TOP - 3),
            (x_out + 5, TOP - 3),
            (x_out + 5, 136),
            (XS[3] + BW * 0.50, 136),
            (XS[3] + BW * 0.50, TOP),
        ])
        T(ax, (x_out + XS[3] + BW * 0.5) / 2 + 8, 137.2, "Карточка заказа", size=7, va="bottom")

        harr(ax, XS[1] + BW, XS[2], YB + 24, "Велосипед после\nдиагностики", "above", 6.2)
        harr(ax, XS[1] + BW, XS[2], YB + 11, "Дефектная\nведомость", "below", 6.4)

        harr(ax, XS[2] + BW, XS[3], YB + 24, "Велосипед\nпосле ремонта", "above", 6.2)
        harr(ax, XS[2] + BW, XS[3], YB + 11, "Перечень\nработ", "below", 6.4)

        harr(ax, XS[3] + BW, 247, YB + 24, "Отремонтированный\nвелосипед", "above", 6.2)
        harr(ax, XS[3] + BW, 247, YB + 12, "Акт выполненных\nработ", "below", 6.2)

    return row_fig(
        "А0    Ремонтировать велосипеды",
        ["Принять\nзаказ", "Диагностировать\nнеисправность", "Выполнить\nремонт", "Выдать\nвелосипед"],
        ["А1", "А2", "А3", None],
        "fig02_a0.png",
        draw,
    )


def fig_a1():
    def draw(ax):
        controls(ax, [
            (0, 0.5, "Регламент\nмастерской"),
            (1, 0.5, "Регламент\nмастерской"),
            (2, 0.5, "Прайс-лист\nуслуг"),
            (3, 0.5, "Регламент\nмастерской"),
        ])
        mechs(ax, [(i, 0.5, "Приёмщик") for i in range(4)])
        harr(ax, 2, XS[0], YB + 22, "Заявка\nклиента", "above", 6.6)
        # велосипед клиента в блоки 2 и 4 нижним коридором
        poly(ax, [(2, 22), (XS[1] - 3, 22), (XS[1] - 3, YB + 8), (XS[1], YB + 8)])
        poly(ax, [(XS[1] - 3, 22), (XS[3] - 3, 22), (XS[3] - 3, YB + 8), (XS[3], YB + 8)])
        T(ax, 4, 23.3, "Велосипед клиента", size=7, ha="left", va="bottom")

        harr(ax, XS[0] + BW, XS[1], YB + 22, "Черновик\nкарточки", "above", 6.4)
        harr(ax, XS[1] + BW, XS[2], YB + 24, "Описание\nсостояния", "above", 6.4)
        harr(ax, XS[1] + BW, XS[2], YB + 12, "Черновик\nкарточки", "below", 6.4)
        harr(ax, XS[2] + BW, XS[3], YB + 22, "Карточка\nзаказа", "above", 6.4)
        harr(ax, XS[3] + BW, 247, YB + 24, "Принятый\nвелосипед", "above", 6.4)
        harr(ax, XS[3] + BW, 247, YB + 12, "Карточка\nзаказа", "below", 6.4)

    return row_fig(
        "А1    Принять заказ",
        [
            "Зарегистрировать\nобращение",
            "Осмотреть\nвнешний вид",
            "Согласовать\nпредварительную\nстоимость",
            "Разместить велосипед\nна хранении",
        ],
        [None, None, None, None],
        "fig03_a1.png",
        draw,
    )


def fig_a2():
    def draw(ax):
        controls(ax, [(i, 0.5, "Регламент\nмастерской") for i in range(4)])
        mechs(ax, [(i, 0.32, "Мастер") for i in range(4)] + [(i, 0.72, "Инструмент") for i in range(3)])
        harr(ax, 2, XS[0], YB + 18, "Принятый\nвелосипед", "above", 6.6)
        # карточка заказа — вход блока 4, нижний коридор, без пересечения соседних стрелок
        poly(ax, [(2, 16), (XS[3] - 3, 16), (XS[3] - 3, YB + 8), (XS[3], YB + 8)])
        T(ax, 4, 17.2, "Карточка заказа", size=7, ha="left", va="bottom")

        harr(ax, XS[0] + BW, XS[1], YB + 24, "Велосипед\nпосле ходовой", "above", 6.2)
        harr(ax, XS[0] + BW, XS[1], YB + 12, "Замечания\nпо ходовой", "below", 6.2)
        harr(ax, XS[1] + BW, XS[2], YB + 24, "Велосипед\nпосле тормозов", "above", 6.2)
        harr(ax, XS[1] + BW, XS[2], YB + 12, "Замечания\nпо тормозам", "below", 6.2)
        harr(ax, XS[2] + BW, XS[3], YB + 24, "Велосипед после\nтрансмиссии", "above", 6.0)
        harr(ax, XS[2] + BW, XS[3], YB + 12, "Замечания по\nтрансмиссии", "below", 6.0)

        harr(ax, XS[3] + BW, 247, YB + 24, "Дефектная\nведомость", "above", 6.4)
        harr(ax, XS[3] + BW, 247, YB + 12, "Велосипед после\nдиагностики", "below", 6.0)

    return row_fig(
        "А2    Диагностировать неисправность",
        [
            "Проверить\nходовую часть",
            "Проверить\nтормоза",
            "Проверить\nтрансмиссию",
            "Составить\nдефектную\nведомость",
        ],
        [None, None, None, None],
        "fig04_a2.png",
        draw,
    )


def fig_a3():
    def draw(ax):
        controls(ax, [(i, 0.5, "Регламент\nмастерской") for i in range(4)])
        mechs(ax, [(i, 0.32, "Мастер") for i in range(4)] + [(i, 0.72, "Инструмент") for i in (1, 2)])
        harr(ax, 2, XS[0], YB + 24, "Комплектующие", "above", 6.6)
        harr(ax, 2, XS[0], YB + 12, "Дефектная\nведомость", "below", 6.4)
        # дефектная ведомость также вход блока 4
        poly(ax, [(10, YB + 12), (10, 22), (XS[3] - 3, 22), (XS[3] - 3, YB + 8), (XS[3], YB + 8)])
        # велосипед после диагностики — вход блока 2
        poly(ax, [(2, 14), (XS[1] - 3, 14), (XS[1] - 3, YB + 8), (XS[1], YB + 8)])
        T(ax, 4, 15.2, "Велосипед после диагностики", size=6.5, ha="left", va="bottom")

        harr(ax, XS[0] + BW, XS[1], YB + 24, "Подобранные\nдетали", "above", 6.4)
        harr(ax, XS[1] + BW, XS[2], YB + 24, "Велосипед\nс заменёнными\nузлами", "above", 6.0)
        # факт замены выходит справа из блока 2 и входит слева в блок 4
        poly(ax, [
            (XS[1] + BW, YB + 10),
            (XS[1] + BW + 4, YB + 10),
            (XS[1] + BW + 4, 28),
            (XS[3] - 4, 28),
            (XS[3] - 4, YB + 16),
            (XS[3], YB + 16),
        ])
        T(ax, (XS[1] + XS[3]) / 2, 29.2, "Факт замены", size=7, va="bottom")

        # велосипед после ремонта — выход блока 3, обходит блок 4 сверху
        poly(ax, [
            (XS[2] + BW, YB + 12),
            (XS[2] + BW + 8, YB + 12),
            (XS[2] + BW + 8, 136),
            (247, 136),
        ])
        T(ax, 188, 137.2, "Велосипед после ремонта", size=6.6, ha="left", va="bottom")
        harr(ax, XS[3] + BW, 247, YB + 22, "Перечень\nработ", "above", 6.4)

    return row_fig(
        "А3    Выполнить ремонт",
        [
            "Подобрать\nкомплектующие",
            "Заменить\nнеисправные\nдетали",
            "Отрегулировать\nузлы",
            "Записать\nвыполненные\nработы",
        ],
        [None, None, None, None],
        "fig05_a3.png",
        draw,
    )


def main():
    for p in (fig_context(), fig_a0(), fig_a1(), fig_a2(), fig_a3()):
        print(p)


if __name__ == "__main__":
    main()
