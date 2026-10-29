# -*- coding: utf-8 -*-
"""Модель данных АИС ЖД по сущностям ЛР6–ЛР8. В текст отчёта не входит."""
from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

OUT = Path(__file__).resolve().parent / "figures" / "er_model.png"
FONT = "Times New Roman"


def entity(ax, x, y, w, h, title, attrs):
    ax.add_patch(
        FancyBboxPatch(
            (x, y),
            w,
            h,
            boxstyle="round,pad=0.02,rounding_size=0.08",
            facecolor="#ffffff",
            edgecolor="black",
            linewidth=1.2,
            zorder=3,
        )
    )
    head = y + h - 0.55
    ax.plot([x + 0.1, x + w - 0.1], [head, head], color="black", lw=1.0, zorder=4)
    ax.text(
        x + w / 2,
        y + h - 0.28,
        title,
        ha="center",
        va="center",
        fontsize=11,
        fontname=FONT,
        fontweight="bold",
        zorder=5,
    )
    for i, attr in enumerate(attrs):
        ax.text(
            x + 0.18,
            head - 0.38 - i * 0.42,
            attr,
            ha="left",
            va="center",
            fontsize=9.5,
            fontname=FONT,
            zorder=5,
        )


def line(ax, xs, ys):
    ax.plot(xs, ys, color="black", lw=1.1, zorder=1, solid_capstyle="round")


def tag(ax, x, y, text):
    ax.text(
        x,
        y,
        text,
        ha="center",
        va="center",
        fontsize=9,
        fontname=FONT,
        style="italic",
        zorder=6,
        bbox=dict(boxstyle="round,pad=0.18", facecolor="white", edgecolor="none", alpha=0.97),
    )


def draw() -> Path:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    # Высокий холст: три яруса с большими промежутками под подписи связей.
    fig, ax = plt.subplots(figsize=(12.5, 14.5), dpi=170)
    ax.set_xlim(0, 32)
    ax.set_ylim(0, 38)
    ax.set_aspect("equal")
    ax.axis("off")
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    ax.text(
        16,
        36.8,
        "Модель данных АИС железной дороги",
        ha="center",
        va="center",
        fontsize=14,
        fontname=FONT,
        fontweight="bold",
    )

    # --- Ярус 1: справочники ---
    y1 = 30.2
    h1 = 4.6
    entity(ax, 0.8, y1, 6.0, h1, "Станция", ["код (PK)", "название"])
    entity(ax, 8.4, y1, 6.2, h1, "Схема вагона", ["тип (PK)", "номера мест"])
    entity(ax, 16.2, y1, 6.6, h1, "Маршрут", ["код (PK)", "название", "время хода, мин"])
    entity(ax, 24.4, y1, 6.6, h1, "Состав маршрута", ["маршрут (FK)", "станция (FK)", "порядок"])

    # станция → состав маршрута (вниз под блоками)
    line(ax, [3.8, 3.8, 27.7, 27.7], [y1, y1 - 1.1, y1 - 1.1, y1])
    tag(ax, 15.8, y1 - 1.1, "1   станция входит в маршрут   *")

    # маршрут → состав
    line(ax, [22.8, 24.4], [y1 + h1 / 2, y1 + h1 / 2])
    tag(ax, 23.6, y1 + h1 / 2 + 0.55, "1 — *")

    # --- Ярус 2: рейс / фонд / билет ---
    y2 = 17.4
    entity(ax, 2.5, y2, 7.4, 6.2, "Рейс", ["номер (PK)", "дата", "маршрут (FK)", "схема (FK)", "время", "открыт / отменён"])
    entity(ax, 12.2, y2 + 0.8, 6.4, 4.8, "Место (фонд)", ["рейс (FK)", "номер места", "свободно"])
    entity(ax, 21.0, y2, 7.6, 6.4, "Билет", ["номер (PK)", "рейс (FK)", "место", "пассажир", "цена", "посадка", "возврат"])

    # схема → рейс
    line(ax, [11.5, 11.5, 6.2, 6.2], [y1, 28.6, 28.6, y2 + 6.2])
    tag(ax, 9.0, 28.6, "схема  1 — *")

    # маршрут → рейс
    line(ax, [19.5, 19.5, 8.5, 8.5], [y1, 27.5, 27.5, y2 + 6.2])
    tag(ax, 14.0, 27.5, "маршрут  1 — *")

    # рейс → место
    line(ax, [9.9, 12.2], [y2 + 3.2, y2 + 3.2])
    tag(ax, 11.05, y2 + 3.75, "1 — *  фонд")

    # рейс → билет
    line(ax, [9.9, 9.9, 21.0], [y2 + 1.8, y2 - 0.9, y2 - 0.9])
    line(ax, [21.0, 21.0], [y2 - 0.9, y2])
    tag(ax, 15.4, y2 - 0.9, "продажа  1 — *")

    # место → билет
    line(ax, [18.6, 21.0], [y2 + 3.2, y2 + 3.2])
    tag(ax, 19.8, y2 + 3.75, "0..1 — 1")

    # --- Ярус 3: смена / обслуживание / ремонт / инвентаризация ---
    y3 = 3.8
    entity(ax, 0.8, y3, 6.4, 5.6, "Смена", ["роль", "сотрудник", "открыта", "черновик продажи"])
    entity(ax, 8.6, y3, 7.2, 5.8, "Обслуживание станции", ["сутки открыты", "кассы", "платформы", "зал ожидания", "проходы"])
    entity(ax, 17.2, y3 + 0.4, 6.4, 5.0, "Заявка на ремонт", ["объект", "описание", "закрыта"])
    entity(ax, 25.0, y3 + 0.4, 6.0, 5.0, "Инвентаризация", ["дата", "объект", "факт"])

    # смена ↔ обслуживание
    line(ax, [7.2, 8.6], [y3 + 2.9, y3 + 2.9])
    tag(ax, 7.9, y3 + 3.5, "разрешает")

    # обслуживание → ремонт
    line(ax, [15.8, 17.2], [y3 + 2.9, y3 + 2.9])
    tag(ax, 16.5, y3 + 3.5, "порождает")

    # обслуживание → инвентаризация (объект станции)
    line(ax, [15.8, 15.8, 25.0], [y3 + 1.4, y3 - 0.85, y3 - 0.85])
    line(ax, [25.0, 25.0], [y3 - 0.85, y3 + 0.4])
    tag(ax, 20.4, y3 - 0.85, "объект станции")

    # смена → билетные операции (внизу)
    line(ax, [4.0, 4.0, 24.8, 24.8], [y3, 1.1, 1.1, y2])
    tag(ax, 14.4, 1.1, "продажа / возврат / посадка только при открытой смене")

    fig.savefig(OUT, dpi=170, bbox_inches="tight", facecolor="white", pad_inches=0.25)
    plt.close(fig)
    return OUT


if __name__ == "__main__":
    print(draw())
