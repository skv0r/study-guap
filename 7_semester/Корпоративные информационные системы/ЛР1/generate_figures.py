"""Рисунки к ЛР1 КИС: Гант, сеть, загрузка ресурсов, затраты."""

from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Patch, Rectangle, RegularPolygon
from matplotlib.lines import Line2D

import project_plan as P

BASE = Path(__file__).resolve().parent
FIG = BASE / "figures"
FIG.mkdir(exist_ok=True)

plt.rcParams["font.family"] = "DejaVu Sans"
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["savefig.dpi"] = 170
plt.rcParams["figure.facecolor"] = "white"


def save(fig, name: str):
    path = FIG / name
    fig.savefig(path, bbox_inches="tight", facecolor="white", pad_inches=0.18)
    plt.close(fig)
    return path


def dnum(d: date) -> float:
    return mdates.date2num(d)


def shade_nonwork(ax, t0: date, t1: date, ymin=0, ymax=1):
    cur = t0
    while cur <= t1:
        if not P.is_workday(cur):
            ax.axvspan(dnum(cur), dnum(cur + timedelta(days=1)), color="#E8E8E8", zorder=0)
        cur += timedelta(days=1)


def fig01_calendar():
    months = [
        (2026, 10, "Октябрь 2026"),
        (2026, 11, "Ноябрь 2026"),
        (2026, 12, "Декабрь 2026"),
        (2027, 1, "Январь 2027"),
    ]
    fig, axes = plt.subplots(1, 4, figsize=(14.2, 3.9))
    weekdays = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
    t = P.TASKS
    p0, p1 = t["1.1"].es, t["4.4"].ef
    for ax, (y, m, title) in zip(axes, months):
        ax.set_xlim(0, 7)
        ax.set_ylim(0, 7)
        ax.set_xticks(range(7))
        ax.set_xticklabels(weekdays, fontsize=8)
        ax.set_yticks([])
        ax.set_title(title, fontsize=10, pad=6)
        ax.invert_yaxis()
        first = date(y, m, 1)
        if m == 12:
            last = date(y + 1, 1, 1) - timedelta(days=1)
        else:
            last = date(y, m + 1, 1) - timedelta(days=1) if m < 12 else date(y + 1, 1, 1) - timedelta(days=1)
        if m == 12:
            last = date(y, 12, 31)
        row0 = first.weekday()
        for day in range(1, last.day + 1):
            d = date(y, m, day)
            col = d.weekday()
            row = (row0 + day - 1) // 7
            x, yy = col, row + 0.12
            in_plan = p0 <= d <= p1
            if d in P.HOLIDAYS:
                fc, ec = "#F4C7C3", "#C0392B"
            elif d.weekday() >= 5:
                fc, ec = "#F0F0F0", "#888888"
            elif in_plan:
                fc, ec = "#D6EAF8", "#2E86AB"
            else:
                fc, ec = "#FFFFFF", "#BBBBBB"
            ax.add_patch(Rectangle((x + 0.08, yy), 0.84, 0.78, facecolor=fc, edgecolor=ec, lw=0.8))
            ax.text(x + 0.5, yy + 0.39, str(day), ha="center", va="center", fontsize=8)
        ax.set_xlim(-0.05, 7.05)
        ax.set_ylim(6.1, -0.15)
        for spine in ax.spines.values():
            spine.set_visible(False)
    fig.suptitle("Календарь проекта «ЛифтКонтур»: рабочие дни, выходные и праздники", fontsize=11, y=1.04)
    handles = [
        Patch(facecolor="#D6EAF8", edgecolor="#2E86AB", label="Рабочий день плана"),
        Patch(facecolor="#F0F0F0", edgecolor="#888888", label="Суббота / воскресенье"),
        Patch(facecolor="#F4C7C3", edgecolor="#C0392B", label="Праздник"),
    ]
    fig.legend(handles=handles, loc="lower center", ncol=3, frameon=False, fontsize=8.5, bbox_to_anchor=(0.5, -0.04))
    save(fig, "fig01_calendar.png")


def fig02_wbs():
    fig, ax = plt.subplots(figsize=(13.6, 6.6))
    ax.set_xlim(0, 14)
    ax.set_ylim(0, 8.2)
    ax.axis("off")

    def box(x, y, w, h, text, fc, fs=8.2):
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.08",
                                    facecolor=fc, edgecolor="black", lw=1.05))
        ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=fs)

    box(4.7, 7.15, 4.6, 0.85, "КИС «ЛифтКонтур»\nразработка информационной системы", "#FFFFFF", 9.2)
    phases = [
        (0.3, "1. Инициация\nи анализ", "#D6EAF8"),
        (3.9, "2. Проектирование", "#FCE4C4"),
        (7.5, "3. Разработка", "#D5F0D0"),
        (11.1, "4. Испытания\nи сдача", "#F7D4D2"),
    ]
    for x, title, fc in phases:
        box(x, 5.55, 2.6, 0.95, title, fc, 8.6)
        ax.annotate("", xy=(x + 1.3, 6.5), xytext=(7.0, 7.15),
                    arrowprops=dict(arrowstyle="-", color="#555", lw=0.9))

    leaves = [
        (0.15, 4.2, "1.1 Обследование бригад"),
        (0.15, 3.25, "1.2 Разбор Excel-журналов"),
        (0.15, 2.3, "1.3 Требования"),
        (0.15, 1.35, "1.4 Техническое задание"),
        (0.15, 0.4, "1.5 Веха: утверждение ТЗ"),
        (3.75, 4.2, "2.1 Архитектура"),
        (3.75, 3.25, "2.2 Модель данных  ||"),
        (3.75, 2.3, "2.3 Интерфейсы АРМ  ||"),
        (3.75, 1.35, "2.4 Интеграция с 1С  ||"),
        (7.35, 4.2, "3.1 Сервер и API  ||"),
        (7.35, 3.25, "3.2 Веб-АРМ  ||"),
        (7.35, 2.3, "3.3 Мобильный клиент  ||"),
        (7.35, 1.35, "3.4 Обмен с 1С"),
        (10.95, 4.2, "4.1 Тестирование"),
        (10.95, 3.25, "4.2 Пилот, 2 дома"),
        (10.95, 2.3, "4.3 Доработка"),
        (10.95, 1.35, "4.4 Веха: сдача"),
    ]
    for x, y, text in leaves:
        phase = text[0]
        fc = {"1": "#EAF3FA", "2": "#FEF3E4", "3": "#E8F6E6", "4": "#FBEAEA"}[phase]
        box(x, y, 2.9, 0.78, text, fc, 7.4)
    ax.text(0.15, 7.95, "Иерархия работ (WBS). Знак || отмечает заранее заложенный параллелизм", fontsize=9)
    save(fig, "fig02_wbs.png")


def fig03_task_list():
    fig, ax = plt.subplots(figsize=(12.8, 7.4))
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 22)
    ax.axis("off")
    ax.set_title("Иерархический список работ (отступ подзадач, как indent в планировщике)", fontsize=11, pad=8)
    headers = ["Код", "Наименование", "Тип", "Длительность", "Предшественники"]
    xs = [0.15, 1.15, 6.4, 8.3, 10.0]
    ax.add_patch(Rectangle((0.1, 20.35), 11.7, 0.9, facecolor="#D9D9D9", edgecolor="black", lw=0.6))
    for x, h in zip(xs, headers):
        ax.text(x, 20.72, h, fontsize=8.5, fontweight="bold", va="center")
    rows = []
    for code in P.ORDER:
        t = P.TASKS[code]
        kind = {"summary": "этап", "milestone": "веха", "leaf": "работа"}[t.kind]
        dur = "—" if t.kind == "summary" else ("0 (веха)" if t.duration == 0 else f"{t.duration} раб. дн.")
        pred = "—" if not t.preds else ", ".join(t.preds)
        indent = 0 if t.kind == "summary" else 1
        rows.append((code, t.name, kind, dur, pred, indent, t.kind == "summary"))
    y = 19.9
    for i, (code, name, kind, dur, pred, indent, bold) in enumerate(rows):
        bg = "#F4F4F4" if i % 2 == 0 else "#FFFFFF"
        ax.add_patch(Rectangle((0.1, y - 0.38), 11.7, 0.82, facecolor=bg, edgecolor="#CCCCCC", lw=0.4))
        vals = [code, ("    " * indent) + name, kind, dur, pred]
        weight = "bold" if bold else "normal"
        for x, v in zip(xs, vals):
            ax.text(x, y + 0.02, v, fontsize=7.6, va="center", fontweight=weight)
        y -= 0.88
    save(fig, "fig03_task_list.png")


def _node_box(ax, x, y, t: P.Task, w=1.55, h=1.05):
    fc = "#F5B7B1" if t.critical else "#D6EAF8"
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.015,rounding_size=0.06",
                                facecolor=fc, edgecolor="#1A1A1A", lw=1.0, zorder=3))
    ax.text(x + w / 2, y + h * 0.78, f"{t.code}", ha="center", va="center", fontsize=7.4, fontweight="bold", zorder=4)
    ax.text(x + w / 2, y + h * 0.52, f"{t.duration}д", ha="center", va="center", fontsize=6.6, zorder=4)
    ax.text(x + w / 2, y + h * 0.22, f"R {P.fmt(t.es)[0:5]}", ha="center", va="center", fontsize=5.8, zorder=4)
    return x + w, y + h / 2, x, y + h / 2


def fig04_network():
    fig, ax = plt.subplots(figsize=(15.4, 7.8))
    ax.set_xlim(-0.2, 16.4)
    ax.set_ylim(-0.3, 8.4)
    ax.axis("off")
    ax.set_title("Сетевой график работ (AON). Красные узлы — критический путь", fontsize=11, pad=8)

    pos = {
        "1.1": (0.15, 5.9),
        "1.2": (0.15, 3.6),
        "1.3": (2.05, 4.75),
        "1.4": (3.95, 4.75),
        "1.5": (5.85, 4.75),
        "2.1": (7.55, 4.75),
        "2.2": (9.35, 6.35),
        "2.3": (9.35, 4.75),
        "2.4": (9.35, 3.15),
        "3.1": (11.15, 6.35),
        "3.2": (11.15, 4.75),
        "3.3": (11.15, 3.15),
        "3.4": (12.95, 6.35),
        "4.1": (12.95, 4.35),
        "4.2": (14.55, 4.35),
        "4.3": (14.55, 2.55),
        "4.4": (14.55, 0.75),
    }
    centers = {}
    for code, (x, y) in pos.items():
        t = P.TASKS[code]
        if t.kind == "milestone":
            ax.add_patch(RegularPolygon((x + 0.55, y + 0.52), 4, radius=0.52, orientation=0.785,
                                        facecolor="#F5B7B1" if t.critical else "#D5F0D0",
                                        edgecolor="black", lw=1.05, zorder=3))
            ax.text(x + 0.55, y + 0.62, t.code, ha="center", va="center", fontsize=7.2, fontweight="bold", zorder=4)
            ax.text(x + 0.55, y + 0.38, "веха", ha="center", va="center", fontsize=6.2, zorder=4)
            centers[code] = (x + 0.55, y + 0.52)
        else:
            _node_box(ax, x, y, t)
            centers[code] = (x + 0.78, y + 0.52)

    def edge(a, b):
        x1, y1 = centers[a]
        x2, y2 = centers[b]
        crit = P.TASKS[a].critical and P.TASKS[b].critical
        ax.annotate("", xy=(x2 - 0.82 if x2 > x1 + 1.2 else x2, y2),
                    xytext=(x1 + 0.82 if x2 > x1 + 1.2 else x1, y1),
                    arrowprops=dict(arrowstyle="-|>", color="#C0392B" if crit else "#555555",
                                    lw=1.35 if crit else 0.9, mutation_scale=9),
                    zorder=2)

    # ручная геометрия стрелок: от правого края к левому
    pairs = [
        ("1.1", "1.3"), ("1.2", "1.3"), ("1.3", "1.4"), ("1.4", "1.5"), ("1.5", "2.1"),
        ("2.1", "2.2"), ("2.1", "2.3"), ("2.1", "2.4"),
        ("2.2", "3.1"), ("2.3", "3.2"), ("2.3", "3.3"),
        ("2.4", "3.4"), ("3.1", "3.4"),
        ("3.1", "4.1"), ("3.2", "4.1"), ("3.3", "4.1"), ("3.4", "4.1"),
        ("4.1", "4.2"), ("4.2", "4.3"), ("4.3", "4.4"),
    ]
    for a, b in pairs:
        x1, y1 = centers[a]
        x2, y2 = centers[b]
        crit = P.TASKS[a].critical and P.TASKS[b].critical
        ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>", mutation_scale=10,
                                     lw=1.4 if crit else 0.85,
                                     color="#C0392B" if crit else "#666666",
                                     connectionstyle="arc3,rad=0.0", zorder=2,
                                     shrinkA=12, shrinkB=12))

    ax.text(0.15, 0.15,
            "Связи типа FS (окончание → начало). Параллельные ветви 2.2 / 2.3 / 2.4 и 3.1 / 3.2 / 3.3 сокращают срок.",
            fontsize=8)
    save(fig, "fig04_network.png")


def draw_gantt(ax, codes, *, with_resources=False, title="", show_summary=True, arrows=True):
    items = []
    for c in codes:
        t = P.TASKS[c]
        if t.kind == "summary" and not show_summary:
            continue
        items.append(t)
    n = len(items)
    t0 = min(t.es for t in items)
    t1 = max(t.ef for t in items) + timedelta(days=2)
    shade_nonwork(ax, t0 - timedelta(days=1), t1, 0, n)
    ymap = {t.code: n - 1 - i for i, t in enumerate(items)}

    if arrows:
        for t in items:
            if t.kind == "summary":
                continue
            for p in t.preds:
                if p not in ymap:
                    continue
                pred = P.TASKS[p]
                ax.annotate(
                    "",
                    xy=(dnum(t.es) + 0.05, ymap[t.code]),
                    xytext=(dnum(pred.ef) + 0.85, ymap[p]),
                    arrowprops=dict(arrowstyle="-|>", color="#444444", lw=0.7, mutation_scale=7,
                                    connectionstyle="arc3,rad=0.08"),
                    zorder=4,
                )

    for t in items:
        y = ymap[t.code]
        x0 = dnum(t.es)
        x1 = dnum(t.ef) + 1
        if t.kind == "summary":
            ax.plot([x0, x1], [y, y], color="#222", lw=4.2, solid_capstyle="butt", zorder=3)
            ax.plot(x0, y, marker="v", color="#222", markersize=7, zorder=4)
            ax.plot(x1, y, marker="v", color="#222", markersize=7, zorder=4)
        elif t.kind == "milestone":
            ax.plot(x0, y, marker="D", color="#C0392B", markersize=9, zorder=5)
        else:
            color = "#C0392B" if t.critical else P.PHASE_COLOR[t.code[0]]
            ax.barh(y, x1 - x0, left=x0, height=0.55, color=color, edgecolor="#222", lw=0.5, zorder=3)
            if with_resources and t.alloc:
                names = ", ".join(P.RESOURCES[rid]["fio"].split()[0] for rid, _ in t.alloc)
                ax.text(x1 + 0.35, y, names, va="center", ha="left", fontsize=6.6, color="#222")
    labels = []
    for t in items:
        prefix = "" if t.kind == "summary" else "  "
        labels.append(f"{prefix}{t.code}  {t.name}")
    ax.set_yticks(range(n))
    ax.set_yticklabels(labels[::-1], fontsize=8)
    ax.set_xlim(dnum(t0) - 1.4, dnum(t1) + (18 if with_resources else 2.2))
    ax.set_ylim(-0.8, n - 0.2)
    ax.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=mdates.MO))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d.%m"))
    ax.tick_params(axis="x", labelsize=7.5)
    ax.grid(axis="x", linestyle=":", color="#AAAAAA", zorder=1)
    ax.set_title(title, fontsize=11, pad=8)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)


def fig05_gantt():
    fig, ax = plt.subplots(figsize=(14.8, 8.2))
    draw_gantt(ax, P.ORDER, title="Диаграмма Ганта проекта КИС «ЛифтКонтур» (критический путь выделен)")
    handles = [
        Patch(facecolor="#C0392B", label="Критическая работа"),
        Patch(facecolor="#4C78A8", label="Инициация"),
        Patch(facecolor="#F58518", label="Проектирование"),
        Patch(facecolor="#54A24B", label="Разработка"),
        Patch(facecolor="#E45756", label="Испытания"),
        Line2D([0], [0], color="#222", lw=4, label="Этап (сводка)"),
        Line2D([0], [0], marker="D", color="#C0392B", lw=0, markersize=8, label="Веха"),
        Patch(facecolor="#E8E8E8", label="Выходные / праздники"),
    ]
    ax.legend(
        handles=handles,
        loc="upper center",
        bbox_to_anchor=(0.5, 1.18),
        fontsize=7.2,
        framealpha=0.94,
        ncol=4,
        borderpad=0.35,
    )
    fig.subplots_adjust(top=0.82)
    fig.autofmt_xdate(rotation=40, ha="right")
    save(fig, "fig05_gantt.png")


def fig06_milestones():
    fig, ax = plt.subplots(figsize=(13.4, 3.6))
    t0, t1 = P.TASKS["1.1"].es, P.TASKS["4.4"].ef
    ax.set_xlim(dnum(t0) - 2, dnum(t1) + 4)
    ax.set_ylim(0, 2.2)
    ax.get_yaxis().set_visible(False)
    ax.plot([dnum(t0), dnum(t1)], [1.0, 1.0], color="#333", lw=2)
    marks = [
        (P.TASKS["1.1"].es, "Старт\n06.10.2026", 1.55),
        (P.TASKS["1.5"].es, "Утверждение ТЗ\n27.10.2026", 0.22),
        (P.TASKS["2.1"].es, "Старт проектирования\n28.10.2026", 1.55),
        (P.TASKS["3.1"].es, "Старт кодирования\n19.11.2026", 0.22),
        (P.TASKS["4.1"].es, "Старт испытаний\n21.12.2026", 1.55),
        (P.TASKS["4.4"].es, "Сдача\n25.01.2027", 0.22),
    ]
    for d, label, ty in marks:
        ax.plot(dnum(d), 1.0, marker="D", color="#C0392B", markersize=11, zorder=3)
        ax.plot([dnum(d), dnum(d)], [1.0, 1.35 if ty > 1 else 0.65], color="#888", lw=0.8)
        ax.text(dnum(d), ty, label, ha="center", va="center", fontsize=8)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d.%m"))
    ax.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=mdates.MO, interval=2))
    ax.set_title("Промежуточные точки (вехи) календарного плана", fontsize=11)
    for spine in ("top", "right", "left"):
        ax.spines[spine].set_visible(False)
    fig.autofmt_xdate(rotation=40, ha="right")
    save(fig, "fig06_milestones.png")


def fig07_team():
    fig, ax = plt.subplots(figsize=(13.2, 6.4))
    ax.set_xlim(0, 12.6)
    ax.set_ylim(0, 8.2)
    ax.axis("off")
    ax.set_title("Команда проекта: роли, контакты, дневные ставки", fontsize=11, pad=8)
    keys = list(P.RESOURCES.keys())
    for i, rid in enumerate(keys):
        r = P.RESOURCES[rid]
        col, row = i % 4, i // 4
        x, y = 0.25 + col * 3.1, 4.4 - row * 3.7
        ax.add_patch(FancyBboxPatch((x, y), 2.9, 3.2, boxstyle="round,pad=0.04,rounding_size=0.12",
                                    facecolor="#F7F9FC", edgecolor="#2C3E50", lw=1.1))
        ax.text(x + 1.45, y + 2.7, r["fio"], ha="center", fontsize=9.2, fontweight="bold")
        ax.text(x + 1.45, y + 2.25, r["role"], ha="center", fontsize=8.0, color="#333")
        ax.text(x + 1.45, y + 1.55, f"{r['rate']:,} ₽/день".replace(",", " "), ha="center", fontsize=9.5)
        ax.text(x + 1.45, y + 0.95, r["email"], ha="center", fontsize=6.6)
        ax.text(x + 1.45, y + 0.5, r["phone"], ha="center", fontsize=7.4)
    save(fig, "fig07_team.png")


def fig08_rates():
    fig, ax = plt.subplots(figsize=(11.6, 4.8))
    names = [P.RESOURCES[k]["fio"] for k in P.RESOURCES]
    rates = [P.RESOURCES[k]["rate"] for k in P.RESOURCES]
    roles = [P.RESOURCES[k]["role"] for k in P.RESOURCES]
    colors = ["#4C78A8", "#F58518", "#54A24B", "#E45756", "#72B7B2", "#EECA3B", "#B279A2", "#FF9DA6"]
    bars = ax.barh(range(len(names)), rates, color=colors, edgecolor="#222", height=0.7)
    ax.set_yticks(range(len(names)))
    ax.set_yticklabels([f"{n}\n{r}" for n, r in zip(names, roles)], fontsize=8)
    ax.set_xlabel("Дневная ставка, ₽")
    ax.set_title("Денежные ресурсы: ставки исполнителей", fontsize=11)
    ax.invert_yaxis()
    for b, v in zip(bars, rates):
        ax.text(v + 80, b.get_y() + b.get_height() / 2, f"{v:,}".replace(",", " "), va="center", fontsize=8)
    ax.set_xlim(0, max(rates) * 1.18)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    save(fig, "fig08_rates.png")


def fig09_resource_load():
    load = P.resource_load_by_day()
    t0, t1 = P.TASKS["1.1"].es, P.TASKS["4.4"].ef
    days = [d for d in P.iter_days(t0, t1) if P.is_workday(d)]
    keys = list(P.RESOURCES)
    data = [[load[rid].get(d, 0.0) * 100 for d in days] for rid in keys]
    fig, ax = plt.subplots(figsize=(14.6, 5.8))
    im = ax.imshow(data, aspect="auto", cmap="YlOrRd", vmin=0, vmax=100, interpolation="nearest")
    ax.set_yticks(range(len(keys)))
    ax.set_yticklabels([P.RESOURCES[k]["fio"] for k in keys], fontsize=8)
    step = max(1, len(days) // 16)
    ax.set_xticks(range(0, len(days), step))
    ax.set_xticklabels([days[i].strftime("%d.%m") for i in range(0, len(days), step)], fontsize=7, rotation=40, ha="right")
    ax.set_title("Занятость ресурсов, % загрузки по рабочим дням", fontsize=11)
    cbar = fig.colorbar(im, ax=ax, fraction=0.025, pad=0.02)
    cbar.set_label("%")
    ax.set_xlabel("Рабочие дни календарного плана")
    save(fig, "fig09_resource_load.png")


def fig10_gantt_resources():
    fig, ax = plt.subplots(figsize=(15.2, 8.4))
    draw_gantt(ax, [c for c in P.ORDER if P.TASKS[c].kind != "summary"],
               with_resources=True, show_summary=False, arrows=False,
               title="Назначение исполнителей на работы (ресурсы на диаграмме Ганта)")
    fig.autofmt_xdate(rotation=40, ha="right")
    save(fig, "fig10_gantt_resources.png")


def fig11_parallel():
    fig, ax = plt.subplots(figsize=(13.8, 5.4))
    codes = ["2.1", "2.2", "2.3", "2.4", "3.1", "3.2", "3.3", "3.4"]
    draw_gantt(ax, codes, show_summary=False, arrows=True,
               title="Фрагмент параллельной разработки: проектирование и кодирование трёх контуров")
    ax.text(0.01, -0.18,
            "2.2 ∥ 2.3 ∥ 2.4 после архитектуры; 3.1 ∥ 3.2 ∥ 3.3 после готовности модели данных и UI. "
            "3.4 ждёт API (3.1) и проект обмена (2.4).",
            transform=ax.transAxes, fontsize=8)
    fig.autofmt_xdate(rotation=40, ha="right")
    fig.subplots_adjust(bottom=0.22)
    save(fig, "fig11_parallel.png")


def fig12_costs():
    codes = [c for c in P.LEAF_CODES if P.TASKS[c].duration > 0]
    vals = [P.task_cost(P.TASKS[c]) / 1000 for c in codes]
    labels = [f"{c} {P.TASKS[c].name}" for c in codes]
    colors = ["#C0392B" if P.TASKS[c].critical else P.PHASE_COLOR[c[0]] for c in codes]
    fig, ax = plt.subplots(figsize=(13.2, 6.6))
    ax.barh(range(len(codes)), vals, color=colors, edgecolor="#222", height=0.72)
    ax.set_yticks(range(len(codes)))
    ax.set_yticklabels(labels, fontsize=8)
    ax.invert_yaxis()
    ax.set_xlabel("Оценка фонда оплаты, тыс. ₽")
    ax.set_title("Стоимость работ (ставка × длительность × загрузка)", fontsize=11)
    for i, v in enumerate(vals):
        ax.text(v + 1.2, i, f"{v:.1f}", va="center", fontsize=7.5)
    total = P.total_cost() / 1000
    ax.axvline(0, color="#222", lw=0.6)
    ax.text(0.99, 0.02, f"Итого ≈ {total:,.0f} тыс. ₽".replace(",", " "),
            transform=ax.transAxes, ha="right", fontsize=9)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    save(fig, "fig12_costs.png")


def fig13_gantt_full():
    fig, ax = plt.subplots(figsize=(16.4, 9.0))
    draw_gantt(ax, P.ORDER, with_resources=True, arrows=False,
               title="Развёрнутая диаграмма Ганта КИС «ЛифтКонтур»: сроки, критический путь, исполнители")
    handles = [
        Patch(facecolor="#C0392B", label="Критический путь"),
        Patch(facecolor="#4C78A8", label="Анализ"),
        Patch(facecolor="#F58518", label="Проектирование"),
        Patch(facecolor="#54A24B", label="Разработка"),
        Patch(facecolor="#E45756", label="Испытания"),
        Patch(facecolor="#E8E8E8", label="Нерабочие дни"),
    ]
    ax.legend(
        handles=handles,
        loc="upper center",
        bbox_to_anchor=(0.5, 1.16),
        fontsize=7.3,
        ncol=3,
        framealpha=0.94,
    )
    fig.subplots_adjust(top=0.84)
    fig.autofmt_xdate(rotation=40, ha="right")
    save(fig, "fig13_gantt_full.png")


def fig14_pipeline():
    fig, ax = plt.subplots(figsize=(12.8, 4.2))
    ax.set_xlim(0, 13)
    ax.set_ylim(0, 4.2)
    ax.axis("off")
    steps = [
        (0.3, "1. WBS\nиерархия работ"),
        (2.8, "2. Календарь\nвыходные, праздники"),
        (5.3, "3. CPM\nкритический путь"),
        (7.8, "4. Ресурсы\nставки и загрузка"),
        (10.3, "5. Диаграммы\nГант, сеть, затраты"),
    ]
    for i, (x, text) in enumerate(steps):
        ax.add_patch(FancyBboxPatch((x, 1.35), 2.3, 1.7, boxstyle="round,pad=0.03,rounding_size=0.1",
                                    facecolor="#EEF4FA", edgecolor="#1A1A1A", lw=1.1))
        ax.text(x + 1.15, 2.2, text, ha="center", va="center", fontsize=9)
        if i < len(steps) - 1:
            ax.annotate("", xy=(x + 2.45, 2.2), xytext=(x + 2.3, 2.2),
                        arrowprops=dict(arrowstyle="-|>", color="#222", lw=1.2, mutation_scale=11))
    ax.set_title("Конвейер планирования в выбранном инструментарии (Python + matplotlib)", fontsize=11)
    ax.text(6.5, 0.55, "Библиотека matplotlib (NumFOCUS / BSD-совместимая лицензия) — построение Ганта, сети и загрузки",
            ha="center", fontsize=8.5)
    save(fig, "fig14_pipeline.png")


def fig15_compare():
    fig, ax = plt.subplots(figsize=(10.4, 4.0))
    labels = ["Последовательный план\n(без параллелизма)", "План с параллельными\nветвями (принят)"]
    vals = [P.sequential_duration(), P.planned_duration()]
    bars = ax.bar(labels, vals, color=["#AAB7B8", "#2471A3"], edgecolor="#222", width=0.55)
    ax.set_ylabel("Рабочих дней")
    ax.set_title("Сокращение срока за счёт распараллеливания работ", fontsize=11)
    for b, v in zip(bars, vals):
        ax.text(b.get_x() + b.get_width() / 2, v + 1.5, f"{v} дн.", ha="center", fontsize=10, fontweight="bold")
    ax.set_ylim(0, max(vals) * 1.18)
    saved = P.sequential_duration() - P.planned_duration()
    ax.text(0.5, 0.08, f"Экономия {saved} раб. дн. ({100 * saved / P.sequential_duration():.0f} % от последовательного срока)",
            transform=ax.transAxes, ha="center", fontsize=9)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    save(fig, "fig15_compare.png")


def fig16_gantt_zoom_resources():
    fig, ax = plt.subplots(figsize=(13.6, 4.8))
    draw_gantt(ax, ["3.1", "3.2", "3.3", "3.4"], with_resources=True, show_summary=False, arrows=True,
               title="Ресурсы на фрагменте Ганта: три параллельных контура кодирования")
    fig.autofmt_xdate(rotation=40, ha="right")
    save(fig, "fig16_gantt_zoom_resources.png")


def main():
    P.schedule()
    fig01_calendar()
    fig02_wbs()
    fig03_task_list()
    fig04_network()
    fig05_gantt()
    fig06_milestones()
    fig07_team()
    fig08_rates()
    fig09_resource_load()
    fig10_gantt_resources()
    fig11_parallel()
    fig12_costs()
    fig13_gantt_full()
    fig14_pipeline()
    fig15_compare()
    fig16_gantt_zoom_resources()
    print("figures:", len(list(FIG.glob("*.png"))))


if __name__ == "__main__":
    main()
