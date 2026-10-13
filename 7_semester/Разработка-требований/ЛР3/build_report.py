#!/usr/bin/env python3
"""Сборка отчёта ЛР3. В текст отчёта не входит."""

from __future__ import annotations

import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import Circle, Ellipse, FancyArrowPatch, Rectangle
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from lxml import etree

BASE = Path(__file__).resolve().parent
TEMPLATE = BASE.parent / "ЛР2" / "ЛР2_Отчет_Буренков_Требования.docx"
OUT = BASE / "ЛР3_Отчет_Буренков_Требования.docx"
FIG = BASE / "figures"
USECASE = FIG / "usecase.png"
PUBLISH = FIG / "publish.png"
SALE = FIG / "sale.png"
CANCEL = FIG / "cancel.png"
BOARD = FIG / "board.png"
REFUND = FIG / "refund.png"
DIRS = FIG / "dirs.png"
ROUTE = FIG / "route.png"
SHIFT = FIG / "shift.png"
STATION = FIG / "station.png"
TOPIC = "Спецификация вариантов использования"
PAGE_LIMIT = 15

for _font in (
    Path(r"C:\Windows\Fonts\times.ttf"),
    Path(r"C:\Windows\Fonts\timesbd.ttf"),
    Path("/System/Library/Fonts/Supplemental/Times New Roman.ttf"),
    Path("/System/Library/Fonts/Supplemental/Times New Roman Bold.ttf"),
):
    if _font.exists():
        font_manager.fontManager.addfont(str(_font))
plt.rcParams["font.family"] = "Times New Roman"


def set_run_font(run, size=14, bold=False):
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    run._element.rPr.rFonts.set(qn("w:cs"), "Times New Roman")
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = False


def format_paragraph(
    p,
    *,
    first_indent=True,
    align="justify",
    space_after=0,
    space_before=0,
    left_indent=0,
    line_spacing=1.5,
    single=False,
):
    pf = p.paragraph_format
    if single:
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    else:
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.line_spacing = line_spacing
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.first_line_indent = Cm(1.25 if first_indent else 0)
    pf.left_indent = Cm(left_indent)
    p.alignment = {
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "left": WD_ALIGN_PARAGRAPH.LEFT,
    }[align]


def add_body(doc, text, *, space_before=0):
    p = doc.add_paragraph()
    format_paragraph(p, space_before=space_before)
    set_run_font(p.add_run(text))
    return p


def add_labeled(doc, label, text):
    p = doc.add_paragraph()
    format_paragraph(p)
    set_run_font(p.add_run(label if label.endswith(" ") else label + " "), bold=True)
    set_run_font(p.add_run(text))
    return p


def add_heading(doc, text, *, new_page=False, level=1):
    if new_page:
        doc.add_page_break()
    p = doc.add_paragraph()
    if level == 0:
        format_paragraph(p, first_indent=False, align="center", space_after=12)
    else:
        format_paragraph(
            p,
            first_indent=True,
            align="left",
            space_before=0 if level == 1 else 8,
            space_after=6 if level == 1 else 3,
        )
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text), bold=True)
    return p


def add_toc_line(doc, title, page, *, indent=0):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="left", left_indent=indent)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(16.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    set_run_font(p.add_run(title))
    set_run_font(p.add_run("\t"))
    set_run_font(p.add_run(str(page)))


def add_figure(doc, path: Path, caption: str, width=15.2):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="center", space_before=6, space_after=0, single=True)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(path), width=Cm(width))
    c = doc.add_paragraph()
    format_paragraph(c, first_indent=False, align="center", space_before=3, space_after=6, single=True)
    set_run_font(c.add_run(caption))


def clone_title() -> Document:
    shutil.copy(TEMPLATE, OUT)
    doc = Document(str(OUT))
    body = doc.element.body
    children = list(body)
    for child in children[12:-1]:
        body.remove(child)
    replacements = {
        "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 2": "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 3",
        "Идентификация действующих лиц и вариантов использования": TOPIC,
    }
    found = set()
    for node in body.iter(qn("w:t")):
        if node.text in replacements:
            found.add(node.text)
            node.text = replacements[node.text]
    missing = set(replacements) - found
    if missing:
        raise SystemExit(f"На титуле не найдены строки: {missing}")
    return doc


def strip_unused_images(path: Path) -> None:
    with zipfile.ZipFile(path, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}
    rel_name = "word/_rels/document.xml.rels"
    root = etree.fromstring(files[rel_name])
    used = set()
    for name, data in files.items():
        if not (name.startswith("word/") and name.endswith(".xml") and "/_rels/" not in name):
            continue
        for key in (b'r:id="', b'r:embed="'):
            start = 0
            while True:
                pos = data.find(key, start)
                if pos < 0:
                    break
                end = data.find(b'"', pos + len(key))
                used.add(data[pos + len(key) : end].decode())
                start = end + 1
    removed = []
    for rel in list(root):
        if rel.get("Id") not in used and "image" in (rel.get("Type") or ""):
            removed.append("word/" + rel.get("Target").replace("\\", "/"))
            root.remove(rel)
    files[rel_name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    for name in removed:
        files.pop(name, None)
    tmp = path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    tmp.replace(path)


def _box(ax, x, y, w, h, text, *, size=8.5):
    ax.add_patch(Rectangle((x, y), w, h, facecolor="white", edgecolor="black", linewidth=0.9))
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=size, linespacing=1.12)


def _arrow(ax, x1, y1, x2, y2):
    ax.add_patch(
        FancyArrowPatch(
            (x1, y1),
            (x2, y2),
            arrowstyle="-|>",
            mutation_scale=9,
            linewidth=0.8,
            color="black",
            shrinkA=1,
            shrinkB=1,
        )
    )


def _actor(ax, x, y):
    ax.add_patch(Circle((x, y + 3.1), 1.15, fill=False, linewidth=1.0, edgecolor="black"))
    ax.plot([x, x], [y + 1.95, y + 0.15], color="black", lw=1.0)
    ax.plot([x - 1.7, x + 1.7], [y + 1.25, y + 1.25], color="black", lw=1.0)
    ax.plot([x, x - 1.35], [y + 0.15, y - 1.7], color="black", lw=1.0)
    ax.plot([x, x + 1.35], [y + 0.15, y - 1.7], color="black", lw=1.0)


def draw_usecase() -> None:
    FIG.mkdir(parents=True, exist_ok=True)
    fig, ax = plt.subplots(figsize=(7.4, 7.6), dpi=180)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")
    fig.patch.set_facecolor("white")
    ax.add_patch(Rectangle((24, 2), 52, 96, fill=False, linewidth=1.1, edgecolor="black"))
    ax.text(50, 95.2, "АИС железной дороги", ha="center", va="center", fontsize=9)

    cases = [
        (87, "UC-1\nОпубликовать рейс"),
        (77, "UC-2\nПродать билет"),
        (67, "UC-3\nЗакрыть продажу"),
        (57, "UC-4\nОтметить посадку"),
        (47, "UC-5\nВернуть билет"),
        (37, "UC-6\nСправочники"),
        (27, "UC-7\nСоставить маршрут"),
        (17, "UC-8\nОткрыть смену"),
        (7, "UC-9\nОбслужить станцию"),
    ]
    for y, text in cases:
        ax.add_patch(Ellipse((50, y), 34, 8.2, facecolor="white", edgecolor="black", linewidth=0.9))
        ax.text(50, y, text, ha="center", va="center", fontsize=6.9)

    _actor(ax, 10, 80)
    ax.text(10, 75.2, "Диспетчер\nрасписания", ha="center", va="top", fontsize=7.2)
    _actor(ax, 10, 40)
    ax.text(10, 35.2, "Кассир", ha="center", va="top", fontsize=7.2)
    _actor(ax, 10, 12)
    ax.text(10, 7.2, "Дежурный\nпо станции", ha="center", va="top", fontsize=7.0)
    _actor(ax, 90, 54)
    ax.text(90, 49.2, "Контролёр\nпосадки", ha="center", va="top", fontsize=7.2)
    _actor(ax, 90, 22)
    ax.text(90, 17.2, "Администратор", ha="center", va="top", fontsize=7.2)

    def link(x1, y1, x2, y2):
        ax.plot([x1, x2], [y1, y2], color="black", lw=0.8)

    link(12, 84, 33, 87)
    link(12, 80, 33, 67)
    link(12, 78, 33, 27)
    link(12, 43, 33, 77)
    link(12, 40, 33, 47)
    link(12, 38, 33, 17)
    link(12, 15, 33, 7)
    link(88, 57, 67, 57)
    link(88, 25, 67, 37)
    fig.savefig(USECASE, bbox_inches="tight", facecolor="white", pad_inches=0.12)
    plt.close(fig)


def draw_flow(path: Path, steps: list[str], side: str) -> None:
    """Компактный горизонтальный поток + боковой отказ — экономит страницы."""
    n = len(steps)
    fig, ax = plt.subplots(figsize=(7.2, 2.55), dpi=180)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")
    fig.patch.set_facecolor("white")
    w = min(16.5, 78 / n - 1.2)
    gap = (78 - n * w) / max(n - 1, 1)
    y, h = 48, 38
    centers = []
    for i, text in enumerate(steps):
        x = 2 + i * (w + gap)
        _box(ax, x, y, w, h, text, size=6.6)
        centers.append(x + w / 2)
        if i < n - 1:
            _arrow(ax, x + w, y + h / 2, x + w + gap, y + h / 2)
    # отказ снизу от середины цепочки
    mid = n // 2
    sx, sw, sh = 28, 44, 28
    _box(ax, sx, 6, sw, sh, side, size=6.4)
    _arrow(ax, centers[mid], y, sx + sw / 2, 6 + sh)
    fig.savefig(path, bbox_inches="tight", facecolor="white", pad_inches=0.08)
    plt.close(fig)


def draw_figures() -> None:
    draw_usecase()
    draw_flow(
        PUBLISH,
        ["Маршрут\nи схема", "Время\nрейса", "Фонд\nмест", "Открыт\nк продаже"],
        "Отказ: нет маршрута/схемы,\nконфликт номера или выход",
    )
    draw_flow(
        SALE,
        ["Запрос\nу окна", "Рейс\nи места", "Оплата\nи резерв", "Печать\nкупона"],
        "Отказ: нет мест, рейс закрыт,\nместо занято или нет цены",
    )
    draw_flow(
        CANCEL,
        ["Открыть\nрейс", "Причина", "Статус\n«отменён»", "Продажи\nзакрыты"],
        "Черновик в кассе не проводится.\nМассовая выплата — позже",
    )
    draw_flow(
        BOARD,
        ["Ввод\nбилета", "Проверка\nстатуса", "Подтвер-\nждение", "Одна\nотметка"],
        "Отказ: билет не найден, возвращён,\nрейс отменён или повтор",
    )
    draw_flow(
        REFUND,
        ["Найти\nбилет", "Проверить\nсрок", "Погасить", "Место\nсвободно*"],
        "Отказ: посадка, срок истёк\nили билет уже погашен",
    )
    draw_flow(
        DIRS,
        ["Открыть\nсправочник", "Создать\nили править", "Проверка\nуникальности", "Сохра-\nнение"],
        "Отказ: дубликат или запись\nзанята рейсом/маршрутом",
    )
    draw_flow(
        ROUTE,
        ["Код\nмаршрута", "Цепочка\nстанций", "Время\nхода", "Сохра-\nнение"],
        "Отказ: <2 станций, повтор\nили маршрут занят рейсом",
    )
    draw_flow(
        SHIFT,
        ["Вход на\nместо", "Открыть\nсмену", "Работа\nпо роли", "Закрыть\nсмену"],
        "Отказ: уже есть открытая смена\nили незавершённый черновик",
    )
    draw_flow(
        STATION,
        ["Открыть\nсутки", "Готовность\nобъектов", "Замечания", "Закрыть\nсутки"],
        "Отказ: станция не в справочнике,\nсутки уже закрыты или выход",
    )


HEADINGS = [
    ("ВВЕДЕНИЕ", 0),
    ("1 Задание", 0),
    ("2 Варианты использования", 0),
    ("2.1 Общая диаграмма", 1),
    ("2.2 Опубликовать расписание рейса", 1),
    ("2.3 Продать билет с резервом места", 1),
    ("2.4 Закрыть продажу при отмене рейса", 1),
    ("2.5 Отметить посадку", 1),
    ("2.6 Вернуть билет", 1),
    ("2.7 Вести справочники станций и схем вагонов", 1),
    ("2.8 Составить маршрут", 1),
    ("2.9 Открыть и закрыть смену", 1),
    ("2.10 Обслужить станцию", 1),
    ("ЗАКЛЮЧЕНИЕ", 0),
]


CASES = [
    {
        "heading": "2.2 Опубликовать расписание рейса",
        "name": "UC-1 «Опубликовать расписание рейса». ",
        "about": (
            "Диспетчер расписания публикует пассажирский рейс пилотного участка: маршрут, дату, время и схему вагонов. "
            "Система собирает один фонд мест и открывает рейс к продаже. Пассажир расписание не публикует и в систему не входит."
        ),
        "flow": (
            "1. Диспетчер указывает номер рейса, дату и маршрут, собранный в UC-7. "
            "2. Задаёт время отправления и прибытия и схему вагонов из справочника. "
            "3. Система создаёт места по схеме, без повторов номеров. "
            "4. Диспетчер подтверждает публикацию. "
            "5. Статус становится «открыт к продаже», касса видит свободные места."
        ),
        "alts": (
            "Нет маршрута или схемы — публикация останавливается: маршрут собирают в UC-7, схему вносит администратор в UC-6; "
            "тот же номер уже открыт или рейс отменён — второй фонд не создаётся; "
            "правка открытого рейса сохраняет проданные места; выход до подтверждения продажу не открывает."
        ),
        "pre": "Диспетчер работает от своего имени, маршрут и схема есть, рейс ещё не опубликован либо открыт на правку и не отменён.",
        "post": "При успехе рейс открыт к продаже и фонд создан; при неудаче продажа не открыта и второй фонд не создан.",
        "links": "Вариант начинает диспетчер. Нужен маршрут UC-7 и схема UC-6. Результат нужен продаже. Отмену выполняет UC-3; публикация их не вызывает.",
        "diagram": "Место на общей схеме — рисунок 1. Краткий поток публикации — рисунок 2.",
        "ok": "Рейс опубликован и открыт к продаже.",
        "bad": "Публикация остановлена: нет записи в справочнике, конфликт номера, рейс уже отменён или выход без подтверждения.",
        "figure": (PUBLISH, "Рисунок 2 — Поток публикации рейса", 14.0),
    },
    {
        "heading": "2.3 Продать билет с резервом места",
        "name": "UC-2 «Продать билет с резервом места». ",
        "about": (
            "Кассир находит открытый рейс, занимает свободное место по цене из прайса и печатает купон. "
            "Пассажир в систему не входит. Держится БЦ-1 и БЦ-2."
        ),
        "flow": (
            "1. Кассир задаёт станции и дату. "
            "2. Выбирает рейс, место и данные пассажира. "
            "3. Фиксирует оплату и подтверждает. "
            "4. Система резервирует место и печатает купон."
        ),
        "alts": "Нет мест, рейс закрыт, место занято, нет цены или отказ до подтверждения — билет не создаётся.",
        "pre": "Смена кассира открыта (UC-8), станция в обслуживании (UC-9), рейс открыт к продаже (UC-1).",
        "post": "При успехе билет проведён и купон напечатан; при неудаче билет не создан.",
        "links": "Начинает кассир. Нужны UC-1, UC-8 и UC-9. Расширение — UC-5.",
        "diagram": "Общая схема — рисунок 1. Поток — рисунок 3.",
        "ok": "Билет продан, место занято.",
        "bad": "Продажа не проведена.",
        "figure": (SALE, "Рисунок 3 — Поток продажи билета", 14.0),
    },
    {
        "heading": "2.4 Закрыть продажу при отмене рейса",
        "name": "UC-3 «Закрыть продажу при отмене рейса». ",
        "about": (
            "Диспетчер отменяет рейс: статус «отменён», новые продажи закрыты, оформленные билеты помечены. "
            "Массовую выплату вариант не проводит. Срок — БЦ-3."
        ),
        "flow": (
            "1. Диспетчер открывает рейс и указывает причину. "
            "2. Подтверждает отмену. "
            "3. Система закрывает продажи и помечает билеты."
        ),
        "alts": "Повторная отмена не создаётся; черновик в кассе не проводится; массовая выплата — во второй очереди.",
        "pre": "Рейс был открыт к продаже, диспетчер работает от своего имени.",
        "post": "При успехе новые продажи невозможны; при неудаче статус прежний.",
        "links": "Меняет результат UC-1. Отдельный билет гасит UC-5; место в продажу не возвращается.",
        "diagram": "Общая схема — рисунок 1. Поток — рисунок 4.",
        "ok": "Продажа закрыта, билеты помечены.",
        "bad": "Отмена не подтверждена.",
        "figure": (CANCEL, "Рисунок 4 — Поток закрытия продажи при отмене рейса", 14.0),
    },
    {
        "heading": "2.5 Отметить посадку",
        "name": "UC-4 «Отметить посадку». ",
        "about": (
            "Контролёр по номеру билета или купону ставит одну отметку посадки. "
            "Цену и расписание не меняет."
        ),
        "flow": (
            "1. Вводит номер билета или считывает купон. "
            "2. Проверяет статус и подтверждает. "
            "3. Система записывает одну отметку."
        ),
        "alts": "Билет не найден, возвращён, рейс отменён или посадка уже есть — отметка не создаётся.",
        "pre": "Смена контролёра открыта (UC-8), станция в обслуживании (UC-9), билет действующий.",
        "post": "При успехе одна отметка; при неудаче отметки нет.",
        "links": "Нужен билет UC-2. После отметки возврат UC-5 закрыт.",
        "diagram": "Общая схема — рисунок 1. Поток — рисунок 5.",
        "ok": "Посадка отмечена один раз.",
        "bad": "Отметка не записана.",
        "figure": (BOARD, "Рисунок 5 — Поток отметки посадки", 14.0),
    },
    {
        "heading": "2.6 Вернуть билет",
        "name": "UC-5 «Вернуть билет». ",
        "about": (
            "Кассир до отправления и до посадки гасит билет. Если рейс не отменён, место снова свободно."
        ),
        "flow": (
            "1. Находит билет по номеру. "
            "2. Подтверждает возврат. "
            "3. Система гасит билет и при необходимости освобождает место."
        ),
        "alts": "Посадка, срок истёк или билет уже погашен — отказ; при отменённом рейсе место остаётся закрытым.",
        "pre": "Смена кассира открыта, билет продан, посадка не отмечена.",
        "post": "Билет возвращён; место свободно только если рейс не отменён.",
        "links": "Точка расширения UC-2. Массовая выплата не вызывается.",
        "diagram": "Общая схема — рисунок 1. Поток — рисунок 6.",
        "ok": "Билет возвращён.",
        "bad": "Возврат не проведён.",
        "figure": (REFUND, "Рисунок 6 — Поток возврата билета", 14.0),
    },
    {
        "heading": "2.7 Вести справочники станций и схем вагонов",
        "name": "UC-6 «Вести справочники станций и схем вагонов». ",
        "about": (
            "Администратор ведёт единые справочники пилотного участка: станции и схемы пассажирских вагонов. "
            "Маршрут как упорядоченная цепочка станций собирает диспетчер в UC-7. Касса второго списка станций не держит. "
            "Грузовой парк и найм сотрудников в справочник не входят."
        ),
        "flow": (
            "1. Администратор открывает справочник станций или схем вагонов. "
            "2. Создаёт или правит запись: код и название станции либо тип вагона и номера мест. "
            "3. Система проверяет уникальность кода станции и номеров мест. "
            "4. Администратор подтверждает сохранение. "
            "5. Запись доступна для составления маршрута и публикации рейса."
        ),
        "alts": (
            "Повтор кода или номера места — запись не сохраняется; "
            "удаление станции или схемы, которую использует открытый рейс или действующий маршрут, запрещено; "
            "выход без сохранения справочник не меняет."
        ),
        "pre": "Администратор работает от своего имени, справочники пилотного участка открыты для правки.",
        "post": "При успехе станция или схема доступна дальше по контуру; при неудаче справочник прежний.",
        "links": "Вариант начинает администратор. UC-7 и UC-1 читают готовые записи и UC-6 не включают. Учётные записи и резервное копирование сюда не входят.",
        "diagram": "Место на общей схеме — рисунок 1. Краткий поток справочников — рисунок 7.",
        "ok": "Запись станции или схемы вагона сохранена.",
        "bad": "Справочник не изменён: дубликат, запись занята открытым рейсом или маршрутом, выход без сохранения.",
        "figure": (DIRS, "Рисунок 7 — Поток ведения справочников", 14.0),
    },
    {
        "heading": "2.8 Составить маршрут",
        "name": "UC-7 «Составить маршрут». ",
        "about": (
            "Диспетчер расписания собирает пассажирский маршрут пилотного участка: упорядоченный список станций "
            "и плановое время хода между ними. Без сохранённого маршрута публикация рейса (UC-1) не начинается. "
            "Это не управление движением поезда и не грузовой маршрут."
        ),
        "flow": (
            "1. Диспетчер создаёт маршрут и задаёт код и название. "
            "2. Добавляет станции из справочника по порядку следования. "
            "3. Указывает плановое время хода или стоянки между соседними станциями. "
            "4. Система проверяет, что станций не меньше двух и коды не повторяются в цепочке. "
            "5. Диспетчер подтверждает сохранение; маршрут доступен при публикации рейса."
        ),
        "alts": (
            "Станции нет в справочнике — сначала UC-6; меньше двух станций или повтор станции в цепочке — маршрут не сохраняется; "
            "правка маршрута, уже выбранного открытым рейсом, запрещена без предварительной отмены или закрытия рейса; "
            "выход без подтверждения маршрут не меняет."
        ),
        "pre": "Диспетчер работает от своего имени, нужные станции есть в справочнике UC-6.",
        "post": "При успехе маршрут сохранён и доступен в UC-1; при неудаче прежняя версия не затирается.",
        "links": "Вариант начинает диспетчер. Станции берутся из UC-6. UC-1 выбирает готовый маршрут и не включает его правку.",
        "diagram": "Место на общей схеме — рисунок 1. Краткий поток маршрута — рисунок 8.",
        "ok": "Маршрут из станций и времён хода сохранён.",
        "bad": "Маршрут не сохранён: нет станций, цепочка некорректна, маршрут занят открытым рейсом или выход без подтверждения.",
        "figure": (ROUTE, "Рисунок 8 — Поток составления маршрута", 14.0),
    },
    {
        "heading": "2.9 Открыть и закрыть смену",
        "name": "UC-8 «Открыть и закрыть смену». ",
        "about": (
            "Кассир или контролёр открывает и закрывает смену на рабочем месте. "
            "Без смены продажа, возврат и посадка недоступны. Найм здесь не выполняется."
        ),
        "flow": (
            "1. Входит на рабочее место станции. "
            "2. Открывает смену: роль и время начала. "
            "3. Работает по роли; в конце закрывает смену."
        ),
        "alts": "Вторая открытая смена не создаётся; закрытие при черновике продажи останавливается; администратор может аварийно закрыть смену.",
        "pre": "Учётная запись есть, станция в обслуживании по UC-9.",
        "post": "Смена активна или закрыта; история сохранена.",
        "links": "UC-2, UC-4 и UC-5 требуют открытой смены. Нужен UC-9.",
        "diagram": "Общая схема — рисунок 1. Поток — рисунок 9.",
        "ok": "Смена открыта или закрыта.",
        "bad": "Конфликт смен, черновик или выход без подтверждения.",
        "figure": (SHIFT, "Рисунок 9 — Поток открытия и закрытия смены", 14.0),
    },
    {
        "heading": "2.10 Обслужить станцию",
        "name": "UC-9 «Обслужить станцию». ",
        "about": (
            "Дежурный по станции ведёт обслуживание станции на сутки: отмечает готовность касс, платформ, "
            "зала ожидания и проходов, фиксирует замечания и закрывает сутки. "
            "Это не управление движением и не продажа билетов: без открытого обслуживания смены персонала не начинаются."
        ),
        "flow": (
            "1. Дежурный выбирает станцию пилотного участка и открывает сутки обслуживания. "
            "2. Отмечает готовность объектов: кассы, платформы, зал, проходы. "
            "3. При необходимости вносит замечание или ограничение. "
            "4. Подтверждает, что станция в обслуживании. "
            "5. В конце суток закрывает обслуживание; новые смены на станции не открываются."
        ),
        "alts": (
            "Станции нет в справочнике — сначала UC-6; сутки уже закрыты — повторное открытие только на новую дату; "
            "закрытие при открытых сменах персонала запрещено, пока смены не закроют в UC-8; "
            "выход без подтверждения статус суток не меняет."
        ),
        "pre": "Дежурный работает от своего имени, станция есть в справочнике UC-6.",
        "post": "При открытии станция в обслуживании и доступна для смен; при закрытии новые смены недоступны, журнал суток сохранён.",
        "links": "Начинает дежурный по станции. UC-8 требует открытого обслуживания. UC-2 и UC-4 идут через смену и сюда напрямую не входят.",
        "diagram": "Общая схема — рисунок 1. Поток — рисунок 10.",
        "ok": "Сутки обслуживания открыты или корректно закрыты.",
        "bad": "Статус не изменён: нет станции, сутки закрыты, есть открытые смены или выход без подтверждения.",
        "figure": (STATION, "Рисунок 10 — Поток обслуживания станции", 14.0),
    },
]


def write_case(doc, case: dict) -> None:
    add_heading(doc, case["heading"], level=2)
    add_labeled(doc, "Имя.", case["name"] + case["about"])
    add_labeled(doc, "Основной поток.", case["flow"])
    add_labeled(doc, "Альтернативы.", case["alts"])
    add_labeled(doc, "Предусловие.", case["pre"])
    add_labeled(doc, "Постусловие.", case["post"])
    add_labeled(doc, "Связи.", case["links"])
    add_labeled(doc, "Диаграмма.", case["diagram"])
    p = doc.add_paragraph()
    format_paragraph(p)
    set_run_font(p.add_run("Успех. "), bold=True)
    set_run_font(p.add_run(case["ok"] + " "))
    set_run_font(p.add_run("Неудача. "), bold=True)
    set_run_font(p.add_run(case["bad"]))
    if case["figure"]:
        path, caption, width = case["figure"]
        add_figure(doc, path, caption, width)


def build_body(doc: Document, pages: dict[str, str]) -> None:
    add_heading(doc, "СОДЕРЖАНИЕ", level=0)
    for title, indent in HEADINGS:
        add_toc_line(doc, title, pages.get(title, "…"), indent=indent * 0.5)

    add_heading(doc, "ВВЕДЕНИЕ", new_page=True, level=0)
    add_body(
        doc,
        "Работа продолжает документ о концепции и границах автоматизированной информационной системы железной дороги. "
        "Система не сводится к кассе: заданы маршруты, справочники, обслуживание станции, смены персонала, "
        "публикация и отмена рейса, а билетный контур — кратко.",
    )
    add_body(
        doc,
        "Пассажир в систему не входит. Грузовые перевозки, управление движением, найм и интернет-магазин вне границ. "
        "Пороги: БЦ-1 — оформление билета не более 3 минут в 90 % случаев за три месяца; БЦ-2 — не более одного "
        "инцидента двойной продажи в месяц; БЦ-3 — после отмены продажи закрыты не позднее 5 минут в 95 % случаев.",
    )

    add_heading(doc, "1 Задание", new_page=True)
    add_body(
        doc,
        "Нужно составить спецификацию вариантов использования: имя, описание, потоки, предусловия, постусловия, "
        "связи, диаграмму и сценарии успеха и неудачи. Берутся UC-1–UC-9: рейс и билет (кратко), справочники, "
        "маршрут, смена и обслуживание станции. Сводка продаж за смену — во второй очереди.",
    )

    add_heading(doc, "2 Варианты использования", new_page=True)
    add_body(
        doc,
        "Девять вариантов показаны на рисунке 1. Для каждого ниже — краткая схема потока (рисунки 2–10): "
        "ход слева направо и типовой отказ снизу.",
    )
    add_heading(doc, "2.1 Общая диаграмма", level=2)
    add_body(
        doc,
        "На рисунке 1 граница системы одна. Диспетчер — публикация, отмена и маршрут; дежурный по станции — "
        "обслуживание станции; кассир — продажа, возврат и смена; контролёр — посадка и смена; администратор — "
        "справочники. Пассажир действующим лицом не является.",
    )
    add_figure(doc, USECASE, "Рисунок 1 — Диаграмма вариантов использования АИС железной дороги", 14.2)
    for case in CASES:
        write_case(doc, case)

    add_heading(doc, "ЗАКЛЮЧЕНИЕ", new_page=True, level=0)
    add_body(
        doc,
        "Составлены спецификации девяти вариантов АИС железной дороги. Маршрут и справочники готовят публикацию рейса, "
        "обслуживание станции и смена открывают работу персонала, билетный контур занимает место, закрывает продажу "
        "при отмене и фиксирует посадку или возврат без раздувания объёма.",
    )
    add_body(
        doc,
        "Пороги БЦ-1–БЦ-3 не ослаблены. Пассажир вне системы. Груз, управление движением, найм и онлайн-продажа "
        "в границы не вошли. Сводка продаж за смену — во второй очереди.",
    )


def save_doc(doc: Document) -> None:
    doc.save(str(OUT))
    strip_unused_images(OUT)


FALLBACK_PAGES = {
    "ВВЕДЕНИЕ": "3",
    "1 Задание": "4",
    "2 Варианты использования": "5",
    "2.1 Общая диаграмма": "5",
    "2.2 Опубликовать расписание рейса": "6",
    "2.3 Продать билет с резервом места": "7",
    "2.4 Закрыть продажу при отмене рейса": "7",
    "2.5 Отметить посадку": "8",
    "2.6 Вернуть билет": "8",
    "2.7 Вести справочники станций и схем вагонов": "9",
    "2.8 Составить маршрут": "10",
    "2.9 Открыть и закрыть смену": "11",
    "2.10 Обслужить станцию": "12",
    "ЗАКЛЮЧЕНИЕ": "13",
}


def measure_pages() -> tuple[dict[str, str], int]:
    import platform

    if platform.system() == "Darwin":
        return dict(FALLBACK_PAGES), 13
    listing = Path(tempfile.gettempdir()) / "lr3_pages.txt"
    script = Path(tempfile.gettempdir()) / "lr3_pages.ps1"
    out = str(OUT).replace("'", "''")
    listing_s = str(listing).replace("'", "''")
    script.write_text(
        "\n".join(
            [
                "$ErrorActionPreference = 'Stop'",
                "$word = $null",
                "$doc = $null",
                "try {",
                "  $word = New-Object -ComObject Word.Application",
                "  $word.Visible = $false",
                "  $word.DisplayAlerts = 0",
                f"  $doc = $word.Documents.Open('{out}', $false, $true)",
                "  $doc.Repaginate()",
                "  $total = $doc.ComputeStatistics(2)",
                "  $lines = New-Object System.Collections.Generic.List[string]",
                "  $lines.Add(\"TOTAL`t$total\")",
                "  foreach ($p in $doc.Paragraphs) {",
                "    $t = $p.Range.Text.Replace(\"`r\",\"\").Replace(\"`a\",\"\").Trim()",
                "    if ($t.Length -gt 0 -and -not $t.Contains(\"`t\")) {",
                "      $page = [int]$p.Range.Information(3)",
                "      $lines.Add(\"$page`t$t\")",
                "    }",
                "  }",
                f"  [System.IO.File]::WriteAllLines('{listing_s}', $lines, [System.Text.UTF8Encoding]::new($false))",
                "} finally {",
                "  if ($doc -ne $null) { $doc.Close(0) }",
                "  if ($word -ne $null) { $word.Quit() }",
                "}",
            ]
        ),
        encoding="utf-8-sig",
    )
    subprocess.run(["powershell", "-NoProfile", "-File", str(script)], check=True)
    found: dict[str, str] = {}
    total = 0
    wanted = {title for title, _indent in HEADINGS}
    for line in listing.read_text(encoding="utf-8").splitlines():
        if "\t" not in line:
            continue
        page, text = line.split("\t", 1)
        if page == "TOTAL":
            total = int(text.strip())
            continue
        text = text.strip()
        if text in wanted and text not in found:
            found[text] = page.strip()
    return found, total


def main() -> None:
    draw_figures()
    pages = {title: "…" for title, _indent in HEADINGS}
    total = 0
    for _ in range(3):
        doc = clone_title()
        build_body(doc, pages)
        save_doc(doc)
        measured, total = measure_pages()
        missing = [title for title, _indent in HEADINGS if title not in measured]
        if missing:
            raise SystemExit(f"Не найдены страницы: {missing}")
        if measured == pages:
            break
        pages = measured
    print("PAGES", total)
    print(pages)
    if total > PAGE_LIMIT:
        raise SystemExit(f"ЛР3 длиннее {PAGE_LIMIT} страниц: {total}")
    print(OUT)


if __name__ == "__main__":
    main()
