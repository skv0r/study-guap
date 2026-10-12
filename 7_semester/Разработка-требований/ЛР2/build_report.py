#!/usr/bin/env python3
"""Сборка отчёта ЛР2. В текст отчёта не входит."""

from __future__ import annotations

import platform
import shutil
import zipfile
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import Circle, Ellipse, FancyArrowPatch, FancyBboxPatch, Rectangle
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Pt
from lxml import etree

BASE = Path(__file__).resolve().parent
REPO = BASE.parents[1]
TEMPLATE = REPO / "Web-программирование" / "ЛР1_Отчет_Буренков_PairWise.docx"
OUT = BASE / "ЛР2_Отчет_Буренков_Требования.docx"
FIG = BASE / "figures"
PROCESS = FIG / "process.png"
USECASE = FIG / "usecase.png"

TOPIC = "Идентификация действующих лиц и вариантов использования"
COURSE = "по курсу: Разработка и анализ требований"
USABLE = Cm(16.5)

_FONT_CANDIDATES = [
    Path("/System/Library/Fonts/Supplemental/Times New Roman.ttf"),
    Path("/System/Library/Fonts/Supplemental/Times New Roman Bold.ttf"),
    Path(r"C:\Windows\Fonts\times.ttf"),
    Path(r"C:\Windows\Fonts\timesbd.ttf"),
]
for _font in _FONT_CANDIDATES:
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
            space_before=0 if level == 1 else 12,
            space_after=12 if level == 1 else 6,
        )
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text), bold=True)
    return p


def add_toc_line(doc, title, page, *, sub=False):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="left", left_indent=0.5 if sub else 0)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(16.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    set_run_font(p.add_run(title))
    set_run_font(p.add_run("\t"))
    set_run_font(p.add_run(str(page)))


def add_table_caption(doc, text):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="left", space_before=12, space_after=6, single=True)
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text))


def add_figure(doc, path: Path, caption: str):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="center", space_before=6, space_after=0, single=True)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(path), width=Cm(16.0))
    c = doc.add_paragraph()
    format_paragraph(c, first_indent=False, align="center", space_before=6, space_after=6, single=True)
    set_run_font(c.add_run(caption))


def _borders(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tcPr.append(borders)


def _set_cell_text(cell, text, *, bold=False, align="left", size=14):
    cell.text = ""
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    lines = str(text).split("\n")
    for i, line in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        format_paragraph(p, first_indent=False, align=align, single=True)
        set_run_font(p.add_run(line if line else " "), size=size, bold=bold)


def add_table(doc, headers, rows, widths):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    total = sum(widths)
    emu = [int(USABLE * (w / total)) for w in widths]
    grid = table._tbl.tblGrid
    for i, child in enumerate(list(grid)):
        child.set(qn("w:w"), str(int(emu[i] / 635)))
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(int(USABLE / 635)))
    tblW.set(qn("w:type"), "dxa")
    jc = tblPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        tblPr.append(jc)
    jc.set(qn("w:val"), "left")
    ind = tblPr.find(qn("w:tblInd"))
    if ind is None:
        ind = OxmlElement("w:tblInd")
        tblPr.append(ind)
    ind.set(qn("w:w"), "0")
    ind.set(qn("w:type"), "dxa")
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            cell.width = Emu(emu[idx])
            _borders(cell)
    for j, header in enumerate(headers):
        _set_cell_text(table.rows[0].cells[j], header, bold=True, align="center")
    for i, row in enumerate(rows):
        for j, value in enumerate(row):
            _set_cell_text(table.rows[i + 1].cells[j], value, align="left")
    return table


def clone_title() -> Document:
    shutil.copy(TEMPLATE, OUT)
    doc = Document(str(OUT))
    body = doc.element.body
    children = list(body)
    for child in children[12:-1]:
        body.remove(child)
    replacements = {
        "ассистент": "доцент, к.т.н.",
        "В.В. Жукалин": "А.В. Бржезовский",
        "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 1": "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 2",
        "Создание проекта Next.js с использованием React, TypeScript и App Router": TOPIC,
        "по курсу: Web-программирование": COURSE,
    }
    for node in body.iter(qn("w:t")):
        if node.text in replacements:
            node.text = replacements[node.text]
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


def _box(ax, x, y, w, h, text, *, fc="white", size=8.5):
    ax.add_patch(
        FancyBboxPatch(
            (x, y),
            w,
            h,
            boxstyle="round,pad=0.15,rounding_size=0.35",
            facecolor=fc,
            edgecolor="black",
            linewidth=0.9,
            mutation_aspect=0.6,
        )
    )
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=size, linespacing=1.15)


def _arrow(ax, x1, y1, x2, y2):
    ax.add_patch(
        FancyArrowPatch(
            (x1, y1),
            (x2, y2),
            arrowstyle="-|>",
            mutation_scale=9,
            linewidth=0.8,
            color="black",
            shrinkA=0,
            shrinkB=0,
        )
    )


def draw_process() -> None:
    FIG.mkdir(parents=True, exist_ok=True)
    fig, ax = plt.subplots(figsize=(6.55, 8.15), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")
    fig.patch.set_facecolor("white")

    lanes = [
        (2, 20, "#F4F4F4", "Внешнее\nсобытие"),
        (24, 24, "#FFFFFF", "Диспетчер\nрасписания"),
        (50, 24, "#F7F7F7", "Кассир"),
        (76, 22, "#FFFFFF", "Контролёр\nпосадки"),
    ]
    for x, w, color, _title in lanes:
        ax.add_patch(Rectangle((x, 6), w, 86, facecolor=color, edgecolor="#666666", linewidth=0.6))
        ax.add_patch(Rectangle((x, 92), w, 7, facecolor="#E6E6E6", edgecolor="black", linewidth=0.7))
    for x, w, _color, title in lanes:
        ax.text(x + w / 2, 95.5, title, ha="center", va="center", fontsize=8, fontweight="bold")

    events = [
        (78, "E1\nНужен\nрейс"),
        (62, "E2\nЗапрос\nпассажира"),
        (46, "E4\nРейс\nотменён"),
        (30, "E5\nПора\nсажать"),
        (14, "E6\nНужна\nсводка"),
    ]
    for y, text in events:
        _box(ax, 3.2, y, 17.6, 10, text, fc="#FFFFFF", size=7.5)

    _box(ax, 25.2, 78, 21.6, 10, "UC-1\nОпубликовать\nрасписание рейса", size=7.5)
    _box(ax, 51.2, 62, 21.6, 10, "UC-2\nПродать билет\nс резервом места", size=7.5)
    _box(ax, 25.2, 46, 21.6, 10, "UC-3\nЗакрыть продажу\nпо отмене рейса", size=7.5)
    _box(ax, 77.1, 30, 19.8, 10, "UC-4\nОтметить\nпосадку", size=7.5)
    _box(ax, 51.2, 14, 21.6, 10, "UC-5\nВернуть билет\nдо отправления", size=7.5)

    _arrow(ax, 20.8, 83, 25.2, 83)
    _arrow(ax, 36, 78, 51.2, 72)
    _arrow(ax, 20.8, 67, 51.2, 67)
    _arrow(ax, 20.8, 51, 25.2, 51)
    _arrow(ax, 36, 46, 77.1, 40)
    _arrow(ax, 20.8, 35, 77.1, 35)
    _arrow(ax, 20.8, 19, 51.2, 19)

    ax.text(
        50,
        2.4,
        "E3 — отказ пассажира до отправления, расширение UC-2 / UC-5.  E6 — сводка за смену, вторая очередь, на цикл не вынесена как отдельный UC.",
        ha="center",
        va="center",
        fontsize=6.4,
    )
    fig.savefig(PROCESS, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def _actor(ax, x, y):
    ax.add_patch(Circle((x, y + 3.1), 1.15, fill=False, linewidth=1.0, edgecolor="black"))
    ax.plot([x, x], [y + 1.95, y + 0.15], color="black", lw=1.0)
    ax.plot([x - 1.7, x + 1.7], [y + 1.25, y + 1.25], color="black", lw=1.0)
    ax.plot([x, x - 1.35], [y + 0.15, y - 1.7], color="black", lw=1.0)
    ax.plot([x, x + 1.35], [y + 0.15, y - 1.7], color="black", lw=1.0)


def draw_usecase() -> None:
    fig, ax = plt.subplots(figsize=(6.55, 7.35), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")
    fig.patch.set_facecolor("white")
    ax.add_patch(Rectangle((26, 8), 48, 84, fill=False, linewidth=1.1, edgecolor="black"))
    ax.text(50, 88.5, "АИС железной дороги\n(пассажирский контур)", ha="center", va="center", fontsize=8)

    cases = [
        (78, "UC-1\nОпубликовать\nрасписание рейса"),
        (64, "UC-2\nПродать билет\nс резервом места"),
        (50, "UC-3\nЗакрыть продажу\nпо отмене рейса"),
        (36, "UC-4\nОтметить посадку"),
        (20, "UC-5\nВернуть билет"),
    ]
    for y, text in cases:
        ax.add_patch(Ellipse((50, y), 30, 11.2, facecolor="white", edgecolor="black", linewidth=0.9))
        ax.text(50, y, text, ha="center", va="center", fontsize=7.2)

    _actor(ax, 10, 70)
    ax.text(10, 65.2, "Диспетчер\nрасписания", ha="center", va="top", fontsize=8)
    _actor(ax, 90, 58)
    ax.text(90, 53.2, "Кассир", ha="center", va="top", fontsize=8)
    _actor(ax, 90, 28)
    ax.text(90, 23.2, "Контролёр\nпосадки", ha="center", va="top", fontsize=8)
    _actor(ax, 10, 28)
    ax.text(10, 23.2, "Бухгалтер\n(участник)", ha="center", va="top", fontsize=8)

    def link(x1, y1, x2, y2, dashed=False):
        ax.plot([x1, x2], [y1, y2], color="black", lw=0.8, linestyle=(0, (3, 2)) if dashed else "solid")

    link(12, 74, 35, 78)
    link(12, 68, 35, 50)
    link(88, 62, 65, 64)
    link(88, 56, 65, 20)
    link(88, 30, 65, 36)
    link(12, 30, 35, 20, dashed=True)
    ax.text(18, 34.5, "участник\nсверки", ha="center", va="bottom", fontsize=6.3)
    fig.savefig(USECASE, bbox_inches="tight", facecolor="white")
    plt.close(fig)


SPECS = [
    (
        "5.1 Опубликовать расписание рейса",
        "Таблица 3 — Эскизная спецификация UC-1",
        [
            ["Идентификатор", "UC-1"],
            ["Имя", "Опубликовать расписание рейса"],
            ["Основное действующее лицо", "Диспетчер расписания"],
            ["Участники", "Нет. Пассажир расписание не публикует."],
            [
                "Предусловия",
                "Диспетчер работает от своего имени. Станции маршрута и схема вагонов уже есть в справочниках. Рейс ещё не опубликован либо открыт на изменение.",
            ],
            [
                "Постусловия",
                "Рейс опубликован: дата, маршрут, состав вагонов и фонд мест доступны кассе. Статус рейса — «открыт к продаже». Одному рейсу соответствует один актуальный фонд мест.",
            ],
            [
                "Основной сценарий",
                "1. Диспетчер создаёт рейс и выбирает маршрут.\n"
                "2. Указывает дату, время отправления и схему состава.\n"
                "3. Система формирует фонд мест по схеме вагонов.\n"
                "4. Диспетчер публикует рейс.\n"
                "5. Система открывает рейс к продаже и делает его видимым кассе.",
            ],
            [
                "Расширения",
                "2a. Станции маршрута или схемы вагона нет. Система не публикует рейс, пока справочник не дополнят.\n"
                "4a. На ту же дату и маршрут уже есть открытый рейс с тем же номером. Система предупреждает о конфликте и не создаёт дубликат.\n"
                "4b. Нужно изменить уже опубликованный рейс без отмены. Диспетчер правит время или состав; система сохраняет новую версию и пересчитывает свободные места, не трогая уже проданные.",
            ],
        ],
    ),
    (
        "5.2 Продать билет с резервом места",
        "Таблица 4 — Эскизная спецификация UC-2",
        [
            ["Идентификатор", "UC-2"],
            ["Имя", "Продать билет с резервом места"],
            ["Основное действующее лицо", "Кассир"],
            ["Участники", "Нет. Пассажир передаёт запрос кассиру и в интерфейс не входит."],
            [
                "Предусловия",
                "Кассир работает от своего имени. Рейс опубликован и открыт к продаже. Тариф позиции есть в прайсе перевозчика.",
            ],
            [
                "Постусловия",
                "Оформлен билет на конкретное место. Место занято и недоступно другим продажам. Напечатан посадочный купон. Повторно продать то же место система не даёт.",
            ],
            [
                "Основной сценарий",
                "1. Кассир ищет рейс по станциям и дате.\n"
                "2. Система показывает свободные места и цену.\n"
                "3. Кассир выбирает место и данные пассажира.\n"
                "4. Кассир подтверждает продажу.\n"
                "5. Система резервирует место, записывает билет и печатает купон.",
            ],
            [
                "Расширения",
                "2a. Свободных мест нет. Система не предлагает занятые места.\n"
                "2b. Рейс отменён или закрыт к продаже. Система не начинает оформление.\n"
                "4a. Пока кассир выбирал место, его уже продали. Система заново показывает свободные места и не проводит устаревший выбор.\n"
                "4b. Пассажир отказывается до подтверждения. Кассир отменяет черновик; место остаётся свободным.\n"
                "5a. Цены в прайсе нет. Строка не проводится, кассир не вписывает сумму произвольно.",
            ],
        ],
    ),
    (
        "5.3 Закрыть продажу при отмене рейса",
        "Таблица 5 — Эскизная спецификация UC-3",
        [
            ["Идентификатор", "UC-3"],
            ["Имя", "Закрыть продажу при отмене рейса"],
            ["Основное действующее лицо", "Диспетчер расписания"],
            ["Участники", "Кассир: после закрытия не может оформить новый билет на этот рейс."],
            [
                "Предусловия",
                "Рейс существует и был открыт к продаже. Диспетчер работает от своего имени. Причина отмены зафиксирована.",
            ],
            [
                "Постусловия",
                "Статус рейса — «отменён». Новые продажи по нему невозможны. Ранее оформленные билеты помечены как билеты на отменённый рейс. Время от отмены до блокировки продажи укладывается в критерий БЦ-3.",
            ],
            [
                "Основной сценарий",
                "1. Диспетчер открывает рейс и выбирает отмену.\n"
                "2. Указывает причину и подтверждает.\n"
                "3. Система меняет статус рейса на «отменён».\n"
                "4. Система блокирует новые продажи и помечает действующие билеты.\n"
                "5. Кассы при следующем поиске не видят рейс как доступный к продаже.",
            ],
            [
                "Расширения",
                "1a. Рейс уже отменён. Система показывает текущий статус и не создаёт вторую отмену.\n"
                "2a. По рейсу идёт незавершённое оформление в кассе. Система отклоняет проведение продажи после смены статуса.\n"
                "4a. Нужен массовый возврат денег — это отдельная операция второй очереди; в UC-3 фиксируется только закрытие продажи и пометка билетов.",
            ],
        ],
    ),
    (
        "5.4 Отметить посадку",
        "Таблица 6 — Эскизная спецификация UC-4",
        [
            ["Идентификатор", "UC-4"],
            ["Имя", "Отметить посадку"],
            ["Основное действующее лицо", "Контролёр посадки"],
            ["Участники", "Нет."],
            [
                "Предусловия",
                "Есть действующий билет на рейс, который не отменён. Контролёр работает от своего имени. Посадка по этому билету ещё не отмечена.",
            ],
            [
                "Постусловия",
                "По билету зафиксирована посадка. Повторная отметка того же билета система отклоняет. Отменённый или возвращённый билет к посадке не принимается.",
            ],
            [
                "Основной сценарий",
                "1. Контролёр вводит номер билета или считывает купон.\n"
                "2. Система показывает рейс, место и статус билета.\n"
                "3. Контролёр подтверждает посадку.\n"
                "4. Система записывает отметку посадки.",
            ],
            [
                "Расширения",
                "2a. Билет не найден. Посадка не отмечается.\n"
                "2b. Рейс отменён или билет возвращён. Система запрещает отметку.\n"
                "3a. Посадка по билету уже есть. Система предупреждает о повторе и не создаёт вторую отметку.",
            ],
        ],
    ),
    (
        "5.5 Вернуть билет",
        "Таблица 7 — Эскизная спецификация UC-5",
        [
            ["Идентификатор", "UC-5"],
            ["Имя", "Вернуть билет"],
            ["Основное действующее лицо", "Кассир"],
            [
                "Участники",
                "Бухгалтер сверяет сумму возврата по сменной выручке. Пассажир передаёт купон кассиру.",
            ],
            [
                "Предусловия",
                "Есть проданный билет, посадка по нему не отмечена. Рейс ещё не отправлен либо действует правило возврата до отправления. Кассир работает от своего имени.",
            ],
            [
                "Постусловия",
                "Билет помечен как возвращённый. Место снова свободно. Сумма возврата доступна для сверки. Повторный возврат того же билета система отклоняет.",
            ],
            [
                "Основной сценарий",
                "1. Кассир находит билет по номеру.\n"
                "2. Система показывает рейс, место, сумму и допустимость возврата.\n"
                "3. Кассир подтверждает возврат.\n"
                "4. Система гасит билет, освобождает место и фиксирует сумму.\n"
                "5. Бухгалтер при сверке смены видит операцию возврата.",
            ],
            [
                "Расширения",
                "2a. Посадка уже отмечена. Возврат не проводится.\n"
                "2b. Срок возврата истёк по правилу перевозчика. Система отказывает.\n"
                "3a. Билет уже возвращён. Система не проводит повтор.\n"
                "4a. Рейс отменён диспетчером. Возврат допускается как отдельный случай; место остаётся закрытым вместе с рейсом, билет гасится.",
            ],
        ],
    ),
]


def build_body(doc: Document, pages: dict[str, str]) -> None:
    add_heading(doc, "СОДЕРЖАНИЕ", level=0)
    toc = [
        ("ВВЕДЕНИЕ", False),
        ("1 Задание", False),
        ("2 Разбор бизнес-процесса", False),
        ("2.1 Ход процесса", True),
        ("2.2 Схема процесса", True),
        ("3 Список внешних событий", False),
        ("4 Действующие лица и варианты использования", False),
        ("4.1 Вывод из процесса и событий", True),
        ("4.2 Диаграмма вариантов использования", True),
        ("5 Эскизные спецификации", False),
        ("5.1 Опубликовать расписание рейса", True),
        ("5.2 Продать билет с резервом места", True),
        ("5.3 Закрыть продажу при отмене рейса", True),
        ("5.4 Отметить посадку", True),
        ("5.5 Вернуть билет", True),
        ("ЗАКЛЮЧЕНИЕ", False),
    ]
    for title, sub in toc:
        add_toc_line(doc, title, pages.get(title, "…"), sub=sub)

    add_heading(doc, "ВВЕДЕНИЕ", new_page=True, level=0)
    add_body(
        doc,
        "Работа продолжает документ о концепции и границах автоматизированной "
        "информационной системы железной дороги. Там зафиксированы разрозненный учёт "
        "расписания и мест, цели БЦ-1–БЦ-3 и функции FEAT1–FEAT11. Здесь по тому же "
        "пассажирскому контуру выделяются действующие лица и варианты использования.",
    )
    add_body(
        doc,
        "Выявление сделано двумя техниками разбора бизнес-процесса: схемой процесса "
        "и списком внешних событий. Из шагов схемы и из событий выведены действующие "
        "лица и варианты использования. Эскизная спецификация составлена для публикации "
        "расписания, продажи билета, закрытия продажи при отмене рейса, отметки посадки "
        "и возврата билета.",
    )

    add_heading(doc, "1 Задание", new_page=True)
    add_body(
        doc,
        "Нужно смоделировать бизнес-процесс продажи билетов и ведения расписания "
        "пассажирского перевозчика, по схеме процесса и по внешним событиям определить "
        "действующих лиц и варианты использования и составить на них эскизные "
        "спецификации. Спецификация краткая: идентификатор, имя, основное действующее "
        "лицо, участники, предусловия, постусловия, основной сценарий и расширения. "
        "Образец уровня детализации — Use-Case Model Survey (RU_FS_UCS_Survey).",
    )

    add_heading(doc, "2 Разбор бизнес-процесса", new_page=True)
    add_heading(doc, "2.1 Ход процесса", level=2)
    add_body(
        doc,
        "Пилот — один участок сети. Фонд мест рейса един для кассы и диспетчера. "
        "Сейчас расписание ведут в таблицах, занятость мест — в журнале кассы, поэтому "
        "одно место продают дважды, а отмена рейса доходит до кассы с опозданием. "
        "В целевом процессе каждое действие оставляет документ в одной системе.",
    )
    add_body(
        doc,
        "Диспетчер публикует рейс: маршрут, время, схему вагонов и фонд мест. "
        "Пассажир просит билет у кассира. Кассир ищет открытый рейс, выбирает свободное "
        "место и оформляет билет; место сразу становится занятым, печатается купон. "
        "Если рейс отменяют, диспетчер закрывает продажу, и касса больше не оформляет "
        "билеты на этот рейс. Перед отправлением контролёр отмечает посадку по "
        "действующему билету. До отправления кассир может вернуть билет и освободить "
        "место; бухгалтер сверяет возвраты со сменной выручкой. Пассажир с системой "
        "напрямую не работает.",
    )
    add_heading(doc, "2.2 Схема процесса", level=2)
    add_body(
        doc,
        "Схема на рисунке 1 разложена по дорожкам ролей. Левая дорожка — внешнее "
        "событие, которое запускает шаг. Стрелки показывают основной контур: UC-1, "
        "UC-2, при необходимости UC-3, затем UC-4. Возврат билета (UC-5) и сводка "
        "за смену на цикл продажи не навязаны: возврат запускается отдельным запросом "
        "пассажира, сводка относится к FEAT10 второй очереди.",
    )
    add_figure(doc, PROCESS, "Рисунок 1 — Схема бизнес-процесса продажи билетов АИС железной дороги")

    add_heading(doc, "3 Список внешних событий", new_page=True)
    add_body(
        doc,
        "Внешнее событие — то, что происходит за границей системы и требует отклика. "
        "Список в таблице 1 собран по шагам раздела 2. Инициатор события и основное "
        "действующее лицо варианта могут не совпадать: пассажир просит билет, а продажу "
        "в системе проводит кассир.",
    )
    add_table_caption(doc, "Таблица 1 — Внешние события")
    add_table(
        doc,
        ["Код", "Событие", "Инициатор", "Отклик", "Вариант"],
        [
            ["E1", "Нужно опубликовать или изменить рейс", "Диспетчер расписания", "Рейс и фонд мест открыты к продаже", "UC-1"],
            ["E2", "Пассажир запросил билет", "Пассажир через кассира", "Продажа с резервом свободного места", "UC-2"],
            ["E3", "Пассажир отказался до отправления", "Пассажир через кассира", "Снятие продажи / возврат", "расширение UC-2, UC-5"],
            ["E4", "Рейс отменён", "Диспетчер расписания", "Блокировка новых продаж", "UC-3"],
            ["E5", "Началась посадка", "Контролёр посадки", "Отметка посадки по билету", "UC-4"],
            ["E6", "Нужна сводка продаж за смену", "Руководитель / бухгалтер", "Сводка по проведённым билетам", "FEAT10, вторая очередь"],
            ["E7", "Пассажир сдал билет до отправления", "Пассажир через кассира", "Гашение билета, место свободно", "UC-5"],
        ],
        [1.5, 4.4, 3.6, 4.0, 3.0],
    )
    add_body(
        doc,
        "Событие E6 видно из потребности в сменной сводке, но функция отнесена ко "
        "второй очереди документа о концепции и границах. Отдельный вариант для неё "
        "в этой работе не специфицируется. Операционный контур, который меняет фонд "
        "мест и статус продажи, закрывают E1–E5 и E7.",
        space_before=6,
    )

    add_heading(doc, "4 Действующие лица и варианты использования", new_page=True)
    add_heading(doc, "4.1 Вывод из процесса и событий", level=2)
    add_body(
        doc,
        "Действующее лицо — роль, которая работает с системой, чтобы получить результат "
        "шага. Пассажир запускает E2, E3 и E7, но документы вводит кассир, поэтому "
        "пассажир действующим лицом системы в первой версии не является. Сопоставление "
        "шагов и событий приведено в таблице 2.",
    )
    add_table_caption(doc, "Таблица 2 — Вывод действующих лиц и вариантов использования")
    add_table(
        doc,
        ["Источник", "Что видно", "Действующее лицо", "Вариант"],
        [
            ["Шаг публикации, E1", "Рейс и фонд мест открыты кассе", "Диспетчер расписания", "UC-1"],
            ["Шаг продажи, E2", "Билет только на свободное место", "Кассир", "UC-2"],
            ["E3", "Отказ до подтверждения / до отправления", "Кассир", "расширение UC-2, UC-5"],
            ["Шаг отмены, E4", "Новые продажи по рейсу закрыты", "Диспетчер расписания", "UC-3"],
            ["Шаг посадки, E5", "Отметка по действующему билету", "Контролёр посадки", "UC-4"],
            ["Участник UC-5", "Сверка суммы возврата, свой вариант не начинает", "Бухгалтер", "участник UC-5"],
            ["Шаг возврата, E7", "Билет погашен, место свободно", "Кассир", "UC-5"],
            ["E6", "Сводка за смену, FEAT10", "Руководитель", "вторая очередь"],
        ],
        [3.6, 5.2, 4.2, 3.5],
    )
    add_body(
        doc,
        "Администратор из документа о концепции и границах ведёт учётные записи и "
        "справочники станций и схем вагонов. В этом процессе он появляется только в "
        "расширении UC-1, если справочника ещё нет. Отдельный вариант ведения "
        "справочника из цикла продажи не выводится. UC-1–UC-3 и печать билета в UC-2 "
        "закрывают первую очередь (расписание, фонд мест, продажа, запрет продажи "
        "занятого и отменённого рейса). UC-4 и UC-5 соответствуют FEAT11 и FEAT9 "
        "второй очереди, но лица и варианты уже видны из посадки и возврата, поэтому "
        "спецификации составлены.",
        space_before=6,
    )
    add_heading(doc, "4.2 Диаграмма вариантов использования", level=2)
    add_body(
        doc,
        "На рисунке 2 сплошная линия связывает вариант с основным действующим лицом. "
        "Пунктир к UC-5 показывает бухгалтера как участника: он сверяет возврат и не "
        "начинает вариант. Пассажир и руководитель на диаграмму не вынесены: пассажир "
        "не входит в систему, сводка руководителя отложена вместе с E6.",
    )
    add_figure(doc, USECASE, "Рисунок 2 — Диаграмма вариантов использования")

    add_heading(doc, "5 Эскизные спецификации", new_page=True)
    add_body(
        doc,
        "Ниже пять кратких спецификаций. В них сохранены правила пилота: один участок "
        "сети, единый фонд мест рейса, нельзя продать занятое место, отмена рейса "
        "закрывает новую продажу, посадка отмечается только по действующему билету.",
    )
    for title, caption, rows in SPECS:
        add_heading(doc, title, level=2)
        add_table_caption(doc, caption)
        add_table(doc, ["Поле", "Содержание"], rows, [4.6, 11.9])

    add_heading(doc, "ЗАКЛЮЧЕНИЕ", new_page=True, level=0)
    add_body(
        doc,
        "По схеме процесса и по списку внешних событий E1–E7 для АИС железной дороги "
        "выделены действующие лица: диспетчер расписания, кассир, контролёр посадки и "
        "бухгалтер как участник возврата. Пассажир остаётся вне системы первой версии. "
        "Составлены эскизные спецификации UC-1 «Опубликовать расписание рейса», UC-2 "
        "«Продать билет с резервом места», UC-3 «Закрыть продажу при отмене рейса», "
        "UC-4 «Отметить посадку» и UC-5 «Вернуть билет». Сводка за смену (E6, FEAT10) "
        "отмечена как событие второй очереди и отдельно не специфицирована.",
    )


HEADINGS = [
    "ВВЕДЕНИЕ",
    "1 Задание",
    "2 Разбор бизнес-процесса",
    "2.1 Ход процесса",
    "2.2 Схема процесса",
    "3 Список внешних событий",
    "4 Действующие лица и варианты использования",
    "4.1 Вывод из процесса и событий",
    "4.2 Диаграмма вариантов использования",
    "5 Эскизные спецификации",
    "5.1 Опубликовать расписание рейса",
    "5.2 Продать билет с резервом места",
    "5.3 Закрыть продажу при отмене рейса",
    "5.4 Отметить посадку",
    "5.5 Вернуть билет",
    "ЗАКЛЮЧЕНИЕ",
]

FALLBACK_PAGES = {
    "ВВЕДЕНИЕ": "3",
    "1 Задание": "4",
    "2 Разбор бизнес-процесса": "5",
    "2.1 Ход процесса": "5",
    "2.2 Схема процесса": "5",
    "3 Список внешних событий": "7",
    "4 Действующие лица и варианты использования": "8",
    "4.1 Вывод из процесса и событий": "8",
    "4.2 Диаграмма вариантов использования": "9",
    "5 Эскизные спецификации": "10",
    "5.1 Опубликовать расписание рейса": "10",
    "5.2 Продать билет с резервом места": "11",
    "5.3 Закрыть продажу при отмене рейса": "12",
    "5.4 Отметить посадку": "12",
    "5.5 Вернуть билет": "13",
    "ЗАКЛЮЧЕНИЕ": "15",
}


def save_doc(doc: Document) -> None:
    doc.save(str(OUT))
    strip_unused_images(OUT)


def measure_pages() -> dict[str, str]:
    import subprocess
    import tempfile

    listing = Path(tempfile.gettempdir()) / "lr2_pages.txt"
    if platform.system() == "Darwin":
        raise RuntimeError("COM Word недоступен на macOS в этой сборке")
    script = Path(tempfile.gettempdir()) / "lr2_pages.ps1"
    script.write_text(
        "\n".join(
            [
                "$ErrorActionPreference = 'Stop'",
                "$word = New-Object -ComObject Word.Application",
                "$word.Visible = $false",
                "$word.DisplayAlerts = 0",
                f"$doc = $word.Documents.Open('{OUT}', $false, $true)",
                "$doc.Repaginate()",
                "$lines = New-Object System.Collections.Generic.List[string]",
                "foreach ($p in $doc.Paragraphs) {",
                '  $t = $p.Range.Text.Replace("`r","").Replace("`a","").Trim()',
                '  if ($t.Length -gt 0 -and -not $t.Contains("`t")) {',
                "    $page = [int]$p.Range.Information(3)",
                '    $lines.Add("$page`t$t")',
                "  }",
                "}",
                "$doc.Close(0)",
                "$word.Quit()",
                f"[System.IO.File]::WriteAllLines('{listing}', $lines, [System.Text.UTF8Encoding]::new($false))",
            ]
        ),
        encoding="utf-8-sig",
    )
    subprocess.run(["powershell", "-NoProfile", "-File", str(script)], check=True)
    found: dict[str, str] = {}
    wanted = set(HEADINGS)
    for line in listing.read_text(encoding="utf-8").splitlines():
        if "\t" not in line:
            continue
        page, text = line.split("\t", 1)
        text = text.strip()
        if text in wanted and text not in found:
            found[text] = page.strip()
    return found


def main() -> None:
    draw_process()
    draw_usecase()
    doc = clone_title()
    build_body(doc, {title: "…" for title in HEADINGS})
    save_doc(doc)
    try:
        pages = measure_pages()
        missing = [title for title in HEADINGS if title not in pages]
        if missing:
            raise RuntimeError(f"Не найдены страницы: {missing}")
    except Exception as exc:
        print(f"Предупреждение: автонумерация страниц не удалась ({exc}); использую оценку.")
        pages = FALLBACK_PAGES
    doc = clone_title()
    build_body(doc, pages)
    save_doc(doc)
    print(pages)
    print(OUT)


if __name__ == "__main__":
    main()
