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
SALE = FIG / "sale.png"
CANCEL = FIG / "cancel.png"
TOPIC = "Спецификация вариантов использования"
PAGE_LIMIT = 15

font_manager.fontManager.addfont(r"C:\Windows\Fonts\times.ttf")
font_manager.fontManager.addfont(r"C:\Windows\Fonts\timesbd.ttf")
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
    fig, ax = plt.subplots(figsize=(7.3, 6.15), dpi=180)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")
    fig.patch.set_facecolor("white")
    ax.add_patch(Rectangle((24, 5), 52, 90, fill=False, linewidth=1.1, edgecolor="black"))
    ax.text(50, 90.2, "АИС железной дороги", ha="center", va="center", fontsize=9)

    cases = [
        (80, "UC-1\nОпубликовать рейс"),
        (67, "UC-2\nПродать билет"),
        (54, "UC-3\nЗакрыть продажу"),
        (41, "UC-4\nОтметить посадку"),
        (28, "UC-5\nВернуть билет"),
        (15, "UC-6\nСправочники\nстанций и схем"),
    ]
    for y, text in cases:
        ax.add_patch(Ellipse((50, y), 34, 10.2, facecolor="white", edgecolor="black", linewidth=0.9))
        ax.text(50, y, text, ha="center", va="center", fontsize=8)

    _actor(ax, 10, 72)
    ax.text(10, 67.2, "Диспетчер\nрасписания", ha="center", va="top", fontsize=8)
    _actor(ax, 10, 28)
    ax.text(10, 23.2, "Кассир", ha="center", va="top", fontsize=8)
    _actor(ax, 90, 41)
    ax.text(90, 36.2, "Контролёр\nпосадки", ha="center", va="top", fontsize=8)
    _actor(ax, 90, 15)
    ax.text(90, 10.2, "Администратор", ha="center", va="top", fontsize=8)

    def link(x1, y1, x2, y2):
        ax.plot([x1, x2], [y1, y2], color="black", lw=0.8)

    link(12, 75, 33, 80)
    link(12, 71, 33, 54)
    link(12, 31, 33, 67)
    link(12, 28, 33, 28)
    link(88, 44, 67, 41)
    link(88, 18, 67, 15)
    fig.savefig(USECASE, bbox_inches="tight", facecolor="white", pad_inches=0.12)
    plt.close(fig)


def draw_flow(path: Path, steps: list[str], side: str, side_at: int) -> None:
    fig, ax = plt.subplots(figsize=(6.6, 4.35), dpi=180)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")
    fig.patch.set_facecolor("white")
    n = len(steps)
    h = 12
    gap = 5.5
    block = n * h + (n - 1) * gap
    top = 50 + block / 2
    xs, w = 8, 50
    centers = []
    y_top = top
    for text in steps:
        y = y_top - h
        _box(ax, xs, y, w, h, text, size=8.2)
        centers.append(y + h / 2)
        y_top = y - gap
    for i in range(n - 1):
        _arrow(ax, xs + w / 2, centers[i] - h / 2, xs + w / 2, centers[i + 1] + h / 2)
    sy = centers[side_at] - 8
    _box(ax, 64, sy, 32, 16, side, size=7.4)
    _arrow(ax, xs + w, centers[side_at], 64, sy + 8)
    fig.savefig(path, bbox_inches="tight", facecolor="white", pad_inches=0.12)
    plt.close(fig)


def draw_figures() -> None:
    draw_usecase()
    draw_flow(
        SALE,
        [
            "Запрос пассажира у окна",
            "Поиск открытого рейса",
            "Свободное место и цена из прайса",
            "Факт оплаты и резерв места",
            "Печать посадочного купона",
        ],
        "Отказ: нет мест,\nрейс закрыт, место\nзанято или нет цены",
        2,
    )
    draw_flow(
        CANCEL,
        [
            "Диспетчер открывает рейс",
            "Причина и подтверждение",
            "Статус «отменён»",
            "Новые продажи закрыты",
            "Билеты помечены",
        ],
        "Черновик в кассе\nне проводится.\nВыплата по всем\nбилетам — позже",
        3,
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
            "1. Диспетчер указывает номер рейса, дату и маршрут из справочника. "
            "2. Задаёт время отправления и прибытия и схему вагонов. "
            "3. Система создаёт места по схеме, без повторов номеров. "
            "4. Диспетчер подтверждает публикацию. "
            "5. Статус становится «открыт к продаже», касса видит свободные места."
        ),
        "alts": (
            "Нет станции или схемы — публикация останавливается, запись вносит администратор в UC-6; "
            "тот же номер уже открыт или рейс отменён — второй фонд не создаётся; "
            "правка открытого рейса сохраняет проданные места; выход до подтверждения продажу не открывает."
        ),
        "pre": "Диспетчер работает от своего имени, нужные записи справочника есть, рейс ещё не опубликован либо открыт на правку и не отменён.",
        "post": "При успехе рейс открыт к продаже и фонд создан; при неудаче продажа не открыта и второй фонд не создан.",
        "links": "Вариант начинает диспетчер. Результат нужен продаже. Нехватку справочника закрывает UC-6, отмену выполняет UC-3; публикация их не вызывает.",
        "diagram": "Вариант показан на рисунке 1.",
        "ok": "Рейс опубликован и открыт к продаже.",
        "bad": "Публикация остановлена: нет записи в справочнике, конфликт номера, рейс уже отменён или выход без подтверждения.",
        "figure": None,
    },
    {
        "heading": "2.3 Продать билет с резервом места",
        "name": "UC-2 «Продать билет с резервом места». ",
        "about": (
            "Кассир по запросу пассажира находит открытый рейс, занимает свободное место по цене из прайса и печатает купон. "
            "Пассажир стоит у окна и в систему не входит. Оформление удерживается в пороге БЦ-1, одно место не закрепляется за двумя действующими билетами (БЦ-2)."
        ),
        "flow": (
            "1. Кассир задаёт станции и дату. "
            "2. Система показывает открытые рейсы, свободные места и цену из прайса. "
            "3. Кассир выбирает место и вносит фамилию и имя. "
            "4. Фиксирует факт оплаты и подтверждает продажу. "
            "5. Система резервирует место и печатает купон."
        ),
        "alts": (
            "Нет мест, рейс закрыт или отменён, место уже занято другой кассой, нет цены в прайсе "
            "или пассажир отказался до подтверждения — билет не создаётся и место этим оформлением не занимается."
        ),
        "pre": "Кассир работает от своего имени, поиск идёт по общему фонду мест, цена берётся из прайса перевозчика.",
        "post": "При успехе билет проведён, место занято и купон напечатан; при неудаче билет не создан.",
        "links": "Вариант начинает кассир, нужен рейс UC-1 со статусом «открыт к продаже». Возврат после продажи — точка расширения UC-5. Платёжный шлюз не используется.",
        "diagram": "Вариант показан на рисунке 1. Поток продажи — на рисунке 2.",
        "ok": "Билет продан, место занято, купон напечатан.",
        "bad": "Продажа не проведена: нет мест, рейс закрыт, место занято, нет цены или отказ до подтверждения.",
        "figure": (SALE, "Рисунок 2 — Поток продажи билета", 12.2),
    },
    {
        "heading": "2.4 Закрыть продажу при отмене рейса",
        "name": "UC-3 «Закрыть продажу при отмене рейса». ",
        "about": (
            "Диспетчер отменяет рейс. Система ставит статус «отменён», блокирует новые продажи и помечает уже оформленные билеты. "
            "Выплату по всем билетам рейса этот вариант не проводит. Срок блокировки задан порогом БЦ-3."
        ),
        "flow": (
            "1. Диспетчер открывает рейс, который был открыт к продаже. "
            "2. Указывает причину отмены. "
            "3. Подтверждает отмену. "
            "4. Система ставит статус «отменён» и запрещает новые продажи. "
            "5. Действующие билеты помечаются, история не стирается."
        ),
        "alts": (
            "Рейс уже отменён — вторая отмена не создаётся; незавершённый черновик в кассе не проводится; "
            "массовая выплата остаётся во второй очереди, отдельный билет гасит UC-5."
        ),
        "pre": "Рейс существует и был открыт к продаже, диспетчер работает от своего имени, причина указана до подтверждения.",
        "post": "При успехе новые продажи невозможны и билеты помечены; при неудаче статус не меняется.",
        "links": "Вариант начинает диспетчер. Он меняет результат UC-1 и не включает продажу или возврат. Помеченный билет можно погасить в UC-5, место в продажу не возвращается.",
        "diagram": "Вариант показан на рисунке 1. Поток отмены — на рисунке 3.",
        "ok": "Продажа закрыта, билеты помечены.",
        "bad": "Отмена не подтверждена, повторная отмена не создана или черновик продажи не проведён.",
        "figure": (CANCEL, "Рисунок 3 — Поток закрытия продажи при отмене рейса", 12.2),
    },
    {
        "heading": "2.5 Отметить посадку",
        "name": "UC-4 «Отметить посадку». ",
        "about": (
            "Контролёр по номеру билета или по купону проверяет, что билет действующий, и фиксирует одну отметку. "
            "Повтор, возвращённый билет и билет отменённого рейса к посадке не принимаются. Цену, место и расписание контролёр здесь не меняет."
        ),
        "flow": (
            "1. Контролёр открывает рабочее место. "
            "2. Вводит номер билета или считывает купон. "
            "3. Система показывает рейс, место, пассажира и статус. "
            "4. Контролёр подтверждает посадку. "
            "5. Система записывает одну отметку."
        ),
        "alts": "Билет не найден, рейс отменён, билет возвращён, посадка уже отмечена или контролёр не подтвердил — новая отметка не создаётся.",
        "pre": "Контролёр работает от своего имени, билет действующий, рейс не отменён, посадка по нему ещё не отмечена.",
        "post": "При успехе по билету одна отметка посадки; при неудаче отметки нет.",
        "links": "Вариант начинает контролёр. Нужен билет UC-2, не погашенный UC-5. После отметки возврат закрыт. Отдельной точки расширения нет.",
        "diagram": "Вариант показан на рисунке 1.",
        "ok": "Посадка отмечена один раз.",
        "bad": "Отметка не записана: билет не найден, рейс отменён, билет возвращён, повтор или отказ от подтверждения.",
        "figure": None,
    },
    {
        "heading": "2.6 Вернуть билет",
        "name": "UC-5 «Вернуть билет». ",
        "about": (
            "Кассир до отправления и до отметки посадки гасит проданный билет. Если рейс не отменён, место снова свободно. "
            "Сумму берёт система из проведённого билета, бухгалтер только видит её при сверке."
        ),
        "flow": (
            "1. Кассир находит билет по номеру. "
            "2. Система показывает рейс, место, сумму и допустимость возврата. "
            "3. Кассир подтверждает возврат. "
            "4. Система гасит билет и фиксирует сумму. "
            "5. Если рейс не отменён, место снова свободно."
        ),
        "alts": (
            "Посадка уже отмечена, срок истёк или билет уже возвращён — отказ; билет не найден — возврат не начинается; "
            "рейс отменён — билет гасится, а место остаётся закрытым."
        ),
        "pre": "Кассир работает от своего имени, билет продан, посадка не отмечена, правило возврата до отправления ещё выполняется.",
        "post": "При успехе билет возвращён и повтор невозможен; место свободно, только если рейс не отменён.",
        "links": "Вариант начинает кассир: это точка расширения уже проданного билета UC-2. Для отменённого рейса используется пометка UC-3. Массовая выплата из UC-5 не вызывается.",
        "diagram": "Вариант показан на рисунке 1.",
        "ok": "Билет возвращён; место свободно, если рейс не отменён.",
        "bad": "Возврат не проведён: посадка отмечена, срок истёк, билет уже погашен или не найден.",
        "figure": None,
    },
    {
        "heading": "2.7 Вести справочники станций и схем вагонов",
        "name": "UC-6 «Вести справочники станций и схем вагонов». ",
        "about": (
            "Администратор ведёт единые справочники пилотного участка: станции, маршруты и схемы пассажирских вагонов. "
            "Диспетчер при публикации выбирает только эти записи, касса второй список станций не держит. "
            "Грузовой парк в справочник не входит."
        ),
        "flow": (
            "1. Администратор открывает справочник станций, маршрутов или схем. "
            "2. Создаёт или правит запись: код станции, список станций маршрута либо тип вагона и номера мест. "
            "3. Система проверяет уникальность кода и номеров и наличие станций маршрута. "
            "4. Администратор подтверждает сохранение. "
            "5. Запись доступна при следующей публикации рейса."
        ),
        "alts": (
            "Повтор кода или номера места и неизвестная станция в маршруте — запись не сохраняется; "
            "удаление того, что использует открытый рейс, запрещено; выход без сохранения справочник не меняет."
        ),
        "pre": "Администратор работает от своего имени, справочники пилотного участка открыты для правки.",
        "post": "При успехе запись доступна для публикации; при неудаче справочник прежний, проданные места открытого рейса не теряются.",
        "links": "Вариант начинает администратор. UC-1 только читает готовые записи и UC-6 не включает. Учётные записи и резервное копирование сюда не входят.",
        "diagram": "Вариант показан на рисунке 1.",
        "ok": "Запись станции, маршрута или схемы вагона сохранена.",
        "bad": "Справочник не изменён: дубликат, неизвестная станция, запись занята открытым рейсом или выход без сохранения.",
        "figure": None,
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
        "Здесь по тому же пассажирскому контуру составлены спецификации шести вариантов: публикация рейса, продажа билета, "
        "закрытие продажи при отмене, посадка, возврат и справочники станций и схем вагонов.",
    )
    add_body(
        doc,
        "Пассажир покупает билет через кассира и в систему не входит. Грузовые перевозки, управление движением и публичный "
        "интернет-магазин билетов остаются вне границ первой версии. Оформление билета от запроса до печати занимает не более "
        "3 минут в 90 % случаев за три месяца после ввода (БЦ-1). Подтверждённых инцидентов двойной продажи одного места — "
        "не более одного в месяц (БЦ-2). После отмены рейса новые продажи невозможны не позднее чем через 5 минут в 95 % случаев (БЦ-3).",
    )

    add_heading(doc, "1 Задание", new_page=True)
    add_body(
        doc,
        "Нужно составить спецификацию вариантов использования: имя, краткое описание, основной и альтернативные потоки, "
        "предусловия, постусловия, связи с другими вариантами, диаграмму и сценарии успеха и неудачи. "
        "Берутся UC-1–UC-5 из контура продажи и расписания и UC-6 «Вести справочники станций и схем вагонов». "
        "Сводка продаж за смену остаётся во второй очереди и отдельным вариантом не задаётся.",
    )

    add_heading(doc, "2 Варианты использования", new_page=True)
    add_body(
        doc,
        "Шесть вариантов одной системы показаны на рисунке 1. Поток продажи билета раскрыт на рисунке 2, "
        "поток закрытия продажи при отмене — на рисунке 3. Шаги остальных вариантов заданы текстом спецификаций.",
    )
    add_heading(doc, "2.1 Общая диаграмма", level=2)
    add_body(
        doc,
        "На рисунке 1 граница системы одна. Диспетчер начинает публикацию и отмену, кассир — продажу и возврат, "
        "контролёр — посадку, администратор — справочники. Пассажир действующим лицом не является.",
    )
    add_figure(doc, USECASE, "Рисунок 1 — Диаграмма вариантов использования АИС железной дороги", 15.0)
    for case in CASES:
        write_case(doc, case)

    add_heading(doc, "ЗАКЛЮЧЕНИЕ", new_page=True, level=0)
    add_body(
        doc,
        "Составлены спецификации шести вариантов автоматизированной информационной системы железной дороги. "
        "Публикация собирает один фонд мест, продажа занимает только свободное место, отмена закрывает новые продажи, "
        "посадка и возврат смотрят один и тот же статус билета, справочники не дают кассе и диспетчеру вести второй список станций.",
    )
    add_body(
        doc,
        "Пороги БЦ-1, БЦ-2 и БЦ-3 из введения спецификациями не ослаблены. Пассажир остаётся вне системы. "
        "Грузовые перевозки и публичная онлайн-продажа в границы не вошли. Сводка за смену по-прежнему относится ко второй очереди.",
    )


def save_doc(doc: Document) -> None:
    doc.save(str(OUT))
    strip_unused_images(OUT)


def measure_pages() -> tuple[dict[str, str], int]:
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
