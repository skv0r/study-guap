#!/usr/bin/env python3
"""Отчёт ЛР1 по СТЗ на производстве: бланк ГУАП, ГОСТ 7.32 / ГОСТ 2.105."""

from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Mm, Pt, RGBColor

BASE = Path(__file__).resolve().parent
BLANK = BASE / "guap_blanks" / "lab.docx"
OUT = BASE / "ЛР1_Отчет_Буренков_СТЗ.docx"
PDF = BASE / "ЛР1_Отчет_Буренков_СТЗ.pdf"
FIG = BASE / "figures" / "scheme.png"

TOC_DEF = [
    ("ВВЕДЕНИЕ", "intro", 0),
    ("1 Постановка задачи", "s1", 0),
    ("2 Паспорт подсистем СТЗ", "s2", 0),
    ("3 Классификационный профиль системы", "s3", 0),
    ("4 Анализ типов данных", "s4", 0),
    ("5 Расчёт параметров информационного процесса", "s5", 0),
    ("6 Выбор интерфейса и проверка временного цикла", "s6", 0),
    ("7 Структурная схема системы", "s7", 0),
    ("ЗАКЛЮЧЕНИЕ", "conc", 0),
    ("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "src", 0),
]

SEARCH = {
    "intro": "ВВЕДЕНИЕ",
    "s1": "1 Постановка задачи",
    "s2": "2 Паспорт подсистем СТЗ",
    "s3": "3 Классификационный профиль системы",
    "s4": "4 Анализ типов данных",
    "s5": "5 Расчёт параметров информационного процесса",
    "s6": "6 Выбор интерфейса и проверка временного цикла",
    "s7": "7 Структурная схема системы",
    "conc": "ЗАКЛЮЧЕНИЕ",
    "src": "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
}


def set_run_font(run, size=14, bold=False, name="Times New Roman", italic=False):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor(0, 0, 0)


def format_paragraph(
    p,
    *,
    first_indent=True,
    align="justify",
    space_after=0,
    space_before=0,
    left_indent=0,
    line_spacing=1.5,
):
    pf = p.paragraph_format
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
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }[align]


def add_body(doc, text):
    p = doc.add_paragraph()
    format_paragraph(p)
    set_run_font(p.add_run(text))
    return p


def add_section_h(doc, text, *, new_page=True):
    if new_page:
        doc.add_page_break()
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=True, align="left", space_before=0, space_after=12)
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text), bold=True)


def add_struct(doc, text, *, new_page=True):
    if new_page:
        doc.add_page_break()
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="center", space_after=18)
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text), bold=True)


def add_toc_line(doc, title, page, *, indent=0):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="left", left_indent=0.75 * indent)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(16.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    set_run_font(p.add_run(title))
    p.add_run("\t")
    set_run_font(p.add_run(str(page)))


def add_caption_table(doc, text):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="left", space_before=10, space_after=4)
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text))


def add_figure(doc, path, caption, w=16.0):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="center", space_before=10)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(path), width=Cm(w))
    c = doc.add_paragraph()
    format_paragraph(c, first_indent=False, align="center", space_before=6, space_after=10)
    set_run_font(c.add_run(caption))


def _set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        tcBorders.append(el)
    tcPr.append(tcBorders)


def set_cell_shading(cell, fill="D9D9D9"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def add_table(doc, headers, rows, widths_cm=None, size=10):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    usable = Cm(16.5)
    if widths_cm:
        total = sum(widths_cm)
        widths = [int(usable * (w / total)) for w in widths_cm]
    else:
        widths = [int(usable / len(headers))] * len(headers)
    grid = t._tbl.tblGrid
    for i, child in enumerate(list(grid)):
        if i < len(widths):
            child.set("w", str(widths[i]))
    for row in t.rows:
        for cell in row.cells:
            _set_cell_border(cell)
            for i, w in enumerate(widths):
                if cell._tc is row.cells[i]._tc:
                    cell.width = Emu(w)
    for j, h in enumerate(headers):
        cell = t.rows[0].cells[j]
        cell.text = ""
        set_cell_shading(cell)
        p = cell.paragraphs[0]
        format_paragraph(p, first_indent=False, align="center", line_spacing=1.15)
        set_run_font(p.add_run(h), bold=True, size=size)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            cell = t.rows[i + 1].cells[j]
            cell.text = ""
            p = cell.paragraphs[0]
            align = "center" if j == 0 else "left"
            format_paragraph(p, first_indent=False, align=align, line_spacing=1.15)
            set_run_font(p.add_run(v), size=size)
    spacer = doc.add_paragraph()
    format_paragraph(spacer, first_indent=False, line_spacing=1.0, space_after=6)
    return t


def add_page_field(paragraph):
    run = paragraph.add_run()
    for kind, val in [("begin", None), ("instr", " PAGE "), ("end", None)]:
        if kind == "instr":
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = val
        else:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
        run._r.append(el)
    set_run_font(run)


def put_cell(cell, text, size=14):
    p = cell.paragraphs[0]
    for r in list(p.runs):
        r.text = ""
    if not p.runs:
        run = p.add_run(text)
    else:
        run = p.runs[0]
        run.text = text
    set_run_font(run, size=size)


def fill_title(doc: Document):
    for p in doc.paragraphs:
        if p.text.startswith("КАФЕДРА"):
            if len(p.runs) >= 2 and p.runs[1].text.strip("_") == "":
                p.runs[1].text = "42"
            else:
                p.clear()
                r = p.add_run("КАФЕДРА № 42")
                set_run_font(r, size=14)
        elif "20__" in p.text:
            for r in p.runs:
                if "20__" in r.text:
                    r.text = r.text.replace("20__", "2026")

    t_prep, t_work, t_stud = doc.tables
    put_cell(t_prep.rows[0].cells[0], "преподаватель")
    put_cell(t_prep.rows[0].cells[4], "М.Д. Яушкина")

    put_cell(t_work.rows[0].cells[0], "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 1")
    put_cell(
        t_work.rows[1].cells[0],
        "Первичное обследование состава системы технического зрения на производстве. "
        "Рассмотрение особенностей типов данных",
    )
    put_cell(t_work.rows[2].cells[0], "по курсу: Системы технического зрения на производстве")

    put_cell(t_stud.rows[0].cells[1], "4321")
    put_cell(t_stud.rows[0].cells[5], "Г.В. Буренков")


def setup_body_section(doc: Document):
    doc.add_section(WD_SECTION_START.NEW_PAGE)
    title_sec, body_sec = doc.sections[0], doc.sections[1]

    for sec in (title_sec, body_sec):
        sec.page_width, sec.page_height = Mm(210), Mm(297)
        sec.left_margin, sec.right_margin = Mm(30), Mm(15)
        sec.top_margin, sec.bottom_margin = Mm(20), Mm(20)

    title_sec.footer.is_linked_to_previous = False
    for p in title_sec.footer.paragraphs:
        p.clear()

    body_sec.footer.is_linked_to_previous = False
    body_sec.header.is_linked_to_previous = False

    sect_pr = body_sec._sectPr
    for old in sect_pr.findall(qn("w:pgNumType")):
        sect_pr.remove(old)
    pg = OxmlElement("w:pgNumType")
    pg.set(qn("w:start"), "2")
    sect_pr.append(pg)

    fp = body_sec.footer.paragraphs[0]
    fp.clear()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_page_field(fp)


def make_scheme():
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyBboxPatch

    FIG.parent.mkdir(parents=True, exist_ok=True)
    plt.rcParams["font.family"] = "DejaVu Sans"
    fig, ax = plt.subplots(figsize=(10.2, 6.4), dpi=160)
    ax.set_xlim(0, 12.3)
    ax.set_ylim(0, 7.4)
    ax.axis("off")

    items = [
        (0.4, 5.5, "Объект\nблистер 10 капсул\nFOV 120×80 мм"),
        (3.2, 5.5, "Освещение\nкупол LED VIS\nстробоскоп"),
        (6.0, 5.5, "Оптика\nобычный объектив\nбез телецентричности"),
        (8.8, 5.5, "Камера CMOS\n640×480 Mono8\n0,307 МБ/кадр"),
        (0.4, 2.2, "GigE Vision\n0,614 МБ/с\nLкаб = 5 м"),
        (3.2, 2.2, "Вычислитель\nPC-based, CV\nГоден/Брак"),
        (6.0, 2.2, "ПЛК\nцифровой I/O\ntIO = 2 мс"),
        (8.8, 2.2, "Исполнитель\nпневмосброс\ntисп ≈ 50 мс"),
    ]
    for x, y, text in items:
        ax.add_patch(
            FancyBboxPatch(
                (x, y),
                2.55,
                1.7,
                boxstyle="round,pad=0.03,rounding_size=0.08",
                facecolor="#f2f2f2",
                edgecolor="black",
                linewidth=1.0,
            )
        )
        ax.text(x + 1.275, y + 0.85, text, ha="center", va="center", fontsize=8.5)
    for x in (2.95, 5.75, 8.55):
        ax.annotate("", xy=(x + 0.25, 6.35), xytext=(x, 6.35), arrowprops=dict(arrowstyle="-|>", color="black", lw=1.1))
    for x in (2.95, 5.75, 8.55):
        ax.annotate("", xy=(x + 0.25, 3.05), xytext=(x, 3.05), arrowprops=dict(arrowstyle="-|>", color="black", lw=1.1))
    ax.plot([11.35, 11.7, 11.7, 1.675], [6.35, 6.35, 4.35, 4.35], color="black", lw=1.1)
    ax.annotate("", xy=(1.675, 3.9), xytext=(1.675, 4.35), arrowprops=dict(arrowstyle="-|>", color="black", lw=1.1))
    ax.text(6.0, 0.55, "Обратная связь: триггер такта / энкодер → экспозиция и стробоскоп", ha="center", fontsize=9)
    fig.tight_layout(pad=0.3)
    fig.savefig(FIG, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def build_body(doc: Document, toc_pages: dict):
    add_struct(doc, "СОДЕРЖАНИЕ", new_page=False)
    for title, key, ind in TOC_DEF:
        add_toc_line(doc, title, toc_pages.get(key, "…"), indent=ind)

    add_struct(doc, "ВВЕДЕНИЕ")
    add_body(
        doc,
        "Целью работы является освоение первичного обследования системы технического "
        "зрения (СТЗ) на производстве: описание состава подсистем, типов данных и "
        "расчёт базовых информационных характеристик процесса [1].",
    )
    add_body(
        doc,
        "Работа выполнена для варианта 1. Объект контроля — блистер на 10 капсул; "
        "контролируемые параметры — полнота заполнения и целостность. Расчёты "
        "выполнены по правилам методических указаний [1]. Численный подбор "
        "фокусного расстояния и глубины резкости не требуется и не выполнялся.",
    )

    add_section_h(doc, "1 Постановка задачи")
    add_body(
        doc,
        "Изделие — фармацевтический блистер на 10 капсул. Контролируются полнота "
        "заполнения ячеек и целостность капсул и фольги. Задача относится к функции "
        "контроля качества; попутно выполняется подсчёт заполненных ячеек. Система "
        "ставится на выходном контроле упаковки до картонирования. Пропуск дефекта "
        "на этом этапе приводит к выдаче неполной или повреждённой упаковки "
        "потребителю и к риску отзыва партии.",
    )
    add_body(
        doc,
        "Количественный критерий брака. Блистер бракуется, если хотя бы одна из "
        "10 ячеек пуста либо содержит дефект целостности линейным размером не менее "
        "2,0 мм (разрыв фольги, скол или деформация капсулы). Допустимое число "
        "пустых ячеек — 0. Особенности сцены: блики фольги и матовый ПВХ.",
    )
    add_body(
        doc,
        "Исходные данные варианта: поле зрения FOV = 120 × 80 мм; минимальный "
        "признак dmin = 2,0 мм; такт Tтакт = 0,5 с; скорость конвейера v = 250 мм/с; "
        "длина линии Lкаб = 5 м. Режим расчёта частоты съёмки — А (дискретная съёмка "
        "по такту); скорость используется для проверки смаза [1].",
    )

    add_section_h(doc, "2 Паспорт подсистем СТЗ")
    add_body(
        doc,
        "Требования к шести подсистемам сведены в таблицу 1. Разрешение камеры, "
        "частота съёмки и интерфейс указаны после расчёта раздела 5.",
    )
    add_caption_table(doc, "Таблица 1 — Паспорт подсистем СТЗ")
    add_table(
        doc,
        ["№", "Подсистема", "Назначение в данной системе", "Ключевой параметр", "Принятое значение / тип", "Обоснование выбора"],
        [
            [
                "1",
                "Освещение",
                "Контраст «капсула / пустая ячейка» и выявление разрыва фольги без зеркальных бликов",
                "геометрия, спектр, режим",
                "диффузный купол; LED VIS; стробоскоп",
                "Фольга даёт блики, ПВХ матовый. tэксп,max = 0,375 мс, поэтому нужна импульсная подсветка",
            ],
            [
                "2",
                "Оптика",
                "Построение изображения блистера на матрице в пределах заданного FOV",
                "FOV, тип объектива, телецентричность",
                "FOV 120 × 80 мм; обычный объектив; телецентричность не требуется",
                "Задача — обнаружение, а не прецизионное измерение. f, WD и ГРИП без паспорта камеры не рассчитываются [1]",
            ],
            [
                "3",
                "Фотоприемник",
                "Цифровой кадр сцены в такте 0,5 с без геометрического смаза",
                "тип матрицы, затвор; разрешение и FPS — после этапа 5",
                "CMOS, глобальный затвор, VIS; 640 × 480; 2 кадр/с",
                "tэксп,max < 1 мс. Бегущий затвор исказит движущийся блистер",
            ],
            [
                "4",
                "Передача данных",
                "Доставка кадра к вычислителю",
                "интерфейс, длина линии",
                "GigE Vision; Lкаб = 5 м",
                "Поток 0,614 МБ/с с запасом > 30 %. USB3 Vision ограничен длиной 5 м",
            ],
            [
                "5",
                "Вычислитель",
                "Сегментация ячеек, проверка заполнения и целостности, сигнал Годен/Брак",
                "архитектура, ускорение",
                "PC-based; классическое CV; GPU не требуется",
                "tобр до 30 мс укладывается в такт 500 мс [1]",
            ],
            [
                "6",
                "Исполнительное устройство",
                "Сброс бракованного блистера с конвейера",
                "тип, шина, задержка",
                "пневмоотбраковщик; дискретный выход ПЛК; tисп ≈ 50 мс",
                "Бинарное решение. Задержка механики в TСТЗ не входит [1]",
            ],
        ],
        widths_cm=[0.9, 2.2, 3.4, 2.6, 3.2, 4.2],
        size=9,
    )

    add_section_h(doc, "3 Классификационный профиль системы")
    add_body(
        doc,
        "Класс проектируемой СТЗ по признакам методических указаний [1] приведён в таблице 2.",
    )
    add_caption_table(doc, "Таблица 2 — Класс системы")
    add_table(
        doc,
        ["Признак", "Возможные категории", "Выбор", "Обоснование"],
        [
            ["Решаемая задача", "контроль / измерение / идентификация / позиционирование", "контроль", "Нужен сигнал Годен/Брак, а не размер или код"],
            ["Размерность данных", "1D / 2D / 2.5D / 3D", "2D", "Пустая ячейка и разрыв фольги проявляются как контраст в плоскости"],
            ["Спектральный диапазон", "UV / VIS / NIR-SWIR / Thermal / X-Ray", "VIS", "Капсула и фольга различимы в видимом свете при купольном освещении"],
            ["Степень интеграции", "оптический датчик / смарт-камера / PC-based / встраиваемая / распределенная", "PC-based", "Несколько признаков (пустота и целостность) удобнее вести на ПК"],
            ["Число камер и схема обзора", "однокамерная / многокамерная / стерео / круговая / сканирующая", "однокамерная", "FOV 120 × 80 мм покрывает блистер одним кадром"],
            ["Подвижность", "стационарная / роботизированная / мобильная", "стационарная", "Камера над конвейером, изделие движется относительно неё"],
            ["Способ освещения", "внешнее / активное контролируемое / структурированное / эмиссионное", "активное контролируемое", "Нужны купол и стробоскоп, внешний цеховой свет недостаточен"],
            ["Характер обработки", "пороговая / классическая CV / ML / DL", "классическая CV", "Ячейки регулярны; нейросеть не требуется"],
            ["Степень участия человека", "автоматическая / автоматизированная / поддержка решений / регистрация", "автоматическая", "Отбраковка в такте без оператора"],
            ["Условия эксплуатации", "лабораторные / цеховые / агрессивные / взрывоопасные / экстремальные", "цеховые", "Упаковочная линия; корпус камеры IP67"],
        ],
        widths_cm=[3.2, 4.6, 2.4, 6.3],
        size=9,
    )

    add_section_h(doc, "4 Анализ типов данных")
    add_body(
        doc,
        "Минимально достаточная размерность — 2D. Одномерный профиль вдоль строки не "
        "гарантирует обнаружение разрыва фольги произвольной ориентации и локального "
        "скола капсулы. Переход к 2.5D или 3D даёт карту высоты, но пустая ячейка и "
        "разрыв фольги уже видны по яркости; объём данных и стоимость растут без "
        "изменения критерия брака.",
    )
    add_body(
        doc,
        "Спектральный диапазон — видимый. Признак не требует UV, ИК или рентгена: "
        "контраст капсулы и фольги формируется освещением. Формат — Mono8: сортировка "
        "по цвету не задана, 256 градаций достаточно после подавления бликов куполом. "
        "Mono12 имел бы смысл только если блики остались бы неустранимыми оптически.",
    )
    add_body(
        doc,
        "Объёмы кадра для одной и той же сцены 640 × 480 пикс. приведены в таблице 3. "
        "Основное сжатие данных происходит не на передаче, а на этапе экстракции: "
        "кадр 0,307 МБ сводится к бинарному решению и, при необходимости, к счётчику "
        "заполненных ячеек.",
    )
    add_caption_table(doc, "Таблица 3 — Сравнение типов данных для одной сцены")
    add_table(
        doc,
        ["Представление", "Формат", "Байт/точку", "Объём кадра, МБ", "Что даёт", "Чего не даёт"],
        [
            ["1D профиль (строка)", "Mono8", "1", "0,00064", "Профиль яркости вдоль линии", "Площадные дефекты вне линии"],
            ["2D полутоновое", "Mono8", "1", "0,307", "Пустые ячейки и разрывы фольги", "Цвет и высоту"],
            ["2D цветное", "RGB8", "3", "0,922", "Цвет капсулы", "Не требуется заданием; поток втрое больше"],
            ["2.5D карта высот", "Coord3D C16", "2", "0,614", "Рельеф фольги и капсулы", "Избыточно для наличия/отсутствия"],
            ["3D облако точек", "Coord3D ABC32f", "12", "3,686", "Полную геометрию", "Не нужно для Годен/Брак; объём ×12 к Mono8"],
        ],
        widths_cm=[3.0, 2.4, 1.8, 2.2, 3.6, 3.5],
        size=9,
    )
    add_body(
        doc,
        "Вывод по «цене» размерности: выбранное 2D Mono8 даёт 0,307 МБ/кадр. Цвет "
        "увеличивает объём в три раза без новой информации о браке. Облако точек "
        "дороже более чем на порядок. Сжатие информации выполняется на сегментации "
        "и решении, а не выбором интерфейса.",
    )

    add_section_h(doc, "5 Расчёт параметров информационного процесса")
    add_body(
        doc,
        "Вариант относится к режиму А. Для обнаружения дефекта поверхности принято "
        "n = 4 пикселя на признак (допустимый диапазон 3–5) [1]. Формат Mono8, "
        "B = 1 байт/пикс. Ход расчёта совпадает с таблицей 4.",
    )
    add_body(
        doc,
        "Требуемый масштабный коэффициент kтреб = dmin / n = 2,0 / 4 = 0,50 мм/пикс. "
        "Требуемое разрешение Wтреб = 120 / 0,50 = 240 пикс, Hтреб = 80 / 0,50 = 160 пикс. "
        "По учебному ряду матричных камер принимается первое значение не ниже этих "
        "порогов: 640 × 480 пикс.",
    )
    add_body(
        doc,
        "Фактический масштаб kx = 120 / 640 = 0,1875 мм/пикс, ky = 80 / 480 = 0,1667 мм/пикс, "
        "k = max(kx, ky) = 0,1875 мм/пикс. Объём кадра V = 640 · 480 · 1 / 10^6 = 0,307 МБ. "
        "FPSтреб = 1 / 0,5 = 2 кадр/с. Поток D = 0,307 · 2 = 0,614 МБ/с.",
    )
    add_caption_table(doc, "Таблица 4 — Расчёт параметров информационного процесса")
    add_table(
        doc,
        ["Параметр", "Обозначение", "Формула", "Результат"],
        [
            ["Требуемый масштабный коэффициент", "kтреб", "dmin / n = 2,0 / 4", "0,50 мм/пикс"],
            ["Требуемое разрешение", "Wтреб × Hтреб", "FOVx / kтреб; FOVy / kтреб", "240 × 160 пикс"],
            ["Принятое разрешение камеры", "W × H", "первый стандарт из ряда А.7 [1]", "640 × 480 пикс"],
            ["Фактический масштаб", "kx; ky; k", "FOVx/W; FOVy/H; max(kx, ky)", "0,1875; 0,1667; 0,1875 мм/пикс"],
            ["Объём кадра", "V", "W · H · B / 10^6", "0,307 МБ"],
            ["Требуемая частота съёмки", "FPSтреб", "режим А: 1 / Tтакт", "2 кадр/с"],
            ["Поток данных", "D", "V · FPS", "0,614 МБ/с"],
            ["Время передачи кадра", "tпер", "1000 · V / Cэфф, Cэфф = 110 МБ/с", "2,8 мс"],
            ["Предельное время экспозиции", "tэксп, max", "1000 · k · pдоп / v, pдоп = 0,5 пикс", "0,375 мс"],
            ["Доступное время для решения", "Tдоп", "режим А: Tтакт", "500 мс"],
            ["Латентность СТЗ", "TСТЗ", "tэксп + tсч + tпер + tобр + tлог + tIO", "38,2 мс"],
            ["Проверка по времени", "—", "TСТЗ ≤ Tдоп", "выполнено"],
        ],
        widths_cm=[4.4, 2.6, 5.8, 3.7],
        size=9,
    )
    add_body(
        doc,
        "Учебные составляющие латентности по [1]: tсч = 2 мс, tлог = 1 мс, tIO = 2 мс; "
        "для классического CV принята верхняя оценка tобр = 30 мс. Тогда "
        "TСТЗ = 0,375 + 2 + 2,8 + 30 + 1 + 2 = 38,2 мс, что меньше Tдоп = 500 мс.",
    )

    add_section_h(doc, "6 Выбор интерфейса и проверка временного цикла")
    add_body(
        doc,
        "Сравнение каналов для потока 0,614 МБ/с и длины 5 м приведено в таблице 5. "
        "Запас по пропускной способности должен быть не менее 30 % [1].",
    )
    add_caption_table(doc, "Таблица 5 — Выбор интерфейса передачи данных")
    add_table(
        doc,
        ["Интерфейс", "Эффективная скорость, МБ/с", "Макс. длина", "Запас к 0,614 МБ/с", "Вывод"],
        [
            ["USB3 Vision", "350–400", "до 5 м", "многократный", "Длина равна Lкаб, запаса по кабелю нет"],
            ["GigE Vision (1G)", "около 110", "до 100 м", "более 100 раз", "Принят: запас и длина"],
            ["10 GigE Vision", "около 1100", "до 100 м", "избыточен", "Не требуется"],
            ["Camera Link (Full)", "около 850", "до 10 м", "избыточен", "Нужна плата захвата"],
            ["CoaXPress CXP-12", "около 1250 на канал", "до 40 м", "избыточен", "Не требуется"],
        ],
        widths_cm=[3.4, 3.4, 2.4, 3.2, 4.1],
        size=9,
    )
    add_body(
        doc,
        "Принят GigE Vision 1G. Условие по времени выполняется: 38,2 мс < 500 мс. "
        "Полный цикл с механикой Tполн = TСТЗ + tисп = 38,2 + 50 = 88,2 мс. Если "
        "отбраковщик стоит ниже по конвейеру, достаточно Lисп ≥ v · Tполн = "
        "250 · 0,0882 ≈ 22 мм.",
    )
    add_body(
        doc,
        "Проверка смаза: tэксп,max = 0,375 мс < 1 мс, поэтому необходимы глобальный "
        "затвор и стробоскопическая подсветка. Конвейеризация алгоритма не нужна: "
        "запас по такту составляет более чем порядок.",
    )

    add_section_h(doc, "7 Структурная схема системы")
    add_body(
        doc,
        "Структура СТЗ от объекта до исполнительного устройства с указанием типа "
        "данных на переходах показана на рисунке 1. Обратная связь — сигнал такта "
        "(или энкодера) на запуск экспозиции и стробоскопа.",
    )
    add_figure(doc, FIG, "Рисунок 1 — Структурная схема СТЗ контроля блистера", w=16.0)
    add_body(
        doc,
        "На входе — оптическое изображение блистера. После камеры данные имеют вид "
        "кадра Mono8 объёмом 0,307 МБ. По GigE Vision передаётся поток 0,614 МБ/с. "
        "Выход вычислителя — бинарный сигнал Годен/Брак на ПЛК; исполнитель получает "
        "дискретный выход и сбрасывает изделие. Ключевой признак СТЗ соблюдён: "
        "результатом является управляющее воздействие в заданном такте, а не "
        "изображение для оператора [1].",
    )

    add_struct(doc, "ЗАКЛЮЧЕНИЕ")
    add_body(
        doc,
        "Выявляемый признак проявляется оптически: пустая ячейка и разрыв фольги "
        "дают контраст яркости при купольном освещении. Рентген, ультразвук или "
        "вихретоковый контроль не требуются. Если блики фольги не удастся подавить "
        "оптически, следует сменить геометрию света, а не тип датчика.",
    )
    add_body(
        doc,
        "Для данной задачи дороже ошибка II рода (пропуск брака): неполная упаковка "
        "может уйти потребителю. Порог лучше смещать в сторону чувствительности, "
        "принимая рост ложной отбраковки (ошибка I рода). Ложный сброс блистера "
        "дешевле отзыва партии.",
    )
    add_body(
        doc,
        "Риски эксплуатации: блики фольги, пыль на куполе, вибрации конвейера, "
        "нестабильный внешний свет. Меры: закрытый купол, IP67, крепление камеры "
        "с виброизоляцией, стробоскоп вместо работы на постоянном свете цеха.",
    )
    add_body(
        doc,
        "Количественный расчёт окупаемости не выполняется: в варианте нет стоимости "
        "оборудования, брака и объёма выпуска. Для срока окупаемости нужны цена СТЗ, "
        "трудозатраты ручного контроля, стоимость пропуска брака и годовой выпуск. "
        "Качественно внедрение оправдано снижением пропуска неполных блистеров, "
        "отказом от сплошного визуального контроля и прослеживаемостью отбраковки.",
    )
    add_body(
        doc,
        "Принятые допущения: n = 4; учебный ряд разрешений; tсч = 2 мс, tлог = 1 мс, "
        "tIO = 2 мс, tобр = 30 мс; Cэфф GigE = 110 МБ/с; tисп = 50 мс; f, WD и ГРИП "
        "не рассчитывались. Условие TСТЗ ≤ Tдоп выполняется с запасом "
        "(38,2 мс против 500 мс).",
    )

    add_struct(doc, "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ")
    sources = [
        "Первичное обследование состава системы технического зрения на производстве. "
        "Рассмотрение особенностей типов данных : метод. указания к лабораторной работе № 1. "
        "– Санкт-Петербург : ГУАП, 2026.",
    ]
    for i, src in enumerate(sources, 1):
        p = doc.add_paragraph()
        format_paragraph(p)
        set_run_font(p.add_run(f"{i}. {src}"))


def build(toc_pages: dict) -> Path:
    make_scheme()
    shutil.copy(BLANK, OUT)
    doc = Document(str(OUT))
    fill_title(doc)
    setup_body_section(doc)
    build_body(doc, toc_pages)
    doc.save(str(OUT))
    return OUT


def convert():
    ps = f"""
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
foreach ($d in @($word.Documents)) {{ $d.Close([ref]0) }}
$src = '{OUT.resolve()}'
$dst = '{PDF.resolve()}'
$doc = $word.Documents.Open($src)
$wdFormatPDF = 17
$doc.SaveAs([ref]$dst, [ref]$wdFormatPDF)
$doc.Close([ref]0)
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
"""
    subprocess.run(
        ["powershell", "-NoProfile", "-Command", ps],
        check=True,
    )


def measure() -> dict[str, int]:
    from pypdf import PdfReader

    reader = PdfReader(str(PDF))
    pages: dict[str, int] = {}
    for i, page in enumerate(reader.pages):
        if i < 2:
            continue
        text = page.extract_text() or ""
        printed = i + 1
        for key, needle in SEARCH.items():
            if key not in pages and needle in text:
                pages[key] = printed
    return pages


def main():
    build({k: "…" for k in SEARCH})
    try:
        convert()
        pages = measure()
        print("pass1", pages)
        if pages:
            build(pages)
            convert()
            print("pass2", measure())
    except Exception as exc:
        print("PDF skipped:", exc)
    print(OUT)


if __name__ == "__main__":
    main()
