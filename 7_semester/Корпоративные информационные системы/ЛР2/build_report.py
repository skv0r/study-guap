# -*- coding: utf-8 -*-
"""Сборка отчёта ЛР2. В текст отчёта не входит."""

from __future__ import annotations

import shutil
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Mm, Pt, RGBColor
from lxml import etree

import generate_figures as figs

BASE = Path(__file__).resolve().parent
SRC = Path(
    r"C:\Users\skvor\OneDrive\Desktop\GitHub\study-guap"
    r"\7_semester\Web-программирование\ЛР1_Отчет_Буренков_PairWise.docx"
)
OUT = BASE / "ЛР2_Отчет_Буренков_КИС.docx"
FIG = BASE / "figures"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

TITLE_REPLACEMENTS = {
    "ассистент": "профессор, д-р техн. наук, профессор",
    "В.В. Жукалин": "В.В. Фомин",
    "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 1": "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 2",
    "Создание проекта Next.js с использованием React, TypeScript и App Router":
        "Изучение методологии моделирования бизнес-процессов IDEF0",
    "по курсу: Web-программирование": "по курсу: Корпоративные информационные системы",
}

TOC = [
    ("1 Задание", "z1", 0),
    ("2 Предметная область", "z2", 0),
    ("3 Модель IDEF0", "z3", 0),
    ("3.1 Контекстная диаграмма А-0", "z31", 1),
    ("3.2 Декомпозиция А0", "z32", 1),
    ("3.3 Приём заказа", "z33", 1),
    ("3.4 Диагностика", "z34", 1),
    ("3.5 Ремонт", "z35", 1),
    ("3.6 Глоссарий стрелок", "z36", 1),
    ("4 Выводы", "z4", 0),
]


def prepare_title(path: Path) -> None:
    shutil.copy(SRC, path)
    import zipfile

    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find(f"{W}body")
    children = list(body)
    if len(children) < 12:
        raise RuntimeError("В образце меньше 12 элементов body")
    p11 = children[11]
    sect = p11.find(f"{W}pPr/{W}sectPr")
    if sect is None:
        raise RuntimeError("В абзаце 11 нет sectPr")
    sect = etree.fromstring(etree.tostring(sect))
    for child in children:
        body.remove(child)
    for child in children[:11]:
        body.append(child)
    body.append(sect)
    replaced = []
    for node in body.iter(f"{W}t"):
        if node.text in TITLE_REPLACEMENTS:
            replaced.append(node.text)
            node.text = TITLE_REPLACEMENTS[node.text]
    missing = [k for k in TITLE_REPLACEMENTS if k not in replaced]
    if missing:
        raise RuntimeError(f"Не найдены строки титула: {missing}")
    report = None
    for node in body.iter(f"{W}t"):
        if node.text and node.text.startswith("ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ"):
            report = node
            break
    rpr = report.getparent().find(f"{W}rPr")
    if rpr is not None and rpr.find(f"{W}b") is not None:
        raise RuntimeError("Строка отчёта на титуле стала жирной")
    xml_out = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = path.with_suffix(".tmp.zip")
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            data = xml_out if item.filename == "word/document.xml" else zin.read(item.filename)
            zout.writestr(item, data)
    tmp.replace(path)


def set_run_font(run, size=14, bold=False):
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = False
    run.font.color.rgb = RGBColor(0, 0, 0)


def format_paragraph(p, *, indent=True, align="justify", before=0, after=0, line=1.5, left=0):
    pf = p.paragraph_format
    if line == 1:
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    else:
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.line_spacing = line
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.first_line_indent = Cm(1.25 if indent else 0)
    pf.left_indent = Cm(left)
    pf.right_indent = Cm(0)
    p.alignment = {
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "left": WD_ALIGN_PARAGRAPH.LEFT,
    }[align]


def add_body(doc, text):
    p = doc.add_paragraph()
    format_paragraph(p)
    set_run_font(p.add_run(text))
    return p


def add_heading(doc, text, bookmark, bid):
    p = doc.add_paragraph()
    format_paragraph(p, indent=True, align="left", before=12, after=6)
    p.paragraph_format.keep_together = True
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text), bold=True)
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bid))
    start.set(qn("w:name"), bookmark)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bid))
    p._p.insert(0, start)
    p._p.append(end)
    return p


def _field_run(paragraph, kind, payload=None):
    run = paragraph.add_run()
    set_run_font(run)
    if kind == "text":
        run.text = payload
        return
    for child in list(run._r):
        if child.tag == qn("w:t"):
            run._r.remove(child)
    if kind == "instr":
        el = OxmlElement("w:instrText")
        el.set(qn("xml:space"), "preserve")
        el.text = payload
    else:
        el = OxmlElement("w:fldChar")
        el.set(qn("w:fldCharType"), kind)
    run._r.append(el)


def add_toc_line(doc, title, bookmark, level):
    p = doc.add_paragraph()
    format_paragraph(p, indent=False, align="left", left=0.75 * level)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(16.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    set_run_font(p.add_run(title))
    set_run_font(p.add_run("\t"))
    _field_run(p, "begin")
    _field_run(p, "instr", f" PAGEREF {bookmark} \\h ")
    _field_run(p, "separate")
    _field_run(p, "text", "0")
    _field_run(p, "end")
    return p


def add_caption_table(doc, text):
    p = doc.add_paragraph()
    format_paragraph(p, indent=False, align="left", before=12, after=6, line=1)
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text))
    return p


def add_caption_figure(doc, text):
    p = doc.add_paragraph()
    format_paragraph(p, indent=False, align="center", before=6, after=0)
    set_run_font(p.add_run(text))
    return p


def add_picture(doc, path, width_cm):
    p = doc.add_paragraph()
    format_paragraph(p, indent=False, align="center", before=6, after=0)
    p.paragraph_format.keep_with_next = True
    run = p.add_run()
    run.add_picture(str(path), width=Cm(width_cm))
    return p


def _set_cell_border(cell):
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


def add_table(doc, headers, rows):
    table = doc.add_table(rows=1 + len(rows), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    twips = [str(int(round(cm / 2.54 * 1440))) for cm in (4.6, 11.9)]
    widths = [Cm(4.6), Cm(11.9)]
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    jc = tblPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        tblPr.append(jc)
    jc.set(qn("w:val"), "left")
    grid = tbl.find(qn("w:tblGrid"))
    for i, child in enumerate(list(grid)):
        child.set(qn("w:w"), twips[i])
    total = sum(int(x) for x in twips)
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(total))
    tblW.set(qn("w:type"), "dxa")
    data = [headers] + rows
    for r, row in enumerate(data):
        for c, text in enumerate(row):
            cell = table.rows[r].cells[c]
            cell.width = widths[c]
            _set_cell_border(cell)
            cell.text = ""
            p = cell.paragraphs[0]
            format_paragraph(p, indent=False, align="left", line=1.5)
            set_run_font(p.add_run(text), bold=(r == 0))
    return table


def apply_section(section, *, landscape: bool, start: int | None):
    if landscape:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Mm(297)
        section.page_height = Mm(210)
    else:
        section.orientation = WD_ORIENT.PORTRAIT
        section.page_width = Mm(210)
        section.page_height = Mm(297)
    section.left_margin = Mm(30)
    section.right_margin = Mm(15)
    section.top_margin = Mm(20)
    section.bottom_margin = Mm(20)
    section.header_distance = Cm(1.25)
    section.footer_distance = Cm(1.25)
    sect = section._sectPr
    for tag in ("w:footerReference", "w:headerReference"):
        for el in list(sect.findall(qn(tag))):
            sect.remove(el)
    ref = OxmlElement("w:footerReference")
    ref.set(qn("w:type"), "default")
    ref.set(qn("r:id"), "rId11")
    sect.insert(0, ref)
    pg = sect.find(qn("w:pgNumType"))
    if start is None:
        if pg is not None:
            sect.remove(pg)
    else:
        if pg is None:
            pg = OxmlElement("w:pgNumType")
            sect.append(pg)
        pg.set(qn("w:start"), str(start))


def build():
    figs.main()
    prepare_title(OUT)
    doc = Document(OUT)
    body = doc.add_section(WD_SECTION.NEW_PAGE)
    apply_section(body, landscape=False, start=2)

    p = doc.add_paragraph()
    format_paragraph(p, indent=False, align="center", before=0, after=12)
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run("СОДЕРЖАНИЕ"), bold=True)
    for i, (title, mark, level) in enumerate(TOC, start=1):
        add_toc_line(doc, title, mark, level)

    add_heading(doc, "1 Задание", "z1", 1)
    add_body(
        doc,
        "Нужно изучить методологию IDEF0 и показать её на сквозном примере своей предметной области. "
        "Модель включает пять диаграмм: контекстную А-0, декомпозицию А0 и три дочерние. "
        "Название блока формулируется глаголом, название стрелки — существительным. "
        "Вход подводится слева, управление сверху, выход справа, механизм снизу. "
        "У каждого блока есть хотя бы одна управляющая и одна выходящая дуга. "
        "На диаграмме декомпозиции от трёх до шести блоков. "
        "Номер в правом нижнем углу блока согласован со ссылкой на дочернюю диаграмму. "
        "Для нескольких стрелок составлен глоссарий. "
        "Цель и точка зрения модели фиксируются на контекстной диаграмме.",
    )

    add_heading(doc, "2 Предметная область", "z2", 2)
    add_body(
        doc,
        "Мастерская «Обод» ремонтирует городские и спортивные велосипеды по отдельным заказам. "
        "Клиент приносит велосипед и сообщает, что нужно исправить. Приёмщик регистрирует обращение, "
        "осматривает внешний вид, по прайс-листу согласовывает предварительную стоимость и оставляет "
        "велосипед на хранении. Мастер по регламенту мастерской проверяет ходовую часть, тормоза и "
        "трансмиссию, записывает замечания и составляет дефектную ведомость. Затем по ведомости со склада "
        "подбирают комплектующие, заменяют неисправные детали, регулируют узлы и фиксируют выполненные работы. "
        "При выдаче клиент получает отремонтированный велосипед и акт выполненных работ. "
        "В границу модели входит один заказ: от заявки и велосипеда до выдачи. Закупка партий у поставщиков, "
        "расчёт оплаты труда и реклама мастерской не рассматриваются.",
    )

    add_heading(doc, "3 Модель IDEF0", "z3", 3)
    add_body(
        doc,
        "Модель описывает один заказ мастерской «Обод». Цель — показать порядок приёма, диагностики, "
        "ремонта и выдачи и состав документов, без которых эти функции не выполняются. "
        "Точка зрения — владелец мастерской: его интересуют границы заказа, роли приёмщика и мастера "
        "и документы, а не конструкция отдельного узла. Дочерние диаграммы построены с той же целью "
        "и с той же точки зрения. На диаграмме А-0 блок имеет номер 0 и ссылку А0. "
        "На диаграмме А0 блоки 1–3 ссылаются на диаграммы А1, А2 и А3, у блока 4 ссылки нет.",
    )
    add_body(
        doc,
        "На контекст входят велосипед клиента, заявка и комплектующие. Управление задают прайс-лист услуг "
        "и регламент мастерской. На выходе — отремонтированный велосипед и акт выполненных работ. "
        "Механизмы — приёмщик, мастер и инструмент. В декомпозиции А0 карточка заказа из приёмки идёт "
        "в диагностику и сверху управляет выдачей. На диаграммах А1, А2 и А3 по четыре блока: внешние "
        "стрелки родительского блока сохранены, у каждого блока есть управление и выход.",
    )

    diagrams = [
        (
            "3.1 Контекстная диаграмма А-0",
            "z31",
            4,
            "Контекстная диаграмма сводит мастерскую к одному блоку «Ремонтировать велосипеды» (рисунок 1). "
            "Цель — описать выполнение одного заказа на ремонт велосипеда. Точка зрения — владелец мастерской.",
            FIG / "fig01_a-0.png",
            "Рисунок 1 — Контекстная диаграмма А-0",
            24.0,
        ),
        (
            "3.2 Декомпозиция А0",
            "z32",
            5,
            "Приём, диагностика, ремонт и выдача сохраняют внешние стрелки контекста (рисунок 2).",
            FIG / "fig02_a0.png",
            "Рисунок 2 — Декомпозиция А0",
            21.0,
        ),
        (
            "3.3 Приём заказа",
            "z33",
            6,
            "Заявка становится карточкой заказа, велосипед клиента — принятым (рисунок 3).",
            FIG / "fig03_a1.png",
            "Рисунок 3 — Диаграмма А1 «Принять заказ»",
            21.0,
        ),
        (
            "3.4 Диагностика",
            "z34",
            7,
            "Проверки идут от ходовой части к тормозам и трансмиссии, затем составляется ведомость (рисунок 4).",
            FIG / "fig04_a2.png",
            "Рисунок 4 — Диаграмма А2 «Диагностировать неисправность»",
            21.0,
        ),
        (
            "3.5 Ремонт",
            "z35",
            8,
            "По ведомости подбирают детали, меняют узлы, регулируют их и записывают работы (рисунок 5).",
            FIG / "fig05_a3.png",
            "Рисунок 5 — Диаграмма А3 «Выполнить ремонт»",
            21.0,
        ),
    ]
    for title, mark, bid, text, image, caption, width in diagrams:
        sec = doc.add_section(WD_SECTION.NEW_PAGE)
        apply_section(sec, landscape=True, start=None)
        add_heading(doc, title, mark, bid)
        add_body(doc, text)
        add_picture(doc, image, width)
        add_caption_figure(doc, caption)

    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    apply_section(sec, landscape=False, start=None)
    add_heading(doc, "3.6 Глоссарий стрелок", "z36", 9)
    add_body(doc, "Смысл документов, которые связывают блоки модели, приведён в таблице 1.")
    add_caption_table(doc, "Таблица 1 — Глоссарий интерфейсных дуг")
    add_table(
        doc,
        ["Дуга", "Содержание"],
        [
            [
                "Карточка заказа",
                "Учётный документ приёмки: клиент, модель велосипеда, описание внешнего вида, "
                "согласованная предварительная стоимость и отметка о хранении.",
            ],
            [
                "Дефектная ведомость",
                "Перечень неисправностей ходовой части, тормозов и трансмиссии и узлы, которые подлежат замене.",
            ],
            [
                "Прайс-лист услуг",
                "Перечень работ мастерской и цен. По нему согласуют предварительную стоимость и закрывают заказ.",
            ],
            [
                "Регламент мастерской",
                "Правила приёмки, диагностики и ремонта, обязательные для приёмщика и мастера.",
            ],
            [
                "Перечень работ",
                "Список фактически выполненных операций и заменённых деталей. Приёмщик переносит его в акт.",
            ],
            [
                "Акт выполненных работ",
                "Документ выдачи: работы, заменённые детали и итоговая сумма.",
            ],
        ],
    )

    add_heading(doc, "4 Выводы", "z4", 10)
    add_body(
        doc,
        "Построена модель IDEF0 мастерской ремонта велосипедов «Обод» из пяти диаграмм. "
        "Контекст А-0 ограничивает модель одним заказом и фиксирует цель и точку зрения владельца. "
        "Декомпозиция А0 делит работу на приём, диагностику, ремонт и выдачу. "
        "Диаграммы А1, А2 и А3 раскрывают первые три функции, выдача остаётся без дальнейшей декомпозиции. "
        "На каждой из этих диаграмм четыре блока, у каждого есть управление и выход, внешние стрелки родителя сохранены. "
        "Глоссарий поясняет карточку заказа, дефектную ведомость, прайс-лист, регламент, перечень работ и акт.",
    )
    doc.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
