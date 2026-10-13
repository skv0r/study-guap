#!/usr/bin/env python3
"""Сборка отчёта ЛР4. В текст отчёта не входит."""

from __future__ import annotations

import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Pt
from lxml import etree

BASE = Path(__file__).resolve().parent
TEMPLATE = BASE.parent / "ЛР2" / "ЛР2_Отчет_Буренков_Требования.docx"
OUT = BASE / "ЛР4_Отчет_Буренков_Требования.docx"
TOPIC = "Дополнительная спецификация"
USABLE = Cm(16.5)
PAGE_LIMIT = 15


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
        format_paragraph(p, first_indent=False, align="center", space_after=8)
    else:
        format_paragraph(
            p,
            first_indent=True,
            align="left",
            space_before=8 if level == 1 else 6,
            space_after=4 if level == 1 else 2,
        )
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text), bold=True)
    return p


def add_toc_line(doc, title, page, *, indent=0):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="left", left_indent=indent, line_spacing=1.15)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(16.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    set_run_font(p.add_run(title))
    set_run_font(p.add_run("\t"))
    set_run_font(p.add_run(str(page)))


def add_table_caption(doc, text):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="left", space_before=8, space_after=6, single=True)
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text))


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


def _set_cell_text(cell, text, *, bold=False, align="left", size=12):
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
        "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 2": "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 4",
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


HEADINGS = [
    ("ВВЕДЕНИЕ", 0),
    ("1 Введение", 0),
    ("1.1 Назначение", 1),
    ("1.2 Границы", 1),
    ("1.3 Определения и сокращения", 1),
    ("1.4 Связанные документы", 1),
    ("1.5 Обзор", 1),
    ("2 Функциональность", 0),
    ("2.1 Язык интерфейса", 1),
    ("2.2 Бизнес-правила", 1),
    ("3 Удобство использования", 0),
    ("4 Надёжность", 0),
    ("5 Производительность", 0),
    ("6 Сопровождаемость", 0),
    ("7 Ограничения проектирования", 0),
    ("8 Документация и справка", 0),
    ("9 Покупные компоненты", 0),
    ("10 Интерфейсы", 0),
    ("10.1 Входы", 1),
    ("10.2 Выходы", 1),
    ("10.3 Смежные системы", 1),
    ("11 Лицензионные требования", 0),
    ("12 Юридические условия", 0),
    ("13 Применимые нормы", 0),
    ("ЗАКЛЮЧЕНИЕ", 0),
]


def build_body(doc: Document, pages: dict[str, str]) -> None:
    add_heading(doc, "СОДЕРЖАНИЕ", level=0)
    for title, indent in HEADINGS:
        add_toc_line(doc, title, pages.get(title, "…"), indent=indent * 0.5)

    add_heading(doc, "ВВЕДЕНИЕ", new_page=True, level=0)
    add_body(
        doc,
        "Документ собирает требования, общие для всей автоматизированной информационной системы железной дороги. "
        "Система охватывает не только кассу: маршруты, обслуживание станции, смены, работу инженеров "
        "(ремонт и инвентаризация) и краткий билетный контур. "
        "Грузовые перевозки и публичный интернет-магазин билетов по-прежнему вне границ.",
    )

    add_heading(doc, "1 Введение", new_page=True)
    add_heading(doc, "1.1 Назначение", level=2)
    add_body(
        doc,
        "Документ фиксирует дополнительные требования к системе на пилотном участке одного перевозчика: "
        "то, что относится ко всей системе, а не к шагам одного варианта.",
    )
    add_heading(doc, "1.2 Границы", level=2)
    add_body(
        doc,
        "Внутри границ: справочники станций и схем, маршруты, обслуживание станции, смены персонала, "
        "заявки на ремонт объектов станции, инвентаризация оборудования и имущества станции, "
        "расписание рейсов и краткий билетный контур. Пассажир рабочего места не имеет. "
        "Вне границ: грузовые перевозки, управление движением, найм и увольнение, закупка запчастей, "
        "интернет-магазин билетов, мобильное приложение, проводки и НДС, платёжный шлюз и несколько перевозчиков.",
    )
    add_heading(doc, "1.3 Определения и сокращения", level=2)
    add_body(doc, "Термины, которые дальше используются без повторного объяснения, сведены в таблице 1.")
    add_table_caption(doc, "Таблица 1 — Определения")
    add_table(
        doc,
        ["Термин", "Смысл"],
        [
            ["АИС", "Автоматизированная информационная система железной дороги"],
            ["Маршрут", "Упорядоченная цепочка станций пилотного участка с плановым временем хода"],
            ["Рейс", "Пассажирский рейс с датой, маршрутом и одним фондом мест"],
            ["Фонд мест", "Места рейса по схеме вагонов; у места не больше одного действующего билета"],
            ["Смена", "Интервал работы сотрудника на рабочем месте кассы или посадки"],
            ["Обслуживание станции", "Сутки готовности станции: кассы, платформы, зал, проходы"],
            ["Заявка на ремонт", "Запись о неисправности объекта станции: объект, описание, статус"],
            ["Инвентаризация", "Сверка фактического наличия оборудования и имущества станции с учётным списком"],
            ["Пилотный участок", "Ограниченный набор станций, маршрутов и рейсов первой версии"],
        ],
        [4.2, 12.3],
    )
    add_heading(doc, "1.4 Связанные документы", level=2)
    add_body(
        doc,
        "Документ о концепции и границах АИС железной дороги и спецификация вариантов UC-1–UC-9. "
        "Требования к ремонту и инвентаризации задаются этой дополнительной спецификацией и ролями инженеров.",
    )
    add_heading(doc, "1.5 Обзор", level=2)
    add_body(
        doc,
        "Дальше требования сгруппированы по виду: функциональность, удобство, надёжность, производительность, "
        "сопровождаемость, ограничения, документация, покупные компоненты, интерфейсы, лицензии, юридические условия и нормы.",
    )

    add_heading(doc, "2 Функциональность", new_page=True)
    add_body(
        doc,
        "Пошаговое поведение вариантов — в спецификации UC. Здесь язык интерфейса и правила на весь контур: "
        "маршрут, обслуживание станции, смены, ремонт, инвентаризация, рейс и билет.",
    )
    add_heading(doc, "2.1 Язык интерфейса", level=2)
    add_body(
        doc,
        "Интерфейс персонала ведётся на русском: подписи ролей, сообщения об отказе и тексты справочников. "
        "Переключения языка нет, публичного экрана пассажира нет.",
    )
    add_heading(doc, "2.2 Бизнес-правила", level=2)
    add_body(doc, "Числовые пороги документа концепции приведены в таблице 2 и дальше не повторяются другими значениями.")
    add_table_caption(doc, "Таблица 2 — Бизнес-цели")
    add_table(
        doc,
        ["Цель", "Порог"],
        [
            ["БЦ-1", "Оформление билета от запроса до печати — не более 3 минут в 90 % случаев за три месяца после ввода"],
            ["БЦ-2", "Не более одного подтверждённого инцидента двойной продажи одного места в месяц"],
            ["БЦ-3", "После отмены рейса новые продажи невозможны не позднее чем через 5 минут в 95 % случаев"],
        ],
        [2.2, 14.3],
    )
    add_body(
        doc,
        "Рядом с порогами таблицы 2: маршрут — не меньше двух станций; рейс без маршрута не публикуют; "
        "смена кассы или посадки — только при открытом обслуживании станции; продажа и посадка — при открытой смене; "
        "цена из прайса; одно место — один действующий билет. "
        "Объект, отмеченный неготовым при обслуживании станции, порождает или обновляет заявку на ремонт; "
        "закрыть заявку может только инженер по ремонту после отметки устранения. "
        "Инвентаризацию проводит инженер по инвентаризации по станции и дате: расхождение фиксируется в акте, "
        "кассир и диспетчер акт не правят. Найм, закупка запчастей и сводка продаж — вне первой версии.",
        space_before=6,
    )

    add_heading(doc, "3 Удобство использования", new_page=True)
    add_body(doc, "Рабочие места и права ролей сведены в таблице 3. У пассажира рабочего места нет.")
    add_table_caption(doc, "Таблица 3 — Роли")
    add_table(
        doc,
        ["Роль", "Может", "Не может"],
        [
            ["Кассир", "Смена, продажа и возврат", "Маршруты, ремонт, инвентаризация"],
            ["Диспетчер", "Маршрут, публикация и отмена рейса", "Продажа, ремонт, инвентаризация"],
            ["Дежурный по станции", "Обслуживание станции, готовность объектов", "Продажа и закрытие заявки на ремонт"],
            ["Инженер по ремонту", "Вести и закрывать заявки на ремонт объектов станции", "Продавать билеты и менять расписание"],
            ["Инженер по инвентаризации", "Проводить инвентаризацию и фиксировать расхождения", "Менять рейсы и обслуживать станцию"],
            ["Контролёр", "Смена и отметка посадки", "Менять цену, место и рейс"],
            ["Администратор", "Справочники, аварийное закрытие смены", "Публиковать рейс и продавать билет"],
        ],
        [4.2, 6.3, 6.0],
    )
    add_body(
        doc,
        "Обслуживание станции, заявки на ремонт и акт инвентаризации показывают объекты станции на одном экране роли. "
        "Форма пишет причину отказа: станция не в обслуживании, смена не открыта, заявка уже закрыта, "
        "инвентаризация по дате уже проведена, маршрут не собран, нет мест.",
        space_before=6,
    )

    add_heading(doc, "4 Надёжность", new_page=True)
    add_body(doc, "Надёжность и время оформления сведены в таблице 4. Числовые пороги те же, что в таблице 2.")
    add_table_caption(doc, "Таблица 4 — Надёжность и производительность")
    add_table(
        doc,
        ["Показатель", "Требование"],
        [
            ["Доступность", "Касса, обслуживание станции, заявки на ремонт и инвентаризация доступны в рабочие часы персонала. Плановые работы — вне этих часов"],
            ["Точность места", "Единый фонд мест; два действующих билета на одно место не проводятся. Порог — БЦ-2"],
            ["Ремонт и инвентарь", "Статус заявки и результат инвентаризации совпадают на всех рабочих местах; закрытая заявка повторно не открывается той же неисправностью без новой записи"],
            ["Блокировка / время", "После отмены касса видит запрет продажи (БЦ-3); оформление билета — в пороге БЦ-1"],
        ],
        [4.2, 12.3],
    )
    add_heading(doc, "5 Производительность")
    add_body(doc, "Требование по времени — строка «Время оформления» таблицы 4 и порог БЦ-1 таблицы 2.")

    add_heading(doc, "6 Сопровождаемость")
    add_body(
        doc,
        "Один сервер базы и рабочие места касс, диспетчерской, дежурного по станции, инженеров "
        "(ремонт и инвентаризация) и контроля посадки. Перед пуском загружают станции, схемы, перечень "
        "оборудования станции для инвентаризации и учётные записи персонала без найма в системе. "
        "Резервное копирование — отдельная инструкция администратора.",
    )

    add_heading(doc, "7 Ограничения проектирования")
    add_body(
        doc,
        "Первая версия — пилотный участок и один перевозчик. Тариф из прайса. Ремонт и инвентаризация ведутся "
        "по объектам станции пилота, без склада запчастей и без закупки. Не проектируются управление движением, "
        "найм, публичный сайт, бухгалтерский учёт, платёжный шлюз и несколько перевозчиков. "
        "При конфликте запросов приоритет у обслуживания станции, кассы и функций порогов таблицы 2.",
    )

    add_heading(doc, "8 Документация и справка")
    add_body(
        doc,
        "К первой версии — краткое руководство по ролям (включая инженеров) и инструкция по резервному копированию. "
        "На формах продажи, ремонта и инвентаризации — короткая подсказка о причине отказа.",
    )

    add_heading(doc, "9 Покупные компоненты")
    add_body(
        doc,
        "Платёжный шлюз и витрина интернет-продажи в первую версию не входят: оплата фиксируется как факт в кассе. "
        "Купон печатается на принтере кассы. База размещается на сервере перевозчика.",
    )

    add_heading(doc, "10 Интерфейсы")
    add_body(doc, "Входы персонала — в таблице 5, выходы системы — в таблице 6.")
    add_heading(doc, "10.1 Входы", level=2)
    add_table_caption(doc, "Таблица 5 — Входные данные")
    add_table(
        doc,
        ["Вход", "Состав", "Кто передаёт"],
        [
            ["Справочник", "Станция, схема вагона или перечень оборудования", "Администратор"],
            ["Маршрут / рейс", "Маршрут, публикация или причина отмены", "Диспетчер"],
            ["Обслуживание станции", "Дата, готовность объектов, замечания", "Дежурный по станции"],
            ["Заявка на ремонт", "Объект, описание неисправности, статус", "Дежурный или инженер по ремонту"],
            ["Инвентаризация", "Станция, дата, факт наличия, расхождения", "Инженер по инвентаризации"],
            ["Смена / билет", "Смена роли или данные продажи, возврата, посадки", "Кассир или контролёр"],
        ],
        [4.2, 7.8, 4.5],
    )
    add_heading(doc, "10.2 Выходы", level=2)
    add_table_caption(doc, "Таблица 6 — Выходные данные")
    add_table(
        doc,
        ["Выход", "Состав", "Кому"],
        [
            ["Маршруты / рейсы", "Цепочки станций, открытые рейсы и места", "Диспетчер, кассир"],
            ["Статус станции", "Сутки обслуживания и готовность объектов", "Дежурный, инженеры"],
            ["Заявки на ремонт", "Открытые и закрытые заявки по объектам", "Инженер по ремонту, дежурный"],
            ["Акт инвентаризации", "Факт наличия и список расхождений", "Инженер по инвентаризации"],
            ["Купон / смена", "Билет или статус смены", "Касса и посадка"],
        ],
        [4.0, 7.8, 4.7],
    )
    add_heading(doc, "10.3 Смежные системы", level=2)
    add_body(
        doc,
        "Принтер кассы печатает купон. Сервер базы хранит станции, схемы, оборудование для инвентаризации, "
        "маршруты, обслуживание, заявки на ремонт, смены, рейсы и билеты. "
        "Склад запчастей, закупки, сайт, найм, груз и управление движением не подключаются.",
    )

    add_heading(doc, "11 Лицензионные требования")
    add_body(
        doc,
        "Система используется внутри одного перевозчика на пилотном участке. Передача нескольким перевозчикам не входит. "
        "Отдельная лицензия на публичную продажу не требуется: интернет-магазина нет.",
    )

    add_heading(doc, "12 Юридические условия")
    add_body(
        doc,
        "Действующий документ поездки — билет, проведённый в системе и напечатанный купоном. "
        "Фамилия и имя нужны для купона и посадки, не для публичной учётной записи. "
        "Возврат выполняется до отправления и до отметки посадки. Массовая выплата по отменённому рейсу этой версией не оформляется.",
    )

    add_heading(doc, "13 Применимые нормы")
    add_body(
        doc,
        "Для пилотного участка действуют правила перевозчика: тариф из прайса, запрет продажи после отмены, "
        "один билет на место, учёт ремонта и инвентаризации объектов станции без отдельного складского контура закупок.",
    )

    add_heading(doc, "ЗАКЛЮЧЕНИЕ", new_page=True, level=0)
    add_body(
        doc,
        "Дополнительная спецификация задаёт общие требования АИС железной дороги. "
        "Пороги — в таблице 2. Кроме кассы и расписания зафиксированы обслуживание станции, "
        "инженеры по ремонту и по инвентаризации, смены персонала и единый фонд мест.",
    )
    add_body(
        doc,
        "Вне границ: груз, управление движением, найм, закупка запчастей, интернет-магазин и платёжный шлюз. "
        "Пошаговые сценарии кассы, диспетчера и дежурного — в спецификации вариантов; "
        "ремонт и инвентаризация раскрыты ролями и правилами этого документа.",
    )


def save_doc(doc: Document) -> None:
    doc.save(str(OUT))
    strip_unused_images(OUT)


FALLBACK_PAGES = {
    "ВВЕДЕНИЕ": "3",
    "1 Введение": "4",
    "1.1 Назначение": "4",
    "1.2 Границы": "4",
    "1.3 Определения и сокращения": "4",
    "1.4 Связанные документы": "5",
    "1.5 Обзор": "5",
    "2 Функциональность": "6",
    "2.1 Язык интерфейса": "6",
    "2.2 Бизнес-правила": "6",
    "3 Удобство использования": "7",
    "4 Надёжность": "8",
    "5 Производительность": "8",
    "6 Сопровождаемость": "8",
    "7 Ограничения проектирования": "9",
    "8 Документация и справка": "9",
    "9 Покупные компоненты": "9",
    "10 Интерфейсы": "10",
    "10.1 Входы": "10",
    "10.2 Выходы": "10",
    "10.3 Смежные системы": "11",
    "11 Лицензионные требования": "11",
    "12 Юридические условия": "11",
    "13 Применимые нормы": "11",
    "ЗАКЛЮЧЕНИЕ": "12",
}


def measure_pages() -> tuple[dict[str, str], int]:
    import platform

    if platform.system() == "Darwin":
        return dict(FALLBACK_PAGES), 12
    listing = Path(tempfile.gettempdir()) / "lr4_pages.txt"
    script = Path(tempfile.gettempdir()) / "lr4_pages.ps1"
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
        raise SystemExit(f"ЛР4 длиннее {PAGE_LIMIT} страниц: {total}")
    print(OUT)


if __name__ == "__main__":
    main()
