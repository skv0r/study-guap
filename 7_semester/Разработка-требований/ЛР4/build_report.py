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
PAGE_LIMIT = 12


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
        "Потоки публикации рейса, продажи, отмены, посадки, возврата и справочников заданы в спецификации вариантов. "
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
        "Внутри границ: расписание, фонд мест, продажа, возврат, посадка и справочники станций и схем вагонов. "
        "Пассажир покупает билет через кассира и рабочего места не имеет. Вне границ: грузовые перевозки, управление "
        "движением, интернет-магазин билетов, мобильное приложение, проводки и НДС, платёжный шлюз и сеть нескольких перевозчиков.",
    )
    add_heading(doc, "1.3 Определения и сокращения", level=2)
    add_body(doc, "Термины, которые дальше используются без повторного объяснения, сведены в таблице 1.")
    add_table_caption(doc, "Таблица 1 — Определения")
    add_table(
        doc,
        ["Термин", "Смысл"],
        [
            ["АИС", "Автоматизированная информационная система железной дороги"],
            ["Рейс", "Пассажирский рейс пилотного участка с датой, маршрутом и одним фондом мест"],
            ["Фонд мест", "Места рейса по схеме вагонов; у места не больше одного действующего билета"],
            ["Посадочный купон", "Печатный документ продажи: рейс, место и пассажир"],
            ["Прайс", "Тариф перевозчика; кассир не подменяет цену вручную"],
            ["Пилотный участок", "Ограниченный набор станций и рейсов первой версии"],
        ],
        [4.2, 12.3],
    )
    add_heading(doc, "1.4 Связанные документы", level=2)
    add_body(
        doc,
        "Документ о концепции и границах автоматизированной информационной системы железной дороги и спецификация "
        "вариантов UC-1–UC-6: публикация рейса, продажа билета, закрытие продажи при отмене, посадка, возврат "
        "и справочники станций и схем вагонов.",
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
        "Пошаговое поведение задано в спецификации вариантов. Здесь остаются язык интерфейса и правила, "
        "которые действуют на весь пассажирский контур.",
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
        "Рядом с порогами таблицы 2 действуют правила контура. Посадка отмечается только по действующему билету, повтор отметки запрещён. "
        "Возврат до отправления и до посадки гасит билет: место снова свободно, если рейс не отменён, и остаётся закрытым, если отменён. "
        "Цена берётся из прайса. Билет действителен только после проведения в системе. Сводка продаж за смену остаётся во второй очереди.",
        space_before=6,
    )

    add_heading(doc, "3 Удобство использования", new_page=True)
    add_body(doc, "Рабочие места и права ролей сведены в таблице 3. У пассажира рабочего места нет.")
    add_table_caption(doc, "Таблица 3 — Роли")
    add_table(
        doc,
        ["Роль", "Может", "Не может"],
        [
            ["Кассир", "Искать рейс, продавать и возвращать билет", "Менять расписание и справочники"],
            ["Диспетчер", "Публиковать, править и отменять рейс", "Продавать билеты"],
            ["Контролёр", "Отмечать посадку по действующему билету", "Менять цену, место и рейс"],
            ["Бухгалтер", "Видеть сумму возврата при сверке", "Начинать возврат"],
            ["Администратор", "Вести станции, маршруты и схемы вагонов", "Публиковать рейс и продавать билет"],
        ],
        [3.6, 6.6, 6.3],
    )
    add_body(
        doc,
        "Поиск рейса, свободные места и цена стоят на одном экране продажи. Форма пишет причину отказа: нет мест, "
        "рейс отменён, место занято, посадка уже отмечена, срок возврата истёк, в справочнике нет записи.",
        space_before=6,
    )

    add_heading(doc, "4 Надёжность", new_page=True)
    add_body(doc, "Надёжность и время оформления сведены в таблице 4. Числовые пороги те же, что в таблице 2.")
    add_table_caption(doc, "Таблица 4 — Надёжность и производительность")
    add_table(
        doc,
        ["Показатель", "Требование"],
        [
            ["Доступность", "Фонд мест, поиск и проведение билета доступны в рабочие часы касс. Плановые работы — вне этих часов. Простой в рабочее время фиксируется как инцидент, параллельный бумажный журнал мест не ведётся"],
            ["Точность места", "Занятость берётся из единого фонда в момент подтверждения. Две кассы не проводят два действующих билета на одно место. Порог — БЦ-2"],
            ["Блокировка после отмены", "Статус, который видит касса, совпадает со статусом диспетчера. Порог срока — БЦ-3"],
            ["Время оформления", "В интервал БЦ-1 входят поиск рейса, выбор места, подтверждение и печать. Поиск показывает только открытые рейсы и свободные места"],
        ],
        [4.2, 12.3],
    )
    add_heading(doc, "5 Производительность")
    add_body(doc, "Требование по времени — строка «Время оформления» таблицы 4 и порог БЦ-1 таблицы 2.")

    add_heading(doc, "6 Сопровождаемость")
    add_body(
        doc,
        "Один сервер базы во внутренней сети и рабочие места касс, диспетчерской и контроля посадки. "
        "На компьютер пассажира ничего не ставится. Перед пуском загружают справочники. "
        "Резервное копирование — отдельная инструкция администратора, не шаг правки станций.",
    )

    add_heading(doc, "7 Ограничения проектирования")
    add_body(
        doc,
        "Первая версия — пилотный участок и один перевозчик, пассажирские перевозки. Тариф из прайса, без отдельной ценовой машины. "
        "Билет не из системы не действует. Не проектируются управление движением, публичный сайт, приложение пассажира, "
        "бухгалтерский учёт, платёжный шлюз и несколько перевозчиков. При конфликте запросов приоритет у кассы, диспетчера "
        "и функций, без которых пороги таблицы 2 не измерить.",
    )

    add_heading(doc, "8 Документация и справка")
    add_body(
        doc,
        "К первой версии готовятся краткое руководство по ролям и инструкция по резервному копированию. "
        "На формах продажи и посадки достаточно короткой подсказки о причине отказа.",
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
            ["Публикация рейса", "Номер, дата, маршрут, время, схема вагонов", "Диспетчер"],
            ["Отмена рейса", "Рейс и причина", "Диспетчер"],
            ["Продажа", "Станции, дата, место, фамилия и имя, факт оплаты", "Кассир"],
            ["Возврат", "Номер билета и подтверждение", "Кассир"],
            ["Посадка", "Номер билета или купон и подтверждение", "Контролёр"],
            ["Справочник", "Станция, маршрут или схема вагона с номерами мест", "Администратор"],
        ],
        [3.8, 8.2, 4.5],
    )
    add_heading(doc, "10.2 Выходы", level=2)
    add_table_caption(doc, "Таблица 6 — Выходные данные")
    add_table(
        doc,
        ["Выход", "Состав", "Кому"],
        [
            ["Рейсы и места", "Открытые рейсы, свободные места, цена из прайса", "Кассир"],
            ["Купон", "Рейс, маршрут, дата, место, пассажир", "Печать на кассе"],
            ["Статус билета", "Действует, возвращён, посадка отмечена или рейс отменён", "Кассир и контролёр"],
            ["Блокировка", "Статус «отменён» и запрет нового оформления", "Касса"],
            ["Возврат", "Номер билета и сумма из проведённого билета", "Бухгалтер"],
        ],
        [3.6, 8.2, 4.7],
    )
    add_heading(doc, "10.3 Смежные системы", level=2)
    add_body(
        doc,
        "Принтер кассы печатает купон. Сервер базы хранит справочники, рейсы, места и билеты. "
        "Платёжные системы, публичный сайт, грузовой контур и управление движением к системе не подключаются.",
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
        "Для пилотного участка система следует правилам перевозчика из документа концепции: тариф из прайса, "
        "возврат до отправления, посадка только по действующему билету, один действующий билет на место, "
        "запрет новой продажи после отмены рейса.",
    )

    add_heading(doc, "ЗАКЛЮЧЕНИЕ", new_page=True, level=0)
    add_body(
        doc,
        "Дополнительная спецификация задаёт общие требования автоматизированной информационной системы железной дороги. "
        "Числовые пороги сведены в таблице 2. Единый фонд мест, русский интерфейс персонала и рабочие места во внутренней сети их поддерживают.",
    )
    add_body(
        doc,
        "В границы не вошли грузовые перевозки, управление движением, публичный интернет-магазин, платёжный шлюз и проводки. "
        "Пошаговые сценарии кассы, диспетчера, контролёра и администратора остаются в спецификации вариантов использования.",
    )


def save_doc(doc: Document) -> None:
    doc.save(str(OUT))
    strip_unused_images(OUT)


def measure_pages() -> tuple[dict[str, str], int]:
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
