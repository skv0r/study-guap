#!/usr/bin/env python3
"""Сборка отчёта ЛР1. В текст отчёта не входит."""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Pt, Twips
from lxml import etree

BASE = Path(__file__).resolve().parent
REPO = BASE.parents[1]
TEMPLATE = REPO / "Web-программирование" / "ЛР1_Отчет_Буренков_PairWise.docx"
OUT = BASE / "ЛР1_Отчет_Буренков_Vision_ЖД.docx"

TOPIC = (
    "Бизнес-требования. Документ о концепции и границах "
    "автоматизированной информационной системы железной дороги"
)
COURSE = "по курсу: Разработка и анализ требований"

TITLE_TEXT = {
    "ассистент": "доцент, к.т.н.",
    "В.В. Жукалин": "А.В. Бржезовский",
    "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 1": "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 1",
    "Создание проекта Next.js с использованием React, TypeScript и App Router": TOPIC,
    "по курсу: Web-программирование": COURSE,
}

USABLE = Cm(16.5)


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
        format_paragraph(p, first_indent=False, align="center", space_after=12, single=False)
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
    return p


def add_table_caption(doc, text):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="left", space_before=12, space_after=6, single=True)
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text))
    return p


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
    tbl = table._tbl
    tblPr = tbl.tblPr
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


def clone_title(number: str) -> Document:
    shutil.copy(TEMPLATE, OUT)
    doc = Document(str(OUT))
    body = doc.element.body
    children = list(body)
    for child in children[12:-1]:
        body.remove(child)
    replacements = dict(TITLE_TEXT)
    replacements["ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 1"] = f"ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № {number}"
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
        if name.startswith("word/") and name.endswith(".xml") and "/_rels/" not in name:
            used.update(etree.fromstring(data).xpath("//@*[contains(local-name(), 'id')]") if False else [])
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
            target = "word/" + rel.get("Target").replace("\\", "/")
            removed.append(target)
            root.remove(rel)
    files[rel_name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    for name in removed:
        files.pop(name, None)
    tmp = path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    tmp.replace(path)


def build_body(doc: Document, pages: dict[str, str]) -> None:
    add_heading(doc, "СОДЕРЖАНИЕ", level=0)
    toc = [
        ("ВВЕДЕНИЕ", False),
        ("1 Бизнес-требования", False),
        ("1.1 Исходные данные и возможность", True),
        ("1.2 Цели и критерии успеха", True),
        ("1.3 Положение о концепции", True),
        ("1.4 Риски, предположения и зависимости", True),
        ("2 Заинтересованные лица и пользователи", False),
        ("3 Рамки проекта", False),
        ("3.1 Основные функции", True),
        ("3.2 Объём версий", True),
        ("3.3 Положение о позиции продукта", True),
        ("4 Бизнес-контекст", False),
        ("ЗАКЛЮЧЕНИЕ", False),
    ]
    for title, sub in toc:
        add_toc_line(doc, title, pages.get(title, "…"), sub=sub)

    add_heading(doc, "ВВЕДЕНИЕ", new_page=True, level=0)
    add_body(
        doc,
        "Первая лабораторная работа по дисциплине «Разработка и анализ требований» "
        "посвящена бизнес-требованиям. Нужно зафиксировать, зачем создаётся система, "
        "кому она нужна, какие у неё границы и как измерить успех. Предмет работы — "
        "автоматизированная информационная система железной дороги (вариант 1).",
    )
    add_body(
        doc,
        "Результат работы — документ о концепции и границах (Vision), а не программа "
        "и не спецификация требований к ПО. Он задаёт общий образ продукта для "
        "следующих разделов: пользовательские требования, модель данных, прототип.",
    )

    add_heading(doc, "1 Бизнес-требования", new_page=True)
    add_heading(doc, "1.1 Исходные данные и возможность", level=2)
    add_body(
        doc,
        "Рассматривается пассажирский перевозчик на железной дороге. Цикл работы: "
        "ведение справочника станций и маршрутов, публикация расписания рейсов, учёт "
        "мест в вагонах, продажа и возврат билетов в кассе, контроль посадки по "
        "оформленным местам. Сейчас расписание ведут в таблицах, занятость мест — "
        "в журнале кассы, а отчёт о продажах сводят вручную в конце смены. Одно и "
        "то же место иногда продают двум пассажирам; при смене рейса кассир узнаёт "
        "об отмене с опозданием.",
    )
    add_body(
        doc,
        "Возможность бизнеса — заменить разрозненный контур одной системой, в которой "
        "кассы, диспетчер расписания и контроль посадки видят один актуальный фонд "
        "мест. Без этого нельзя гарантировать билет на конкретное место и вовремя "
        "закрывать продажи при отмене рейса. Система внутренняя для персонала "
        "перевозчика; публичный сайт бронирования в первую версию не входит. "
        "Формулировка проблемы приведена в таблице 1.",
    )
    add_table_caption(doc, "Таблица 1 — Формулировка проблемы")
    add_table(
        doc,
        ["Элемент", "Описание"],
        [
            [
                "Проблема",
                "Разрозненный учёт расписания, мест и продаж: касса и диспетчер работают с разными данными.",
            ],
            ["Затрагивает", "Кассы, диспетчерскую службу расписания, контроль посадки, руководство."],
            ["Последствие", "Двойная продажа мест, претензии пассажиров, потери выручки, срывы посадки."],
            [
                "Успешное решение",
                "Единый фонд мест и билетов; нельзя продать занятое место; отмена рейса сразу закрывает продажу; сводка продаж за смену из системы.",
            ],
        ],
        [4.0, 12.5],
    )

    add_heading(doc, "1.2 Цели и критерии успеха", level=2)
    add_body(
        doc,
        "Цели заданы в измеримом виде (таблица 2). Формулировки вида «улучшить учёт» не используются.",
    )
    add_table_caption(doc, "Таблица 2 — Бизнес-цели")
    add_table(
        doc,
        ["Код", "Цель"],
        [
            [
                "БЦ-1",
                "Оформление билета от запроса пассажира до печати — не более 3 минут в 90 % случаев за три месяца после ввода.",
            ],
            [
                "БЦ-2",
                "Двойная продажа одного места — не более одного подтверждённого инцидента в месяц.",
            ],
            [
                "БЦ-3",
                "После отмены рейса новые продажи по нему невозможны не позднее чем через 5 минут в 95 % случаев.",
            ],
        ],
        [2.2, 14.3],
    )
    add_body(
        doc,
        "Проект считают успешным, если первая версия стала основным контуром продажи "
        "билетов на выбранном участке сети (набор станций и рейсов пилота). БЦ-1 "
        "проверяют по журналу операций кассы, БЦ-2 — по журналу инцидентов, БЦ-3 — "
        "по времени между записью об отмене рейса и блокировкой продажи. Если кассы "
        "продолжат параллельные журналы на бумаге, БЦ-2 измерить нельзя.",
        space_before=6,
    )

    add_heading(doc, "1.3 Положение о концепции", level=2)
    add_body(
        doc,
        "Положение составлено по шаблону ключевых слов. Для кассиров и диспетчеров "
        "расписания, которым нужно продавать билеты с контролем занятости мест, "
        "система является автоматизированной информационной системой железной дороги. "
        "Она даёт единый доступ к станциям, маршрутам, расписанию рейсов, фонду мест "
        "и документам продажи. В отличие от таблиц и бумажных журналов продукт не "
        "допускает продажу уже занятого места и формирует сводку продаж из одной базы, "
        "без ручного сведения сменных отчётов.",
    )

    add_heading(doc, "1.4 Риски, предположения и зависимости", level=2)
    add_body(doc, "Основные бизнес-риски приведены в таблице 3.")
    add_table_caption(doc, "Таблица 3 — Бизнес-риски")
    add_table(
        doc,
        ["Риск", "Реагирование"],
        [
            [
                "Параллельная продажа «в обход» системы",
                "Билет действителен только из системы; назначить владельца данных по фонду мест",
            ],
            [
                "Неполные или устаревшие справочники станций и схем вагонов",
                "Нормализация справочников и инвентаризация мест перед опытной эксплуатацией",
            ],
            [
                "Расширение границ (груз, онлайн-витрина, управление движением)",
                "Фиксировать исключения в подразделе 3.2; лишнее откладывать на выпуск 2",
            ],
        ],
        [7.0, 9.5],
    )
    add_body(
        doc,
        "Предположения: пилот — один участок сети с ограниченным числом станций; "
        "пассажирские перевозки, без грузового контура; тарифы берутся из прайса "
        "перевозчика без сложной динамической ценовой машины; работа касс в рабочие "
        "часы станций пилота. Зависимости: согласованные справочники станций, "
        "маршрутов и схем вагонов к запуску; сервер СУБД во внутренней сети; "
        "кассы принимают печать билета из системы. Если нужен публичный сайт "
        "бронирования или несколько перевозчиков, в первую версию это не включают — "
        "это изменение границ.",
        space_before=6,
    )

    add_heading(doc, "2 Заинтересованные лица и пользователи", new_page=True)
    add_body(
        doc,
        "Пользователя нельзя считать одной группой. Заказчик и владелец "
        "бизнес-требований — руководитель перевозчика: он утверждает цели и разрешает "
        "споры о границах. Пассажир — косвенный участник первой версии: в интерфейс "
        "не входит, покупает билет через кассира.",
    )
    add_body(
        doc,
        "Классы пользователей различаются задачами и правами (таблица 4). "
        "Привилегированные классы — кассир и диспетчер расписания: без их "
        "представителей нельзя закрыть БЦ-1–БЦ-3. Для каждого из них назначается "
        "сторонник продукта, который передаёт и согласовывает пользовательские "
        "требования после утверждения этого документа.",
    )
    add_table_caption(doc, "Таблица 4 — Заинтересованные лица и классы пользователей")
    add_table(
        doc,
        ["Роль", "Тип", "Задачи и права"],
        [
            ["Руководитель перевозчика", "Заказчик", "Утверждает цели и границы; смотрит сводки продаж"],
            ["Кассир", "Пользователь", "Поиск рейса, продажа и возврат билета; расписание не меняет"],
            ["Диспетчер расписания", "Пользователь", "Публикация, изменение и отмена рейсов; билеты не продаёт"],
            ["Контролёр посадки", "Пользователь", "Проверка билета и отметка посадки; цены не меняет"],
            ["Бухгалтер", "Пользователь", "Сверка сменной выручки и печатных форм"],
            ["Администратор", "Пользователь", "Учётные записи, справочники, резервные копии"],
        ],
        [4.5, 3.2, 8.8],
    )

    add_heading(doc, "3 Рамки проекта", new_page=True)
    add_heading(doc, "3.1 Основные функции", level=2)
    add_body(doc, "Функции пронумерованы, чтобы на них ссылаться в следующих работах (таблица 5).")
    add_table_caption(doc, "Таблица 5 — Функции системы")
    add_table(
        doc,
        ["Код", "Функция", "Версия"],
        [
            ["FEAT1", "Справочник станций и маршрутов", "1"],
            ["FEAT2", "Схемы вагонов и фонд мест рейса", "1"],
            ["FEAT3", "Расписание рейсов (публикация и изменение)", "1"],
            ["FEAT4", "Поиск рейса и свободных мест", "1"],
            ["FEAT5", "Продажа билета с резервом места", "1"],
            ["FEAT6", "Запрет продажи занятого места и по отменённому рейсу", "1"],
            ["FEAT7", "Печать билета / посадочного купона", "1"],
            ["FEAT8", "Права по ролям из таблицы 4", "1"],
            ["FEAT9", "Возврат билета", "2"],
            ["FEAT10", "Сводка продаж за смену / период", "2"],
            ["FEAT11", "Отметка посадки контролёром", "2"],
        ],
        [2.4, 11.3, 2.8],
    )

    add_heading(doc, "3.2 Объём версий", level=2)
    add_body(
        doc,
        "Первая версия закрывает БЦ-1–БЦ-3 на пилотном участке: FEAT1–FEAT8. Она "
        "должна надёжно вести фонд мест и продажу; одно место нельзя закрепить за "
        "двумя действующими билетами. Удобство отдельных форм можно упростить. "
        "Во вторую версию откладываются FEAT9–FEAT11, самообслуживание пассажира "
        "и, при необходимости, льготные категории тарифов.",
    )
    add_body(
        doc,
        "В границы не входят: грузовые перевозки и вагонный парк как отдельный "
        "контур; управление движением поездов и диспетчеризация пути; публичный "
        "интернет-магазин билетов и мобильное приложение пассажира; "
        "регламентированный бухгалтерский учёт (проводки, НДС); интеграция с "
        "платёжными системами сверх фиксации факта оплаты в кассе; сеть нескольких "
        "независимых перевозчиков в первой версии. Исключения записаны, чтобы "
        "границы не расползлись.",
    )

    add_heading(doc, "3.3 Положение о позиции продукта", level=2)
    add_body(doc, "Положение о позиции приведено в таблице 6.")
    add_table_caption(doc, "Таблица 6 — Положение о позиции продукта")
    add_table(
        doc,
        ["Элемент", "Формулировка"],
        [
            ["Для", "пассажирского железнодорожного перевозчика"],
            ["которая", "теряет время на сверку мест и допускает двойную продажу"],
            ["система", "автоматизированная информационная система железной дороги"],
            ["которая", "ведёт единое расписание, фонд мест и документы продажи билетов"],
            ["В отличие от", "учёта в таблицах и бумажных журналах кассы"],
            [
                "наш продукт",
                "не даёт продать занятое место и даёт сводку продаж из одной базы",
            ],
        ],
        [3.6, 12.9],
    )

    add_heading(doc, "4 Бизнес-контекст", new_page=True)
    add_body(
        doc,
        "Ведущий фактор первой версии — достоверность фонда мест и запрет продажи "
        "занятого или отменённого рейса. Ограничение — пилотный участок и исключения "
        "из подраздела 3.2. Степень свободы — состав отчётов сверх сменной сводки и "
        "удобство отдельных форм: их можно упростить, чтобы удержать срок опытной "
        "эксплуатации. При конфликте запросов приоритет у кассира, диспетчера "
        "расписания и функций, без которых не измерить БЦ-1–БЦ-3.",
    )
    add_body(
        doc,
        "Развёртывание — во внутренней сети, один сервер СУБД, рабочие места касс "
        "и диспетчерской пилотного участка. Перед опытной эксплуатацией загружают "
        "справочники станций, маршрутов и схем вагонов; параллельные журналы мест "
        "с этой даты не ведут. Обучение — короткий сценарий по ролям таблицы 4. "
        "Документация первой версии — краткое руководство пользователя и инструкция "
        "по резервному копированию.",
    )

    add_heading(doc, "ЗАКЛЮЧЕНИЕ", new_page=True, level=0)
    add_body(
        doc,
        "Составлен документ концепции и границ для автоматизированной информационной "
        "системы железной дороги. Зафиксированы проблема, цели БЦ-1–БЦ-3, критерии "
        "успеха, положение о концепции, риски и предположения, роли и границы первой "
        "версии (FEAT1–FEAT8). Исключены грузовой контур, управление движением и "
        "публичная онлайн-продажа. Документ задаёт общий образ продукта и не заменяет "
        "спецификацию требований к ПО.",
    )


HEADINGS = [
    "ВВЕДЕНИЕ",
    "1 Бизнес-требования",
    "1.1 Исходные данные и возможность",
    "1.2 Цели и критерии успеха",
    "1.3 Положение о концепции",
    "1.4 Риски, предположения и зависимости",
    "2 Заинтересованные лица и пользователи",
    "3 Рамки проекта",
    "3.1 Основные функции",
    "3.2 Объём версий",
    "3.3 Положение о позиции продукта",
    "4 Бизнес-контекст",
    "ЗАКЛЮЧЕНИЕ",
]


def save_doc(doc: Document) -> None:
    doc.save(str(OUT))
    strip_unused_images(OUT)


def measure_pages() -> dict[str, str]:
    import platform
    import subprocess
    import tempfile

    listing = Path(tempfile.gettempdir()) / "lr1_pages.txt"
    if platform.system() == "Darwin":
        script = Path(tempfile.gettempdir()) / "lr1_pages.applescript"
        # wdActiveEndPageNumber = 3
        script.write_text(
            "\n".join(
                [
                    f'set outPath to POSIX file "{listing}"',
                    f'set docPath to POSIX file "{OUT}"',
                    "tell application \"Microsoft Word\"",
                    "  set wasRunning to running",
                    "  open docPath",
                    "  set theDoc to active document",
                    "  set pageLines to {}",
                    "  set paraCount to count of paragraphs of theDoc",
                    "  repeat with i from 1 to paraCount",
                    "    set p to paragraph i of theDoc",
                    "    set t to content of text object of p",
                    "    set AppleScript's text item delimiters to {return, linefeed, character id 7}",
                    "    set tParts to text items of t",
                    "    set AppleScript's text item delimiters to \"\"",
                    "    set t to tParts as text",
                    "    set t to my trim(t)",
                    "    if (t is not \"\") and (t does not contain tab) then",
                    "      set pg to (get information for (create range theDoc start (start of content of text object of p) end (end of content of text object of p)) information type active end adjusted page number)",
                    "      set end of pageLines to ((pg as text) & tab & t)",
                    "    end if",
                    "  end repeat",
                    "  close theDoc saving no",
                    "  if not wasRunning then quit",
                    "end tell",
                    "set text item delimiters to linefeed",
                    "set pageText to pageLines as text",
                    "set text item delimiters to \"\"",
                    "do shell script \"printf %s \" & quoted form of pageText & \" > \" & quoted form of POSIX path of outPath",
                    "on trim(s)",
                    "  set s to s as text",
                    "  repeat while s begins with \" \" or s begins with tab",
                    "    set s to text 2 thru -1 of s",
                    "  end repeat",
                    "  repeat while s ends with \" \" or s ends with tab",
                    "    if (count of s) < 1 then exit repeat",
                    "    set s to text 1 thru -2 of s",
                    "  end repeat",
                    "  return s",
                    "end trim",
                ]
            ),
            encoding="utf-8",
        )
        subprocess.run(["osascript", str(script)], check=True)
    else:
        script = Path(tempfile.gettempdir()) / "lr1_pages.ps1"
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


FALLBACK_PAGES = {
    "ВВЕДЕНИЕ": "3",
    "1 Бизнес-требования": "4",
    "1.1 Исходные данные и возможность": "4",
    "1.2 Цели и критерии успеха": "4",
    "1.3 Положение о концепции": "5",
    "1.4 Риски, предположения и зависимости": "5",
    "2 Заинтересованные лица и пользователи": "7",
    "3 Рамки проекта": "8",
    "3.1 Основные функции": "8",
    "3.2 Объём версий": "8",
    "3.3 Положение о позиции продукта": "8",
    "4 Бизнес-контекст": "10",
    "ЗАКЛЮЧЕНИЕ": "11",
}


def main() -> None:
    doc = clone_title("1")
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
    doc = clone_title("1")
    build_body(doc, pages)
    save_doc(doc)
    print(pages)
    print(OUT)


if __name__ == "__main__":
    main()
