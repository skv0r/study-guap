#!/usr/bin/env python3
"""ЛР1: сжатый документ концепции и границ по ГОСТ 7.32 / 2.105."""

from __future__ import annotations

import shutil
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt, RGBColor

BASE = Path(__file__).resolve().parent
BLANK = BASE / "guap_blanks" / "lab.docx"
OUT = BASE / "ЛР1_Отчет_Буренков_Vision_склад.docx"


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


def add_sub_h(doc, text):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=True, align="left", space_before=12, space_after=6)
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text), bold=True)


def add_struct(doc, text, *, new_page=True):
    if new_page:
        doc.add_page_break()
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="center", space_after=18)
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text), bold=True)


def add_toc_line(doc, title, page, *, sub=False):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="left", left_indent=0.5 if sub else 0)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(16.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    set_run_font(p.add_run(title))
    p.add_run("\t")
    set_run_font(p.add_run(str(page)))


def add_caption_table(doc, text):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="left", space_before=10, space_after=4)
    set_run_font(p.add_run(text))


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


def add_table(doc, headers, rows, col_widths=None):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    if col_widths:
        for row in t.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Cm(w)
    for row in t.rows:
        for cell in row.cells:
            _set_cell_border(cell)
    for j, h in enumerate(headers):
        cell = t.rows[0].cells[j]
        cell.text = ""
        p = cell.paragraphs[0]
        format_paragraph(p, first_indent=False, align="center")
        set_run_font(p.add_run(h), bold=True, size=12)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            cell = t.rows[i + 1].cells[j]
            cell.text = ""
            p = cell.paragraphs[0]
            format_paragraph(p, first_indent=False, align="left")
            set_run_font(p.add_run(v), size=12)
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
    run = p.add_run(text)
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
    put_cell(t_prep.rows[0].cells[0], "доцент, к.т.н.")
    put_cell(t_prep.rows[0].cells[4], "А.В. Бржезовский")
    put_cell(t_work.rows[0].cells[0], "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 1")
    put_cell(
        t_work.rows[1].cells[0],
        "Бизнес-требования. Документ о концепции и границах информационной системы учета реализации товаров со склада",
    )
    put_cell(t_work.rows[2].cells[0], "по курсу: Разработка и анализ требований")
    put_cell(t_stud.rows[0].cells[1], "4321")
    put_cell(t_stud.rows[0].cells[5], "Г.В. Буренков")


def setup_body_section(doc: Document):
    doc.add_section(WD_SECTION_START.NEW_PAGE)
    title_sec, body_sec = doc.sections[0], doc.sections[1]
    title_sec.footer.is_linked_to_previous = False
    for p in title_sec.footer.paragraphs:
        p.clear()
    body_sec.page_width, body_sec.page_height = Mm(210), Mm(297)
    body_sec.left_margin, body_sec.right_margin = Mm(30), Mm(15)
    body_sec.top_margin, body_sec.bottom_margin = Mm(20), Mm(20)
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


def build_body(doc: Document, toc: list[tuple[str, int, bool]]):
    add_struct(doc, "СОДЕРЖАНИЕ", new_page=False)
    for title, page, sub in toc:
        add_toc_line(doc, title, page, sub=sub)

    add_struct(doc, "ВВЕДЕНИЕ")
    add_body(
        doc,
        "Первая лабораторная работа по дисциплине «Разработка и анализ "
        "требований» посвящена бизнес-требованиям. Нужно зафиксировать, "
        "зачем создаётся система, кому она нужна, какие у неё границы и "
        "как измерить успех. Основа — главы 5 и 6 и приложение В книги [1]. "
        "Образец таблиц проблемы и позиции продукта — документ Vision [2]. "
        "Вариант 1 приложения к пособию [3]: информационная система учета "
        "реализации (продажи) товаров со склада.",
    )
    add_body(
        doc,
        "Результат работы — документ о концепции и границах, а не программа "
        "и не спецификация требований к ПО. Он задаёт общий образ продукта "
        "для следующих работ: пользовательские требования, модель данных, "
        "прототип [1], [3].",
    )

    add_section_h(doc, "1 Бизнес-требования")
    add_sub_h(doc, "1.1 Исходные данные и возможность")
    add_body(
        doc,
        "Рассматривается торговая организация со складским запасом. Цикл "
        "работы: приёмка от поставщика, хранение, заказ покупателя, резерв, "
        "отгрузка, оформление реализации, сверка остатков. Сейчас номенклатура, "
        "цены, контрагенты и отгрузки ведутся в электронных таблицах и "
        "бумажных накладных. Остаток склада и данные отдела продаж не "
        "совпадают: товар, уже обещанный одному покупателю, предлагают "
        "другому; сводку продаж за период собирают вручную в конце недели.",
    )
    add_body(
        doc,
        "Возможность бизнеса — заменить этот контур одной системой, в которой "
        "склад и продажи видят один остаток. Без этого нельзя надёжно обещать "
        "срок отгрузки и видеть вклад позиций в выручку. Система внутренняя, "
        "не витрина интернет-магазина. Формулировка проблемы по образцу [2] "
        "приведена в таблице 1.",
    )
    add_caption_table(doc, "Таблица 1 — Формулировка проблемы")
    add_table(
        doc,
        ["Элемент", "Описание"],
        [
            [
                "Проблема",
                "Разрозненный учёт остатков и продаж: данные склада и отдела продаж не совпадают.",
            ],
            ["Затрагивает", "Отдел продаж, склад, бухгалтерию, руководство."],
            ["Последствие", "Срывы отгрузки, претензии покупателей, ошибки в выручке."],
            [
                "Успешное решение",
                "Единый остаток и резерв; нельзя отгрузить сверх доступного количества; сводка продаж за период из системы.",
            ],
        ],
        col_widths=[4.0, 12.5],
    )

    add_sub_h(doc, "1.2 Цели и критерии успеха")
    add_body(
        doc,
        "Цели заданы в измеримом виде, как требует глава 5 книги [1] (таблица 2). "
        "Формулировки вида «улучшить учёт» не используются.",
    )
    add_caption_table(doc, "Таблица 2 — Бизнес-цели")
    add_table(
        doc,
        ["Код", "Цель"],
        [
            [
                "БЦ-1",
                "Оформление реализации от заявки до накладной — не более 10 минут в 80 % случаев за три месяца после ввода.",
            ],
            [
                "БЦ-2",
                "Расхождения учётного и фактического остатка — не более 3 % позиций к концу первого полугодия.",
            ],
            [
                "БЦ-3",
                "Повторная продажа уже зарезервированного товара — не более одного инцидента в месяц.",
            ],
        ],
        col_widths=[2.2, 14.3],
    )
    add_body(
        doc,
        "Проект считают успешным, если первая версия стала основным контуром "
        "учёта на одном складе. БЦ-1 проверяют по выборке оформленных "
        "документов, БЦ-2 — по акту инвентаризации, БЦ-3 — по журналу "
        "инцидентов отдела продаж. Если сотрудники продолжат параллельные "
        "таблицы, БЦ-2 измерить нельзя.",
    )

    add_sub_h(doc, "1.3 Положение о концепции")
    add_body(
        doc,
        "Положение составлено по шаблону ключевых слов, приведённому в главе 5 "
        "книги [1]. Для сотрудников склада и отдела продаж, которым нужно "
        "оформлять реализацию с контролем остатка, система является "
        "информационной системой учёта продаж со склада. Она даёт единый "
        "доступ к номенклатуре, доступному остатку, резерву и документам "
        "реализации. В отличие от таблиц и бумажных накладных продукт не "
        "допускает отгрузки сверх доступного количества и формирует сводку "
        "продаж без ручной сборки.",
    )

    add_sub_h(doc, "1.4 Риски, предположения и зависимости")
    add_body(doc, "Основные бизнес-риски приведены в таблице 3.")
    add_caption_table(doc, "Таблица 3 — Бизнес-риски")
    add_table(
        doc,
        ["Риск", "Реагирование"],
        [
            [
                "Параллельный учёт в таблицах",
                "Отгрузка только по документу в системе; назначить владельца данных",
            ],
            [
                "Неполные начальные остатки",
                "Инвентаризация перед опытной эксплуатацией",
            ],
            [
                "Расширение границ (магазин, WMS, производство)",
                "Фиксировать исключения в разделе 3.3; лишнее откладывать на выпуск 2",
            ],
        ],
        col_widths=[7.5, 9.0],
    )
    add_body(
        doc,
        "Предположения: один основной склад; номенклатура — тысячи позиций, "
        "не миллионы; цены берутся из прайса организации, без сложной "
        "скидочной машины; работа во внутренней сети в рабочие дни. "
        "Зависимости: согласованные справочники номенклатуры и контрагентов "
        "к запуску; сервер СУБД во внутренней сети; бухгалтерия принимает "
        "печать накладной из системы или выгрузку реквизитов. Если складов "
        "несколько, в первую версию это не включают — это изменение границ.",
    )

    add_section_h(doc, "2 Заинтересованные лица и пользователи")
    add_body(
        doc,
        "Глава 6 книги [1] требует не считать «пользователя» одной группой. "
        "Заказчик и владелец бизнес-требований — руководитель организации: "
        "он утверждает цели и разрешает споры о границах. Покупатель — "
        "косвенный участник: в интерфейс не входит, получает отгрузку через "
        "менеджера.",
    )
    add_body(
        doc,
        "Классы пользователей различаются задачами и правами (таблица 4). "
        "Привилегированные классы — менеджер по продажам и кладовщик: без "
        "их представителей нельзя закрыть БЦ-1 и БЦ-3. Для каждого из них "
        "назначается сторонник продукта, который передаёт и согласовывает "
        "пользовательские требования после утверждения этого документа [1].",
    )
    add_caption_table(doc, "Таблица 4 — Заинтересованные лица и классы пользователей")
    add_table(
        doc,
        ["Роль", "Тип", "Задачи и права"],
        [
            [
                "Руководитель",
                "Заказчик",
                "Утверждает цели и границы; смотрит сводки",
            ],
            [
                "Менеджер по продажам",
                "Пользователь",
                "Заказ, резерв, реализация; остаток вручную не правит",
            ],
            [
                "Кладовщик",
                "Пользователь",
                "Приход и отгрузка по резерву; цены не меняет",
            ],
            [
                "Заведующий складом",
                "Пользователь",
                "Инвентаризация и корректировка остатка по акту",
            ],
            [
                "Бухгалтер",
                "Пользователь",
                "Печать и сверка документов реализации",
            ],
            [
                "Администратор",
                "Пользователь",
                "Учётные записи, справочники, резервные копии",
            ],
        ],
        col_widths=[4.4, 3.2, 8.9],
    )

    add_section_h(doc, "3 Рамки проекта")
    add_sub_h(doc, "3.1 Основные функции")
    add_body(
        doc,
        "Функции пронумерованы как FEATn по образцу [2], чтобы на них "
        "ссылаться в следующих работах (таблица 5).",
    )
    add_caption_table(doc, "Таблица 5 — Функции системы")
    add_table(
        doc,
        ["Код", "Функция", "Версия"],
        [
            ["FEAT1", "Справочник номенклатуры", "1"],
            ["FEAT2", "Справочник покупателей", "1"],
            ["FEAT3", "Приход на склад", "1"],
            ["FEAT4", "Доступный и зарезервированный остаток", "1"],
            ["FEAT5", "Заказ покупателя и резерв", "1"],
            ["FEAT6", "Реализация с запретом отгрузки сверх остатка", "1"],
            ["FEAT7", "Печать накладной / счёта", "1"],
            ["FEAT8", "Права по ролям из таблицы 4", "1"],
            ["FEAT9", "Возврат товара", "2"],
            ["FEAT10", "Сводка реализации за период", "2"],
            ["FEAT11", "Инвентаризация по акту", "2"],
        ],
        col_widths=[2.2, 11.5, 2.8],
    )

    add_sub_h(doc, "3.2 Объём версий")
    add_body(
        doc,
        "Первая версия закрывает БЦ-1 и БЦ-3 на одном складе: FEAT1–FEAT8. "
        "Она должна надёжно вести остаток и резерв; документ реализации "
        "нельзя провести дважды по тем же строкам заказа. Удобство отдельных "
        "форм можно упростить. Во вторую версию откладываются FEAT9–FEAT11, "
        "штрихкодирование и, при необходимости, партионный учёт.",
    )
    add_body(
        doc,
        "В границы не входят: интернет-магазин и кабинет покупателя; "
        "производство и модуль закупок; регламентированный бухгалтерский учёт "
        "(проводки, НДС); адресное хранение и мобильное место кладовщика; "
        "несколько складов в первой версии. Исключения записаны, чтобы "
        "границы не расползлись [1].",
    )

    add_sub_h(doc, "3.3 Положение о позиции продукта")
    add_body(doc, "Положение о позиции составлено по шаблону [2] и приведено в таблице 6.")
    add_caption_table(doc, "Таблица 6 — Положение о позиции продукта")
    add_table(
        doc,
        ["Элемент", "Формулировка"],
        [
            ["Для", "торговой организации со складским запасом"],
            ["которая", "теряет время на сверку остатков и срывает отгрузки"],
            ["система", "информационная система учета реализации товаров со склада"],
            ["которая", "ведёт единый остаток, резерв и документы продажи"],
            ["В отличие от", "учёта в таблицах и бумажных накладных"],
            ["наш продукт", "не даёт отгрузить сверх остатка и даёт сводку продаж из одной базы"],
        ],
        col_widths=[3.5, 13.0],
    )

    add_section_h(doc, "4 Бизнес-контекст")
    add_body(
        doc,
        "Ведущий фактор первой версии — достоверность остатка и запрет "
        "отгрузки сверх доступного количества. Ограничение — один склад и "
        "исключения из подраздела 3.2. Степень свободы — состав отчётов сверх "
        "сводки и удобство отдельных форм: их можно упростить, чтобы удержать "
        "срок опытной эксплуатации [1]. При конфликте запросов приоритет у "
        "менеджера по продажам, кладовщика и функций, без которых не измерить "
        "БЦ-1 и БЦ-3.",
    )
    add_body(
        doc,
        "Развёртывание — во внутренней сети, один сервер СУБД, рабочие места "
        "склада и отдела продаж. Перед опытной эксплуатацией проводят "
        "инвентаризацию и загружают начальные остатки; параллельные таблицы "
        "с этой даты не ведут. Обучение — короткий сценарий по ролям "
        "таблицы 4. Документация первой версии — краткое руководство "
        "пользователя и инструкция по резервному копированию.",
    )

    add_struct(doc, "ЗАКЛЮЧЕНИЕ")
    add_body(
        doc,
        "Составлен документ концепции и границ для учёта продаж со склада "
        "(вариант 1). Зафиксированы проблема, цели БЦ-1–БЦ-3, критерии "
        "успеха, положение о концепции, риски и предположения, роли и "
        "границы первой версии (FEAT1–FEAT8). Исключены интернет-магазин, "
        "производство и полноценный бухгалтерский контур. Документ закрывает "
        "шаблон глав 5–6 и приложения В книги [1] и приёмы примера [2]. "
        "Он служит основой следующих работ, но не заменяет спецификацию "
        "требований к ПО.",
    )

    add_struct(doc, "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ")
    sources = [
        "Вигерс К., Битти Дж. Разработка требований к программному обеспечению. "
        "3-е изд., доп. / пер. с англ. — М. : Русская редакция ; СПб. : БХВ-Петербург, 2014. — 736 с.",
        "RU Financial Services. Vision. Version 1.5. RU e-st Case Study 3. — IBM Corp., 2003, 2010. — 16 с.",
        "Богословская Н. В., Бржезовский А. В., Семененко Т. В. Разработка и анализ требований: "
        "средства прототипирования : учеб.-метод. пособие. — СПб. : ГУАП, 2022. — 70 с.",
    ]
    for i, src in enumerate(sources, 1):
        p = doc.add_paragraph()
        format_paragraph(p)
        set_run_font(p.add_run(f"{i}. {src}"))


TOC = [
    ("ВВЕДЕНИЕ", 3, False),
    ("1 Бизнес-требования", 4, False),
    ("1.1 Исходные данные и возможность", 4, True),
    ("1.2 Цели и критерии успеха", 4, True),
    ("1.3 Положение о концепции", 5, True),
    ("1.4 Риски, предположения и зависимости", 5, True),
    ("2 Заинтересованные лица и пользователи", 7, False),
    ("3 Рамки проекта", 8, False),
    ("3.1 Основные функции", 8, True),
    ("3.2 Объём версий", 8, True),
    ("3.3 Положение о позиции продукта", 9, True),
    ("4 Бизнес-контекст", 10, False),
    ("ЗАКЛЮЧЕНИЕ", 11, False),
    ("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", 12, False),
]


def main():
    shutil.copy(BLANK, OUT)
    doc = Document(str(OUT))
    fill_title(doc)
    setup_body_section(doc)
    build_body(doc, TOC)
    doc.save(str(OUT))
    print(OUT)


if __name__ == "__main__":
    main()
