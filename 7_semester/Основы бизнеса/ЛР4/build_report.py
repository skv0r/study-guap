#!/usr/bin/env python3
"""Практическая работа № 4. Основы бизнеса: расширение мощности ООО «Жар-короб»."""

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
BLANK = Path(r"C:\Users\skvor\OneDrive\Desktop\GitHub\study-guap\7_semester\_титулы\ТИТУЛ_ОБ_ПР1.docx")
OUT = BASE / "ПР4_Отчет_Буренков_ОБ.docx"
PDF = BASE / "ПР4_Отчет_Буренков_ОБ.pdf"
FIG_GANTT = BASE / "figures" / "gantt.png"

BODY = 12
TBL = 10

TOC_DEF = [
    ("ВВЕДЕНИЕ", "intro", 0),
    ("1 Цель работы", "s1", 0),
    ("2 Действующая мощность и прогноз спроса", "s2", 0),
    ("3 Три сценария расширения", "s3", 0),
    ("4 Дополнительный выпуск и выручка", "s4", 0),
    ("5 Дополнительные расходы и амортизация", "s5", 0),
    ("6 Денежный поток и чистая приведённая стоимость", "s6", 0),
    ("7 Срок окупаемости", "s7", 0),
    ("8 Риски сценариев", "s8", 0),
    ("9 Выбор сценария", "s9", 0),
    ("10 Календарный план реализации", "s10", 0),
    ("ЗАКЛЮЧЕНИЕ", "conc", 0),
]

SEARCH = {key: title for title, key, _ in TOC_DEF}


def set_run_font(run, size=BODY, bold=False, name="Times New Roman", italic=False):
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


def add_struct(doc, text, *, new_page=True):
    if new_page:
        doc.add_page_break()
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="center", space_after=12)
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text), bold=True)


def add_section_h(doc, text, *, new_page=True):
    if new_page:
        doc.add_page_break()
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=True, align="left", space_before=0, space_after=6)
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text), bold=True)


def add_caption_table(doc, text):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="left", space_before=8, space_after=2, line_spacing=1.0)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text))


def add_figure(doc, path, caption, w=17.5):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="center", space_before=8)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(path), width=Cm(w))
    c = doc.add_paragraph()
    format_paragraph(c, first_indent=False, align="center", space_before=6, space_after=8)
    set_run_font(c.add_run(caption))


def add_toc_line(doc, title, page, *, indent=0):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="left", left_indent=0.75 * indent)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(17.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    set_run_font(p.add_run(title))
    p.add_run("\t")
    set_run_font(p.add_run(str(page)))


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


def add_table(doc, headers, rows, widths_cm=None, size=TBL):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    usable = Cm(18.4)
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
        trPr = row._tr.get_or_add_trPr()
        cant = OxmlElement("w:cantSplit")
        trPr.append(cant)
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
        format_paragraph(p, first_indent=False, align="center", line_spacing=1.0)
        set_run_font(p.add_run(h), bold=True, size=size)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            cell = t.rows[i + 1].cells[j]
            cell.text = ""
            p = cell.paragraphs[0]
            align = "center" if j > 0 else "left"
            format_paragraph(p, first_indent=False, align=align, line_spacing=1.0)
            set_run_font(p.add_run(v), size=size, bold=(i == len(rows) - 1 and "Итого" in str(row[0])))
    spacer = doc.add_paragraph()
    format_paragraph(spacer, first_indent=False, line_spacing=1.0, space_after=4)
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


def fill_title(doc: Document):
    t_work = doc.tables[1]
    t_work.rows[0].cells[0].paragraphs[0].runs[0].text = "ОТЧЕТ О ПРАКТИЧЕСКОЙ РАБОТЕ № 4"
    t_work.rows[1].cells[0].paragraphs[0].runs[0].text = (
        "План расширения производства ООО «Жар-короб»: три сценария"
    )


def setup_body_section(doc: Document):
    doc.add_section(WD_SECTION_START.NEW_PAGE)
    title_sec, body_sec = doc.sections[0], doc.sections[1]
    title_sec.page_width, title_sec.page_height = Mm(210), Mm(297)
    title_sec.left_margin, title_sec.right_margin = Mm(30), Mm(10)
    title_sec.top_margin, title_sec.bottom_margin = Mm(20), Mm(20)
    body_sec.page_width, body_sec.page_height = Mm(210), Mm(297)
    body_sec.left_margin = Mm(12.7)
    body_sec.right_margin = Mm(12.7)
    body_sec.top_margin = Mm(12.7)
    body_sec.bottom_margin = Mm(16)
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


def make_gantt():
    import matplotlib.pyplot as plt

    FIG_GANTT.parent.mkdir(parents=True, exist_ok=True)
    plt.rcParams["font.family"] = "DejaVu Sans"
    tasks = [
        ("Техническое задание на раскройный комплекс", 0, 2),
        ("Запрос предложений и выбор поставщика", 2, 2),
        ("Договор и оплата", 4, 3),
        ("Поставка и монтаж", 7, 5),
        ("Набор и обучение второй смены", 8, 6),
        ("Пробная партия и приёмка", 12, 2),
        ("Промышленный запуск двух смен", 14, 2),
    ]
    fig, ax = plt.subplots(figsize=(12.2, 6.4), dpi=170)
    ax.set_xlim(0, 17)
    ax.set_ylim(-0.6, len(tasks) - 0.4)
    colors = ["#d9d9d9", "#d9d9d9", "#d9d9d9", "#c5d5c5", "#c5d5c5", "#f2e2b8", "#a8c4a8"]
    for i, (name, start, dur) in enumerate(tasks):
        y = len(tasks) - 1 - i
        ax.barh(y, dur, left=start, height=0.55, color=colors[i], edgecolor="black", linewidth=0.9)
        ax.text(-0.15, y, name, ha="right", va="center", fontsize=8.2)
    ax.axvline(16, color="black", linestyle=(0, (3, 2)), linewidth=0.9)
    ax.text(16.1, len(tasks) - 0.35, "запуск", fontsize=8)
    ax.set_yticks([])
    ax.set_xlabel("Недели от начала проекта")
    ax.set_xticks(range(0, 17, 2))
    for spine in ("top", "right", "left"):
        ax.spines[spine].set_visible(False)
    fig.tight_layout(pad=0.3)
    fig.savefig(FIG_GANTT, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def build_body(doc: Document, toc_pages: dict):
    add_struct(doc, "СОДЕРЖАНИЕ", new_page=False)
    for title, key, ind in TOC_DEF:
        add_toc_line(doc, title, toc_pages.get(key, "…"), indent=ind)

    add_struct(doc, "ВВЕДЕНИЕ")
    add_body(
        doc,
        "ООО «Жар-короб» шьёт двухконтурные термокороба. В предыдущих работах зафиксированы "
        "форма общества, поставщик листа вспененного полиэтилена и организационная структура. "
        "Сейчас ограничение смещается на мощность цеха. Действующий выпуск 400 коробов в месяц "
        "при мощности одной смены 480 коробов. План роста спроса 35 процентов выводит пиковые "
        "месяцы выше 480 изделий, поэтому без расширения часть заявок будет потеряна.",
    )
    add_body(
        doc,
        "Мощность — наибольший выпуск, который цех может дать за месяц при заданном режиме "
        "работы. Загрузка — отношение фактического выпуска к мощности: 400 / 480 = 0,833, "
        "то есть 83 процента. Сравниваются три сценария: экстенсивный, интенсивный и гибридный. "
        "Критерии выбора — дополнительная выручка при ограничении спросом, денежный поток, "
        "чистая приведённая стоимость за три года, срок окупаемости и производственный риск.",
    )

    add_section_h(doc, "1 Цель работы")
    add_body(
        doc,
        "Цель — выбрать сценарий расширения производственной мощности ООО «Жар-короб» "
        "по чистой приведённой стоимости и рискам и составить календарный план реализации.",
    )
    add_body(doc, "Задачи работы:")
    add_body(doc, "1) задать действующую мощность, цену изделия и помесячный прогноз спроса;")
    add_body(doc, "2) описать три сценария с капитальными затратами и приростом мощности;")
    add_body(doc, "3) рассчитать дополнительный выпуск как минимум из прогноза и новой мощности минус текущий спрос;")
    add_body(doc, "4) рассчитать дополнительную выручку, операционные расходы и амортизацию;")
    add_body(doc, "5) рассчитать денежный поток, чистую приведённую стоимость за три года при ставке 18 процентов и срок окупаемости;")
    add_body(doc, "6) сопоставить риски и построить диаграмму Ганта выбранного сценария.")

    add_section_h(doc, "2 Действующая мощность и прогноз спроса")
    add_body(
        doc,
        "Цех 340 квадратных метров, одна смена. Оборудование: раскройный пресс листа "
        "вспененного полиэтилена толщиной 30 миллиметров, два термопресса контура, ручной "
        "инструмент сборки. Цена короба 8 000 рублей: годовая выручка 38 400 000 / 4 800 "
        "изделий = 8 000. Себестоимость около 4 800 рублей, маржа около 40 процентов; "
        "в расчёте расширения используется цена, потому что дополнительные короба продаются "
        "по той же цене, а переменные затраты на материал уже сидят в действующей марже. "
        "Дополнительные операционные расходы сценариев задаются отдельно и не смешиваются "
        "с полной себестоимостью.",
    )
    add_body(
        doc,
        "Сезонность не мебельная: летом растёт спрос на короба для пикника и дачи, зимой — "
        "заявки кухонь агрегаторов. Текущий годовой спрос 4 800 коробов. Прогноз: каждый "
        "месяц умножается на 1,35, декабрь подогнан так, чтобы сумма года равнялась "
        "4 800 × 1,35 = 6 480.",
    )
    add_caption_table(doc, "Таблица 1 — Текущий спрос и прогноз, штук")
    add_table(
        doc,
        ["Месяц", "Текущий спрос", "Прогноз +35 %", "Месяц", "Текущий спрос", "Прогноз +35 %"],
        [
            ["Январь", "370", "500", "Июль", "440", "594"],
            ["Февраль", "350", "473", "Август", "450", "608"],
            ["Март", "380", "513", "Сентябрь", "400", "540"],
            ["Апрель", "390", "527", "Октябрь", "380", "513"],
            ["Май", "410", "553", "Ноябрь", "360", "486"],
            ["Июнь", "430", "581", "Декабрь", "440", "592"],
            ["Итого", "4 800", "6 480", "—", "—", "—"],
        ],
        widths_cm=[3.0, 3.2, 3.2, 3.0, 3.2, 2.8],
    )
    add_body(
        doc,
        "Проверка года прогноза: 500 + 473 + 513 + 527 + 553 + 581 + 594 + 608 + 540 + 513 + "
        "486 + 592 = 6 480. При мощности 480 коробов в месяц без расширения не закрываются "
        "январь и все месяцы с марта по декабрь, кроме февраля (473 меньше 480). Пик — "
        "август, 608 коробов.",
    )

    add_section_h(doc, "3 Три сценария расширения")
    add_body(
        doc,
        "Сценарий 1, экстенсивный: аренда дополнительных 80 квадратных метров по 1 050 рублей "
        "за метр в месяц и второй термопресс контура за 620 000 рублей, найм двух сборщиков "
        "по 52 000 рублей. Аренда 80 × 1 050 = 84 000. Фонд оплаты двух сборщиков "
        "52 000 × 2 = 104 000. Дополнительные операционные расходы 84 000 + 104 000 = "
        "188 000 рублей в месяц. Мощность 580 коробов в месяц. Капитальные затраты 620 000 рублей.",
    )
    add_body(
        doc,
        "Сценарий 2, интенсивный: автоматический раскройный комплекс листа за 1 850 000 рублей "
        "и переход на две смены. Надбавка второй смены 20 процентов к окладам раскроя и сборки: "
        "58 000 × 2 × 0,20 + 52 000 × 4 × 0,20 = 23 200 + 41 600 = 64 800. Дополнительная "
        "электроэнергия и отопление второй смены 77 200. Итого 64 800 + 77 200 = 142 000 рублей "
        "в месяц. Мощность 620 коробов в месяц. Капитальные затраты 1 850 000 рублей.",
    )
    add_body(
        doc,
        "Сценарий 3, гибридный: раскрой листа передаётся подрядчику по 110 рублей за лист, "
        "освободившееся время пресса идёт на сборку. При среднем расходе около 650 листов "
        "в месяц после роста: 110 × 650 = 71 500, доставка кроя 6 500, итого 78 000 рублей "
        "в месяц. Капитальных затрат нет. Мощность 560 коробов в месяц.",
    )
    add_caption_table(doc, "Таблица 2 — Параметры сценариев")
    add_table(
        doc,
        ["Показатель", "Сценарий 1", "Сценарий 2", "Сценарий 3"],
        [
            ["Суть", "Площадь, пресс, два сборщика", "Комплекс и две смены", "Раскрой на подряде"],
            ["Мощность, шт./мес", "580", "620", "560"],
            ["Капитальные затраты, ₽", "620 000", "1 850 000", "0"],
            ["Доп. операционные расходы, ₽/мес", "188 000", "142 000", "78 000"],
        ],
        widths_cm=[4.6, 4.6, 4.6, 4.6],
        size=9,
    )

    add_section_h(doc, "4 Дополнительный выпуск и выручка")
    add_body(
        doc,
        "Дополнительный выпуск месяца равен минимуму из прогноза спроса и новой мощности "
        "минус текущий спрос. Формула не позволяет продать больше, чем просят, и не позволяет "
        "выпустить больше, чем тянет цех. Пример, август: текущий спрос 450, прогноз 608. "
        "Сценарий 1: минимум из 608 и 580 равен 580, дополнительный выпуск 580 − 450 = 130. "
        "Сценарий 2: минимум из 608 и 620 равен 608, 608 − 450 = 158. Сценарий 3: минимум "
        "из 608 и 560 равен 560, 560 − 450 = 110.",
    )
    add_caption_table(doc, "Таблица 3 — Дополнительный выпуск по месяцам, штук")
    add_table(
        doc,
        ["Месяц", "Сценарий 1", "Сценарий 2", "Сценарий 3"],
        [
            ["Январь", "130", "130", "130"],
            ["Февраль", "123", "123", "123"],
            ["Март", "133", "133", "133"],
            ["Апрель", "137", "137", "137"],
            ["Май", "143", "143", "143"],
            ["Июнь", "150", "151", "130"],
            ["Июль", "140", "154", "120"],
            ["Август", "130", "158", "110"],
            ["Сентябрь", "140", "140", "140"],
            ["Октябрь", "133", "133", "133"],
            ["Ноябрь", "126", "126", "126"],
            ["Декабрь", "140", "152", "120"],
            ["Итого", "1 625", "1 680", "1 545"],
        ],
        widths_cm=[4.6, 4.6, 4.6, 4.6],
    )
    add_body(
        doc,
        "Проверка сценария 2: 130 + 123 + 133 + 137 + 143 + 151 + 154 + 158 + 140 + 133 + "
        "126 + 152 = 1 680. Дополнительная выручка = дополнительный выпуск × 8 000. "
        "Сценарий 1: 1 625 × 8 000 = 13 000 000 рублей. Сценарий 2: 1 680 × 8 000 = "
        "13 440 000 рублей. Сценарий 3: 1 545 × 8 000 = 12 360 000 рублей. Только сценарий 2 "
        "закрывает августовский пик 608: мощность 620 больше 608.",
    )

    add_section_h(doc, "5 Дополнительные расходы и амортизация")
    add_body(
        doc,
        "Операционные расходы — регулярные платежи каждого месяца: аренда, оплата труда, "
        "электроэнергия, подряд. Амортизация — списание капитальных затрат равными долями "
        "за пять лет. Формула линейная: капитальные затраты / 5. Амортизация входит в "
        "расходы и в прибыль, но не является оттоком денег: станок уже оплачен в момент покупки.",
    )
    add_body(
        doc,
        "Сценарий 1. Операционные расходы за год: 188 000 × 12 = 2 256 000. Амортизация: "
        "620 000 / 5 = 124 000. Расходы всего: 2 256 000 + 124 000 = 2 380 000.",
    )
    add_body(
        doc,
        "Сценарий 2. Операционные расходы: 142 000 × 12 = 1 704 000. Амортизация: "
        "1 850 000 / 5 = 370 000. Расходы всего: 1 704 000 + 370 000 = 2 074 000.",
    )
    add_body(
        doc,
        "Сценарий 3. Операционные расходы: 78 000 × 12 = 936 000. Амортизация 0, потому что "
        "капитальных затрат нет. Расходы всего 936 000.",
    )
    add_caption_table(doc, "Таблица 4 — Дополнительные расходы, рублей в год")
    add_table(
        doc,
        ["Показатель", "Сценарий 1", "Сценарий 2", "Сценарий 3"],
        [
            ["Операционные расходы", "2 256 000", "1 704 000", "936 000"],
            ["Амортизация", "124 000", "370 000", "0"],
            ["Расходы всего", "2 380 000", "2 074 000", "936 000"],
            ["Дополнительная прибыль", "10 620 000", "11 366 000", "11 424 000"],
        ],
        widths_cm=[5.2, 4.4, 4.4, 4.4],
    )
    add_body(
        doc,
        "Прибыль = дополнительная выручка − расходы всего. Сценарий 1: 13 000 000 − 2 380 000 = "
        "10 620 000. Сценарий 2: 13 440 000 − 2 074 000 = 11 366 000. Сценарий 3: "
        "12 360 000 − 936 000 = 11 424 000.",
    )

    add_section_h(doc, "6 Денежный поток и чистая приведённая стоимость")
    add_body(
        doc,
        "Денежный поток года — приток минус отток денег. Амортизация в отток не входит. "
        "Денежный поток = дополнительная выручка − операционные расходы. Сценарий 1: "
        "13 000 000 − 2 256 000 = 10 744 000. Сценарий 2: 13 440 000 − 1 704 000 = "
        "11 736 000. Сценарий 3: 12 360 000 − 936 000 = 11 424 000.",
    )
    add_body(
        doc,
        "Дисконтирование приводит будущий рубль к сегодняшнему. Ставка 18 процентов годовых "
        "означает: рубль через год стоит 1 / 1,18 = 0,8475 сегодняшнего рубля. Через два года: "
        "1 / 1,18² = 1 / 1,3924 = 0,7182. Через три года: 1 / 1,18³ = 1 / 1,6430 = 0,6086. "
        "Сумма коэффициентов за три года: 0,8475 + 0,7182 + 0,6086 = 2,1743.",
    )
    add_body(
        doc,
        "Чистая приведённая стоимость за три года = − капитальные затраты + денежный поток × "
        "2,1743. Поток каждого из трёх лет принят одинаковым: мощность выходит на режим "
        "в горизонте оценки, отдельный разгон в модель не закладывался.",
    )
    add_body(
        doc,
        "Сценарий 1: 10 744 000 × 0,847458 = 9 105 089; 10 744 000 × 0,718184 = 7 716 169; "
        "10 744 000 × 0,608631 = 6 539 130. Сумма 23 360 388. 23 360 388 − 620 000 = "
        "22 740 388 рублей.",
    )
    add_body(
        doc,
        "Сценарий 2: 11 736 000 × 0,847458 = 9 945 767; 11 736 000 × 0,718184 = 8 428 608; "
        "11 736 000 × 0,608631 = 7 142 892. Сумма 25 517 267. 25 517 267 − 1 850 000 = "
        "23 667 267 рублей.",
    )
    add_body(
        doc,
        "Сценарий 3: 11 424 000 × 0,847458 = 9 681 360; 11 424 000 × 0,718184 = 8 204 534; "
        "11 424 000 × 0,608631 = 6 953 000. Сумма 24 838 894 рублей; капитальные затраты "
        "равны нулю. По чистой приведённой стоимости формально лидирует сценарий 3.",
    )
    add_caption_table(doc, "Таблица 5 — Денежные потоки и чистая приведённая стоимость")
    add_table(
        doc,
        ["Показатель", "Сценарий 1", "Сценарий 2", "Сценарий 3"],
        [
            ["Денежный поток, ₽/год", "10 744 000", "11 736 000", "11 424 000"],
            ["Поток года 0, ₽", "−620 000", "−1 850 000", "0"],
            ["NPV за 3 года, ₽", "22 740 388", "23 667 267", "24 838 894"],
        ],
        widths_cm=[5.2, 4.4, 4.4, 4.4],
    )

    add_section_h(doc, "7 Срок окупаемости")
    add_body(
        doc,
        "Простой срок окупаемости = капитальные затраты / годовой денежный поток. "
        "Это число лет, за которое накопленные потоки вернут вложение без учёта ставки.",
    )
    add_body(
        doc,
        "Сценарий 1: 620 000 / 10 744 000 = 0,058 года, 0,058 × 12 ≈ 0,7 месяца. "
        "Сценарий 2: 1 850 000 / 11 736 000 = 0,158 года, 0,158 × 12 ≈ 1,9 месяца. "
        "Сценарий 3: капитальных затрат нет, срок окупаемости вложений равен нулю. "
        "Короткие сроки объясняются тем, что дополнительные 1 500–1 700 коробов в год "
        "по 8 000 рублей дают выручку больше 12 миллионов при операционных расходах "
        "меньше 2,3 миллиона. Это не ошибка округления, а следствие высокой цены короба "
        "относительно аренды, смены и подряда.",
    )

    add_section_h(doc, "8 Риски сценариев")
    add_body(
        doc,
        "Сценарий 1. Риск найма: два сборщика с нужной выработкой 100 коробов на человека "
        "в месяц можно искать дольше месяца, аренда 80 метров при этом уже идёт. "
        "Мощность 580 не покрывает август 608, июль 594 и июнь 581: часть пикового спроса "
        "теряется даже после вложений.",
    )
    add_body(
        doc,
        "Сценарий 2. Риск отказа раскройного комплекса: при одной машине на две смены "
        "простой останавливает обе смены. Риск второй смены: надбавка 20 процентов не "
        "гарантирует набор ночных раскройщиков. Электроэнергия может оказаться выше 77 200. "
        "Зато мощность 620 покрывает пик 608, раскрой листа остаётся в цехе, толщина контура "
        "не зависит от подрядчика.",
    )
    add_body(
        doc,
        "Сценарий 3. Риск качества: разнотолщинность кроя у подрядчика возвращает конфликт "
        "входного контроля, уже разобранный в работе по структуре. Риск цены: ставка 110 рублей "
        "за лист может вырасти, операционные расходы 78 000 не фиксированы договором на три года. "
        "Мощность 560 не закрывает июнь, июль, август и декабрь прогноза. Чистая приведённая "
        "стоимость здесь наибольшая именно потому, что станок не покупается, а не потому, что "
        "сценарий лучше обслуживает спрос.",
    )

    add_section_h(doc, "9 Выбор сценария")
    add_body(
        doc,
        "Оптимальный сценарий — второй, интенсивный. Чистая приведённая стоимость "
        "23 667 267 рублей ниже, чем у третьего, на 1 171 627 рублей, но второй сценарий "
        "единственный закрывает все двенадцать месяцев прогноза. Раскрой двухконтурного "
        "короба остаётся внутри цеха. Срок окупаемости около двух месяцев. Первый сценарий "
        "отпадает из-за дыры в пиковые месяцы. Третий отпадает из-за качества кроя и той же "
        "дыры в пике, несмотря на формальное лидерство по чистой приведённой стоимости.",
    )

    add_section_h(doc, "10 Календарный план реализации")
    add_body(
        doc,
        "Диаграмма Ганта показывает работы выбранного сценария по шкале недель. Полоса — "
        "длительность, вертикальная пунктирная линия — промышленный запуск двух смен. "
        "Набор второй смены начинается параллельно монтажу, иначе запуск ждал бы людей "
        "после приёмки станка.",
    )
    add_figure(doc, FIG_GANTT, "Рисунок 1 — Диаграмма Ганта сценария 2, недели от старта")
    add_body(
        doc,
        "Порядок: техническое задание две недели, запрос предложений две недели, договор "
        "и оплата три недели, поставка и монтаж пять недель, набор и обучение второй смены "
        "шесть недель с наложением на монтаж, пробная партия две недели, промышленный запуск "
        "две недели. Горизонт 16 недель, около четырёх месяцев. Вехи: оплаченный договор, "
        "акт монтажа, акт пробной партии, приказ о двух сменах.",
    )

    add_struct(doc, "ЗАКЛЮЧЕНИЕ")
    add_body(
        doc,
        "При прогнозе 6 480 коробов в год действующая мощность одной смены 480 изделий "
        "в месяц не закрывает пик. Дополнительная выручка сценариев: 13 000 000, 13 440 000 "
        "и 12 360 000 рублей. Чистая приведённая стоимость за три года при ставке 18 процентов: "
        "22 740 388, 23 667 267 и 24 838 894 рубля. Выбран сценарий 2: раскройный комплекс "
        "и две смены, мощность 620, капитальные затраты 1 850 000 рублей, окупаемость около "
        "двух месяцев, календарный план 16 недель. Сценарий с наибольшей чистой приведённой "
        "стоимостью не выбран, потому что не покрывает августовский спрос и выносит раскрой "
        "на подряд.",
    )


def build(toc_pages: dict) -> Path:
    make_gantt()
    shutil.copy(BLANK, OUT)
    doc = Document(str(OUT))
    fill_title(doc)
    setup_body_section(doc)
    build_body(doc, toc_pages)
    doc.save(str(OUT))
    return OUT


def convert():
    src = str(OUT.resolve()).replace("'", "''")
    dst = str(PDF.resolve()).replace("'", "''")
    ps = f"""
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
foreach ($d in @($word.Documents)) {{ $d.Close([ref]0) }}
$doc = $word.Documents.Open('{src}')
$wdFormatPDF = 17
$doc.SaveAs([ref] '{dst}', [ref]$wdFormatPDF)
$doc.Close([ref]0)
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
"""
    subprocess.run(["powershell", "-NoProfile", "-Command", ps], check=True)


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
        from pypdf import PdfReader

        print("PDF pages", len(PdfReader(str(PDF)).pages))
        if pages:
            build(pages)
            convert()
            print("pass2", measure())
            print("PDF pages", len(PdfReader(str(PDF)).pages))
    except Exception as exc:
        print("PDF skipped:", exc)
    print(OUT)


if __name__ == "__main__":
    main()
