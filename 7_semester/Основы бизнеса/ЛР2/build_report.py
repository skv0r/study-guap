#!/usr/bin/env python3
"""Практическая работа № 2. Основы бизнеса: рейтинговый выбор поставщика."""

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
OUT = BASE / "ПР2_Отчет_Буренков_ОБ.docx"
PDF = BASE / "ПР2_Отчет_Буренков_ОБ.pdf"
FIG = BASE / "figures" / "rating_scheme.png"

BODY = 12
TBL = 10

TOC_DEF = [
    ("ВВЕДЕНИЕ", "intro", 0),
    ("1 Цель работы", "s1", 0),
    ("2 Исходные данные кейса", "s2", 0),
    ("3 Критерии и веса", "s3", 0),
    ("4 Расчёт полной цены поставки", "s4", 0),
    ("5 Баллы и итоговый рейтинг", "s5", 0),
    ("6 Схема метода", "s6", 0),
    ("7 Выбор победителя", "s7", 0),
    ("8 Анализ чувствительности", "s8", 0),
    ("9 Поставка «точно в срок»", "s9", 0),
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


def add_sub_h(doc, text):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=True, align="left", space_before=6, space_after=4)
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
            format_paragraph(p, first_indent=False, align="left", line_spacing=1.0)
            set_run_font(p.add_run(v), size=size)
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
    put_cell(t_prep.rows[0].cells[0], "канд. экон. наук, доцент")
    put_cell(t_prep.rows[0].cells[4], "Г.С. Армашова-Тельник")
    put_cell(t_work.rows[0].cells[0], "ОТЧЕТ О ПРАКТИЧЕСКОЙ РАБОТЕ № 2")
    put_cell(t_work.rows[1].cells[0], "Рейтинговый выбор поставщика сырья для производства термокоробов")
    put_cell(t_work.rows[2].cells[0], "по курсу: Основы бизнеса")
    put_cell(t_stud.rows[0].cells[1], "4321")
    put_cell(t_stud.rows[0].cells[5], "Г.В. Буренков")


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


def _box(ax, x, y, w, h, text, *, dash=False, fill="#f2f2f2", size=8.2):
    from matplotlib.patches import FancyBboxPatch

    ax.add_patch(
        FancyBboxPatch(
            (x, y),
            w,
            h,
            boxstyle="round,pad=0.02,rounding_size=0.08",
            facecolor=fill,
            edgecolor="black",
            linewidth=1.05,
            linestyle=(0, (3, 2)) if dash else "-",
        )
    )
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=size)


def _arrow(ax, x1, y1, x2, y2):
    ax.annotate(
        "",
        xy=(x2, y2),
        xytext=(x1, y1),
        arrowprops=dict(arrowstyle="-|>", color="black", lw=1.05, mutation_scale=9),
    )


def make_scheme():
    import matplotlib.pyplot as plt

    FIG.parent.mkdir(parents=True, exist_ok=True)
    plt.rcParams["font.family"] = "DejaVu Sans"
    fig, ax = plt.subplots(figsize=(11.6, 7.2), dpi=170)
    ax.set_xlim(0, 12.4)
    ax.set_ylim(0, 8.0)
    ax.axis("off")
    _box(ax, 4.3, 7.15, 3.8, 0.65, "Кейс: выбрать поставщика ППЭ", fill="#ffffff", size=9)
    _arrow(ax, 6.2, 7.15, 6.2, 6.6)
    _box(ax, 3.9, 5.85, 4.6, 0.7, "7 критериев, веса сумма 100 %", fill="#ffffff", size=9)
    _arrow(ax, 6.2, 5.85, 6.2, 5.3)
    _box(ax, 3.7, 4.5, 5.0, 0.75, "Балл 1–10 × вес → рейтинг", fill="#ffffff", size=9)
    _arrow(ax, 6.2, 4.5, 6.2, 3.95)
    _box(ax, 3.5, 3.15, 5.4, 0.75, "Чувствительность: вес цены 40 %", fill="#ffffff", size=9)
    _arrow(ax, 6.2, 3.15, 6.2, 2.6)
    _box(ax, 3.7, 1.8, 5.0, 0.75, "База: Б. При «точно в срок»: тоже Б", fill="#f2f2f2", size=9)
    _box(
        ax,
        0.2,
        0.2,
        3.2,
        1.4,
        "Комментарий:\nвес — доля важности\nбалл — оценка факта\nне путать с ценой листа",
        dash=True,
        fill="#ffffff",
        size=7.4,
    )
    _box(
        ax,
        9.0,
        0.2,
        3.15,
        1.4,
        "Комментарий:\nсамый дешёвый — В\nрейтинг выбирает Б\nэто разные вопросы",
        dash=True,
        fill="#ffffff",
        size=7.4,
    )
    fig.tight_layout(pad=0.2)
    fig.savefig(FIG, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def build_body(doc: Document, toc_pages: dict):
    add_struct(doc, "СОДЕРЖАНИЕ", new_page=False)
    for title, key, ind in TOC_DEF:
        add_toc_line(doc, title, toc_pages.get(key, "…"), indent=ind)

    add_struct(doc, "ВВЕДЕНИЕ")
    add_body(
        doc,
        "ООО «Жар-короб» шьёт двухконтурные термокороба. Основной материал — лист "
        "вспененного полиэтилена (ППЭ) толщиной 30 мм. Прежний поставщик поднял цену "
        "на 16 %, маржа сборки падает. На запрос откликнулись три новых поставщика. "
        "Решение принимается методом взвешенной рейтинговой оценки: у каждого критерия "
        "есть вес (доля важности, сумма весов равна 100 %) и балл от 1 до 10. "
        "Ниже показаны сами арифметические шаги, а не только итог.",
    )

    add_section_h(doc, "1 Цель работы")
    add_body(
        doc,
        "Цель — выбрать поставщика листа ППЭ для ООО «Жар-короб» методом взвешенной "
        "рейтинговой оценки и проверить устойчивость выбора.",
    )
    add_body(doc, "Задачи работы:")
    add_body(doc, "1) задать семь критериев и веса, сумма весов 100 %;")
    add_body(doc, "2) проставить баллы 1–10 и пояснить каждый балл;")
    add_body(doc, "3) посчитать рейтинг как сумму «балл × вес / 100»;")
    add_body(doc, "4) сравнить победителя рейтинга с самым дешёвым по полной цене листа;")
    add_body(doc, "5) поднять вес цены до 40 % и проверить, сменится ли победитель;")
    add_body(
        doc,
        "6) указать поставщика при работе «точно в срок» (just-in-time): сборка идёт "
        "с колёс, склад крупных партий держать нельзя.",
    )

    add_section_h(doc, "2 Исходные данные кейса")
    add_body(
        doc,
        "Потребность цеха — 480 листов в месяц. Это не «на глаз»: на один термокороб "
        "уходит 1,2 листа (два контура плюс обрезь 20 %), план выпуска 400 коробов, "
        "480 = 400 × 1,2. Закупка идёт один раз в месяц, поэтому число рейсов равно 1, "
        "если минимальная партия не больше 480 листов.",
    )
    add_caption_table(doc, "Таблица 1 — Условия трёх поставщиков")
    add_table(
        doc,
        ["Параметр", "А. Поролон-Север, Вологда", "Б. Теплолист СПб", "В. ПеноПласт-НН, Нижний Новгород"],
        [
            ["Цена листа ППЭ 30 мм", "1 180 ₽", "1 410 ₽", "1 050 ₽"],
            ["Срок поставки", "10 дней", "2 дня", "18 дней"],
            ["Минимальная партия", "160 листов", "40 листов", "400 листов"],
            ["Оплата", "предоплата 40 %, остаток по факту", "постоплата 10 дней (сначала товар, деньги через 10 дней)", "предоплата 100 %"],
            ["Доставка", "самовывоз, 29 000 ₽ за рейс", "6 500 ₽ за рейс, бесплатно от 80 листов", "доставка 48 000 ₽ за рейс"],
            ["Репутация", "5 лет, отзывов мало, конфликтов за год нет", "12 лет, 1 конфликт по браку за год", "4 года, отзывы без конфликтов"],
            ["Возврат брака", "по согласованию, до 1,5 % партии", "в течение 7 дней, до 4 % партии", "замена в следующей партии, до 2,5 %"],
            ["Сертификация", "ЕАС (единый знак обращения ЕАЭС)", "ЕАС и ISO 9001 (система качества)", "ЕАС"],
        ],
        widths_cm=[3.2, 5.0, 5.1, 5.1],
    )
    add_body(
        doc,
        "Поставщик А ближе к «дешевле, но ехать самим». Поставщик Б — городской, "
        "быстрее и с постоплатой, лист дороже. Поставщик В даёт самую низкую цену "
        "листа, но везёт 18 дней и просит 100 % денег заранее.",
    )

    add_section_h(doc, "3 Критерии и веса")
    add_body(
        doc,
        "Вес — это доля важности критерия в решении, не процент скидки и не доля в уставном "
        "капитале. Семь критериев взяты из порядка работы. Сумма 100 %.",
    )
    add_caption_table(doc, "Таблица 2 — Веса критериев")
    add_table(
        doc,
        ["Критерий", "Вес, %", "Почему такой вес"],
        [
            ["Цена с логистикой", "25", "Маржа уже падает после роста цен прежнего поставщика на 16 %"],
            ["Срок поставки", "20", "Без листа линия сборки стоит; простой дороже разницы в 200 ₽ на листе"],
            ["Условия оплаты", "15", "100 % предоплата замораживает оборот на 2–3 недели"],
            ["Минимальная партия", "10", "Партия 400 листов при плане 480 почти не оставляет гибкости"],
            ["Репутация", "10", "Срыв серии коробов для агрегатора бьёт сильнее, чем разовый брак"],
            ["Возврат брака", "10", "Важно не только «сколько процентов», но и как быстро меняют"],
            ["Сертификация", "10", "ЕАС обязателен для обращения; ISO 9001 снижает риск серии брака"],
        ],
        widths_cm=[4.2, 2.0, 12.2],
    )
    add_body(doc, "Проверка: 25 + 20 + 15 + 10 + 10 + 10 + 10 = 100.")

    add_section_h(doc, "4 Расчёт полной цены поставки")
    add_body(
        doc,
        "Цена листа в прайсе — не полная цена. Полная цена листа = (цена × 480 листов + "
        "доставка за рейс) / 480. Для Б при месячной закупке 480 листов доставка равна "
        "нулю: 480 ≥ 80. Порог 80 действует только на мелкие партии (см. раздел 9).",
    )
    add_body(doc, "Поставщик А: 1 180 × 480 = 566 400 ₽. Доставка 29 000 ₽. Итого 595 400 ₽. Полная цена листа 595 400 / 480 = 1 240,42 ₽.")
    add_body(doc, "Поставщик Б: 1 410 × 480 = 676 800 ₽. Доставка 0 ₽. Полная цена листа 1 410 ₽.")
    add_body(doc, "Поставщик В: 1 050 × 480 = 504 000 ₽. Доставка 48 000 ₽. Итого 552 000 ₽. Полная цена листа 552 000 / 480 = 1 150 ₽.")
    add_caption_table(doc, "Таблица 3 — Полная стоимость месячной закупки")
    add_table(
        doc,
        ["Показатель", "А. Поролон-Север", "Б. Теплолист СПб", "В. ПеноПласт-НН"],
        [
            ["Материал, ₽", "566 400", "676 800", "504 000"],
            ["Доставка, ₽", "29 000", "0", "48 000"],
            ["Итого за месяц, ₽", "595 400", "676 800", "552 000"],
            ["Полная цена листа, ₽", "1 240,42", "1 410,00", "1 150,00"],
        ],
        widths_cm=[4.6, 4.6, 4.6, 4.6],
    )
    add_body(
        doc,
        "По деньгам за месяц самый дешёвый — В (552 000 ₽). А дороже на 43 400 ₽, "
        "Б дороже В на 124 800 ₽. Если смотреть только на это, выбор очевиден. "
        "Рейтинг проверяет, не слишком ли В проигрывает по сроку, оплате и партии.",
    )

    add_section_h(doc, "5 Баллы и итоговый рейтинг")
    add_body(
        doc,
        "Балл 10 — лучший из трёх по этому критерию, не «идеал в природе». "
        "Худший из трёх получает 3 или 4, чтобы разрыв был виден, но не обнулял остальные критерии.",
    )
    add_caption_table(doc, "Таблица 4 — Баллы 1–10 и вклад в рейтинг")
    add_table(
        doc,
        ["Критерий", "Вес", "А", "Б", "В", "А × вес", "Б × вес", "В × вес"],
        [
            ["Цена с логистикой", "0,25", "7", "4", "10", "1,75", "1,00", "2,50"],
            ["Срок поставки", "0,20", "6", "10", "3", "1,20", "2,00", "0,60"],
            ["Условия оплаты", "0,15", "6", "10", "3", "0,90", "1,50", "0,45"],
            ["Минимальная партия", "0,10", "7", "10", "4", "0,70", "1,00", "0,40"],
            ["Репутация", "0,10", "6", "8", "5", "0,60", "0,80", "0,50"],
            ["Возврат брака", "0,10", "6", "8", "5", "0,60", "0,80", "0,50"],
            ["Сертификация", "0,10", "7", "10", "7", "0,70", "1,00", "0,70"],
            ["Рейтинг", "1,00", "—", "—", "—", "6,45", "8,10", "5,65"],
        ],
        widths_cm=[3.6, 1.6, 1.4, 1.4, 1.4, 2.0, 2.0, 2.0],
    )
    add_body(
        doc,
        "Пример строки «цена»: у В полная цена 1 150 ₽ — балл 10; у А 1 240,42 ₽ — балл 7; "
        "у Б 1 410 ₽ — балл 4. Вклад В: 10 × 0,25 = 2,50. Вклад А: 7 × 0,25 = 1,75. "
        "Вклад Б: 4 × 0,25 = 1,00.",
    )
    add_body(
        doc,
        "Срок: 2 дня → 10, 10 дней → 6, 18 дней → 3. Оплата: постоплата 10 дней → 10; "
        "предоплата 40 % → 6; предоплата 100 % → 3. Партия при 22 рабочих днях "
        "(480 / 22 ≈ 22 листа в день): 40 листов ≈ 2 дня запаса → 10; "
        "160 листов ≈ 7 дней → 7; 400 листов ≈ 18 дней запаса → 4.",
    )
    add_body(
        doc,
        "Репутация: 12 лет и ISO 9001, но один конфликт — не 10, а 8. Возврат: "
        "7 календарных дней и понятный лимит 4 % лучше, чем «по согласованию» "
        "и чем замена только в следующей машине. Сертификация: ЕАС есть у всех (7), "
        "ISO 9001 добавляет Б до 10.",
    )
    add_body(doc, "Итог: А = 6,45; Б = 8,10; В = 5,65. Первое место — поставщик Б.")

    add_section_h(doc, "6 Схема метода")
    add_body(
        doc,
        "На рисунке 1 — порядок шагов. Пунктир справа и слева — пояснение терминов, "
        "не отдельные этапы.",
    )
    add_figure(doc, FIG, "Рисунок 1 — Порядок рейтингового выбора поставщика")

    add_section_h(doc, "7 Выбор победителя")
    add_body(
        doc,
        "Интуитивно берут самого дешёвого. Это поставщик В: 1 150 ₽ за лист и "
        "552 000 ₽ за месяц. Рейтинг ставит Б (8,10 против 5,65). Разница не в «вкусе», "
        "а в том, что В забирает 2,50 балла ценой и сразу отдаёт их сроком (0,60), "
        "оплатой (0,45) и партией (0,40).",
    )
    add_body(
        doc,
        "Для цеха это значит: экономия 124 800 ₽ в месяц против Б покупается "
        "предоплатой всей суммы, ожиданием 18 дней и складом почти на весь месяц. "
        "Если за эти 18 дней агрегатор сдвигает заказ, материал уже оплачен и едет. "
        "Б дороже, но постоплата 10 дней и поставка за 2 дня оставляют оборот в кассе.",
    )

    add_section_h(doc, "8 Анализ чувствительности")
    add_body(
        doc,
        "Анализ чувствительности здесь — проверка, не сломается ли выбор, если цену "
        "сочтём важнее. Вес цены поднимаем с 25 % до 40 %. Остальные 60 % делим "
        "пропорционально старым весам: коэффициент 60 / 75 = 0,8. Тогда срок 16 %, "
        "оплата 12 %, партия, репутация, брак и сертификат по 8 %. Сумма: "
        "40 + 16 + 12 + 8 + 8 + 8 + 8 = 100.",
    )
    add_caption_table(doc, "Таблица 5 — Рейтинг при весе цены 40 %")
    add_table(
        doc,
        ["Критерий", "Вес", "А × вес", "Б × вес", "В × вес"],
        [
            ["Цена с логистикой", "0,40", "2,80", "1,60", "4,00"],
            ["Срок поставки", "0,16", "0,96", "1,60", "0,48"],
            ["Условия оплаты", "0,12", "0,72", "1,20", "0,36"],
            ["Минимальная партия", "0,08", "0,56", "0,80", "0,32"],
            ["Репутация", "0,08", "0,48", "0,64", "0,40"],
            ["Возврат брака", "0,08", "0,48", "0,64", "0,40"],
            ["Сертификация", "0,08", "0,56", "0,80", "0,56"],
            ["Рейтинг", "1,00", "6,56", "7,28", "6,52"],
        ],
        widths_cm=[4.4, 2.2, 3.6, 3.6, 3.6],
    )
    add_body(
        doc,
        "Победитель не меняется: Б = 7,28; А = 6,56; В = 6,52. Разрыв Б и В сжимается "
        "с 2,45 до 0,76, но В всё ещё не обходит Б. Даже когда цена «важнее всего», "
        "18 дней и стопроцентная предоплата не дают В первое место.",
    )

    add_section_h(doc, "9 Поставка «точно в срок»")
    add_body(
        doc,
        "«Точно в срок» (just-in-time) — лист приезжает к дню сборки, большой склад "
        "не держат. Тогда критичны минимальная партия и срок, а не полная цена листа.",
    )
    add_body(
        doc,
        "Суточный расход: 480 / 22 рабочих дня ≈ 22 листа. Партия А 160 листов — запас "
        "на 160 / 22 ≈ 7 дней. Партия В 400 листов — запас на 400 / 22 ≈ 18 дней, "
        "плюс 18 дней в пути: это уже не «с колёс». Партия Б 40 листов — запас на "
        "40 / 22 ≈ 2 дня, срок 2 дня, можно возить дважды в неделю. Партия 40 меньше "
        "порога бесплатной доставки 80, поэтому каждый рейс Б стоит 6 500 ₽: "
        "6 500 / 40 = 162,50 ₽ к цене листа, полная цена 1 572,50 ₽. Это дороже "
        "месячной закупки, но склад на 18 дней всё равно не нужен.",
    )
    add_body(
        doc,
        "При «точно в срок» оптимален поставщик Б: партия 40 листов и поставка за 2 дня. "
        "Поставщик В непригоден: и партия, и срок измеряются почти месяцем. Поставщик А "
        "возможен как запасной, если Б срывает отгрузку, но не как основной при работе "
        "без склада.",
    )

    add_struct(doc, "ЗАКЛЮЧЕНИЕ")
    add_body(
        doc,
        "По базовому рейтингу выбран поставщик Б «Теплолист СПб»: 8,10 против 6,45 у А "
        "и 5,65 у В. Самый дешёвый по полной цене листа — В (1 150 ₽ и 552 000 ₽ в месяц). "
        "Рейтинг с ценой не совпал, потому что у В предоплата 100 %, срок 18 дней и "
        "партия 400 листов.",
    )
    add_body(
        doc,
        "При весе цены 40 % победитель тот же (Б = 7,28). При работе «точно в срок» "
        "снова Б: партия 40 листов и поставка за 2 дня. Арифметика месячной закупки "
        "и вкладов в рейтинг совпадает с расчётной таблицей Excel.",
    )


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
