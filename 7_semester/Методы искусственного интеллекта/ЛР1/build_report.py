#!/usr/bin/env python3
"""Отчёт ЛР1 МИИ: линейная регрессия в KNIME (курьерская доставка)."""
from __future__ import annotations

import shutil
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
OUT = BASE / "ЛР1_Отчет_Буренков_МИИ.docx"
FIG = BASE / "figures"
SHOTS = BASE / "screenshots"

TOC_DEF = [
    ("ВВЕДЕНИЕ", "intro", 0),
    ("1 Постановка задачи и исходные данные", "s1", 0),
    ("2 Среда KNIME Analytics Platform", "s2", 0),
    ("3 Построение модели линейной регрессии", "s3", 0),
    ("4 Оценка качества и визуализация", "s4", 0),
    ("ЗАКЛЮЧЕНИЕ", "conc", 0),
    ("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "src", 0),
]


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


def add_section_h(doc, text, *, new_page=True):
    if new_page:
        doc.add_page_break()
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=True, align="left", space_after=12)
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


def add_figure(doc, path, caption, w=15.5):
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


def add_table(doc, headers, rows, widths_cm=None, size=11):
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
            format_paragraph(p, first_indent=False, align="left", line_spacing=1.15)
            set_run_font(p.add_run(v), size=size)
    spacer = doc.add_paragraph()
    format_paragraph(spacer, first_indent=False, line_spacing=1.0, space_after=6)


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
    put_cell(t_prep.rows[0].cells[4], "В.В. Фомин")
    put_cell(t_work.rows[0].cells[0], "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 1")
    put_cell(t_work.rows[1].cells[0], "Метод линейной регрессии")
    put_cell(t_work.rows[2].cells[0], "по курсу: Методы искусственного интеллекта")
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


def build_body(doc: Document, toc_pages: dict):
    add_struct(doc, "СОДЕРЖАНИЕ", new_page=False)
    for title, key, ind in TOC_DEF:
        add_toc_line(doc, title, toc_pages.get(key, "…"), indent=ind)

    add_struct(doc, "ВВЕДЕНИЕ")
    add_body(
        doc,
        "Целью работы является построение модели линейной регрессии в среде "
        "KNIME Analytics Platform 5.12 и оценка её качества на отложенной выборке. "
        "По методическому практикуму используются узлы CSV Reader, Color Manager, "
        "Partitioning, Linear Regression Learner, Regression Predictor, Numeric Scorer "
        "и Scatter Plot [1].",
    )
    add_body(
        doc,
        "В качестве предметной области выбран прогноз времени курьерской доставки "
        "по расстоянию до адреса. Набор данных — 100 заказов. Обучающая доля — 20 % "
        "записей, способ разбиения — линейная выборка (первые строки файла), "
        "как задано в работе № 2 практикума [1].",
    )

    add_section_h(doc, "1 Постановка задачи и исходные данные")
    add_body(
        doc,
        "Рассматривается зависимость времени доставки time_min от расстояния "
        "distance_km. Модель обычной линейной регрессии с константой:",
    )
    add_body(
        doc,
        "time_min = β₀ + β₁ · distance_km + ε.",
    )
    add_body(
        doc,
        "Файл courier_delivery.csv содержит 100 строк и три столбца: идентификатор "
        "заказа, расстояние в километрах и фактическое время в минутах "
        "(таблица 1). Расстояния лежат в диапазоне примерно 1,2…15 км.",
    )
    add_caption_table(doc, "Таблица 1 — Структура набора данных courier_delivery.csv")
    add_table(
        doc,
        ["Столбец", "Тип", "Роль"],
        [
            ["order_id", "целое", "идентификатор, в модель не входит"],
            ["distance_km", "вещественное", "независимая переменная"],
            ["time_min", "вещественное", "зависимая переменная (отклик)"],
        ],
        widths_cm=[4.0, 4.0, 8.5],
    )
    add_body(
        doc,
        "Узел Partitioning выделяет 20 % строк в обучающую выборку (верхний порт) "
        "и 80 % — в тестовую (нижний порт). При линейной выборке обучение идёт "
        "по первым 20 заказам, оценка качества — по оставшимся 80.",
    )

    add_section_h(doc, "2 Среда KNIME Analytics Platform")
    add_body(
        doc,
        "Работа выполнена в KNIME Analytics Platform 5.12.0 LTS. Репозиторий узлов "
        "содержит группу IO с CSV Reader, а также разделы Manipulation и Views "
        "(рисунок 1).",
    )
    knime_ui = SHOTS / "knime_typed_csv.png"
    if not knime_ui.exists():
        knime_ui = SHOTS / "knime_created.png"
    if knime_ui.exists():
        add_figure(doc, knime_ui, "Рисунок 1 — KNIME Analytics Platform: репозиторий узлов", w=15.5)
    add_body(
        doc,
        "Данные читаются узлом CSV Reader из файла courier_delivery.csv.",
    )

    add_section_h(doc, "3 Построение модели линейной регрессии")
    add_body(
        doc,
        "Узлы соединены по схеме практикума (рисунок 2.1 методички [1]): "
        "CSV Reader → Color Manager → Partitioning; обучающий порт Partitioning "
        "подаётся на Linear Regression Learner; модель и тестовый порт — на "
        "Regression Predictor; с выхода предиктора — Numeric Scorer и Scatter Plot "
        "(рисунок 2).",
    )
    add_figure(doc, FIG / "workflow.png", "Рисунок 2 — Схема workflow линейной регрессии", w=16.0)
    add_caption_table(doc, "Таблица 2 — Назначение узлов модели")
    add_table(
        doc,
        ["Узел", "Назначение", "Репозиторий"],
        [
            ["CSV Reader", "чтение courier_delivery.csv", "IO"],
            ["Color Manager", "цветовая кодировка distance_km", "Views"],
            ["Partitioning", "20 % / линейная выборка", "Manipulation"],
            ["Linear Regression Learner", "оценка β₀, β₁ по обучающей выборке", "Analytics > Mining > Linear/Polynomial Regression"],
            ["Regression Predictor", "столбец Prediction (time_min) на тесте", "Analytics > Mining > Linear/Polynomial Regression"],
            ["Numeric Scorer", "R², MAE, MSE, RMSE, MAPE, adj. R²", "Analytics > Mining > Scoring"],
            ["Scatter Plot", "диаграмма рассеяния distance_km–time_min", "Views"],
        ],
        widths_cm=[4.2, 6.3, 6.0],
        size=10,
    )
    add_body(
        doc,
        "В Linear Regression Learner целевой столбец — time_min, в список Include "
        "включён только distance_km. Константа (intercept) сохранена. "
        "По обучающим 20 строкам получено уравнение",
    )
    add_body(
        doc,
        "time_min = 6,792 + 3,462 · distance_km.",
    )
    add_body(
        doc,
        "Интерпретация: при нулевом расстоянии базовая длительность около 6,8 мин "
        "(приём заказа, выход курьера); каждый дополнительный километр увеличивает "
        "время примерно на 3,5 мин.",
    )

    add_section_h(doc, "4 Оценка качества и визуализация")
    add_body(
        doc,
        "Numeric Scorer сравнивает фактический time_min и столбец прогноза на "
        "тестовых 80 объектах. Метрики приведены в таблице 3 и на рисунке 3.",
    )
    add_caption_table(doc, "Таблица 3 — Метрики Numeric Scorer на тестовой выборке")
    add_table(
        doc,
        ["Показатель", "Значение", "Смысл"],
        [
            ["R²", "0,915", "доля объяснённой дисперсии времени"],
            ["Adjusted R²", "0,914", "R² с поправкой на число предикторов"],
            ["MAE", "2,472 мин", "средняя абсолютная ошибка"],
            ["MSE", "9,482", "средний квадрат ошибки"],
            ["RMSE", "3,079 мин", "корень из MSE, в единицах отклика"],
            ["Mean signed difference", "2,054 мин", "средняя разница со знаком"],
            ["MAPE", "7,20 %", "средняя абсолютная процентная ошибка"],
        ],
        widths_cm=[4.6, 3.4, 8.5],
        size=10,
    )
    add_figure(doc, FIG / "numeric_scorer.png", "Рисунок 3 — Сводка Numeric Scorer", w=13.0)
    add_body(
        doc,
        "Коэффициент детерминации 0,915 означает, что около 91,5 % разброса "
        "времени доставки на тесте объясняется расстоянием. MAPE около 7 % "
        "приемлема для оперативного прогноза слота доставки. Скорректированный R² "
        "почти совпадает с R², так как предиктор один.",
    )
    add_figure(
        doc,
        FIG / "scatter.png",
        "Рисунок 4 — Диаграмма рассеяния: обучение (20 %), тест (80 %) и оценённая прямая",
        w=14.5,
    )
    add_body(
        doc,
        "На рисунке 4 видна линейная зависимость: точки теста группируются вдоль "
        "прямой, оценённой только по первым 20 заказам. Выбросы вверх соответствуют "
        "задержкам (пробки, ожидание у двери) и не описываются одной переменной "
        "расстояния.",
    )

    add_struct(doc, "ЗАКЛЮЧЕНИЕ")
    add_body(
        doc,
        "Построена модель линейной регрессии времени курьерской доставки по "
        "расстоянию в KNIME 5.12 по узлам методички. При разбиении 20 % / линейная "
        "выборка получено уравнение time_min = 6,792 + 3,462 · distance_km. "
        "На тесте R² = 0,915, RMSE = 3,08 мин, MAPE = 7,2 %. Зависимость на "
        "диаграмме рассеяния линейная, модель пригодна как простой базовый прогноз; "
        "для учёта пробок и типа здания потребуются дополнительные признаки.",
    )

    add_struct(doc, "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ")
    add_body(
        doc,
        "1. Лабораторный практикум KNIME. Работа № 2. Метод линейной регрессии. "
        "СПб.: ГУАП, кафедра 42, 2026.",
    )
    add_body(
        doc,
        "2. KNIME AG. KNIME Analytics Platform 5.12 LTS. Zurich, 2026. "
        "URL: https://www.knime.com (дата обращения: 25.09.2026).",
    )
    add_body(
        doc,
        "3. Draper N., Smith H. Applied Regression Analysis. 3rd ed. New York: Wiley, 1998.",
    )


def main():
    shutil.copy2(BLANK, OUT)
    doc = Document(str(OUT))
    fill_title(doc)
    setup_body_section(doc)
    toc_pages = {
        "intro": 3,
        "s1": 4,
        "s2": 5,
        "s3": 6,
        "s4": 7,
        "conc": 8,
        "src": 8,
    }
    build_body(doc, toc_pages)
    doc.save(str(OUT))
    print("wrote", OUT, "size", OUT.stat().st_size)


if __name__ == "__main__":
    main()
