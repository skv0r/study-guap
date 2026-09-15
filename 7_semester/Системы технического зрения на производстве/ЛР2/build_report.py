#!/usr/bin/env python3
"""Отчёт ЛР2 по СТЗ: подключение камеры и захват в MATLAB."""

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
OUT = BASE / "ЛР2_Отчет_Буренков_СТЗ.docx"
PDF = BASE / "ЛР2_Отчет_Буренков_СТЗ.pdf"
SHOTS = BASE / "screenshots"
CODE = (BASE / "lr2_photo_live.m").read_text(encoding="utf-8")

TOC_DEF = [
    ("ВВЕДЕНИЕ", "intro", 0),
    ("1 Средства и объект эксперимента", "s1", 0),
    ("2 Работа в Image Acquisition Explorer", "s2", 0),
    ("3 Программный захват и настройка параметров", "s3", 0),
    ("4 Запись видеопотока на диск (Logging)", "s4", 0),
    ("5 Расчёт потока данных и выбор интерфейса", "s5", 0),
    ("6 Детектирование движения и влияние разрешения", "s6", 0),
    ("ЗАКЛЮЧЕНИЕ", "conc", 0),
    ("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "src", 0),
    ("ПРИЛОЖЕНИЕ А Листинг программы захвата", "app", 0),
]

SEARCH = {
    "intro": "ВВЕДЕНИЕ",
    "s1": "1 Средства и объект эксперимента",
    "s2": "2 Работа в Image Acquisition Explorer",
    "s3": "3 Программный захват и настройка параметров",
    "s4": "4 Запись видеопотока на диск (Logging)",
    "s5": "5 Расчёт потока данных и выбор интерфейса",
    "s6": "6 Детектирование движения и влияние разрешения",
    "conc": "ЗАКЛЮЧЕНИЕ",
    "src": "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
    "app": "ПРИЛОЖЕНИЕ А",
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
            format_paragraph(p, first_indent=False, align="left", line_spacing=1.15)
            set_run_font(p.add_run(v), size=size)
    spacer = doc.add_paragraph()
    format_paragraph(spacer, first_indent=False, line_spacing=1.0, space_after=6)


def add_listing(doc, title: str, code: str):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="left", space_before=10, space_after=4)
    set_run_font(p.add_run(title), bold=True)
    for line in code.strip("\n").splitlines():
        lp = doc.add_paragraph()
        format_paragraph(lp, first_indent=False, align="left", line_spacing=1.0, space_after=0)
        set_run_font(lp.add_run(line if line else " "), size=9, name="Courier New")


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
    put_cell(t_work.rows[0].cells[0], "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 2")
    put_cell(
        t_work.rows[1].cells[0],
        "Подключение и настройка видеокамеры. Захват изображений и видеокадров на производстве",
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


def build_body(doc: Document, toc_pages: dict):
    add_struct(doc, "СОДЕРЖАНИЕ", new_page=False)
    for title, key, ind in TOC_DEF:
        add_toc_line(doc, title, toc_pages.get(key, "…"), indent=ind)

    add_struct(doc, "ВВЕДЕНИЕ")
    add_body(
        doc,
        "Целью работы является подключение источника видео к MATLAB, определение "
        "доступных форматов, захват и запись видеопотока, расчёт потока данных "
        "и обработка последовательности кадров разностным методом [1], [2].",
    )
    add_body(
        doc,
        "Выполнены оба варианта: Image Acquisition Explorer (взамен удалённого "
        "в R2024a приложения imaqtool) и программный захват. Источник — Color Device "
        "адаптера mwdemoimaq, DeviceID = 1 [3]. В поле зрения — оператор в наушниках. "
        "Последовательность 180 кадров (6 с при 30 кадр/с, RGB_NTSC 640×480) записана "
        "на диск. Исследованы режимы RGB_NTSC, S-Video 160×120 и CCIR 768×576. "
        "Разностный детектор построен по соседним кадрам видеопотока.",
    )

    add_section_h(doc, "1 Средства и объект эксперимента")
    add_body(
        doc,
        "Среда: MATLAB R2024a (24.1), Windows 11. Установлены Image Acquisition "
        "Toolbox 24.1 и Image Processing Toolbox. Адаптер mwdemoimaq "
        "зарегистрирован командой imaqregister из Adaptor Kit. imaqhwinfo "
        "возвращает три устройства (рисунок 1, таблица 1).",
    )
    add_caption_table(doc, "Таблица 1 — Обнаруженные устройства mwdemoimaq")
    add_table(
        doc,
        ["DeviceID", "Имя", "Форматы", "Назначение в работе"],
        [
            ["1", "Color Device", "RGB_NTSC, S-Video", "цветной захват, Hue/Saturation, Logging"],
            ["2", "Monochrome Device", "RS170, CCIR", "третий режим, полутоновый поток"],
            ["3", "Digital Device", "файл устройства", "не использовалось"],
        ],
        widths_cm=[2.2, 3.6, 4.2, 6.5],
    )
    add_figure(doc, SHOTS / "imaqhwinfo.png", "Рисунок 1 — Вывод imaqhwinfo по адаптеру mwdemoimaq", w=15.0)
    add_body(
        doc,
        "Рабочий формат Color Device — RGB_NTSC, 640×480, 30 кадр/с. Для выбора "
        "промышленного канала по методичке считается несжатый поток "
        "D = W × H × B × FPS / 10^6 при тех же W, H, B и FPS [1]. Объект "
        "съёмки — оператор в наушниках на фоне помещения.",
    )

    add_section_h(doc, "2 Работа в Image Acquisition Explorer")
    add_body(
        doc,
        "Команда imaqtool в R2024a удалена. Использовано приложение Image "
        "Acquisition Explorer [4]. В Hardware Browser отображаются Color Device, "
        "Monochrome Device и Digital Device, DeviceID = 1…3, индикатор зелёный "
        "(рисунок 2).",
    )
    add_figure(doc, SHOTS / "imaq_explorer_crop.png", "Рисунок 2 — Image Acquisition Explorer: список устройств", w=15.5)
    add_body(
        doc,
        "Предварительный просмотр Color Device, формат RGB_NTSC. Окно показывает "
        "640×480, 30 кадр/с и строку Waiting for START (рисунок 3). В кадре — "
        "оператор в наушниках. Это штатное состояние: изображение уже идёт, "
        "сбор кадров ещё не запущен [2].",
    )
    add_figure(doc, SHOTS / "preview.png", "Рисунок 3 — Окно preview Color Device, формат RGB_NTSC", w=14.0)

    add_section_h(doc, "3 Программный захват и настройка параметров")
    add_body(
        doc,
        "Объект видеовхода — Color Device, DeviceID = 1, формат RGB_NTSC, "
        "ReturnedColorSpace = rgb. Из свойств источника изменены два параметра: "
        "Hue и Saturation. Exposure и WhiteBalance у устройства недоступны.",
    )
    add_body(
        doc,
        "Исходные значения: Hue = 0,5, Saturation = 50. Затем Hue = 0,12 и "
        "Saturation = 92: стена уходит в красно-оранжевый (рисунки 4, 5). "
        "Затем Hue = 0,82 и Saturation = 12: кадр бледнеет, появляется "
        "зеленоватый оттенок.",
    )
    add_figure(doc, SHOTS / "source_props.png", "Рисунок 4 — Hue и Saturation после getselectedsource", w=14.5)
    add_figure(doc, SHOTS / "prop_hue_sat_1.png", "Рисунок 5 — Preview после Hue = 0,12, Saturation = 92", w=14.0)

    add_section_h(doc, "4 Запись видеопотока на диск (Logging)")
    add_body(
        doc,
        "Заданы LoggingMode = disk, FramesPerTrigger = 180. Запись 6,0 с при "
        "30 кадр/с, файл capture\\log_6s.avi, контейнер Motion JPEG AVI, объём "
        "7,88 МБ (рисунок 6). FramesAcquired = DiskLoggerFrameCount = 180. "
        "Требование методички 5–10 с выполнено, потерь кадров нет.",
    )
    add_figure(doc, SHOTS / "logging.png", "Рисунок 6 — Параметры LoggingMode = disk после start(vid)", w=14.5)
    add_caption_table(doc, "Таблица 2 — Параметры записи на диск")
    add_table(
        doc,
        ["Параметр", "Значение"],
        [
            ["LoggingMode", "disk"],
            ["FramesPerTrigger", "180"],
            ["FPS", "30"],
            ["Длительность, с", "6,0"],
            ["FramesAcquired", "180"],
            ["DiskLoggerFrameCount", "180"],
            ["Файл", "log_6s.avi"],
            ["Объём", "7,88 МБ"],
        ],
        widths_cm=[6.0, 10.5],
    )

    add_section_h(doc, "5 Расчёт потока данных и выбор интерфейса")
    add_body(
        doc,
        "Исследованы три режима. Частота кадров у источника равна 30 кадр/с. "
        "Поток несжатых данных D = W × H × B × FPS / 10^6 МБ/с, где для RGB "
        "B = 3, для полутонового CCIR B = 1 [1]. Результаты — в таблице 3.",
    )
    add_caption_table(doc, "Таблица 3 — Режимы и расчётный поток")
    add_table(
        doc,
        ["№", "Устройство", "Формат", "Разрешение", "FPS", "B, байт/пикс", "Поток, МБ/с"],
        [
            ["1", "Color Device", "RGB_NTSC", "640 × 480", "30", "3", "27,648"],
            ["2", "Color Device", "S-Video", "160 × 120", "30", "3", "1,728"],
            ["3", "Monochrome Device", "CCIR", "768 × 576", "30", "1", "13,271"],
        ],
        widths_cm=[1.2, 3.4, 2.6, 2.6, 1.5, 2.4, 2.8],
    )
    add_body(
        doc,
        "Максимальный оценочный поток 27,648 МБ/с. С запасом 30 % требуется "
        "около 36 МБ/с. GigE Vision 1G (около 120 МБ/с, линия до 100 м) достаточен "
        "и принят как минимально достаточный промышленный канал. USB3 Vision "
        "(около 400 МБ/с) избыточен по скорости и ограничен длиной ~5 м. "
        "Camera Link избыточен и требует плату захвата. Фактический канал "
        "стенда — память процесса MATLAB; GigE — оценка для промышленной камеры "
        "с тем же расчётным потоком [1].",
    )
    add_caption_table(doc, "Таблица 4 — Сравнение интерфейсов для потока 27,648 МБ/с")
    add_table(
        doc,
        ["Интерфейс", "Полезная скорость", "Длина линии", "Вывод"],
        [
            ["GigE Vision 1G", "около 120 МБ/с", "до 100 м", "минимально достаточный"],
            ["USB3 Vision", "около 400 МБ/с", "около 5 м", "хватает, кабель короткий"],
            ["Camera Link", "до 850 МБ/с", "до 10 м", "избыточен, нужна плата"],
        ],
        widths_cm=[3.6, 3.8, 3.4, 5.7],
    )

    add_section_h(doc, "6 Детектирование движения и влияние разрешения")
    add_body(
        doc,
        "Разностный метод: два кадра видеопоследовательности переводятся в "
        "полутоновые, берётся imabsdiff, порог отделяет изменившиеся пиксели [1]. "
        "Сравнивались кадры № 10 и № 22 (смещение головы оператора относительно "
        "фона). Пороги 5, 10 и 25 выделяют контуры лица и наушников "
        "(рисунки 7–9). При пороге 25 маска реже. Рабочий порог — 10. "
        "На рисунке 10 — окно детектора, FramesAcquired = 180.",
    )
    add_figure(doc, SHOTS / "motion_RGB_NTSC_th5.png", "Рисунок 7 — Маска движения, RGB_NTSC, порог 5", w=11.0)
    add_figure(doc, SHOTS / "motion_RGB_NTSC_th10.png", "Рисунок 8 — Маска движения, RGB_NTSC, порог 10", w=11.0)
    add_figure(doc, SHOTS / "motion_RGB_NTSC_th25.png", "Рисунок 9 — Маска движения, RGB_NTSC, порог 25", w=11.0)
    add_figure(doc, SHOTS / "motion_live.png", "Рисунок 10 — Окно детектора движения, порог 10, FramesAcquired = 180", w=13.0)
    add_body(
        doc,
        "Сравнение нагрузки при одинаковом алгоритме и 20 парах кадров: "
        "RGB_NTSC 640×480 (307 200 пикс.) и S-Video 160×120 (19 200 пикс., в 16 раз "
        "меньше). Среднее время цикла 0,23 мс и 0,10 мс соответственно (таблица 5). "
        "Маска S-Video грубее из-за малого числа пикселей (рисунок 11).",
    )
    add_caption_table(doc, "Таблица 5 — Время цикла разностной обработки, 20 пар кадров")
    add_table(
        doc,
        ["Формат", "Разрешение", "Пар кадров", "FramesAcquired", "Среднее, мс", "Макс., мс"],
        [
            ["RGB_NTSC", "640 × 480", "20", "40", "0,23", "0,54"],
            ["S-Video", "160 × 120", "20", "40", "0,10", "0,33"],
        ],
        widths_cm=[2.8, 2.8, 2.4, 3.2, 2.6, 2.4],
    )
    add_figure(doc, SHOTS / "motion_S_Video_th10.png", "Рисунок 11 — Маска движения, S-Video 160×120, порог 10", w=8.0)

    add_struct(doc, "ЗАКЛЮЧЕНИЕ")
    add_body(
        doc,
        "К MATLAB подключён Color Device (DeviceID = 1, mwdemoimaq). В кадре — "
        "оператор в наушниках. Выполнены preview 640×480 при 30 кадр/с, изменение "
        "Hue и Saturation, запись 180 кадров на диск (6 с, log_6s.avi), разностный "
        "детектор и сравнение двух разрешений. Максимальный поток 27,648 МБ/с "
        "покрывается GigE Vision 1G. Команда imaqtool в R2024a удалена, использован "
        "Image Acquisition Explorer.",
    )

    add_struct(doc, "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ")
    sources = [
        "Подключение и настройка видеокамеры. Захват изображений и видеокадров на производстве : метод. указания к лабораторной работе № 2. – Санкт-Петербург : ГУАП, 2026.",
        "Image Acquisition Toolbox : Basic Image Acquisition Procedure. – URL: https://www.mathworks.com/help/imaq/basic-image-acquisition-procedure.html (дата обращения: 14.09.2026).",
        "Creating a Custom Adaptor using the Adaptor Kit. – URL: https://www.mathworks.com/help/imaq/creating-a-custom-adaptor-using-the-adaptor-kit.html (дата обращения: 14.09.2026).",
        "Image Acquisition Explorer. – URL: https://www.mathworks.com/help/imaq/image-acquisition-explorer.html (дата обращения: 14.09.2026).",
    ]
    for i, src in enumerate(sources, 1):
        p = doc.add_paragraph()
        format_paragraph(p)
        set_run_font(p.add_run(f"{i}. {src}"))

    add_struct(doc, "ПРИЛОЖЕНИЕ А")
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="center", space_after=12)
    set_run_font(p.add_run("Листинг программы захвата"), bold=True)
    add_body(
        doc,
        "В приложении приведён m-файл lr2_photo_live.m: видеопоследовательность "
        "Color Device, preview, Hue/Saturation, запись 6 с на диск, расчёт потока "
        "и разностный детектор по соседним кадрам.",
    )
    add_listing(doc, "Листинг А.1 — Файл lr2_photo_live.m", CODE)


def build(toc_pages: dict) -> Path:
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
        if pages:
            build(pages)
            convert()
            print("pass2", measure())
    except Exception as exc:
        print("PDF skipped:", exc)
    print(OUT)


if __name__ == "__main__":
    main()
