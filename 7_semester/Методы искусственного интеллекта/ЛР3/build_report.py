# -*- coding: utf-8 -*-
"""Сборка отчёта ЛР3 МИИ. В текст отчёта этот файл не попадает."""
from __future__ import annotations

import json
import shutil
import zipfile
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from lxml import etree

BASE = Path(__file__).resolve().parent
REPO = BASE.parents[1]
SRC = REPO / "Web-программирование" / "ЛР1_Отчет_Буренков_PairWise.docx"
OUT = BASE / "ЛР3_Отчет_Буренков_МИИ.docx"
FIG = BASE / "figures"
DATA = BASE / "data"

CLASSES = ["сеть", "доступ", "ПО"]
FEATURES = ["users_affected", "downtime_min", "latency_ms", "error_rate_pct"]
FEATURE_RU = {
    "users_affected": "Пользователи",
    "downtime_min": "Простой, мин",
    "latency_ms": "Задержка, мс",
    "error_rate_pct": "Ошибки, %",
}


def num(value: float, digits: int = 2) -> str:
    text = f"{value:.{digits}f}".replace(".", ",")
    if text.startswith("-"):
        text = "−" + text[1:]
    return text


def pct(value: float) -> str:
    return num(value * 100, 2)


def set_run_font(run, size=14, bold=False):
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = False
    run.font.color.rgb = RGBColor(0, 0, 0)


def format_paragraph(p, *, first_indent=True, align="justify", space_after=0, space_before=0, line_spacing=1.5):
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = line_spacing
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.first_line_indent = Cm(1.25 if first_indent else 0)
    p.alignment = {
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "left": WD_ALIGN_PARAGRAPH.LEFT,
    }[align]


def add_text(doc, text, *, indent=True, align="justify", bold=False, space_after=0, space_before=0, size=14):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=indent, align=align, space_after=space_after, space_before=space_before)
    set_run_font(p.add_run(text), size=size, bold=bold)
    return p


def add_heading(doc, text):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=True, align="left", space_after=12, space_before=0)
    p.paragraph_format.keep_with_next = True
    if text in {"СОДЕРЖАНИЕ"}:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        set_run_font(p.add_run(text), bold=True)
        return p
    number, title = text.split(" ", 1)
    set_run_font(p.add_run(number), bold=True)
    space = p.add_run(" ")
    set_run_font(space, bold=True)
    space._r.find(qn("w:t")).set(qn("xml:space"), "preserve")
    set_run_font(p.add_run(title), bold=True)
    return p


def add_toc(doc, title, page):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="left")
    p.paragraph_format.tab_stops.add_tab_stop(Cm(16.5), 2, 1)
    set_run_font(p.add_run(title))
    set_run_font(p.add_run("\t"))
    set_run_font(p.add_run(str(page)))
    return p


def set_cell_border(cell):
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


def shade_header(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), "F2F2F2")
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell(cell, text, *, bold=False, size=11, align="center"):
    cell.text = ""
    p = cell.paragraphs[0]
    format_paragraph(p, first_indent=False, align=align, line_spacing=1.0, space_after=0, space_before=0)
    set_run_font(p.add_run(str(text)), size=size, bold=bold)
    set_cell_border(cell)
    tcPr = cell._tc.get_or_add_tcPr()
    vAlign = OxmlElement("w:vAlign")
    vAlign.set(qn("w:val"), "center")
    tcPr.append(vAlign)


def add_table(doc, headers, rows, widths, *, font=11, body_align="center"):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tblPr.append(borders)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    total = int(round(sum(widths) * 567))
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(total))
    tblW.set(qn("w:type"), "dxa")
    tblPr.append(tblW)
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        tblPr.addnext(grid)
    else:
        for child in list(grid):
            grid.remove(child)
    for width in widths:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(int(round(width * 567))))
        grid.append(gc)
    for j, header_text in enumerate(headers):
        set_cell(table.rows[0].cells[j], header_text, bold=True, size=font)
        shade_header(table.rows[0].cells[j])
        table.rows[0].cells[j].width = Cm(widths[j])
    trPr = table.rows[0]._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:tblHeader"))
    aligns = body_align if isinstance(body_align, (list, tuple)) else [body_align] * len(headers)
    for i, row in enumerate(rows):
        for j, value in enumerate(row):
            set_cell(table.rows[i + 1].cells[j], value, size=font, align=aligns[j])
            table.rows[i + 1].cells[j].width = Cm(widths[j])
    return table


def add_table_caption(doc, text, width_cm):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="left", space_after=6, space_before=12, line_spacing=1.0)
    p.paragraph_format.left_indent = Cm(max(0.0, (16.5 - width_cm) / 2))
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text))
    return p


def add_picture(doc, path, width_cm):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="center", space_before=6, space_after=0)
    p.add_run().add_picture(str(path), width=Cm(width_cm))
    return p


def replace_title_text(doc: Document) -> None:
    replacements = {
        "ассистент": "профессор, д-р техн. наук, профессор",
        "В.В. Жукалин": "В.В. Фомин",
        "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 1": "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 3",
        "Создание проекта Next.js с использованием React, TypeScript и App Router": "МЕТОД ДЕРЕВА РЕШЕНИЙ",
        "по курсу: Web-программирование": "по курсу: Методы искусственного интеллекта",
    }
    found = {key: 0 for key in replacements}
    for child in list(doc.element.body)[:12]:
        for node in child.iter(qn("w:t")):
            if node.text in replacements:
                found[node.text] += 1
                node.text = replacements[node.text]
    missing = [key for key, count in found.items() if count != 1]
    if missing:
        raise RuntimeError(f"title replace failed: {missing} {found}")


def drop_body_after_title(doc: Document) -> None:
    body = doc.element.body
    children = list(body)
    if children[-1].tag != qn("w:sectPr"):
        raise RuntimeError("body sectPr missing")
    for child in children[12:-1]:
        body.remove(child)


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
    for rel in list(root):
        if rel.get("Id") not in used and "image" in (rel.get("Type") or ""):
            target = "word/" + rel.get("Target").replace("\\", "/")
            files.pop(target, None)
            root.remove(rel)
    files[rel_name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    tmp.replace(path)


def fmt_int(value: float) -> str:
    return str(int(round(value)))


def build() -> None:
    df = pd.read_csv(DATA / "tickets.csv")
    m = json.loads((DATA / "metrics.json").read_text(encoding="utf-8"))
    assert m["n"] == 90 and m["n_train"] == 63 and m["n_test"] == 27
    assert m["correct"] == 25
    assert abs(m["accuracy"] - 25 / 27) < 1e-9

    if OUT.exists():
        OUT.unlink()
    shutil.copy2(SRC, OUT)
    doc = Document(str(OUT))
    replace_title_text(doc)
    drop_body_after_title(doc)

    pages = {
        "1 Задание": "3",
        "2 Входные данные": "4",
        "3 Ход выполнения работы": "7",
        "4 Выводы": "10",
    }

    add_heading(doc, "СОДЕРЖАНИЕ")
    for title in ("1 Задание", "2 Входные данные", "3 Ход выполнения работы", "4 Выводы"):
        add_toc(doc, title, pages[title])

    doc.add_page_break()
    add_heading(doc, "1 Задание")
    add_text(
        doc,
        "Цель работы — по четырём числовым признакам обращения в IT-поддержку определить "
        "тип инцидента методом дерева решений и проверить ответ на тикетах, которые не "
        "входили в обучение.",
    )
    add_text(
        doc,
        "По практикуму нужно прочитать таблицу тикетов, назначить классам цвет для "
        "диаграммы, отделить 30 % строк в тест стратифицированной выборкой по классу, "
        "а 70 % оставить для обучения. Затем построить дерево решений, получить прогноз "
        "на тесте, посчитать матрицу ошибок и метрики качества, выписать логические "
        "правила листьев и построить диаграмму рассеяния.",
    )
    add_text(
        doc,
        "Расчёт выполнен в Python. Класс — тип инцидента: сеть, доступ или ПО. Признаки — "
        "число затронутых пользователей, длительность простоя, задержка ответа сервиса и "
        "доля ошибок. Номер тикета в модель не входит. Этот же набор данных рассчитан "
        "как общий для последующих работ по классификации.",
    )

    doc.add_page_break()
    add_heading(doc, "2 Входные данные")
    add_text(
        doc,
        "Исходный набор — 90 обращений в IT-поддержку. В каждой строке шесть полей: "
        "ticket_id — номер тикета; users_affected — число затронутых пользователей; "
        "downtime_min — длительность простоя, мин; latency_ms — задержка ответа сервиса, мс; "
        "error_rate_pct — доля ошибочных ответов, %; ticket_type — тип инцидента.",
    )
    add_text(
        doc,
        "Классов три, в каждом по 30 тикетов. Пользователи — от "
        f"{fmt_int(df['users_affected'].min())} до {fmt_int(df['users_affected'].max())}, "
        f"простой — от {fmt_int(df['downtime_min'].min())} до {fmt_int(df['downtime_min'].max())} мин, "
        f"задержка — от {fmt_int(df['latency_ms'].min())} до {fmt_int(df['latency_ms'].max())} мс, "
        f"доля ошибок — от {num(df['error_rate_pct'].min(), 1)} до {num(df['error_rate_pct'].max(), 1)} %. "
        "Пропусков нет. Одним порогом класс не отделить: у «сети» и «ПО» пересекаются "
        "интервалы простоя, а часть сетевых тикетов имеет пониженную задержку.",
    )
    add_text(
        doc,
        "Классы различаются типичной картиной инцидента и тем, какие признаки у него "
        "выражены сильнее. Краткое описание приведено в таблице 1.",
    )
    add_table_caption(doc, "Таблица 1 — Описание классов инцидентов", 16.5)
    add_table(
        doc,
        ["Класс", "Что это за инцидент", "Типичные признаки", "Тикетов"],
        [
            [
                "доступ",
                "Сбой входа, пароля или прав пользователя. Обычно точечный: страдает один человек или небольшая группа.",
                "Мало пользователей, короткий простой, низкая задержка и доля ошибок",
                "30",
            ],
            [
                "ПО",
                "Ошибка приложения или сервиса: падения, 5xx, неверный ответ бизнес-логики. Сервис «живой», но отвечает плохо.",
                "Средний охват, длинный простой, умеренная задержка, высокая доля ошибок",
                "30",
            ],
            [
                "сеть",
                "Деградация канала или инфраструктуры: таймауты, потери пакетов, медленный ответ у многих клиентов сразу.",
                "Много пользователей, заметный простой, высокая задержка, умеренные ошибки",
                "30",
            ],
        ],
        [2.2, 5.8, 6.3, 2.2],
        font=10,
        body_align=["center", "left", "left", "center"],
    )

    data_rows = []
    for row in df.itertuples(index=False):
        data_rows.append(
            [
                str(int(row.ticket_id)),
                fmt_int(row.users_affected),
                fmt_int(row.downtime_min),
                fmt_int(row.latency_ms),
                num(row.error_rate_pct, 1),
                row.ticket_type,
            ]
        )
    add_table_caption(doc, "Таблица 2 — Обращения в IT-поддержку", 16.5)
    add_table(
        doc,
        ["№", "Польз.", "Простой, мин", "Задержка, мс", "Ошибки, %", "Тип"],
        data_rows,
        [1.4, 2.2, 3.0, 3.2, 2.6, 2.6],
        font=10,
    )

    doc.add_page_break()
    add_heading(doc, "3 Ход выполнения работы")
    add_text(
        doc,
        "Последовательность расчёта показана на рисунке 1. Сначала читается таблица и "
        "проверяются типы полей. Классам назначены цвета для диаграммы: синий — сеть, "
        "зелёный — доступ, красный — ПО. Цвет в оценку дерева не входит.",
    )
    add_picture(doc, FIG / "scheme.png", 15.5)
    add_text(doc, "Рисунок 1 — Схема расчёта", indent=False, align="center", space_before=6, space_after=12)

    add_text(
        doc,
        "В обучение отобрано 70 % строк, то есть 63 из 90, в тест — 27. Разбиение "
        "стратифицировано по классу: в каждой части по 21 и по 9 тикетов каждого типа. "
        "Так тест не оказывается «пустыми» по одному из классов.",
    )
    add_text(
        doc,
        "Дерево решений построено по критерию Джини, глубина не больше трёх, в листе не "
        "меньше трёх обучающих тикетов. Ограничение глубины нужно, чтобы правила оставались "
        "читаемыми и дерево не запоминало шум. По 63 обучающим тикетам получено дерево "
        "рисунка 2.",
    )
    add_picture(doc, FIG / "tree.png", 16.0)
    add_text(doc, "Рисунок 2 — Дерево решений", indent=False, align="center", space_before=6, space_after=12)

    add_text(doc, "Логические правила листьев:")
    for rule in m["rules"]:
        add_text(doc, rule)
    add_text(
        doc,
        "Правила согласуются с предметным смыслом. Короткий простой отделяет сбои доступа. "
        "Среди более длинных простоев высокая задержка указывает на сеть, а меньшая — на "
        "ошибку приложения.",
        space_before=6,
    )

    cm = m["confusion"]
    add_text(
        doc,
        f"На 27 тестовых тикетах верны {m['correct']}, ошибочны {m['n_test'] - m['correct']}. "
        f"Accuracy = {pct(m['accuracy'])} %. Матрица ошибок приведена в таблице 2: строки — "
        "фактический класс, столбцы — прогноз дерева.",
        space_before=6,
    )
    add_table_caption(doc, "Таблица 3 — Матрица ошибок на тесте", 12.0)
    add_table(
        doc,
        ["Факт \\ прогноз", "сеть", "доступ", "ПО"],
        [
            ["сеть", str(cm[0][0]), str(cm[0][1]), str(cm[0][2])],
            ["доступ", str(cm[1][0]), str(cm[1][1]), str(cm[1][2])],
            ["ПО", str(cm[2][0]), str(cm[2][1]), str(cm[2][2])],
        ],
        [3.6, 2.4, 2.4, 2.4],
        font=11,
    )

    per = {row["class"]: row for row in m["per_class"]}
    add_table_caption(doc, "Таблица 4 — Метрики качества по классам", 14.5)
    add_table(
        doc,
        ["Класс", "Precision, %", "Recall, %", "F-measure, %", "Тикетов"],
        [
            [
                name,
                pct(per[name]["precision"]),
                pct(per[name]["recall"]),
                pct(per[name]["f1"]),
                str(per[name]["support"]),
            ]
            for name in CLASSES
        ],
        [2.8, 3.0, 2.8, 3.2, 2.4],
        font=11,
    )
    add_text(
        doc,
        f"Средние по трём классам: precision {pct(m['macro_precision'])} %, "
        f"recall {pct(m['macro_recall'])} %, F-measure {pct(m['macro_f1'])} %. "
        f"«Доступ» и «ПО» на тесте узнаны полностью. У «сети» 7 верных из 9: два тикета "
        f"с относительно низкой задержкой дерево отнесло к «ПО».",
        space_before=6,
    )

    err_ids = m["error_ids"]
    add_text(
        doc,
        "Диаграмма рассеяния на рисунке 3 построена по всем 90 тикетам. По горизонтали — "
        "задержка, по вертикали — доля ошибок. Цвет точки — фактический класс. Ошибки "
        f"теста (тикеты № {err_ids[0]} и № {err_ids[1]}) обведены чёрным: это пограничные "
        "сетевые обращения, у которых задержка ближе к зоне приложения.",
    )
    add_picture(doc, FIG / "scatter.png", 15.2)
    add_text(
        doc,
        "Рисунок 3 — Классы инцидентов в осях задержки и доли ошибок",
        indent=False,
        align="center",
        space_before=6,
        space_after=12,
    )

    doc.add_page_break()
    add_heading(doc, "4 Выводы")
    add_text(
        doc,
        "Тип IT-инцидента определён методом дерева решений по числу пользователей, "
        "простою, задержке и доле ошибок. В обучение вошло 63 тикета, в тест — 27, в "
        "каждом классе 21 и 9 соответственно. Дерево глубины 3 дало три читаемых правила: "
        "короткий простой — «доступ»; при большем простое низкая задержка — «ПО», высокая — «сеть».",
    )
    add_text(
        doc,
        f"На 27 тестовых тикетах верны {m['correct']}, accuracy = {pct(m['accuracy'])} %. "
        f"Средние precision, recall и F-measure равны {pct(m['macro_precision'])} %, "
        f"{pct(m['macro_recall'])} % и {pct(m['macro_f1'])} %. Набор пригоден как базовый "
        "для сравнения с другими методами классификации на тех же признаках и классах.",
    )

    doc.save(str(OUT))
    strip_unused_images(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
