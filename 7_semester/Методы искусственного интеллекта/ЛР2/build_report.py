# -*- coding: utf-8 -*-
"""Сборка отчёта ЛР2. В текст отчёта этот файл не попадает."""
from __future__ import annotations

import json
import shutil
import zipfile
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from lxml import etree
import pandas as pd

BASE = Path(__file__).resolve().parent
SRC = Path(
    r"C:\Users\skvor\OneDrive\Desktop\GitHub\study-guap\7_semester\Web-программирование\ЛР1_Отчет_Буренков_PairWise.docx"
)
OUT = BASE / "ЛР2_Отчет_Буренков_МИИ.docx"
FIG = BASE / "figures"
DATA = BASE / "data"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CLASSES = ["прокол", "тормоза", "трансмиссия"]
FEATURES = ["pressure_bar", "pad_mm", "chain_pct", "lever_mm"]
DIGITS = {"pressure_bar": 2, "pad_mm": 2, "chain_pct": 2, "lever_mm": 1}
FEATURE_RU = {
    "pressure_bar": "Давление, бар",
    "pad_mm": "Колодки, мм",
    "chain_pct": "Цепь, %",
    "lever_mm": "Ход ручки, мм",
}
COLORS = {"прокол": "#1f77b4", "тормоза": "#d62728", "трансмиссия": "#2ca02c"}


def num(value: float, digits: int) -> str:
    return f"{value:.{digits}f}".replace(".", ",")


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
    if text == "СОДЕРЖАНИЕ":
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


def add_toc(doc, title, key):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="left")
    p.paragraph_format.tab_stops.add_tab_stop(Cm(16.5), 2, 1)  # right, dots
    set_run_font(p.add_run(title))
    tab = p.add_run("\t")
    set_run_font(tab)
    set_run_font(p.add_run(key))
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


def set_cell(cell, text, *, bold=False, size=12, align="center"):
    cell.text = ""
    p = cell.paragraphs[0]
    format_paragraph(p, first_indent=False, align=align, line_spacing=1.0, space_after=0, space_before=0)
    set_run_font(p.add_run(str(text)), size=size, bold=bold)
    set_cell_border(cell)
    tcPr = cell._tc.get_or_add_tcPr()
    vAlign = OxmlElement("w:vAlign")
    vAlign.set(qn("w:val"), "center")
    tcPr.append(vAlign)


def add_table(doc, headers, rows, widths, *, font=12, header=True):
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
    tblW = OxmlElement("w:tblW")
    total = int(round(sum(widths) * 567))
    tblW.set(qn("w:w"), str(total))
    tblW.set(qn("w:type"), "dxa")
    tblPr.append(tblW)
    tblInd = OxmlElement("w:tblInd")
    tblInd.set(qn("w:w"), "0")
    tblInd.set(qn("w:type"), "dxa")
    tblPr.append(tblInd)
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for child in list(grid):
            grid.remove(child)
    else:
        grid = OxmlElement("w:tblGrid")
        tblPr.addnext(grid)
    for width in widths:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(int(round(width * 567))))
        grid.append(gc)
    for j, header_text in enumerate(headers):
        set_cell(table.rows[0].cells[j], header_text, bold=True, size=font)
        shade_header(table.rows[0].cells[j])
        table.rows[0].cells[j].width = Cm(widths[j])
    if header:
        trPr = table.rows[0]._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:tblHeader"))
    for i, row in enumerate(rows):
        for j, value in enumerate(row):
            set_cell(table.rows[i + 1].cells[j], value, size=font, align="center")
            table.rows[i + 1].cells[j].width = Cm(widths[j])
    return table


def add_table_caption(doc, text, width_cm):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="left", space_after=6, space_before=12, line_spacing=1.0)
    p.paragraph_format.left_indent = Cm(max(0.0, (16.5 - width_cm) / 2))
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text))
    return p


def add_picture(doc, path, width_cm):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="center", space_before=6, space_after=0)
    p.add_run().add_picture(str(path), width=Cm(width_cm))
    return p


def add_figure_caption(doc, text):
    return add_text(doc, text, indent=False, align="center", space_before=6, space_after=12)


def draw_scheme(path: Path) -> None:
    plt.rcParams["font.family"] = "Times New Roman"
    labels = [
        "Чтение таблицы заявок",
        "Цвет класса для диаграммы",
        "Разбиение 70 % / 30 % по классу",
        "Общий масштаб четырёх признаков",
        "Выбор k по обучающей выборке",
        "Классификация k ближайшими соседями",
        "Матрица ошибок и метрики качества",
        "Диаграмма рассеяния",
    ]
    n = len(labels)
    box_h = 0.72
    gap = 0.42
    top = 0.15
    fig_h = top + n * box_h + (n - 1) * gap + 0.12
    fig, ax = plt.subplots(figsize=(7.6, fig_h))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, fig_h)
    ax.axis("off")
    y = fig_h - top - box_h
    for i, text in enumerate(labels):
        ax.add_patch(
            plt.Rectangle((0.55, y), 8.9, box_h, fill=False, edgecolor="black", linewidth=1.0)
        )
        ax.text(5.0, y + box_h / 2, text, ha="center", va="center", fontsize=12, color="black")
        if i < n - 1:
            ax.annotate(
                "",
                xy=(5.0, y - gap + 0.02),
                xytext=(5.0, y - 0.01),
                arrowprops=dict(arrowstyle="-|>", color="black", lw=1.05, mutation_scale=11),
            )
        y -= box_h + gap
    fig.savefig(path, dpi=160, bbox_inches="tight", facecolor="white", pad_inches=0.12)
    plt.close()


def draw_scatter(df: pd.DataFrame, error_ids: list[int], path: Path) -> None:
    plt.rcParams["font.family"] = "Times New Roman"
    fig, ax = plt.subplots(figsize=(7.5, 5.7))
    for name in CLASSES:
        part = df[df["fault"] == name]
        ax.scatter(
            part["chain_pct"],
            part["pad_mm"],
            c=COLORS[name],
            s=42,
            label=name,
            zorder=3,
            edgecolors="white",
            linewidths=0.4,
        )
    err = df[df["request_id"].isin(error_ids)]
    ax.scatter(
        err["chain_pct"],
        err["pad_mm"],
        facecolors="none",
        edgecolors="black",
        s=130,
        linewidths=1.15,
        label="ошибка на тесте",
        zorder=4,
    )
    from matplotlib.ticker import FuncFormatter

    def comma_tick(value, _pos):
        return f"{value:g}".replace(".", ",")

    ax.xaxis.set_major_formatter(FuncFormatter(comma_tick))
    ax.yaxis.set_major_formatter(FuncFormatter(comma_tick))
    ax.set_xlabel("Удлинение цепи, %", fontsize=12)
    ax.set_ylabel("Остаток колодок, мм", fontsize=12)
    ax.tick_params(labelsize=11)
    ax.grid(True, color="#d0d0d0", linewidth=0.6)
    ax.set_axisbelow(True)
    ax.legend(frameon=True, fontsize=11)
    fig.tight_layout()
    fig.savefig(path, dpi=160, bbox_inches="tight", facecolor="white")
    plt.close()


def span(df, fault, col):
    part = df.loc[df["fault"] == fault, col]
    return float(part.min()), float(part.max())


def votes(neighbors):
    counts = {name: 0 for name in CLASSES}
    for nb in neighbors:
        counts[nb["fault"]] += 1
    return counts


def vote_phrase(neighbors) -> str:
    counts = votes(neighbors)
    items = [(name, counts[name]) for name in CLASSES if counts[name]]
    items.sort(key=lambda item: (-item[1], CLASSES.index(item[0])))
    return ", ".join(f"«{name}» — {count}" for name, count in items)


def class_stats(cm):
    rows = []
    n = int(sum(sum(row) for row in cm))
    for i, name in enumerate(CLASSES):
        tp = cm[i][i]
        fp = sum(cm[r][i] for r in range(3)) - tp
        fn = sum(cm[i]) - tp
        tn = n - tp - fp - fn
        precision = tp / (tp + fp) if tp + fp else 0.0
        recall = tp / (tp + fn) if tp + fn else 0.0
        f1 = 2 * precision * recall / (precision + recall) if precision + recall else 0.0
        rows.append(
            {
                "name": name,
                "tp": tp,
                "fp": fp,
                "tn": tn,
                "fn": fn,
                "precision": precision,
                "recall": recall,
                "f1": f1,
            }
        )
    return rows


def replace_title_text(doc: Document) -> None:
    replacements = {
        "ассистент": "профессор, д-р техн. наук, профессор",
        "В.В. Жукалин": "В.В. Фомин",
        "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 1": "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 2",
        "Создание проекта Next.js с использованием React, TypeScript и App Router": "МЕТОД K-БЛИЖАЙШИХ СОСЕДЕЙ",
        "по курсу: Web-программирование": "по курсу: Методы искусственного интеллекта",
    }
    body = doc.element.body
    children = list(body)
    found = {key: 0 for key in replacements}
    for child in children[:12]:
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


def prepare_document() -> Document:
    if OUT.exists():
        OUT.unlink()
    shutil.copy2(SRC, OUT)
    doc = Document(str(OUT))
    replace_title_text(doc)
    drop_body_after_title(doc)
    return doc


def build() -> None:
    df = pd.read_csv(DATA / "requests.csv")
    metrics = json.loads((DATA / "metrics.json").read_text(encoding="utf-8"))
    by_id = {int(row.request_id): row for row in df.itertuples(index=False)}
    pred = {item["request_id"]: item for item in metrics["predictions"]}
    cm = metrics["confusion"]
    stats = class_stats(cm)
    k = metrics["k"]
    errors = [item for item in metrics["predictions"] if not item["correct"]]
    error_ids = [item["request_id"] for item in errors]

    assert metrics["n"] == 90 and metrics["n_train"] == 63 and metrics["n_test"] == 27
    assert k == 11
    assert metrics["correct"] == 23 and len(errors) == 4
    assert all(item["fault"] == "трансмиссия" for item in errors)
    assert set(metrics["train_ids"]).isdisjoint(metrics["test_ids"])
    for item in metrics["predictions"]:
        ids = [nb["request_id"] for nb in item["neighbors"]]
        assert ids == sorted(ids, key=lambda i: next(nb["distance"] for nb in item["neighbors"] if nb["request_id"] == i))
        assert all(i in metrics["train_ids"] for i in ids)
        assert len(ids) == k

    stored = {row["class"]: row for row in metrics["per_class"]}
    for row in stats:
        src = stored[row["name"]]
        assert abs(row["precision"] - src["precision"]) < 1e-9
        assert abs(row["recall"] - src["recall"]) < 1e-9
        assert abs(row["f1"] - src["f1"]) < 1e-9

    p_min, p_max = span(df, "прокол", "pressure_bar")
    b_min, b_max = span(df, "тормоза", "pressure_bar")
    c_p_min, c_p_max = span(df, "прокол", "chain_pct")
    c_t_min, c_t_max = span(df, "трансмиссия", "chain_pct")
    assert p_max > b_min and c_t_min < c_p_max

    r32 = pred[32]
    r7 = pred[7]
    r63 = pred[63]
    r89 = pred[89]
    assert r32["correct"] and not r32["tie"] and votes(r32["neighbors"])["прокол"] == 9
    assert r7["tie"] and r7["correct"] and r7["fault"] == "прокол"
    assert r63["tie"] and not r63["correct"]
    assert r89["correct"] and r89["fault"] == "трансмиссия"
    assert votes(r7["neighbors"])["прокол"] == 4 and votes(r7["neighbors"])["тормоза"] == 4
    assert votes(r63["neighbors"])["прокол"] == 4 and votes(r63["neighbors"])["тормоза"] == 4
    assert votes(r63["neighbors"])["трансмиссия"] == 3

    FIG.mkdir(parents=True, exist_ok=True)
    draw_scheme(FIG / "scheme.png")
    draw_scatter(df, error_ids, FIG / "scatter.png")

    doc = prepare_document()
    macro_p = sum(row["precision"] for row in stats) / 3
    macro_r = sum(row["recall"] for row in stats) / 3
    macro_f = sum(row["f1"] for row in stats) / 3
    loo = {row["k"]: row for row in metrics["loo"]}
    test_ids = ", ".join(str(i) for i in metrics["test_ids"])
    mu = metrics["mu"]
    sd = metrics["sd"]
    cent = metrics["centroids"]

    add_heading(doc, "СОДЕРЖАНИЕ")
    add_toc(doc, "1 Задание", "@@1@@")
    add_toc(doc, "2 Входные данные", "@@2@@")
    add_toc(doc, "3 Ход выполнения работы", "@@3@@")
    add_toc(doc, "4 Выводы", "@@4@@")

    doc.add_page_break()
    add_heading(doc, "1 Задание")
    add_text(
        doc,
        "Цель работы — по четырём измерениям велосипеда, принятого в мастерскую, определить тип поломки методом k ближайших соседей и проверить этот ответ на заявках, которые не входили в обучение.",
    )
    add_text(
        doc,
        "По практикуму нужно прочитать таблицу заявок, назначить классам цвет для диаграммы рассеяния, отделить 30 % строк в тест стратифицированной выборкой по классу, а 70 % оставить для обучения. Затем отнести тестовые заявки к классу по k ближайшим соседям, посчитать матрицу ошибок, accuracy, precision, recall и F-measure и построить диаграмму рассеяния.",
    )
    add_text(
        doc,
        "Расчёт выполнен в Python. Класс — тип поломки: прокол, тормоза или трансмиссия. Признаки — давление в покрышке, остаток тормозных колодок, удлинение цепи и свободный ход тормозной ручки. Номер заявки в расстояние не входит.",
    )

    doc.add_page_break()
    add_heading(doc, "2 Входные данные")
    add_text(
        doc,
        "Исходный набор — 90 заявок мастерской ремонта велосипедов. В каждой строке шесть полей: request_id — номер заявки, pressure_bar — давление в передней покрышке, бар, pad_mm — остаточная толщина колодок, мм, chain_pct — удлинение цепи, %, lever_mm — свободный ход тормозной ручки, мм, fault — тип поломки. Признаки модели — pressure_bar, pad_mm, chain_pct и lever_mm, все вещественные. Класс — fault. Номер заявки нужен только чтобы отличать строки.",
    )
    add_text(
        doc,
        f"Классов три, в каждом по 30 заявок. Давление лежит в интервале от {num(metrics['ranges']['pressure_bar']['min'], 2)} до {num(metrics['ranges']['pressure_bar']['max'], 2)} бар, остаток колодок — от {num(metrics['ranges']['pad_mm']['min'], 2)} до {num(metrics['ranges']['pad_mm']['max'], 2)} мм, удлинение цепи — от {num(metrics['ranges']['chain_pct']['min'], 2)} до {num(metrics['ranges']['chain_pct']['max'], 2)} %, ход ручки — от {num(metrics['ranges']['lever_mm']['min'], 1)} до {num(metrics['ranges']['lever_mm']['max'], 1)} мм. Пропусков нет.",
    )
    add_text(
        doc,
        f"Одним порогом класс не отделить. Давление у «прокол» занимает от {num(p_min, 2)} до {num(p_max, 2)} бар, у «тормоза» — от {num(b_min, 2)} до {num(b_max, 2)} бар, интервалы накладываются. Удлинение цепи у «трансмиссия» — от {num(c_t_min, 2)} до {num(c_t_max, 2)} %, у «прокол» — от {num(c_p_min, 2)} до {num(c_p_max, 2)} %, эти промежутки тоже пересекаются. Полный набор приведён в таблице 1.",
    )
    data_widths = [1.6, 2.8, 2.8, 2.3, 3.3, 3.7]
    add_table_caption(doc, "Таблица 1 — Заявки мастерской ремонта велосипедов", sum(data_widths))
    data_rows = []
    for row in df.itertuples(index=False):
        data_rows.append(
            [
                str(int(row.request_id)),
                num(row.pressure_bar, 2),
                num(row.pad_mm, 2),
                num(row.chain_pct, 2),
                num(row.lever_mm, 1),
                row.fault,
            ]
        )
    add_table(
        doc,
        ["№", "Давление, бар", "Колодки, мм", "Цепь, %", "Ход ручки, мм", "Тип поломки"],
        data_rows,
        data_widths,
        font=12,
    )

    doc.add_page_break()
    add_heading(doc, "3 Ход выполнения работы")
    add_text(
        doc,
        "Последовательность расчёта показана на рисунке 1. Сначала читается таблица и проверяются типы: номер заявки — целое число, четыре измерения — вещественные, тип поломки — одно из трёх значений. Пропусков нет.",
    )
    add_picture(doc, FIG / "scheme.png", 14.2)
    add_figure_caption(doc, "Рисунок 1 — Схема расчёта")
    add_text(
        doc,
        "Классу на диаграмме назначен цвет: синий — прокол, красный — тормоза, зелёный — трансмиссия. Цвет не подставляется в расстояние и нужен только чтобы на рисунке было видно, где классы расходятся и где накладываются.",
    )
    add_text(
        doc,
        f"В обучение отобрано 70 % строк, то есть {metrics['n_train']} из {metrics['n']}, в тест — 30 %, то есть {metrics['n_test']}. Отбор стратифицирован по классу: в каждом классе 21 заявка уходит в обучение и 9 в тест, доли классов не смещаются. Номера заявок теста: {test_ids}. Остальные номера составляют обучение.",
    )
    add_text(
        doc,
        "Четыре признака измерены в разных единицах. Ход ручки измеряется миллиметрами, удлинение цепи — долями процента. Если считать расстояние по сырым числам, ход ручки перевесит остальные признаки. Перед поиском соседей каждый признак приведён к нулевому среднему и единичному разбросу по 63 обучающим заявкам. Разброс — корень из среднего квадрата отклонения от среднего. Те же среднее и разброс применены к тесту, тест в эти две величины не входит. Округлённые значения даны в таблице 2.",
    )
    scale_w = [6.4, 4.2, 4.2]
    add_table_caption(doc, "Таблица 2 — Среднее и разброс обучающей выборки", sum(scale_w))
    add_table(
        doc,
        ["Признак", "Среднее", "Разброс"],
        [
            [FEATURE_RU[col], num(mu[col], 4), num(sd[col], 4)]
            for col in FEATURES
        ],
        scale_w,
        font=12,
    )
    add_text(
        doc,
        "Расстояние между заявками — евклидово: корень из суммы квадратов разностей четырёх приведённых признаков. Класс заявки — тот, который среди k ближайших обучающих заявок встречается чаще других. Если два класса набрали поровну голосов, берётся класс более близкого из этих соседей. При трёх классах такая ничья возможна и при нечётном k.",
        space_before=12,
    )
    add_text(
        doc,
        "Число соседей выбрано по обучающим 63 заявкам, тестовые заявки в этом выборе не участвовали. Для каждого нечётного k от 1 до 15 обучающая заявка по очереди откладывалась: среднее и разброс считались по остальным 62, соседи искались среди них, ответ сравнивался с её классом. Доля верных ответов приведена в таблице 3.",
    )
    loo_w = [3.2, 4.6, 4.4]
    add_table_caption(doc, "Таблица 3 — Число соседей на обучающей выборке", sum(loo_w))
    add_table(
        doc,
        ["k", "Верных ответов", "Доля, %"],
        [[str(row["k"]), f"{row['correct']} из {row['n']}", pct(row["accuracy"])] for row in metrics["loo"]],
        loo_w,
        font=12,
    )
    add_text(
        doc,
        f"Наибольшая доля — {loo[11]['correct']} из {loo[11]['n']}, то есть {pct(loo[11]['accuracy'])} %, при k = 11. При k = 1 верных только {loo[1]['correct']}: один сосед повторяет случайный шум измерения. При k = 15 верных {loo[15]['correct']}, это меньше, чем при k = 11. В зоне пересечения классов дальний сосед уже относится к другой поломке и портит большинство. Дальше используется k = 11.",
        space_before=6,
    )
    nn32 = r32["neighbors"][0]
    add_text(
        doc,
        f"Как считается ответ, видно на тестовой заявке 32. Факт — «прокол»: давление {num(r32['pressure_bar'], 2)} бар, колодки {num(r32['pad_mm'], 2)} мм, цепь {num(r32['chain_pct'], 2)} %, ход ручки {num(r32['lever_mm'], 1)} мм. Одиннадцать соседей перечислены в таблице 4 в порядке возрастания расстояния. Девять из них — «прокол», поэтому поставлен тот же класс, и он совпал с фактом. Ближайшая заявка — {nn32['request_id']}, расстояние {num(nn32['distance'], 3)}.",
        space_before=12,
    )
    nb_w = [3.4, 4.2, 4.4]
    add_table_caption(doc, "Таблица 4 — Соседи заявки 32", sum(nb_w))
    add_table(
        doc,
        ["Заявка", "Класс", "Расстояние"],
        [[str(nb["request_id"]), nb["fault"], num(nb["distance"], 3)] for nb in r32["neighbors"]],
        nb_w,
        font=12,
    )
    n7_p = r7["neighbors"][0]
    n7_b = next(nb for nb in r7["neighbors"] if nb["fault"] == "тормоза")
    n63_drive = [nb for nb in r63["neighbors"] if nb["fault"] == "трансмиссия"][:2]
    n63_p = next(nb for nb in r63["neighbors"] if nb["fault"] == "прокол")
    add_text(
        doc,
        f"Ничья на тесте возникла дважды. У заявки 7 голоса разделились: «прокол» — 4, «тормоза» — 4, «трансмиссия» — 3. Ближе из двух лидирующих заявка {n7_p['request_id']} класса «прокол», расстояние {num(n7_p['distance'], 3)}, против {num(n7_b['distance'], 3)} у заявки {n7_b['request_id']} класса «тормоза». Поставлен «прокол», и это совпало с фактом. У заявки 63 факт — «трансмиссия», а голоса такие же по счёту: «прокол» — 4, «тормоза» — 4, «трансмиссия» — 3. Два ближайших соседа — заявки {n63_drive[0]['request_id']} и {n63_drive[1]['request_id']} класса «трансмиссия», расстояния {num(n63_drive[0]['distance'], 3)} и {num(n63_drive[1]['distance'], 3)}, но третьего голоса этому классу не хватило. Из лидеров ближе заявка {n63_p['request_id']} класса «прокол», расстояние {num(n63_p['distance'], 3)}, поэтому поставлен «прокол».",
        space_before=12,
    )
    add_text(
        doc,
        "Ответы по всем 27 тестовым заявкам собраны в таблице 5. По ней же считается матрица ошибок: строка — фактический класс, столбец — поставленный.",
    )
    test_w = [2.4, 4.2, 4.4, 3.6]
    add_table_caption(doc, "Таблица 5 — Классификация тестовых заявок", sum(test_w))
    test_rows = []
    for item in metrics["predictions"]:
        test_rows.append(
            [
                str(item["request_id"]),
                item["fault"],
                item["pred"],
                "верно" if item["correct"] else "ошибка",
            ]
        )
    add_table(doc, ["№", "Факт", "Прогноз", "Итог"], test_rows, test_w, font=12)
    conf_w = [4.2, 3.6, 3.6, 3.6]
    add_table_caption(doc, "Таблица 6 — Матрица ошибок на тестовых заявках", sum(conf_w))
    add_table(
        doc,
        ["Факт \\ прогноз", "прокол", "тормоза", "трансмиссия"],
        [[CLASSES[i], *[str(cm[i][j]) for j in range(3)]] for i in range(3)],
        conf_w,
        font=12,
    )
    add_text(
        doc,
        f"Верных ответов {metrics['correct']} из {metrics['n_test']}, ошибок {metrics['errors']}. Accuracy = {metrics['correct']}/{metrics['n_test']} = {pct(metrics['accuracy'])} %. Все 9 проколов и все 9 тормозов узнаны. Из 9 трансмиссий верны 5: одна отнесена к проколу, три — к тормозам.",
        space_before=12,
    )
    add_text(
        doc,
        "Precision класса — доля верных среди заявок, которым поставлен этот класс. Recall — доля найденных заявок класса среди тех, у которых он действительно был; в статистике точности ту же величину называют sensitivity. F-measure — гармоническое среднее precision и recall: удвоенное произведение, делённое на сумму. Числители и знаменатели берутся из таблицы 6, сами величины — в таблице 7.",
    )
    met_w = [3.2, 1.5, 1.5, 1.5, 1.5, 2.6, 2.3, 2.4]
    add_table_caption(doc, "Таблица 7 — Precision, recall и F-measure по классам", sum(met_w))
    add_table(
        doc,
        ["Класс", "TP", "FP", "TN", "FN", "Precision, %", "Recall, %", "F-measure, %"],
        [
            [
                row["name"],
                str(row["tp"]),
                str(row["fp"]),
                str(row["tn"]),
                str(row["fn"]),
                pct(row["precision"]),
                pct(row["recall"]),
                pct(row["f1"]),
            ]
            for row in stats
        ],
        met_w,
        font=12,
    )
    err_chains = [item["chain_pct"] for item in errors]
    add_text(
        doc,
        f"Средние по трём классам: precision {pct(macro_p)} %, recall {pct(macro_r)} %, F-measure {pct(macro_f)} %. Средний recall совпадает с accuracy, потому что в тесте по 9 заявок каждого класса. У «трансмиссия» precision = {pct(stats[2]['precision'])} %: все {stats[2]['tp']} заявок, которым поставлен этот класс, действительно им оказались. Recall этого класса — {pct(stats[2]['recall'])} %, F-measure — {pct(stats[2]['f1'])} %. Четыре ошибки — заявки {', '.join(str(i) for i in error_ids)}. У них удлинение цепи от {num(min(err_chains), 2)} до {num(max(err_chains), 2)} %. Это не порог: заявка 89 с цепью {num(r89['chain_pct'], 2)} % и колодками {num(r89['pad_mm'], 2)} мм тоже лежит в том же промежутке, но {votes(r89['neighbors'])['трансмиссия']} из 11 соседей — «трансмиссия», и ответ верный. Ошибка определяется соседством в четырёх признаках, а не одним столбцом.",
        space_before=12,
    )
    add_text(
        doc,
        "Диаграмма рассеяния на рисунке 2 построена по всем 90 заявкам. По горизонтали — удлинение цепи, по вертикали — остаток колодок. Цвет точки — фактический класс. Четыре ошибки теста обведены чёрным. Тормоза в основном ниже: колодки тоньше. Трансмиссия смещена вправо, к большему удлинению, но левый край зелёного облака накрывает синие и красные точки. Обведённые заявки лежат в этом пересечении. Давление и ход ручки на рисунок не вынесены, а расстояние считается по всем четырём признакам, поэтому соседство на рисунке не равно соседству для классификатора. Заявка 89 на плоскости рядом с ошибками, а в четырёхмерном расстоянии у неё другое окружение.",
    )
    add_picture(doc, FIG / "scatter.png", 15.0)
    add_figure_caption(doc, "Рисунок 2 — Классы поломок в осях цепи и колодок")

    doc.add_page_break()
    add_heading(doc, "4 Выводы")
    add_text(
        doc,
        f"Тип поломки велосипеда определён методом k ближайших соседей по давлению в покрышке, остатку колодок, удлинению цепи и ходу тормозной ручки. В обучение вошло 63 заявки, в тест — 27, в каждом классе 21 и 9. Признаки перед расчётом расстояния приведены к одному масштабу по обучающей выборке. Число соседей k = 11: на обучающей части оно дало {loo[11]['correct']} верных ответов из 63, а k = 1 и k = 15 дали меньше, {loo[1]['correct']} и {loo[15]['correct']}.",
    )
    add_text(
        doc,
        f"На 27 тестовых заявках верны {metrics['correct']}, accuracy = {pct(metrics['accuracy'])} %. Средние precision, recall и F-measure равны {pct(macro_p)} %, {pct(macro_r)} % и {pct(macro_f)} %. Прокол и тормоза на тесте узнаны полностью, трансмиссия — в 5 случаях из 9. Когда модель ставит «трансмиссия», на этих заявках ошибки нет: precision класса 100 %. Четыре пропущенные трансмиссии лежат в пересечении классов, это видно и по голосам соседей, и по диаграмме. Для приёмки велосипеда ответ годится как подсказка по проколу и тормозам. Спорную трансмиссию с умеренным удлинением цепи мастер должен проверять отдельно: по этим данным безошибочного разделения нет.",
    )

    doc.save(str(OUT))
    # by_id kept so a missing request fails above if the example id disappears
    assert by_id[32].fault == "прокол"


def strip_orphan_media(path: Path) -> None:
    tmp = path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(path, "r") as zin:
        rels = zin.read("word/_rels/document.xml.rels")
        root = etree.fromstring(rels)
        used = set()
        for blob_name in zin.namelist():
            if not (blob_name.startswith("word/") and blob_name.endswith(".xml")):
                continue
            if "/_rels/" in blob_name:
                continue
            for el in etree.fromstring(zin.read(blob_name)).iter():
                for attr, value in el.attrib.items():
                    local = attr.split("}")[-1]
                    if local in ("id", "embed"):
                        used.add(value)
        keep_rels = []
        drop_targets = set()
        for rel in root:
            rid = rel.get("Id")
            target = rel.get("Target")
            if rid in used or (target and not target.startswith("media/")):
                keep_rels.append(rel)
            else:
                if target:
                    drop_targets.add("word/" + target.replace("\\", "/"))
        for rel in list(root):
            if rel not in keep_rels:
                root.remove(rel)
        new_rels = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename in drop_targets:
                    continue
                data = new_rels if item.filename == "word/_rels/document.xml.rels" else zin.read(item.filename)
                zout.writestr(item, data)
    tmp.replace(path)


def patch_toc(pages: dict[str, str]) -> None:
    tmp = OUT.with_suffix(".tmp.docx")
    with zipfile.ZipFile(OUT, "r") as zin, zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                text = data.decode("utf-8")
                for key, value in pages.items():
                    needle = f"@@{key}@@"
                    if text.count(needle) != 1:
                        raise RuntimeError(f"toc marker {needle} count {text.count(needle)}")
                    text = text.replace(needle, value)
                data = text.encode("utf-8")
            zout.writestr(item, data)
    tmp.replace(OUT)


def read_heading_pages() -> dict[str, str]:
    import win32com.client

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = word.Documents.Open(str(OUT), ReadOnly=True)
    found = {}
    wanted = {
        "1 Задание": "1",
        "2 Входные данные": "2",
        "3 Ход выполнения работы": "3",
        "4 Выводы": "4",
    }
    for paragraph in doc.Paragraphs:
        raw = paragraph.Range.Text.replace("\r", "").replace("\x07", "").strip()
        if raw in wanted:
            page = int(paragraph.Range.Information(3))
            found[wanted[raw]] = str(page)
    doc.Close(False)
    word.Quit()
    if set(found) != set(wanted.values()):
        raise RuntimeError(f"pages not found: {found}")
    return found


def main() -> None:
    build()
    strip_orphan_media(OUT)
    pages = read_heading_pages()
    patch_toc(pages)
    print("pages", pages)
    print(OUT)


if __name__ == "__main__":
    main()
