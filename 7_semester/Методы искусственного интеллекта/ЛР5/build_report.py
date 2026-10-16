# -*- coding: utf-8 -*-
"""Сборка отчёта ЛР5 МИИ. В текст отчёта этот файл не попадает."""
from __future__ import annotations

import importlib.util
import json
import shutil
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

BASE = Path(__file__).resolve().parent
SRC_BUILD = BASE.parent / "ЛР2" / "build_report.py"
spec = importlib.util.spec_from_file_location("lr2_docx", SRC_BUILD)
docx = importlib.util.module_from_spec(spec)
spec.loader.exec_module(docx)

OUT = BASE / "ЛР5_Отчет_Буренков_МИИ.docx"
docx.OUT = OUT
FIG = BASE / "figures"
DATA = BASE / "data"


def pct(value: float) -> str:
    return f"{float(value) * 100:.2f}".replace(".", ",")


def replace_title(doc: Document) -> None:
    replacements = {
        "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 1": "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 5",
        "ЛИНЕЙНАЯ РЕГРЕССИЯ": "МЕТОД ОПОРНЫХ ВЕКТОРОВ",
    }
    found = {key: 0 for key in replacements}
    for child in list(doc.element.body)[:12]:
        for node in child.iter(qn("w:t")):
            if node.text in replacements:
                found[node.text] += 1
                node.text = replacements[node.text]
    missing = [key for key, count in found.items() if count != 1]
    if missing:
        raise RuntimeError(f"title replace failed: {found}")


def build() -> None:
    m = json.loads((DATA / "metrics.json").read_text(encoding="utf-8"))
    src = BASE.parent / "ЛР1" / "ЛР1_Отчет_Буренков_МИИ.docx"
    if OUT.exists():
        OUT.unlink()
    shutil.copy2(src, OUT)
    doc = Document(str(OUT))
    replace_title(doc)
    docx.drop_body_after_title(doc)
    docx.add_heading(doc, "СОДЕРЖАНИЕ")
    docx.add_toc(doc, "1 Задание", "@@1@@")
    docx.add_toc(doc, "2 Входные данные", "@@2@@")
    docx.add_toc(doc, "3 Ход выполнения работы", "@@3@@")
    docx.add_toc(doc, "4 Выводы", "@@4@@")

    doc.add_page_break()
    docx.add_heading(doc, "1 Задание")
    docx.add_text(
        doc,
        "Цель работы — отнести розыгрыш мяча к одному из трёх классов методом опорных векторов "
        "и сравнить точность с методами лабораторных 2, 3 и 4 на тех же строках.",
    )
    docx.add_text(
        doc,
        "По практикуму читается своя таблица, модель обучается на обучающей части, для теста "
        "считаются класс и вероятности классов. Ядро — радиальная базисная функция, штраф за "
        "неверную точку равен 1, параметр сигма равен 0,5. Затем считаются матрица ошибок, "
        "accuracy, precision, recall и F-measure. В вывод входит сравнение четырёх методов.",
    )

    doc.add_page_break()
    docx.add_heading(doc, "2 Входные данные")
    docx.add_text(
        doc,
        "Таблица, классы и разбиение те же, что в лабораторных 2, 3 и 4. Девяносто розыгрышей, "
        "по 30 футбольных, баскетбольных и волейбольных. Обучение — 63 строки, тест — 27, "
        "в каждом классе 21 и 9. Пять кодов: динамика владения, игровая зона, удержание контроля, "
        "вектор атаки и контакт с землёй. Обычные наборы, по 8 строк в классе: футбольный 0, 0, 1, 0, 2; "
        "баскетбольный 1, 2, 2, 1, 0; волейбольный 2, 1, 0, 1, 0. Два других набора встречаются "
        "сразу у нескольких мячей, а код зоны 0, 1 и 2 есть у всех трёх классов.",
    )

    doc.add_page_break()
    docx.add_heading(doc, "3 Ход выполнения работы")
    docx.add_text(
        doc,
        "Расчёт выполнен на Python. Схема приведена на рисунке 1. Штраф равен 1, ядро RBF, сигма равна 0,5.",
    )
    docx.add_picture(doc, FIG / "scheme.png", 15.5)
    docx.add_figure_caption(doc, "Рисунок 1 — Схема расчёта")
    counts = m["n_support"]
    docx.add_text(
        doc,
        f"Опорных векторов {m['n_support_total']}: футбольный — {counts['футбольный']}, "
        f"баскетбольный — {counts['баскетбольный']}, волейбольный — {counts['волейбольный']}. "
        "Граница проходит по нескольким разным розыгрышам, а не по одному обычному набору.",
    )
    docx.add_table_caption(doc, "Таблица 1 — Число опорных векторов", 12)
    docx.add_table(
        doc,
        ["Класс", "Опорных векторов"],
        [[name, str(counts[name])] for name in ("футбольный", "баскетбольный", "волейбольный")],
        [6.0, 4.5],
    )
    example = m["example"]
    probs = example["probs"]
    fx = example["features"]
    docx.add_text(
        doc,
        f"Для строки {example['id']} коды {fx['net_share']}, {fx['play_zone']}, {fx['hold']}, "
        f"{fx['attack']}, {fx['ground']}. Факт — «{example['ball']}», прогноз — «{example['pred']}». "
        f"Вероятности: футбольный {pct(probs['футбольный'])} %, "
        f"баскетбольный {pct(probs['баскетбольный'])} %, волейбольный {pct(probs['волейбольный'])} %.",
    )
    cm = m["confusion"]
    docx.add_table_caption(doc, "Таблица 2 — Матрица ошибок на 27 тестовых строках", 12)
    docx.add_table(
        doc,
        ["Факт \\ прогноз", "футбольный", "баскетбольный", "волейбольный"],
        [[name, *[str(cm[i][j]) for j in range(3)]] for i, name in enumerate(("футбольный", "баскетбольный", "волейбольный"))],
        [4.0, 2.8, 3.4, 3.2],
        font=11,
    )
    docx.add_text(
        doc,
        f"Верных ответов {m['correct']} из 27. Accuracy, средние precision, recall и F-measure равны "
        f"{pct(m['accuracy'])} %.",
    )
    docx.add_picture(doc, FIG / "scatter.png", 14)
    docx.add_figure_caption(doc, "Рисунок 2 — Игровая зона и вектор атаки")
    docx.add_text(
        doc,
        "Сравнение сделано на одном и том же тесте из 27 строк. Лабораторная 2 — k ближайших соседей "
        "при k = 1. Лабораторная 3 — дерево решений глубиной не больше 3. Лабораторная 4 — наивный Байес. "
        "Лабораторная 5 — опорные векторы. Числа собраны в таблице 3.",
    )
    docx.add_table_caption(doc, "Таблица 3 — Точность четырёх методов на одном тесте", 16.5)
    docx.add_table(
        doc,
        ["Работа", "Метод", "Верных", "Accuracy, %", "F-measure, %"],
        [
            [item["lab"], item["method"], f"{item['correct']} из 27", pct(item["accuracy"]), pct(item["f1"])]
            for item in m["comparison"]
        ],
        [2.2, 4.8, 2.8, 3.3, 3.4],
        font=12,
    )

    doc.add_page_break()
    docx.add_heading(doc, "4 Выводы")
    ranked = sorted(m["comparison"], key=lambda item: (-item["correct"], item["lab"]))
    best = ranked[0]
    rest = ", ".join(f"{item['method']} — {item['correct']} из 27" for item in ranked[1:])
    docx.add_text(
        doc,
        f"Опорные векторы на этом разбиении верно назвали {m['correct']} из 27. "
        f"Accuracy равна {pct(m['accuracy'])} %. Опорных векторов {m['n_support_total']}: "
        "граница класса держится на разных розыгрышах, включая те, что совпадают у нескольких мячей.",
    )
    docx.add_text(
        doc,
        f"На одном тесте точнее остальных метод «{best['method']}»: {best['correct']} из 27, "
        f"accuracy {pct(best['accuracy'])} %. Дальше {rest}. "
        "Стопроцентного ответа нет ни у одного метода. Два набора из пяти кодов встречаются "
        "у разных мячей, и эти строки нельзя отнести к классу без ошибки. Игровая зона больше "
        "не отделяет мячи сама по себе.",
    )
    doc.save(str(OUT))


def assert_clean() -> None:
    import zipfile
    from lxml import etree

    with zipfile.ZipFile(OUT) as zf:
        text = zf.read("word/document.xml").decode("utf-8")
    plain = "".join(etree.fromstring(text.encode("utf-8")).itertext())
    banned = ["тикет", "велосипед", "iris", "sklearn", "matplotlib", "KNIME", "build_report", "compute_", "seed"]
    hits = [word for word in banned if word.lower() in plain.lower()]
    if hits:
        raise RuntimeError(f"лишнее в отчёте: {hits}")
    if "@@" in plain:
        raise RuntimeError("маркер содержания не заменён")
    if plain.count("Python") != 1:
        raise RuntimeError(f"Python упомянут {plain.count('Python')} раз")
    for needle in ("МЕТОД ОПОРНЫХ ВЕКТОРОВ", "Таблица 3 — Точность четырёх методов", "точнее остальных", "дерево решений"):
        if needle not in plain:
            raise RuntimeError(f"нет {needle}")


def main() -> None:
    build()
    docx.strip_unused_images(OUT)
    pages = docx.read_heading_pages()
    docx.patch_toc(pages)
    again = docx.read_heading_pages()
    if again != pages:
        raise RuntimeError(f"страницы съехали: {pages} -> {again}")
    assert_clean()
    print("pages", pages)
    print(OUT)


if __name__ == "__main__":
    main()
