# -*- coding: utf-8 -*-
"""Сборка отчёта ЛР4 МИИ. В текст отчёта этот файл не попадает."""
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

OUT = BASE / "ЛР4_Отчет_Буренков_МИИ.docx"
docx.OUT = OUT
FIG = BASE / "figures"
DATA = BASE / "data"
FEATURE_RU = {
    "net_share": "Динамика владения",
    "play_zone": "Игровая зона",
    "hold": "Удержание контроля",
    "attack": "Вектор атаки",
    "ground": "Контакт с землёй",
}


def num(value: float, digits: int = 4) -> str:
    return f"{float(value):.{digits}f}".replace(".", ",")


def pct(value: float) -> str:
    return f"{float(value) * 100:.2f}".replace(".", ",")


def replace_title(doc: Document) -> None:
    replacements = {
        "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 1": "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 4",
        "ЛИНЕЙНАЯ РЕГРЕССИЯ": "МЕТОД НАИВНОГО БАЙЕСА",
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
        "Цель работы — по пяти кодам розыгрыша определить вид мяча наивным байесовским методом "
        "и проверить ответ на наблюдениях, которые не входили в обучение.",
    )
    docx.add_text(
        doc,
        "По практикуму читается своя таблица, 70 % строк остаются для обучения и 30 % откладываются "
        "в тест со стратификацией по классу. По обучающим строкам для каждого класса и каждого кода "
        "считаются среднее и стандартное отклонение. Числовой код описывается нормальным распределением. "
        "Для тестовой строки вероятность класса равна произведению вероятностей её кодов и вероятности "
        "самого класса. Ответ — класс с наибольшей вероятностью. Затем считаются матрица ошибок, "
        "accuracy, precision, recall и F-measure.",
    )

    doc.add_page_break()
    docx.add_heading(doc, "2 Входные данные")
    docx.add_text(
        doc,
        "Объект тот же, что в лабораторных 2 и 3: розыгрыш мяча. Классы — футбольный, баскетбольный "
        "и волейбольный. Пять кодов принимают значения 0, 1 или 2. Всего 90 строк, по 30 в классе. "
        "В обучении 63 строки, в тесте 27. В каждом классе это 21 и 9 строк. Номер строки в расчёт не входит.",
    )
    docx.add_text(doc, "Расшифровка кодов собрана в таблице 1. Обычные наборы — в таблице 2.")
    docx.add_table_caption(doc, "Таблица 1 — Расшифровка кодов", 16.5)
    docx.add_table(
        doc,
        ["Код", "0", "1", "2"],
        [
            ["Динамика владения", "пас партнёру", "удар или бросок в цель", "через сетку"],
            ["Игровая зона", "цель у земли", "на уровне игрока", "высокая цель"],
            ["Удержание контроля", "краткий удар", "контакт ногой", "контакт рукой"],
            ["Вектор атаки", "вдоль площадки", "круто вверх", "по дуге"],
            ["Контакт с землёй", "в воздухе", "один отскок", "катится"],
        ],
        [4.2, 4.1, 4.1, 4.1],
        font=12,
    )
    docx.add_table_caption(doc, "Таблица 2 — Обычные коды трёх мячей", 16.5)
    docx.add_table(
        doc,
        ["Мяч", "Владение", "Зона", "Контроль", "Вектор", "Земля"],
        [
            ["футбольный", "0", "0", "1", "0", "2"],
            ["баскетбольный", "1", "2", "2", "1", "0"],
            ["волейбольный", "2", "1", "0", "1", "0"],
        ],
        [3.5, 2.6, 2.6, 2.6, 2.6, 2.6],
    )
    docx.add_text(
        doc,
        "Обычный набор — самый частый в классе, по 8 строк, а не все 30. Остальные розыгрыши отличаются. "
        "Два полных набора кодов принадлежат сразу нескольким мячам. Коды 1, 2, 0, 1, 0 — высокий "
        "короткий удар вверх: футбольный удар головой, баскетбольное добивание и волейбольная атака. "
        "Коды 0, 1, 2, 0, 0 — передача руками на уровне игрока. Игровая зона у всех трёх мячей "
        "принимает и 0, и 1, и 2, поэтому по зоне класс не называется.",
    )

    doc.add_page_break()
    docx.add_heading(doc, "3 Ход выполнения работы")
    docx.add_text(
        doc,
        "Расчёт выполнен на Python. Схема приведена на рисунке 1. Цвета классов на диаграмме рассеяния: "
        "футбольный синий, баскетбольный оранжевый, волейбольный зелёный.",
    )
    docx.add_picture(doc, FIG / "scheme.png", 15.5)
    docx.add_figure_caption(doc, "Рисунок 1 — Схема расчёта")
    docx.add_text(
        doc,
        "Априорная вероятность каждого класса равна 21/63, то есть одной трети. Для числового кода "
        "внутри класса берётся нормальное распределение со средним и стандартным отклонением обучающих "
        "строк этого класса. Если бы отклонение равнялось нулю, в плотность подставилось бы 0,0001. "
        "В этой выборке отклонение каждого кода внутри класса больше нуля, подмена не понадобилась.",
    )
    rows = []
    for stat in m["stats"]:
        rows.append(
            [
                stat["class"],
                FEATURE_RU[stat["feature"]],
                num(stat["mean"], 4),
                num(stat["sd"], 4),
                num(stat["sd_used"], 4),
            ]
        )
    docx.add_table_caption(doc, "Таблица 3 — Среднее и стандартное отклонение по обучающим строкам", 16.5)
    docx.add_table(
        doc,
        ["Класс", "Код", "Среднее", "Отклонение", "В расчёте"],
        rows,
        [3.6, 4.3, 2.6, 3.0, 3.0],
        font=11,
    )
    docx.add_text(
        doc,
        "Средние в таблице 3 — это частоты кодов, а не постоянная метка класса. Зона внутри класса "
        "уже не равна одному числу, поэтому чужой класс не получает нулевую вероятность автоматически.",
    )
    example = m["example"]
    probs = example["probs"]
    fx = example["features"]
    twin = next(
        item
        for item in m["predictions"]
        if (not item["correct"]) and item["features"] == example["features"]
    )
    docx.add_text(
        doc,
        f"Строка {example['id']} — коды {fx['net_share']}, {fx['play_zone']}, {fx['hold']}, "
        f"{fx['attack']}, {fx['ground']}. Факт и прогноз — «{example['ball']}». Вероятности: "
        f"футбольный {pct(probs['футбольный'])} %, баскетбольный {pct(probs['баскетбольный'])} %, "
        f"волейбольный {pct(probs['волейбольный'])} %.",
    )
    docx.add_text(
        doc,
        f"Строка {twin['id']} имеет те же пять кодов, но факт — «{twin['ball']}», прогноз — "
        f"«{twin['pred']}». Это передача руками на уровне игрока: такой набор есть и у футбольного, "
        "и у баскетбольного мяча, поэтому метод ставит более частый в обучении класс и ошибается.",
    )
    cm = m["confusion"]
    docx.add_table_caption(doc, "Таблица 4 — Матрица ошибок на 27 тестовых строках", 12)
    docx.add_table(
        doc,
        ["Факт \\ прогноз", "футбольный", "баскетбольный", "волейбольный"],
        [[name, *[str(cm[i][j]) for j in range(3)]] for i, name in enumerate(["футбольный", "баскетбольный", "волейбольный"])],
        [4.0, 2.8, 3.4, 3.2],
        font=11,
    )
    docx.add_text(
        doc,
        f"Верных ответов {m['correct']} из 27, ошибок {27 - m['correct']}. Accuracy, средние precision, "
        f"recall и F-measure равны {pct(m['accuracy'])} %, {pct(m['macro_precision'])} %, "
        f"{pct(m['macro_recall'])} % и {pct(m['macro_f1'])} %.",
    )
    docx.add_picture(doc, FIG / "scatter.png", 14)
    docx.add_figure_caption(doc, "Рисунок 2 — Игровая зона и вектор атаки")
    docx.add_text(
        doc,
        "На рисунке 2 одинаковые коды лежат в одной точке, и в одной точке бывают разные мячи. "
        "Игровая зона классы не разводит.",
    )

    doc.add_page_break()
    docx.add_heading(doc, "4 Выводы")
    docx.add_text(
        doc,
        "Наивный байесовский метод на 90 розыгрышах и том же разбиении 63 и 27, что в лабораторных 2 и 3, "
        f"верно назвал {m['correct']} из 27. Accuracy равна {pct(m['accuracy'])} %. "
        "Ошибки сидят на наборах, которые встречаются у разных мячей: одинаковые коды нельзя "
        "отнести к одному классу без промаха.",
    )
    doc.save(str(OUT))


def assert_clean() -> None:
    import zipfile
    from lxml import etree

    with zipfile.ZipFile(OUT) as zf:
        text = zf.read("word/document.xml").decode("utf-8")
    plain = "".join(etree.fromstring(text.encode("utf-8")).itertext())
    banned = ["тикет", "велосипед", "iris", "sklearn", "matplotlib", "KNIME", "build_report", "compute_", "seed", "numpy"]
    hits = [word for word in banned if word.lower() in plain.lower()]
    if hits:
        raise RuntimeError(f"лишнее в отчёте: {hits}")
    if "@@" in plain:
        raise RuntimeError("маркер содержания не заменён")
    if plain.count("Python") != 1:
        raise RuntimeError(f"Python упомянут {plain.count('Python')} раз")
    if "МЕТОД НАИВНОГО БАЙЕСА" not in plain or "0,0001" not in plain:
        raise RuntimeError("нет темы или порога отклонения")


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
