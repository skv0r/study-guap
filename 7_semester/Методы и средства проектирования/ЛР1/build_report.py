#!/usr/bin/env python3
"""Сборка отчёта ЛР1 МиСП: планы до/после индексов и при параметрах планировщика."""

from __future__ import annotations

import shutil
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

BASE = Path(__file__).resolve().parent
SRC = BASE / "ЛР1_Отчет_Буренков_МиСП.docx"
OUT = BASE / "ЛР1_Отчет_Буренков_МиСП.docx"


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
    mono=False,
):
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.line_spacing = 1.5
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.first_line_indent = Cm(1.25 if first_indent and not mono else 0)
    p.alignment = {
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "left": WD_ALIGN_PARAGRAPH.LEFT,
    }[align]


def add_body(doc, text, *, first_indent=True, mono=False, space_before=0):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=first_indent, mono=mono, space_before=space_before, align="left" if mono else "justify")
    set_run_font(p.add_run(text), size=12 if mono else 14)
    return p


def add_heading(doc, text, *, level=1):
    p = doc.add_paragraph()
    format_paragraph(
        p,
        first_indent=level > 0,
        align="left",
        space_before=6 if level > 1 else 0,
        space_after=6,
    )
    set_run_font(p.add_run(text), bold=True)
    return p


def add_toc_line(doc, title, page):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="left")
    from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER

    p.paragraph_format.tab_stops.add_tab_stop(Cm(16.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    set_run_font(p.add_run(title))
    set_run_font(p.add_run("\t"))
    set_run_font(p.add_run(str(page)))


def add_caption(doc, text):
    p = doc.add_paragraph()
    format_paragraph(p, first_indent=False, align="center", space_before=6, space_after=3)
    set_run_font(p.add_run(text))


def add_table(doc, headers, rows, widths):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        format_paragraph(p, first_indent=False, align="center")
        set_run_font(p.add_run(h), bold=True)
        cell.width = Cm(widths[i])
    for r_i, row in enumerate(rows):
        for c_i, val in enumerate(row):
            cell = table.rows[r_i + 1].cells[c_i]
            cell.text = ""
            p = cell.paragraphs[0]
            format_paragraph(p, first_indent=False, align="left")
            set_run_font(p.add_run(str(val)))
            cell.width = Cm(widths[c_i])
    doc.add_paragraph()


def strip_after_title(doc: Document) -> None:
    body = doc.element.body
    children = list(body)
    # оставить титул до «Санкт-Петербург 2026» включительно (index 33), убрать с СОДЕРЖАНИЕ
    keep_until = None
    for i, child in enumerate(children):
        texts = [t.text for t in child.iter(qn("w:t")) if t.text]
        if "".join(texts).startswith("Санкт-Петербург"):
            keep_until = i
            break
    if keep_until is None:
        raise SystemExit("Не найден титул")
    for child in children[keep_until + 1 :]:
        if child.tag == qn("w:sectPr"):
            continue
        body.remove(child)


def build_body(doc: Document) -> None:
    add_heading(doc, "СОДЕРЖАНИЕ", level=0)
    # level 0 centered — fix
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.paragraphs[-1].paragraph_format.first_line_indent = Cm(0)
    for title, page in [
        ("1 Цель работы", 2),
        ("2 Задание", 3),
        ("3 Ход работы", 4),
        ("4 Выводы", 8),
    ]:
        add_toc_line(doc, title, page)

    add_heading(doc, "1 Цель работы")
    add_body(
        doc,
        "Получить навык подбора индексов и сравнения планов выполнения запросов в PostgreSQL на базе варианта 5: "
        "вуз, факультет, группа, студент, конференция, программа конференции, выступление и участие студента в выступлении.",
    )

    add_heading(doc, "2 Задание")
    add_body(
        doc,
        "Наполнить базу тестовыми данными сопоставимого с эксплуатацией объёма. Выполнить запросы выборки, соединения, "
        "агрегации и подзапроса NOT EXISTS, а также поиск и сортировку по ФИО. Зафиксировать план и время, создать индексы "
        "по полям отбора, сортировки и соединения и повторить замеры. Сопоставить планы планировщика в трёх случаях: "
        "до дополнительных индексов, после их создания и при изменённых параметрах enable_seqscan, enable_indexscan, "
        "enable_nestloop и enable_hashjoin.",
    )

    add_heading(doc, "3 Ход работы")
    add_body(
        doc,
        "Замеры выполнены в PostgreSQL 17.11. В таблицах ниже — Execution Time третьего прогона. "
        "Страницы в планах читались из буферного кэша.",
    )

    add_heading(doc, "3.1 Объём данных", level=2)
    add_body(
        doc,
        "Схема варианта 5 сохранена. Для замера добавлены группы факультетов ГУАП, студенты этих групп, доклады и участия. "
        "Строки группы 4226 не удалялись. Объём после загрузки приведён в таблице 1.",
    )
    add_caption(doc, "Таблица 1 – Объём данных")
    add_table(
        doc,
        ["Таблица", "Строк", "Размер"],
        [
            ["Вуз", "1", "40 kB"],
            ["Факультет", "4", "40 kB"],
            ["Группа", "206", "72 kB"],
            ["Студент", "100 008", "8904 kB"],
            ["Конференция", "3", "40 kB"],
            ["Программа_конференции", "3", "40 kB"],
            ["Выступление", "12 008", "1400 kB"],
            ["Выступление_студент", "84 772", "4896 kB"],
        ],
        [7.0, 4.0, 4.5],
    )

    add_heading(doc, "3.2 Запросы", level=2)
    add_body(
        doc,
        "Поиск по ФИО возвращает одну строку: Алексеев Иван Сергеевич, группа 4226. Темы докладов группы 4226: "
        "«Модели данных и CASE-средства» и «Совместный доклад межфакультетской группы». Запрос (а) отбирает 123 ФИО "
        "студентов факультета 1 на конференции «Информатика».",
    )
    add_body(
        doc,
        'SELECT "Студент"."ФИО", "Группа"."Номер_группы"\n'
        'FROM "Студент"\n'
        'JOIN "Группа" ON "Группа"."ID_группы" = "Студент"."ID_группы"\n'
        "WHERE \"Студент\".\"ФИО\" = 'Алексеев Иван Сергеевич'",
        first_indent=False,
        mono=True,
    )

    add_heading(doc, "3.3 Индексы", level=2)
    add_body(
        doc,
        "Первичные ключи и UNIQUE уже имели индексы. Добавлены пять индексов по полям отбора, сортировки и соединения.",
    )
    for line in [
        'CREATE INDEX "Студент_ФИО" ON "Студент" ("ФИО");',
        'CREATE INDEX "Студент_ID_группы" ON "Студент" ("ID_группы");',
        'CREATE INDEX "Выступление_ID_программы" ON "Выступление" ("ID_программы");',
        'CREATE INDEX "Выступление_студент_ID_студента" ON "Выступление_студент" ("ID_студента");',
        'CREATE INDEX "Группа_Номер_группы" ON "Группа" ("Номер_группы");',
    ]:
        add_body(doc, line, first_indent=False, mono=True)

    add_heading(doc, "3.4 Время до и после индексов", level=2)
    add_body(
        doc,
        "Таблица 2 сравнивает время третьего прогона до дополнительных индексов и после них.",
    )
    add_caption(doc, "Таблица 2 – Время третьего прогона, мс")
    add_table(
        doc,
        ["Запрос", "До индексов", "После индексов"],
        [
            ["Поиск по ФИО", "6,189", "0,056"],
            ["Запрос (а)", "26,583", "23,993"],
            ["Запрос (б)", "28,432", "7,101"],
            ["Запрос (в)", "123,218", "112,230"],
            ["Сортировка по ФИО", "131,346", "33,835"],
            ["BETWEEN", "0,018", "0,016"],
            ["Среднее число участников", "15,789", "14,321"],
            ["Запрос (ж)", "169,029", "160,780"],
        ],
        [8.0, 4.0, 4.5],
    )

    add_heading(doc, "3.5 Сопоставление планов: до и после индексов", level=2)
    add_body(
        doc,
        "Планировщик меняет не только время, но и форму плана. В таблице 3 для характерных запросов сопоставлены "
        "основные узлы плана до и после индексов.",
    )
    add_caption(doc, "Таблица 3 – План до и после индексов")
    add_table(
        doc,
        ["Запрос", "До индексов", "После индексов"],
        [
            [
                "Поиск по ФИО",
                "Nested Loop + Seq Scan «Студент» (отброшено 100 007 строк)",
                "Hash Join + Index Scan «Студент_ФИО»",
            ],
            [
                "Сортировка по ФИО",
                "Sort (external merge, Disk 4608 kB) + Seq Scan «Студент»",
                "Nested Loop: Index Scan «Студент_ФИО» + Index Scan «Группа_pkey» без сортировки на диске",
            ],
            [
                "Запрос (б)",
                "Unique + Sort; Seq Scan «Выступление_студент» (~84 тыс. строк)",
                "Тот же каркас, но отбор группы через индекс; время 28,4 → 7,1 мс",
            ],
            [
                "Запрос (а)",
                "HashAggregate + Hash Join; Seq Scan участий",
                "Тот же каркас; Seq Scan участий остаётся, выигрыш мал (26,6 → 24,0 мс)",
            ],
            [
                "Запрос (в)",
                "HashAggregate + Parallel Hash Join / Index Only Scan PK участий",
                "План почти тот же; время 123 → 112 мс",
            ],
            [
                "Запрос (ж)",
                "Unique + Sort + Hash Right Anti Join; Seq Scan участий и студентов",
                "Anti Join сохранён; часть соединений через Nested Loop; 169 → 161 мс",
            ],
        ],
        [3.2, 6.5, 6.8],
    )
    add_body(
        doc,
        "Фрагмент плана поиска по ФИО до индекса:",
        space_before=6,
    )
    for line in [
        'Seq Scan on "Студент"',
        "Rows Removed by Filter: 100007",
        "Execution Time: 6.189 ms",
    ]:
        add_body(doc, line, first_indent=False, mono=True)
    add_body(doc, "После индекса:")
    for line in [
        'Index Scan using "Студент_ФИО" on "Студент"',
        "Execution Time: 0.056 ms",
    ]:
        add_body(doc, line, first_indent=False, mono=True)
    add_body(
        doc,
        "Вывод по этому случаю: индексы резко меняют план там, где есть равенство или упорядочивание по индексируемому "
        "полю (ФИО, номер группы). Запросы, которые вынуждены читать почти все участия, планировщик оставляет на "
        "последовательном просмотре или хеш-соединении — индекс даёт малый выигрыш.",
    )

    add_heading(doc, "3.6 Сопоставление планов: параметры планировщика", level=2)
    add_body(
        doc,
        "После создания индексов те же запросы (поиск по ФИО, (а), (в), (ж)) выполнены ещё в четырёх случаях: "
        "по очереди выключались enable_seqscan, enable_indexscan, enable_nestloop и enable_hashjoin. "
        "Время — в таблице 4, изменение узлов плана — в таблице 5.",
    )
    add_caption(doc, "Таблица 4 – Время при параметрах планировщика, мс")
    add_table(
        doc,
        ["Случай", "ФИО", "(а)", "(в)", "(ж)"],
        [
            ["После индексов (база)", "0,056", "23,993", "112,230", "160,780"],
            ["enable_seqscan = off", "0,031", "25,817", "118,328", "166,868"],
            ["enable_indexscan = off", "0,055", "23,774", "181,445", "160,876"],
            ["enable_nestloop = off", "0,062", "21,923", "133,172", "170,387"],
            ["enable_hashjoin = off", "0,024", "31,553", "234,273", "43,981"],
        ],
        [5.0, 2.5, 2.8, 3.0, 3.2],
    )
    add_caption(doc, "Таблица 5 – Как меняется план при параметрах")
    add_table(
        doc,
        ["Случай", "Что меняется в плане"],
        [
            [
                "enable_seqscan = off",
                "Поиск по ФИО: Nested Loop из Index Scan «Студент_ФИО» и Index Scan «Группа_pkey». "
                "Тяжёлые запросы почти не ускоряются: планировщик и так мало опирался на Seq Scan ключевых таблиц.",
            ],
            [
                "enable_indexscan = off",
                "Поиск по ФИО: вместо Index Scan — Bitmap Index Scan + Bitmap Heap Scan по тому же индексу "
                "(последовательный просмотр к «Студент» не возвращается). Запрос (в) заметно медленнее: "
                "без Index Only Scan по PK участий план раздувается хеш-соединениями (181 мс).",
            ],
            [
                "enable_nestloop = off",
                "Запрос (а): Nested Loop заменяется Merge Join / Hash Join; время чуть лучше базы (21,9 мс). "
                "Поиск по ФИО остаётся на индексе через Hash Join.",
            ],
            [
                "enable_hashjoin = off",
                "Запрос (в): Hash Join → Merge Join + Sort, время хуже (234 мс). "
                "Запрос (ж): Anti Join без хеша переходит на схему с Index Only Scan "
                "«Выступление_студент_ID_студента» и ускоряется до 44 мс.",
            ],
        ],
        [4.5, 12.0],
    )
    add_body(
        doc,
        "Сопоставление случаев. Базовый план после индексов — отправная точка. Выключение seqscan почти не меняет "
        "тяжёлые запросы, но подтверждает, что точечный поиск уже идёт по индексу. Выключение indexscan заставляет "
        "искать обходные пути: для ФИО это bitmap по тому же индексу, для (в) — более дорогие хеш-соединения. "
        "Выключение nestloop и hashjoin показывает альтернативы соединений: merge вместо nested loop иногда полезен, "
        "а отказ от hashjoin для (в) вреден, для (ж) — неожиданно выгоден на этих данных. "
        "Итог: один и тот же запрос планировщик собирает по-разному в зависимости от доступных методов доступа и соединений; "
        "сравнивать нужно и время, и узлы плана, а не только одну цифру Execution Time.",
    )

    add_heading(doc, "4 Выводы")
    add_body(
        doc,
        "На базе варианта 5 (100 008 студентов, 12 008 выступлений, 84 772 участий) сопоставлены планы в трёх группах случаев: "
        "до индексов, после индексов и при четырёх параметрах планировщика. "
        "Индексы меняют план точечного поиска по ФИО (Seq Scan 6,189 мс → Index Scan 0,056 мс) и убирают сортировку "
        "на диске при ORDER BY ФИО (131 → 34 мс). Запросы, читающие почти все участия, ускоряются слабо. "
        "Параметры enable_seqscan / enable_indexscan / enable_nestloop / enable_hashjoin меняют способ доступа и соединения: "
        "отказ от hashjoin замедляет запрос (в) и ускоряет запрос (ж). Для отчёта достаточно видеть, что планировщик "
        "выбирает разный план под разные ограничения, а индекс полезен не на каждом запросе.",
    )


def main() -> None:
    if not SRC.exists():
        raise SystemExit(f"нет файла {SRC}")
    # работаем с копией в памяти поверх существующего
    doc = Document(str(SRC))
    strip_after_title(doc)
    build_body(doc)
    tmp = OUT.with_suffix(".tmp.docx")
    doc.save(str(tmp))
    tmp.replace(OUT)
    print(OUT)


if __name__ == "__main__":
    main()
