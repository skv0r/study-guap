# -*- coding: utf-8 -*-
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from docx_common import (
    add_caption, add_code, add_contents_title, add_h, add_p, add_table, add_toc_item, measure_pages, new_doc,
)

BASE = Path(__file__).resolve().parent
DATA = json.loads((BASE / "ЛР1" / "results_lr1.json").read_text(encoding="utf-8"))
OUT = BASE / "ЛР1" / "ЛР1_Отчет_Буренков_МиСП.docx"


def ms(value):
    return f"{value:.3f}".replace(".", ",")


def nsp(value):
    return f"{value:,}".replace(",", " ")


def by_id(items, qid):
    for item in items:
        if item["id"] == qid:
            return item
    raise KeyError(qid)


def third(phase, qid):
    return by_id(phase, qid)["runs"][-1]


def sql_of(qid):
    return by_id(DATA["queries"], qid)["sql"]


def excerpt(qid, text):
    def norm(value):
        return " ".join(value.split())

    if norm(text) not in norm(sql_of(qid)):
        raise SystemExit(f"Отрывок запроса {qid} не найден в JSON")
    return text


def param_run(qid, name):
    for item in DATA["params"]:
        if item["id"] != qid:
            continue
        for variant in item["variants"]:
            if variant["name"] == name:
                return variant["run"]
    raise KeyError((qid, name))


def plan_has(plan, marker):
    if marker not in plan:
        raise SystemExit(f"В плане нет фрагмента: {marker}")


def exec_line(plan, value):
    for line in plan.splitlines():
        stripped = line.strip()
        if stripped.startswith("Execution Time:"):
            number = stripped.split()[2]
            if abs(float(number) - value) > 0.0001:
                raise SystemExit(f"Время плана {number} не совпало с JSON {value}")
            return stripped
    raise SystemExit("В плане нет Execution Time")


LABELS = {
    "fio": "Поиск по ФИО",
    "a": "Запрос (а)",
    "b": "Запрос (б)",
    "v": "Запрос (в)",
    "order": "Сортировка по ФИО",
    "between": "BETWEEN",
    "g": "Среднее число участников",
    "zh": "Запрос (ж)",
}

TOC = [
    "1 Цель работы",
    "2 Задание",
    "3 Ход работы",
    "4 Выводы",
]


def build(pages):
    doc = new_doc("1", "Индексация данных и планы выполнения запросов")
    add_contents_title(doc)
    for title in TOC:
        add_toc_item(doc, title, pages[title])

    add_h(doc, "1 Цель работы", page_break=True)
    add_p(
        doc,
        "Получить навык подбора индексов и сравнения планов выполнения запросов в PostgreSQL "
        "на базе варианта 5: вуз, факультет, группа, студент, конференция, программа конференции, "
        "выступление и участие студента в выступлении.",
    )

    add_h(doc, "2 Задание", page_break=True)
    add_p(
        doc,
        "Наполнить базу тестовыми данными сопоставимого с эксплуатацией объёма. Выполнить запросы "
        "выборки, соединения, агрегации и подзапроса NOT EXISTS, а также поиск и сортировку по ФИО. "
        "Зафиксировать план и время, создать индексы по полям отбора, сортировки и соединения "
        "и повторить замеры. Сравнить план, который строит планировщик, с планами при параметрах "
        "enable_seqscan, enable_indexscan, enable_nestloop и enable_hashjoin.",
    )

    fio_b = third(DATA["before"], "fio")
    fio_a = third(DATA["after"], "fio")
    plan_has(fio_b["plan"], 'Seq Scan on "Студент"')
    plan_has(fio_b["plan"], "Rows Removed by Filter: 100007")
    plan_has(fio_a["plan"], 'Index Scan using "Студент_ФИО"')
    plan_has(param_run("fio", "indexscan_off")["plan"], 'Bitmap Index Scan on "Студент_ФИО"')
    plan_has(param_run("zh", "hashjoin_off")["plan"], 'Index Only Scan using "Выступление_студент_ID_студента"')
    if 'Index Scan using "Группа_Номер_группы"' in third(DATA["after"], "b")["plan"]:
        raise SystemExit("Индекс номера группы оказался в плане, текст отчёта устарел")

    add_h(doc, "3 Ход работы", page_break=True)
    add_p(
        doc,
        f"Замеры выполнены в PostgreSQL {DATA['server_version']}. В таблице 2 приведено "
        "Execution Time третьего прогона. Страницы в планах читались из буферного кэша.",
    )

    add_h(doc, "3.1 Объём данных")
    add_p(
        doc,
        "Схема варианта 5 сохранена. Для замера добавлены группы факультетов ГУАП, студенты этих групп, "
        "доклады и участия. Строки группы 4226 не удалялись. Объём после загрузки приведён в таблице 1. "
        "Размер включает индексы первичного ключа и ограничений UNIQUE.",
    )
    add_caption(doc, "Таблица 1 – Объём данных")
    order = [
        "Вуз", "Факультет", "Группа", "Студент",
        "Конференция", "Программа_конференции", "Выступление", "Выступление_студент",
    ]
    add_table(
        doc,
        ["Таблица", "Строк", "Размер"],
        [
            [name, nsp(DATA["counts"][name]), DATA["sizes"][name]]
            for name in order
        ],
    )

    add_h(doc, "3.2 Запросы")
    add_p(
        doc,
        "Поиск по ФИО возвращает одну строку: Алексеев Иван Сергеевич, группа 4226.",
    )
    add_code(
        doc,
        excerpt(
            "fio",
            'SELECT "Студент"."ФИО", "Группа"."Номер_группы"\n'
            'FROM "Студент"\n'
            'JOIN "Группа"\n'
            '  ON "Группа"."ID_группы" = "Студент"."ID_группы"\n'
            "WHERE \"Студент\".\"ФИО\" = 'Алексеев Иван Сергеевич'",
        ),
    )
    add_p(
        doc,
        "Темы докладов группы 4226: «Модели данных и CASE-средства» и "
        "«Совместный доклад межфакультетской группы».",
    )
    add_code(
        doc,
        excerpt(
            "b",
            'SELECT DISTINCT "Выступление"."Название_темы"\n'
            'FROM "Выступление"\n'
            'JOIN "Выступление_студент"\n'
            '  ON "Выступление_студент"."ID_выступления" = "Выступление"."ID_выступления"\n'
            'JOIN "Студент"\n'
            '  ON "Студент"."ID_студента" = "Выступление_студент"."ID_студента"\n'
            'JOIN "Группа"\n'
            '  ON "Группа"."ID_группы" = "Студент"."ID_группы"\n'
            "WHERE \"Группа\".\"Номер_группы\" = '4226'",
        ),
    )
    add_p(
        doc,
        f"Запрос с соединением отбирает {third(DATA['after'], 'a')['nrows']} различных ФИО "
        "студентов факультета с номером 1, выступавших на конференции «Информатика». "
        "Среди них Алексеев Иван Сергеевич, Борисова Мария Олеговна и Журавлёва Ольга Николаевна. "
        "Дальше запрос соединяет выступление, программу и конференцию. Отрывок:",
    )
    add_code(
        doc,
        excerpt(
            "a",
            'SELECT DISTINCT "Студент"."ФИО"\n'
            'FROM "Студент"\n'
            'JOIN "Группа"\n'
            '  ON "Группа"."ID_группы" = "Студент"."ID_группы"\n'
            'JOIN "Факультет"\n'
            '  ON "Факультет"."ID_факультета" = "Группа"."ID_факультета"\n'
            'JOIN "Выступление_студент"\n'
            '  ON "Выступление_студент"."ID_студента" = "Студент"."ID_студента"',
        ),
    )

    add_h(doc, "3.3 Индексы")
    add_p(
        doc,
        "Первичные ключи и ограничения UNIQUE уже имели индексы. Внешний ключ отдельный индекс не создаёт. "
        "Добавлены пять индексов по полям отбора, сортировки и соединения.",
    )
    add_code(doc, DATA["indexes_sql"].strip())
    add_p(
        doc,
        "«Студент_ФИО» обслуживает равенство и сортировку по ФИО, «Студент_ID_группы» — соединение "
        "студента с группой, «Выступление_ID_программы» — соединение доклада с программой, "
        "«Выступление_студент_ID_студента» — поиск участий студента, «Группа_Номер_группы» — отбор "
        "по номеру группы. Индекс номера группы планировщик не взял. Сортировка по ФИО и запрос (ж) "
        "после индексов читают «Группа» по индексу первичного ключа.",
    )

    add_h(doc, "3.4 Время и план")
    add_p(doc, "Таблица 2 собрана из Execution Time третьего прогона: до дополнительных индексов и после них.")
    add_caption(doc, "Таблица 2 – Время третьего прогона, мс")
    add_table(
        doc,
        ["Запрос", "До индексов", "После индексов"],
        [
            [LABELS[qid], ms(third(DATA["before"], qid)["execution_ms"]), ms(third(DATA["after"], qid)["execution_ms"])]
            for qid in ["fio", "a", "b", "v", "order", "between", "g", "zh"]
        ],
    )
    add_p(
        doc,
        f"Поиск по ФИО до индекса шёл последовательным просмотром «Студент» и отбрасывал 100 007 строк, "
        f"третий прогон занял {ms(fio_b['execution_ms'])} мс. После индекса «Студент_ФИО» условие стало "
        f"Index Cond, время — {ms(fio_a['execution_ms'])} мс. Фрагмент плана:",
    )
    add_code(
        doc,
        "\n".join(
            [
                'Seq Scan on "Студент"',
                "Rows Removed by Filter: 100007",
                exec_line(fio_b["plan"], fio_b["execution_ms"]),
                'Index Scan using "Студент_ФИО" on "Студент"',
                exec_line(fio_a["plan"], fio_a["execution_ms"]),
            ]
        ),
    )

    add_h(doc, "3.5 Параметры планировщика")
    seq = [ms(param_run(qid, "seqscan_off")["execution_ms"]) for qid in ("fio", "a", "v", "zh")]
    idx = [ms(param_run(qid, "indexscan_off")["execution_ms"]) for qid in ("fio", "a", "v", "zh")]
    nest = [ms(param_run(qid, "nestloop_off")["execution_ms"]) for qid in ("fio", "a", "v", "zh")]
    hsh = [ms(param_run(qid, "hashjoin_off")["execution_ms"]) for qid in ("fio", "a", "v", "zh")]
    add_p(
        doc,
        "После индексов поиск по ФИО и запросы (а), (в) и (ж) повторены с одним изменённым параметром. "
        f"При enable_seqscan = off время {seq[0]}, {seq[1]}, {seq[2]} и {seq[3]} мс: поиск по ФИО остаётся "
        "Index Scan по «Студент_ФИО». При enable_indexscan = off время "
        f"{idx[0]}, {idx[1]}, {idx[2]} и {idx[3]} мс; таблица «Студент» в поиске по ФИО читается через "
        "Bitmap Index Scan того же индекса, последовательный просмотр к ней не возвращается. "
        f"При enable_nestloop = off время {nest[0]}, {nest[1]}, {nest[2]} и {nest[3]} мс. "
        f"При enable_hashjoin = off время {hsh[0]}, {hsh[1]}, {hsh[2]} и {hsh[3]} мс: запрос (в) становится "
        "дольше, а запрос (ж) переходит на Index Only Scan по индексу «Выступление_студент_ID_студента» "
        "и ускоряется. Порядок чисел в каждой четвёрке: поиск по ФИО, запрос (а), запрос (в), запрос (ж).",
    )

    add_h(doc, "4 Выводы", page_break=True)
    order_b = third(DATA["before"], "order")
    order_a = third(DATA["after"], "order")
    b_b = third(DATA["before"], "b")
    b_a = third(DATA["after"], "b")
    add_p(
        doc,
        f"На базе варианта 5 ({nsp(DATA['counts']['Студент'])} студентов, "
        f"{nsp(DATA['counts']['Выступление'])} выступлений, {nsp(DATA['counts']['Выступление_студент'])} участий) "
        "пять дополнительных индексов изменили точечный поиск по ФИО: последовательный просмотр "
        f"{ms(fio_b['execution_ms'])} мс сменился индексным {ms(fio_a['execution_ms'])} мс. "
        f"Сортировка по ФИО ускорилась с {ms(order_b['execution_ms'])} мс до {ms(order_a['execution_ms'])} мс, "
        f"темы группы 4226 — с {ms(b_b['execution_ms'])} мс до {ms(b_a['execution_ms'])} мс. "
        "Запросы, которые читают все участия, ускорились слабо. Параметры enable_seqscan, "
        "enable_indexscan, enable_nestloop и enable_hashjoin меняют способ доступа и соединение. "
        "Выключение хеш-соединения замедляет запрос (в) и ускоряет запрос (ж).",
    )
    doc.save(OUT)


def main():
    build({title: "…" for title in TOC})
    pages = measure_pages(OUT, TOC)
    build(pages)
    print(pages)
    print(OUT)


if __name__ == "__main__":
    main()
