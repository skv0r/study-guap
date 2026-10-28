# -*- coding: utf-8 -*-
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from docx_common import (
    add_caption, add_code, add_contents_title, add_figure, add_h, add_p, add_table, add_toc_item,
    measure_pages, new_doc,
)

BASE = Path(__file__).resolve().parent
FIG = BASE / "ЛР3" / "figures"
OUT = BASE / "ЛР3" / "ЛР3_Отчет_Буренков_МиСП.docx"
TOC = ["1 Цель работы", "2 Задание", "3 Ход работы", "4 Выводы"]


def build(pages):
    doc = new_doc("3", "Транзакционная репликация")
    add_contents_title(doc)
    for title in TOC:
        add_toc_item(doc, title, pages[title])

    add_h(doc, "1 Цель работы", page_break=True)
    add_p(
        doc,
        "Получить навык переноса строки студента между издателем и узлом в PostgreSQL "
        "на базе варианта 5 и чтения ответа сервера: и когда оператор выполнен, и когда сервер его отклонил.",
    )

    add_h(doc, "2 Задание", page_break=True)
    add_p(
        doc,
        "Собрать объекты репликации, выполнить вставку, изменение и удаление, показать содержимое таблиц "
        "на издателе и на узле и изменить схему издателя. Запросы выполняются отдельными блоками: "
        "следующий начинается, когда предыдущий уже выполнен.",
    )

    add_h(doc, "3 Ход работы", page_break=True)
    add_p(
        doc,
        "Среда работы — PostgreSQL. База одна: издатель лежит в схеме pub, узел — в схеме sub. "
        "Триггер repl после вставки, изменения и удаления копирует строку таблицы «Студент». "
        "Публикация pub_vuz объявляет эту таблицу объектом репликации. "
        "Подписка на ту же базу сервером не завершается, поэтому в отчёте те ответы, которые сервер печатает сразу.",
    )

    add_h(doc, "3.1 Объекты")
    add_p(doc, "Ключ решения — триггер на каждую строку и публикация таблицы:")
    add_code(doc, 'CREATE TRIGGER repl')
    add_code(doc, 'AFTER INSERT OR UPDATE OR DELETE ON pub."Студент"')
    add_code(doc, "FOR EACH ROW EXECUTE FUNCTION pub.copy_student();")
    add_code(doc, "")
    add_code(doc, 'CREATE PUBLICATION pub_vuz FOR TABLE pub."Студент";')
    add_p(
        doc,
        "Функция при вставке пишет новую строку в sub.\"Студент\", при изменении обновляет ту же строку по ID, "
        "при удалении убирает её с узла вместе с издателем.",
    )

    add_h(doc, "3.2 Перенос строки")
    add_p(
        doc,
        "Сначала записан Иванов Иван, факультет 4. Затем имя изменено на Иванов Пётр. "
        "На рисунке 1 обе стороны возвращают одну и ту же строку: ID 1, Иванов Пётр, факультет 4. "
        "Эти значения сведены в таблице 1.",
    )
    add_figure(doc, FIG / "both.png", "Рисунок 1 — строка после изменения на издателе и на узле")
    add_caption(doc, "Таблица 1 — Строка после UPDATE, до DELETE")
    add_table(
        doc,
        ["Сторона", "ID", "ФИО", "Факультет"],
        [["издатель pub", "1", "Иванов Пётр", "4"], ["узел sub", "1", "Иванов Пётр", "4"]],
        [4, 2, 6, 4.5],
    )
    add_p(doc, "DELETE по ID 1 убирает строку с обеих сторон. На рисунке 2 счётчик узла равен 0.")
    add_figure(doc, FIG / "zero.png", "Рисунок 2 — узел после удаления строки")

    add_h(doc, "3.3 Что сервер отклонил")
    add_p(
        doc,
        "Первая вставка студента 1 снова проходит. Вторая, с тем же ID и именем Петров Пётр, не проходит. "
        "Строка Петрова не появляется ни у издателя, ни на узле. Ответ сервера на рисунке 3:",
    )
    add_code(doc, 'ERROR:  duplicate key value violates unique constraint "Студент_pkey"')
    add_code(doc, 'ПОДРОБНОСТИ:  Key ("ID")=(1) already exists.')
    add_figure(doc, FIG / "dup.png", "Рисунок 3 — отказ повторной вставки с тем же ID")
    add_p(doc, "Повторное создание публикации pub_vuz сервер тоже отклоняет. Ответ на рисунке 4:")
    add_code(doc, 'ERROR:  publication "pub_vuz" already exists')
    add_figure(doc, FIG / "pub.png", "Рисунок 4 — отказ повторного создания публикации")

    add_h(doc, "3.4 Изменение схемы")
    add_p(
        doc,
        "После вставки, которая прошла, у издателя добавляется столбец «Группа» и в него пишется 4321. "
        "На рисунке 5 у издателя четыре столбца, студент 1 — Иванов Иван. "
        "Триггер копирует только ID, ФИО и Факультет, поэтому на узле столбец не появляется. "
        "На рисунке 6 список столбцов узла состоит из трёх позиций.",
    )
    add_figure(doc, FIG / "group.png", "Рисунок 5 — столбец «Группа» у издателя")
    add_figure(doc, FIG / "cols.png", "Рисунок 6 — столбцы узла после изменения схемы")

    add_h(doc, "4 Выводы", page_break=True)
    add_p(
        doc,
        "Объекты репликации собраны в одной базе PostgreSQL: триггер repl и публикация pub_vuz. "
        "После UPDATE обе стороны вернули «Иванов Пётр», факультет 4. После DELETE счётчик узла стал равен 0. "
        "Повтор ID=1 сервер отклонил сообщением duplicate key, повтор публикации — сообщением publication \"pub_vuz\" already exists. "
        "Столбец «Группа» со значением 4321 есть только у издателя. На узле остались ID, ФИО и Факультет.",
    )
    doc.save(OUT)


def main():
    build({title: "…" for title in TOC})
    pages = measure_pages(OUT, TOC)
    build(pages)
    print(pages)


if __name__ == "__main__":
    main()
