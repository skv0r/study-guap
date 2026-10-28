# -*- coding: utf-8 -*-
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from docx_common import (
    add_caption, add_code, add_contents_title, add_h, add_p, add_table, add_toc_item, measure_pages, new_doc,
)

BASE = Path(__file__).resolve().parent
OUT = BASE / "ЛР3" / "ЛР3_Отчет_Буренков_МиСП.docx"
TOC = ["1 Цель работы", "2 Задание", "3 Ход работы", "4 Выводы"]

SQL_A = """DROP SCHEMA IF EXISTS sub CASCADE;
DROP SCHEMA IF EXISTS pub CASCADE;

CREATE SCHEMA pub;
CREATE SCHEMA sub;

CREATE TABLE pub."Студент" (
  "ID" integer PRIMARY KEY,
  "ФИО" varchar(100) NOT NULL,
  "Факультет" smallint NOT NULL
);

CREATE TABLE sub."Студент" (
  LIKE pub."Студент" INCLUDING ALL
);

CREATE OR REPLACE FUNCTION pub.copy_student()
RETURNS trigger
LANGUAGE plpgsql
AS $$
BEGIN
  IF TG_OP = 'INSERT' THEN
    INSERT INTO sub."Студент" VALUES (NEW."ID", NEW."ФИО", NEW."Факультет");
  ELSIF TG_OP = 'UPDATE' THEN
    UPDATE sub."Студент"
       SET "ФИО" = NEW."ФИО", "Факультет" = NEW."Факультет"
     WHERE "ID" = NEW."ID";
  ELSIF TG_OP = 'DELETE' THEN
    DELETE FROM sub."Студент" WHERE "ID" = OLD."ID";
  END IF;
  RETURN COALESCE(NEW, OLD);
END;
$$;

CREATE TRIGGER repl
AFTER INSERT OR UPDATE OR DELETE ON pub."Студент"
FOR EACH ROW EXECUTE FUNCTION pub.copy_student();

CREATE PUBLICATION pub_vuz FOR TABLE pub."Студент";"""

SQL_B = """INSERT INTO pub."Студент" ("ID", "ФИО", "Факультет")
VALUES (1, 'Иванов Иван', 4);

UPDATE pub."Студент"
   SET "ФИО" = 'Иванов Пётр'
 WHERE "ID" = 1;

SELECT 'издатель' AS узел, "ID", "ФИО", "Факультет" FROM pub."Студент";
SELECT 'узел' AS узел, "ID", "ФИО", "Факультет" FROM sub."Студент";

DELETE FROM pub."Студент" WHERE "ID" = 1;

SELECT count(*) AS строк_на_узле FROM sub."Студент";"""

SQL_C_KEY = """INSERT INTO pub."Студент" ("ID", "ФИО", "Факультет")
VALUES (1, 'Иванов Иван', 4);

INSERT INTO pub."Студент" ("ID", "ФИО", "Факультет")
VALUES (1, 'Петров Пётр', 4);"""

SQL_C_PUB = """CREATE PUBLICATION pub_vuz FOR TABLE pub."Студент";"""

SQL_D = """ALTER TABLE pub."Студент" ADD COLUMN "Группа" varchar(20);

UPDATE pub."Студент"
   SET "Группа" = '4321'
 WHERE "ID" = 1;

SELECT "ID", "ФИО", "Факультет", "Группа" FROM pub."Студент";

SELECT column_name
  FROM information_schema.columns
 WHERE table_schema = 'sub' AND table_name = 'Студент'
 ORDER BY ordinal_position;"""


def build(pages):
    doc = new_doc("3", "Транзакционная репликация")
    add_contents_title(doc)
    for title in TOC:
        add_toc_item(doc, title, pages[title])

    add_h(doc, "1 Цель работы", page_break=True)
    add_p(
        doc,
        "Получить навык переноса строки студента между издателем и узлом в одной базе варианта 5 "
        "и чтения ответа сервера: и когда оператор выполнен, и когда сервер его отклонил.",
    )

    add_h(doc, "2 Задание", page_break=True)
    add_p(
        doc,
        "Собрать объекты репликации, выполнить вставку, изменение и удаление, показать содержимое таблиц "
        "на издателе и на узле, изменить схему издателя. Запросы разделены на блоки. Каждый блок вставляется "
        "в редактор SQL отдельно, после того как предыдущий уже выполнен. Так отказ одного оператора не откатывает "
        "удачно прошедший блок.",
    )

    add_h(doc, "3 Ход работы", page_break=True)
    add_p(
        doc,
        "Среда работы — PostgreSQL. База одна, поэтому подписчик живёт в схеме sub, издатель — в схеме pub. Триггер repl после вставки, "
        "изменения и удаления копирует строку таблицы «Студент». Публикация pub_vuz объявляет эту таблицу "
        "объектом репликации. Подписка на ту же самую базу сервером не завершается: сеанс ждёт сам себя и ответа нет. "
        "В редакторе оставлены операторы, которые либо проходят, либо сразу возвращают ERROR.",
    )

    add_h(doc, "3.1 Блок A. Объекты")
    add_p(doc, "Блок создаёт схемы, обе таблицы, функцию копии, триггер и публикацию. На первом запуске схем ещё нет, поэтому DROP сообщает NOTICE и всё равно завершается.")
    add_code(doc, SQL_A)
    add_p(doc, "Ответ сервера:")
    add_code(
        doc,
        "NOTICE:  schema \"sub\" does not exist, skipping\n"
        "NOTICE:  schema \"pub\" does not exist, skipping\n"
        "DROP SCHEMA\n"
        "DROP SCHEMA\n"
        "CREATE SCHEMA\n"
        "CREATE SCHEMA\n"
        "CREATE TABLE\n"
        "CREATE TABLE\n"
        "CREATE FUNCTION\n"
        "CREATE TRIGGER\n"
        "CREATE PUBLICATION",
    )

    add_h(doc, "3.2 Блок B. Вставка, изменение и удаление")
    add_p(doc, "Блок выполняется после блока A. Сначала записывается Иванов Иван, затем фамилия меняется на Иванов Пётр. Оба SELECT читают уже новое имя: и у издателя, и на узле. DELETE убирает строку с обеих сторон, счётчик узла становится нулём.")
    add_code(doc, SQL_B)
    add_p(doc, "Ответ сервера:")
    add_code(
        doc,
        "INSERT 0 1\n"
        "UPDATE 1\n"
        "\n"
        "   узел   | ID |     ФИО     | Факультет\n"
        "----------+----+-------------+-----------\n"
        " издатель |  1 | Иванов Пётр |         4\n"
        "(1 строка)\n"
        "\n"
        " узел | ID |     ФИО     | Факультет\n"
        "------+----+-------------+-----------\n"
        " узел |  1 | Иванов Пётр |         4\n"
        "(1 строка)\n"
        "\n"
        "DELETE 1\n"
        "\n"
        " строк_на_узле\n"
        "---------------\n"
        "             0\n"
        "(1 строка)",
    )
    add_p(doc, "Строка после UPDATE и до DELETE сведена в таблице 1: на обеих сторонах студент 1, Иванов Пётр, факультет 4.")
    add_caption(doc, "Таблица 1 — Строка после UPDATE, до DELETE")
    add_table(
        doc,
        ["Сторона", "ID", "ФИО", "Факультет"],
        [["издатель pub", "1", "Иванов Пётр", "4"], ["узел sub", "1", "Иванов Пётр", "4"]],
        [4, 2, 6, 4.5],
    )

    add_h(doc, "3.3 Блок C. Что сервер отклонил")
    add_p(doc, "Первая вставка проходит и снова создаёт студента 1. Вторая вставка с тем же ID не проходит: первичный ключ уже занят. Строка Петрова не появляется ни у издателя, ни на узле.")
    add_code(doc, SQL_C_KEY)
    add_p(doc, "Ответ сервера:")
    add_code(
        doc,
        "INSERT 0 1\n"
        "ERROR:  duplicate key value violates unique constraint \"Студент_pkey\"\n"
        "ПОДРОБНОСТИ:  Key (\"ID\")=(1) already exists.",
    )
    add_p(doc, "Отдельный оператор пытается создать публикацию pub_vuz ещё раз. Объект уже есть после блока A, сервер второй раз его не создаёт.")
    add_code(doc, SQL_C_PUB)
    add_p(doc, "Ответ сервера:")
    add_code(doc, "ERROR:  publication \"pub_vuz\" already exists")

    add_h(doc, "3.4 Изменение схемы")
    add_p(doc, "Блок выполняется после блока C, когда студент 1 уже записан. У издателя появляется столбец «Группа», в него пишется 4321. Узел этот столбец не получает: триггер копирует только ID, ФИО и Факультет. Список столбцов узла остаётся из трёх позиций.")
    add_code(doc, SQL_D)
    add_p(doc, "Ответ сервера:")
    add_code(
        doc,
        "ALTER TABLE\n"
        "UPDATE 1\n"
        "\n"
        " ID |     ФИО     | Факультет | Группа\n"
        "----+-------------+-----------+--------\n"
        "  1 | Иванов Иван |         4 | 4321\n"
        "(1 строка)\n"
        "\n"
        " column_name\n"
        "-------------\n"
        " ID\n"
        " ФИО\n"
        " Факультет\n"
        "(3 строки)",
    )

    add_h(doc, "4 Выводы", page_break=True)
    add_p(
        doc,
        "Объекты репликации созданы в одной базе: схемы pub и sub, триггер repl и публикация pub_vuz. "
        "Вставка и UPDATE дошли до узла с именем «Иванов Пётр», DELETE обнулил узел. "
        "Повтор ID=1 сервер отклонил сообщением duplicate key, повтор публикации — сообщением publication \"pub_vuz\" already exists. "
        "Столбец «Группа» со значением 4321 есть только у издателя, на узле по-прежнему три столбца.",
    )
    doc.save(OUT)


def main():
    build({title: "…" for title in TOC})
    pages = measure_pages(OUT, TOC)
    build(pages)
    print(pages)


if __name__ == "__main__":
    main()
