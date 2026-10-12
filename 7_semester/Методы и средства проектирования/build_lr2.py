# -*- coding: utf-8 -*-
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from docx_common import (
    add_caption, add_code, add_contents_title, add_h, add_p, add_table, add_toc_item, measure_pages, new_doc,
)

BASE = Path(__file__).resolve().parent
DATA = json.loads((BASE / "ЛР2" / "results_lr2.json").read_text(encoding="utf-8"))
OUT = BASE / "ЛР2" / "ЛР2_Отчет_Буренков_МиСП.docx"
IDS = DATA["ids"]


def yn(value):
    return "да" if value else "нет"


TOC = [
    "1 Цель работы",
    "2 Задание",
    "3 Ход работы",
    "4 Выводы",
]


def deadlock_quote(text):
    kept = []
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("HINT") or stripped.startswith("CONTEXT"):
            continue
        kept.append(stripped)
    quote = "\n".join(kept)
    if "обнаружена взаимоблокировка" not in quote:
        raise SystemExit("В сообщении сервера нет фразы о взаимоблокировке")
    return quote


def lock_rows():
    picked = []
    seen = set()
    for row in DATA["locks"]["rows"]:
        pid, app, locktype, mode, granted, relation, _page, _tuple, wait_type, wait_event, _query = row
        rel = "" if relation is None else str(relation).replace('"', "")
        key = (app, locktype, rel, mode, granted)
        take = False
        if app == "lock_holder" and locktype == "relation" and rel == "Студент" and mode == "RowExclusiveLock":
            take = True
        elif app == "lock_holder" and locktype == "transactionid" and mode == "ExclusiveLock":
            take = True
        elif app == "lock_waiter" and locktype == "relation" and rel == "Студент" and mode == "RowShareLock":
            take = True
        elif app == "lock_waiter" and locktype == "tuple" and mode == "RowShareLock":
            take = True
        elif app == "lock_waiter" and locktype == "transactionid" and mode == "ShareLock" and not granted:
            take = True
        if take and key not in seen:
            seen.add(key)
            role = "держатель" if app == "lock_holder" else "ожидающий"
            wait = "" if not wait_type else f"{wait_type}/\n{wait_event}"
            if mode.endswith("Lock") and len(mode) > len("ShareLock"):
                mode = mode[:-4] + "\nLock"
            picked.append([role, locktype, mode, yn(granted), rel, wait])
    if len(picked) < 4:
        raise SystemExit(f"Мало строк блокировок: {len(picked)}")
    return picked


def build(pages):
    doc = new_doc("2", "Транзакции, уровни изоляции и блокировки")
    add_contents_title(doc)
    for title in TOC:
        add_toc_item(doc, title, pages[title])

    add_h(doc, "1 Цель работы", page_break=True)
    add_p(
        doc,
        "Получить навык наблюдения аномалий параллельного доступа к базе варианта 5, "
        "их появления и исключения уровнями изоляции, чтения отчёта о блокировках "
        "и сообщения сервера о взаимоблокировке.",
    )

    add_h(doc, "2 Задание", page_break=True)
    add_p(
        doc,
        "Двумя сеансами смоделировать грязное чтение, неповторяемое чтение и фантомные строки. "
        "Повторить опыты на уровнях read uncommitted, read committed, repeatable read и serializable. "
        "Пока одна транзакция держит изменённую строку, снять блокировки и пояснить режимы. "
        "Двумя транзакциями получить взаимоблокировку и привести текст, который вернул сервер.",
    )

    dirty = DATA["dirty"]
    rc = DATA["nonrepeatable_read_committed"]
    rr = DATA["nonrepeatable_repeatable_read"]
    ph_rc = DATA["phantom_read_committed"]
    ph_rr = DATA["phantom_repeatable_read"]
    ph_se = DATA["phantom_serializable"]
    locks = DATA["locks"]
    dead = DATA["deadlock"]
    victim = next(name for name, info in dead.items() if info["status"] == "error")
    survivor = next(name for name, info in dead.items() if info["status"] == "commit")

    add_h(doc, "3 Ход работы", page_break=True)
    add_p(
        doc,
        f"Опыты идут в базе варианта 5. Для них использованы студент {IDS['student_id']} "
        f"(Соколов Игорь Павлович) и студент {IDS['student2_id']} (Алексеев Иван Сергеевич) "
        f"группы 4226, а также выступление {IDS['talk_id']} с темой «{IDS['theme']}». "
        "Каждый сеанс — отдельное соединение с сервером.",
    )

    add_h(doc, "3.1 Грязное чтение")
    add_p(
        doc,
        "Сеанс A в уровне read committed меняет тему на «НЕЗАФИКСИРОВАННАЯ ПРАВКА ТЕМЫ» "
        "и не фиксирует транзакцию. Сеанс B читает ту же строку:",
    )
    add_code(
        doc,
        "BEGIN ISOLATION LEVEL READ UNCOMMITTED;\n"
        'SELECT "Название_темы"\n'
        'FROM "Выступление"\n'
        f"WHERE \"ID_выступления\" = {IDS['talk_id']};",
    )
    add_p(
        doc,
        f"SHOW transaction_isolation в сеансе B вернул {dirty['reader_isolation']}. "
        f"Сеанс A уже видел «{dirty['writer_uncommitted']}», сеанс B прочитал «{dirty['reader_saw']}». "
        "Незафиксированной правки в ответе нет. После ROLLBACK сеанса A повторное чтение снова вернуло "
        f"«{dirty['reader_saw_after_rollback']}». Грязное чтение не воспроизвелось: в PostgreSQL "
        "read uncommitted ведёт себя как read committed и показывает только зафиксированную версию строки.",
    )

    add_h(doc, "3.2 Неповторяемое чтение")
    add_p(
        doc,
        "Сеанс A дважды читает тему одного выступления. Между чтениями сеанс B меняет тему "
        "на «Черновик темы после фиксации правки» и выполняет COMMIT. Сеанс A:",
    )
    add_code(
        doc,
        'SELECT "Название_темы"\n'
        'FROM "Выступление"\n'
        f"WHERE \"ID_выступления\" = {IDS['talk_id']};",
    )
    add_p(
        doc,
        f"На уровне read committed первое чтение вернуло «{rc['first']}», второе — «{rc['second']}». "
        "Одна и та же строка в одной транзакции прочитана по-разному. "
        f"На уровне repeatable read оба чтения вернули «{rr['first']}»: зафиксированная правка "
        "второго сеанса в снимок первой транзакции не вошла.",
    )

    add_h(doc, "3.3 Фантомные строки")
    add_p(
        doc,
        "Сеанс A дважды считает студентов группы 4226. Между подсчётами сеанс B вставляет новую строку "
        "и фиксирует её. После опыта вставленная строка удалена.",
    )
    add_code(
        doc,
        "SELECT count(*)\n"
        'FROM "Студент" s\n'
        'JOIN "Группа" g\n'
        '  ON g."ID_группы" = s."ID_группы"\n'
        "WHERE g.\"Номер_группы\" = '4226';",
    )
    add_p(
        doc,
        f"На уровне read committed первый подсчёт дал {ph_rc['count_before']}, второй — {ph_rc['count_after']}. "
        "Фантом воспроизведён, ошибки при COMMIT читающей транзакции не было. "
        f"На repeatable read оба подсчёта равны {ph_rr['count_before']} и {ph_rr['count_after']}, "
        f"на serializable — {ph_se['count_before']} и {ph_se['count_after']}. "
        "Фантом 3 затем 4 получен только на read committed: repeatable read уже держит снимок "
        "на начало транзакции и отсекает новую строку.",
    )

    add_h(doc, "3.4 Блокировки")
    add_p(
        doc,
        f"Сеанс-держатель (pid {locks['pid_holder']}) внутри открытой транзакции выполняет "
        f"UPDATE студента {IDS['student_id']}, присваивая ФИО текущее значение: строка блокируется, "
        f"данные не меняются. Сеанс-ожидающий (pid {locks['pid_waiter']}) читает ту же строку с FOR SHARE "
        f"и останавливается. pg_blocking_pids вернул {locks['blocking_pids'][0]}.",
    )
    add_code(
        doc,
        'SELECT "ФИО"\n'
        'FROM "Студент"\n'
        f"WHERE \"ID_студента\" = {IDS['student_id']}\n"
        "FOR SHARE;",
    )
    add_p(
        doc,
        "В таблице 1 несколько строк отчёта о блокировках, не полный список. "
        "Пустой объект — блокировка идентификатора транзакции, а не таблицы.",
    )
    add_caption(doc, "Таблица 1 – Блокировки двух сеансов")
    add_table(
        doc,
        ["Сеанс", "Тип", "Режим", "Выдана", "Объект", "Ожидание"],
        lock_rows(),
        widths=[2.9, 2.8, 3.4, 2.2, 2.0, 3.2],
    )
    add_p(
        doc,
        f"Держатель {locks['pid_holder']} получил RowExclusiveLock на таблицу «Студент» и ExclusiveLock "
        "на идентификатор своей транзакции; он не ждёт. Ожидающий получил RowShareLock на таблицу и на строку, "
        "но ShareLock на транзакцию держателя не выдан: событие ожидания — Lock/transactionid.",
    )

    add_h(doc, "3.5 Взаимоблокировка")
    add_p(
        doc,
        f"Два сеанса обновляют студентов {IDS['student_id']} и {IDS['student2_id']} в противоположном порядке. "
        "Присваивание ФИО самому себе берёт блокировку записи и не меняет имя. "
        f"Сеанс {survivor.replace('session_', '')} ниже идёт от первого студента ко второму; "
        "второй сеанс выполняет те же два оператора наоборот.",
    )
    add_code(
        doc,
        'UPDATE "Студент" SET "ФИО" = "ФИО"\n'
        f"WHERE \"ID_студента\" = {IDS['student_id']};\n"
        'UPDATE "Студент" SET "ФИО" = "ФИО"\n'
        f"WHERE \"ID_студента\" = {IDS['student2_id']};",
    )
    add_p(
        doc,
        f"Сеанс {survivor.replace('session_', '')} зафиксировал транзакцию. "
        f"Сеанс {victim.replace('session_', '')} получил отказ, его транзакция отменена. "
        "Сообщение сервера:",
    )
    add_code(doc, deadlock_quote(dead[victim]["server"]))
    add_p(
        doc,
        "Сервер сообщает, что обнаружена взаимоблокировка: каждый из двух процессов ждёт ShareLock "
        "на транзакцию другого. Жертвой выбран один сеанс, второй после этого выполняет COMMIT.",
    )

    add_h(doc, "4 Выводы", page_break=True)
    add_p(
        doc,
        "Грязное чтение не получено. Уровень сеанса был read uncommitted, "
        f"SHOW вернул {dirty['reader_isolation']}, но SELECT увидел зафиксированную тему "
        f"«{dirty['reader_saw']}», а не незафиксированную правку. "
        "В PostgreSQL read uncommitted ведёт себя как read committed. "
        f"Неповторяемое чтение на read committed воспроизведено: второе чтение вернуло «{rc['second']}». "
        f"На repeatable read оба чтения вернули «{rr['first']}». "
        f"Фантом 3 затем 4 получен только на read committed "
        f"({ph_rc['count_before']} и {ph_rc['count_after']}). "
        f"На repeatable read и serializable повторный подсчёт остался равен {ph_rr['count_after']}. "
        f"Пока сеанс {locks['pid_holder']} держал студента {IDS['student_id']}, сеанс {locks['pid_waiter']} "
        f"с FOR SHARE ждал ShareLock. Встречные UPDATE дали сообщение «обнаружена взаимоблокировка»: "
        f"сеанс {survivor.replace('session_', '')} зафиксирован, сеанс {victim.replace('session_', '')} отменён.",
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
