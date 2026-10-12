# -*- coding: utf-8 -*-
"""Два сеанса: изоляция, блокировки, тупик. Не для текста отчёта."""
import json
import threading
import time
from pathlib import Path

import psycopg

PASSWORD = "Lab4321pg"
DB = "vuz_var5"
OUT = Path(__file__).resolve().parent / "results_lr2.json"


def connect():
    return psycopg.connect(
        host="127.0.0.1", port=5432, dbname=DB,
        user="postgres", password=PASSWORD, autocommit=False,
    )


def fetch_all(conn, sql, args=None):
    with conn.cursor() as cur:
        cur.execute(sql, args or ())
        if cur.description is None:
            return []
        cols = [d.name for d in cur.description]
        rows = cur.fetchall()
    return cols, [list(r) for r in rows]


def one(conn, sql, args=None):
    with conn.cursor() as cur:
        cur.execute(sql, args or ())
        if cur.description is None:
            return None
        return cur.fetchone()


def scalar(conn, sql, args=None):
    row = one(conn, sql, args)
    return None if row is None else row[0]


def exc_text(err: BaseException) -> str:
    parts = []
    diag = getattr(err, "diag", None)
    if diag is not None:
        for label, attr in (
            ("ERROR", "message_primary"),
            ("DETAIL", "message_detail"),
            ("HINT", "message_hint"),
            ("CONTEXT", "context"),
        ):
            val = getattr(diag, attr, None)
            if val:
                parts.append(f"{label}:  {val}")
    if not parts:
        parts.append(str(err))
    return "\n".join(parts)


def setup(conn):
    conn.execute("BEGIN")
    # служебные строки для опытов, в предметной схеме
    fid = scalar(
        conn,
        """
        SELECT g."ID_группы"
        FROM "Группа" g
        JOIN "Факультет" f ON f."ID_факультета" = g."ID_факультета"
        JOIN "Вуз" v ON v."ID_вуза" = f."ID_вуза"
        WHERE v."Название" = 'ГУАП' AND g."Номер_группы" = '4226'
        """,
    )
    conn.execute(
        """
        INSERT INTO "Студент" ("ID_группы", "ФИО")
        SELECT %s, 'Соколов Игорь Павлович'
        WHERE NOT EXISTS (
          SELECT 1 FROM "Студент" WHERE "ФИО" = 'Соколов Игорь Павлович'
        )
        """,
        (fid,),
    )
    conn.execute(
        """
        INSERT INTO "Выступление" ("ID_программы", "Название_темы")
        SELECT p."ID_программы", 'Черновик темы для проверки изоляции'
        FROM "Конференция" k
        JOIN "Программа_конференции" p ON p."ID_конференции" = k."ID_конференции"
        WHERE k."Название" = 'Информатика'
          AND NOT EXISTS (
            SELECT 1 FROM "Выступление"
            WHERE "Название_темы" = 'Черновик темы для проверки изоляции'
          )
        """,
    )
    conn.commit()
    sid = scalar(conn, """SELECT "ID_студента" FROM "Студент" WHERE "ФИО" = 'Соколов Игорь Павлович'""")
    sid2 = scalar(conn, """SELECT "ID_студента" FROM "Студент" WHERE "ФИО" = 'Алексеев Иван Сергеевич'""")
    vid = scalar(
        conn,
        """SELECT "ID_выступления" FROM "Выступление" WHERE "Название_темы" = 'Черновик темы для проверки изоляции'""",
    )
    theme = scalar(conn, """SELECT "Название_темы" FROM "Выступление" WHERE "ID_выступления" = %s""", (vid,))
    return {"group_id": fid, "student_id": sid, "student2_id": sid2, "talk_id": vid, "theme": theme}


def restore_theme(talk_id, theme):
    c = connect()
    c.execute(
        'UPDATE "Выступление" SET "Название_темы" = %s WHERE "ID_выступления" = %s',
        (theme, talk_id),
    )
    c.commit()
    c.close()


def dirty(ids):
    a = connect()
    b = connect()
    a.execute("BEGIN ISOLATION LEVEL READ COMMITTED")
    a.execute("SET application_name = 'dirty_writer'")
    a.execute(
        """UPDATE "Выступление" SET "Название_темы" = 'НЕЗАФИКСИРОВАННАЯ ПРАВКА ТЕМЫ' WHERE "ID_выступления" = %s""",
        (ids["talk_id"],),
    )
    seen_by_writer = scalar(
        a,
        """SELECT "Название_темы" FROM "Выступление" WHERE "ID_выступления" = %s""",
        (ids["talk_id"],),
    )
    b.execute("BEGIN ISOLATION LEVEL READ UNCOMMITTED")
    b.execute("SET application_name = 'dirty_reader'")
    level = scalar(b, "SHOW transaction_isolation")
    seen = scalar(
        b,
        """SELECT "Название_темы" FROM "Выступление" WHERE "ID_выступления" = %s""",
        (ids["talk_id"],),
    )
    a.rollback()
    after = scalar(
        b,
        """SELECT "Название_темы" FROM "Выступление" WHERE "ID_выступления" = %s""",
        (ids["talk_id"],),
    )
    b.rollback()
    a.close()
    b.close()
    return {
        "writer_uncommitted": seen_by_writer,
        "reader_isolation": level,
        "reader_saw": seen,
        "reader_saw_after_rollback": after,
        "original": ids["theme"],
    }


def nonrepeatable(ids, level):
    a = connect()
    b = connect()
    a.execute(f"BEGIN ISOLATION LEVEL {level}")
    first = scalar(
        a,
        """SELECT "Название_темы" FROM "Выступление" WHERE "ID_выступления" = %s""",
        (ids["talk_id"],),
    )
    b.execute("BEGIN")
    b.execute(
        """UPDATE "Выступление" SET "Название_темы" = 'Черновик темы после фиксации правки' WHERE "ID_выступления" = %s""",
        (ids["talk_id"],),
    )
    b.commit()
    second = scalar(
        a,
        """SELECT "Название_темы" FROM "Выступление" WHERE "ID_выступления" = %s""",
        (ids["talk_id"],),
    )
    a.rollback()
    b.close()
    a.close()
    restore_theme(ids["talk_id"], ids["theme"])
    return {"level": level, "first": first, "second": second}


def count_group(conn):
    return scalar(
        conn,
        """
        SELECT count(*)
        FROM "Студент" s
        JOIN "Группа" g ON g."ID_группы" = s."ID_группы"
        WHERE g."Номер_группы" = '4226'
        """,
    )


def phantom(ids, level):
    a = connect()
    b = connect()
    a.execute(f"BEGIN ISOLATION LEVEL {level}")
    c1 = count_group(a)
    b.execute("BEGIN")
    b.execute(
        """
        INSERT INTO "Студент" ("ID_группы", "ФИО")
        VALUES (%s, 'Фантомный Студент Сеанса')
        """,
        (ids["group_id"],),
    )
    b.commit()
    c2 = count_group(a)
    err = None
    try:
        a.commit()
    except Exception as e:
        err = exc_text(e)
        a.rollback()
    # убрать вставленную строку
    c = connect()
    c.execute("""DELETE FROM "Студент" WHERE "ФИО" = 'Фантомный Студент Сеанса'""")
    c.commit()
    c.close()
    b.close()
    a.close()
    return {"level": level, "count_before": c1, "count_after": c2, "commit_error": err}


def lock_report(ids):
    a = connect()
    b = connect()
    watcher = psycopg.connect(
        host="127.0.0.1", port=5432, dbname=DB,
        user="postgres", password=PASSWORD, autocommit=True,
    )
    a.execute("BEGIN")
    a.execute("SET application_name = 'lock_holder'")
    pid_a = scalar(a, "SELECT pg_backend_pid()")
    a.execute(
        """UPDATE "Студент" SET "ФИО" = "ФИО" WHERE "ID_студента" = %s""",
        (ids["student_id"],),
    )
    pid_b_box = {}
    held = threading.Event()

    def waiter():
        b.execute("BEGIN")
        b.execute("SET application_name = 'lock_waiter'")
        pid_b_box["pid"] = scalar(b, "SELECT pg_backend_pid()")
        held.set()
        b.execute(
            """SELECT "ФИО" FROM "Студент" WHERE "ID_студента" = %s FOR SHARE""",
            (ids["student_id"],),
        )
        b.rollback()

    t = threading.Thread(target=waiter)
    t.start()
    held.wait(10)
    time.sleep(0.4)
    pid_b = pid_b_box["pid"]
    cols, rows = fetch_all(
        watcher,
        """
        SELECT l.pid,
               c.application_name,
               l.locktype,
               l.mode,
               l.granted,
               CASE WHEN l.relation IS NULL THEN NULL ELSE l.relation::regclass::text END AS relation,
               l.page,
               l.tuple,
               c.wait_event_type,
               c.wait_event,
               left(c.query, 180) AS query
        FROM pg_locks l
        JOIN pg_stat_activity c ON c.pid = l.pid
        WHERE l.pid = ANY(%s)
        ORDER BY l.pid, l.granted DESC, l.locktype, l.mode
        """,
        ([pid_a, pid_b],),
    )
    blocking = scalar(watcher, "SELECT pg_blocking_pids(%s)", (pid_b,))
    a.rollback()
    t.join(15)
    a.close()
    b.close()
    watcher.close()
    return {
        "pid_holder": pid_a,
        "pid_waiter": pid_b,
        "blocking_pids": list(blocking) if blocking is not None else [],
        "columns": cols,
        "rows": rows,
    }


def deadlock(ids):
    a = connect()
    b = connect()
    barrier = threading.Barrier(2)
    out = {}

    def work(name, conn, first_id, second_id):
        try:
            conn.execute("BEGIN")
            conn.execute(f"SET application_name = '{name}'")
            conn.execute("SET deadlock_timeout = '1s'")
            conn.execute(
                """UPDATE "Студент" SET "ФИО" = "ФИО" WHERE "ID_студента" = %s""",
                (first_id,),
            )
            barrier.wait(20)
            conn.execute(
                """UPDATE "Студент" SET "ФИО" = "ФИО" WHERE "ID_студента" = %s""",
                (second_id,),
            )
            conn.commit()
            out[name] = {"status": "commit"}
        except Exception as e:
            out[name] = {"status": "error", "server": exc_text(e)}
            try:
                conn.rollback()
            except Exception:
                pass

    t1 = threading.Thread(target=work, args=("session_1", a, ids["student_id"], ids["student2_id"]))
    t2 = threading.Thread(target=work, args=("session_2", b, ids["student2_id"], ids["student_id"]))
    t1.start()
    t2.start()
    t1.join(30)
    t2.join(30)
    a.close()
    b.close()
    return out


def main():
    boot = connect()
    ids = setup(boot)
    boot.close()
    result = {
        "ids": {k: (int(v) if hasattr(v, "__int__") and not isinstance(v, str) else v) for k, v in ids.items()},
        "dirty": dirty(ids),
        "nonrepeatable_read_committed": nonrepeatable(ids, "READ COMMITTED"),
        "nonrepeatable_repeatable_read": nonrepeatable(ids, "REPEATABLE READ"),
        "phantom_read_committed": phantom(ids, "READ COMMITTED"),
        "phantom_repeatable_read": phantom(ids, "REPEATABLE READ"),
        "phantom_serializable": phantom(ids, "SERIALIZABLE"),
        "locks": lock_report(ids),
        "deadlock": deadlock(ids),
    }
    # ids may contain int already
    result["ids"] = {
        "group_id": int(ids["group_id"]),
        "student_id": int(ids["student_id"]),
        "student2_id": int(ids["student2_id"]),
        "talk_id": int(ids["talk_id"]),
        "theme": ids["theme"],
    }
    OUT.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(OUT)


if __name__ == "__main__":
    main()
