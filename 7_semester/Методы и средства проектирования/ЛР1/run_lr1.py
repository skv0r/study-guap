# -*- coding: utf-8 -*-
"""Замеры планов и времени для лабораторной по индексации. Не для текста отчёта."""
import json
import re
import subprocess
import sys
from pathlib import Path

import psycopg

ROOT6 = Path(r"C:\Users\skvor\OneDrive\Desktop\GitHub\study-guap\6_semester\МиCПИСиТ")
SCHEMA = ROOT6 / "lr1" / "lab1_variant5_postgresql.sql"
SEED = ROOT6 / "lr2" / "lab2_variant5_data.sql"
OUT = Path(__file__).resolve().parent / "results_lr1.json"
PSQL = Path(r"C:\Program Files\PostgreSQL\17\bin\psql.exe")
PASSWORD = "Lab4321pg"
DB = "vuz_var5"

BULK = r"""
INSERT INTO "Группа" ("ID_факультета", "Номер_группы")
SELECT f."ID_факультета",
       'Т' || f."Номер_факультета"::text || '-' || lpad(gs::text, 3, '0')
FROM "Факультет" f
JOIN "Вуз" v ON v."ID_вуза" = f."ID_вуза"
CROSS JOIN generate_series(1, 50) AS gs
WHERE v."Название" = 'ГУАП';

INSERT INTO "Студент" ("ID_группы", "ФИО")
SELECT g."ID_группы",
       'Студент ' || f."Номер_факультета"::text || ' ' || g."Номер_группы" || ' ' || lpad(n::text, 3, '0')
FROM "Группа" g
JOIN "Факультет" f ON f."ID_факультета" = g."ID_факультета"
JOIN "Вуз" v ON v."ID_вуза" = f."ID_вуза" AND v."Название" = 'ГУАП'
CROSS JOIN generate_series(1, 500) AS n
WHERE g."Номер_группы" LIKE 'Т%';

INSERT INTO "Выступление" ("ID_программы", "Название_темы")
SELECT p."ID_программы",
       'Доклад ' || k."Название" || ' №' || gs::text
FROM "Конференция" k
JOIN "Программа_конференции" p ON p."ID_конференции" = k."ID_конференции"
CROSS JOIN generate_series(1, 4000) AS gs;

INSERT INTO "Выступление_студент" ("ID_выступления", "ID_студента")
SELECT v."ID_выступления", s."ID_студента"
FROM (
  SELECT "ID_студента", row_number() OVER (ORDER BY "ID_студента") AS rn
  FROM "Студент"
  WHERE "ФИО" LIKE 'Студент %'
) s
JOIN "Выступление" v
  ON ((v."ID_выступления" + s.rn) % 17) = 0
WHERE s.rn <= 120;
"""

INDEXES = """
CREATE INDEX "Студент_ФИО" ON "Студент" ("ФИО");
CREATE INDEX "Студент_ID_группы" ON "Студент" ("ID_группы");
CREATE INDEX "Выступление_ID_программы" ON "Выступление" ("ID_программы");
CREATE INDEX "Выступление_студент_ID_студента" ON "Выступление_студент" ("ID_студента");
CREATE INDEX "Группа_Номер_группы" ON "Группа" ("Номер_группы");
"""

QUERIES = [
    {
        "id": "fio",
        "title": "Поиск студента по ФИО",
        "source": "Отбор по полю, которое в запросах лабораторных 3–5 участвует в сортировке и в условиях",
        "sql": """
SELECT "Студент"."ФИО", "Группа"."Номер_группы"
FROM "Студент"
JOIN "Группа" ON "Группа"."ID_группы" = "Студент"."ID_группы"
WHERE "Студент"."ФИО" = 'Алексеев Иван Сергеевич'
""",
        "params": True,
    },
    {
        "id": "a",
        "title": "Студенты первого факультета на конференции «Информатика»",
        "source": "Запрос (а) лабораторной 3",
        "sql": """
SELECT DISTINCT "Студент"."ФИО"
FROM "Студент"
JOIN "Группа" ON "Группа"."ID_группы" = "Студент"."ID_группы"
JOIN "Факультет" ON "Факультет"."ID_факультета" = "Группа"."ID_факультета"
JOIN "Выступление_студент" ON "Выступление_студент"."ID_студента" = "Студент"."ID_студента"
JOIN "Выступление" ON "Выступление"."ID_выступления" = "Выступление_студент"."ID_выступления"
JOIN "Программа_конференции" ON "Программа_конференции"."ID_программы" = "Выступление"."ID_программы"
JOIN "Конференция" ON "Конференция"."ID_конференции" = "Программа_конференции"."ID_конференции"
WHERE "Факультет"."Номер_факультета" = 1
  AND "Конференция"."Название" = 'Информатика'
""",
        "params": True,
    },
    {
        "id": "b",
        "title": "Темы докладов студентов группы 4226",
        "source": "Запрос (б) лабораторной 3",
        "sql": """
SELECT DISTINCT "Выступление"."Название_темы"
FROM "Выступление"
JOIN "Выступление_студент" ON "Выступление_студент"."ID_выступления" = "Выступление"."ID_выступления"
JOIN "Студент" ON "Студент"."ID_студента" = "Выступление_студент"."ID_студента"
JOIN "Группа" ON "Группа"."ID_группы" = "Студент"."ID_группы"
WHERE "Группа"."Номер_группы" = '4226'
""",
        "params": False,
    },
    {
        "id": "v",
        "title": "Выступления студентов разных факультетов",
        "source": "Запрос (в) лабораторной 3",
        "sql": """
SELECT DISTINCT "Выступление"."Название_темы"
FROM "Выступление_студент" AS "Участие_1"
JOIN "Выступление_студент" AS "Участие_2"
  ON "Участие_2"."ID_выступления" = "Участие_1"."ID_выступления"
 AND "Участие_2"."ID_студента" > "Участие_1"."ID_студента"
JOIN "Выступление" ON "Выступление"."ID_выступления" = "Участие_1"."ID_выступления"
JOIN "Студент" AS "Участник_1" ON "Участник_1"."ID_студента" = "Участие_1"."ID_студента"
JOIN "Студент" AS "Участник_2" ON "Участник_2"."ID_студента" = "Участие_2"."ID_студента"
JOIN "Группа" AS "Группа_1" ON "Группа_1"."ID_группы" = "Участник_1"."ID_группы"
JOIN "Группа" AS "Группа_2" ON "Группа_2"."ID_группы" = "Участник_2"."ID_группы"
JOIN "Факультет" AS "Факультет_1" ON "Факультет_1"."ID_факультета" = "Группа_1"."ID_факультета"
JOIN "Факультет" AS "Факультет_2" ON "Факультет_2"."ID_факультета" = "Группа_2"."ID_факультета"
WHERE "Факультет_1"."ID_факультета" <> "Факультет_2"."ID_факультета"
""",
        "params": True,
    },
    {
        "id": "order",
        "title": "Список студентов по алфавиту ФИО",
        "source": "Запрос с ORDER BY лабораторной 3",
        "sql": """
SELECT "Студент"."ФИО", "Группа"."Номер_группы"
FROM "Студент"
JOIN "Группа" ON "Группа"."ID_группы" = "Студент"."ID_группы"
ORDER BY "Студент"."ФИО" ASC
""",
        "params": False,
    },
    {
        "id": "between",
        "title": "Факультеты ГУАП с номерами от 2 до 4",
        "source": "Запрос с BETWEEN лабораторной 3",
        "sql": """
SELECT "Факультет"."Номер_факультета", "Факультет"."Название"
FROM "Факультет"
JOIN "Вуз" ON "Вуз"."ID_вуза" = "Факультет"."ID_вуза"
WHERE "Вуз"."Название" = 'ГУАП'
  AND "Факультет"."Номер_факультета" BETWEEN 2 AND 4
""",
        "params": False,
    },
    {
        "id": "g",
        "title": "Среднее число участников выступления на конференции «Информатика»",
        "source": "Запрос (г) лабораторной 4",
        "sql": """
SELECT avg(t."Участников_на_выступление")
FROM (
  SELECT count("Выступление_студент"."ID_студента")::numeric AS "Участников_на_выступление"
  FROM "Выступление"
  JOIN "Программа_конференции"
    ON "Программа_конференции"."ID_программы" = "Выступление"."ID_программы"
  JOIN "Конференция"
    ON "Конференция"."ID_конференции" = "Программа_конференции"."ID_конференции"
  JOIN "Выступление_студент"
    ON "Выступление_студент"."ID_выступления" = "Выступление"."ID_выступления"
  WHERE "Конференция"."Название" = 'Информатика'
  GROUP BY "Выступление"."ID_выступления"
) AS t
""",
        "params": False,
    },
    {
        "id": "zh",
        "title": "Студенты четвёртого факультета, не выступавшие на конференциях",
        "source": "Запрос (ж) лабораторной 5",
        "sql": """
SELECT DISTINCT s."ФИО" AS "ФИО"
FROM "Студент" AS s
JOIN "Группа" AS g ON g."ID_группы" = s."ID_группы"
JOIN "Факультет" AS f ON f."ID_факультета" = g."ID_факультета"
WHERE f."Номер_факультета" = 4
  AND NOT EXISTS (
    SELECT 1
    FROM "Выступление_студент" AS vs
    WHERE vs."ID_студента" = s."ID_студента"
  )
ORDER BY s."ФИО"
""",
        "params": True,
    },
]

PARAM_SETS = [
    ("seqscan_off", {"enable_seqscan": "off"}),
    ("indexscan_off", {"enable_indexscan": "off"}),
    ("nestloop_off", {"enable_nestloop": "off"}),
    ("hashjoin_off", {"enable_hashjoin": "off"}),
]


def psql_file(path: Path):
    env = dict(**subprocess.os.environ)
    env["PGPASSWORD"] = PASSWORD
    env["PGCLIENTENCODING"] = "UTF8"
    cmd = [
        str(PSQL),
        "-h", "127.0.0.1",
        "-p", "5432",
        "-U", "postgres",
        "-d", DB,
        "-v", "ON_ERROR_STOP=1",
        "-f", str(path),
    ]
    print("psql", path.name, flush=True)
    r = subprocess.run(cmd, env=env, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.returncode != 0:
        print(r.stdout)
        print(r.stderr)
        raise SystemExit(f"psql failed {path}")


def admin_conn():
    return psycopg.connect(
        host="127.0.0.1", port=5432, dbname="postgres",
        user="postgres", password=PASSWORD, autocommit=True,
    )


def db_conn():
    return psycopg.connect(
        host="127.0.0.1", port=5432, dbname=DB,
        user="postgres", password=PASSWORD, autocommit=True,
    )


def recreate():
    with admin_conn() as conn:
        conn.execute(
            "SELECT pg_terminate_backend(pid) FROM pg_stat_activity WHERE datname = %s AND pid <> pg_backend_pid()",
            (DB,),
        )
        conn.execute(f'DROP DATABASE IF EXISTS "{DB}"')
        conn.execute(f'CREATE DATABASE "{DB}"')
    print("database created", flush=True)


def exec_script(conn, sql):
    parts = [p.strip() for p in sql.split(";") if p.strip()]
    for part in parts:
        conn.execute(part)


def counts(conn):
    names = [
        "Вуз", "Факультет", "Группа", "Студент",
        "Конференция", "Программа_конференции", "Выступление", "Выступление_студент",
    ]
    out = {}
    for name in names:
        out[name] = conn.execute(f'SELECT count(*) FROM "{name}"').fetchone()[0]
    sizes = {}
    for name in names:
        sizes[name] = conn.execute(
            "SELECT pg_size_pretty(pg_total_relation_size(%s::regclass))",
            (f'"{name}"',),
        ).fetchone()[0]
    return out, sizes


def parse_times(plan: str):
    ex = re.search(r"Execution Time: ([0-9.]+) ms", plan)
    pl = re.search(r"Planning Time: ([0-9.]+) ms", plan)
    return {
        "execution_ms": float(ex.group(1)) if ex else None,
        "planning_ms": float(pl.group(1)) if pl else None,
    }


def root_rows(plan: str):
    first = plan.splitlines()[0]
    m = re.search(r"actual time=[0-9.]+..[0-9.]+ rows=(\d+)", first)
    return int(m.group(1)) if m else None


def explain(conn, sql, settings=None):
    if settings:
        for key, val in settings.items():
            conn.execute(f"SET {key} = {val}")
    try:
        rows = conn.execute("EXPLAIN (ANALYZE, BUFFERS) " + sql).fetchall()
        plan = "\n".join(r[0] for r in rows)
    finally:
        if settings:
            for key in settings:
                conn.execute(f"RESET {key}")
    info = parse_times(plan)
    info["plan"] = plan
    info["nrows"] = root_rows(plan)
    return info


def sample_query(conn, sql, limit=8):
    rows = conn.execute(f"SELECT * FROM ({sql}) AS q LIMIT {limit}").fetchall()
    return [[None if v is None else str(v) for v in row] for row in rows]


def measure_phase(conn, runs=3):
    phase = []
    for q in QUERIES:
        print(" query", q["id"], flush=True)
        attempts = []
        for i in range(runs):
            print("  run", i + 1, flush=True)
            attempts.append(explain(conn, q["sql"]))
        phase.append({"id": q["id"], "runs": attempts})
    return phase


def measure_params(conn):
    out = []
    for q in QUERIES:
        if not q["params"]:
            continue
        print(" params", q["id"], flush=True)
        variants = []
        for name, settings in PARAM_SETS:
            print("  ", name, flush=True)
            info = explain(conn, q["sql"], settings)
            variants.append({"name": name, "settings": settings, "run": info})
        out.append({"id": q["id"], "variants": variants})
    return out


def main():
    recreate()
    psql_file(SCHEMA)
    psql_file(SEED)
    with db_conn() as conn:
        conn.execute("SET statement_timeout = '180s'")
        print("bulk", flush=True)
        exec_script(conn, BULK)
        print("analyze", flush=True)
        conn.execute("ANALYZE")
        row_counts, sizes = counts(conn)
        print(row_counts, flush=True)
        print("BEFORE", flush=True)
        before = measure_phase(conn)
        print("indexes", flush=True)
        exec_script(conn, INDEXES)
        samples = []
        for q in QUERIES:
            print(" sample", q["id"], flush=True)
            samples.append({"id": q["id"], "rows": sample_query(conn, q["sql"])})
        idx_rows = conn.execute(
            """
            SELECT indexname, indexdef
            FROM pg_indexes
            WHERE schemaname = 'public'
            ORDER BY tablename, indexname
            """
        ).fetchall()
        print("AFTER", flush=True)
        after = measure_phase(conn)
        print("PARAMS", flush=True)
        params = measure_params(conn)
        version = conn.execute("SHOW server_version").fetchone()[0]
    payload = {
        "server_version": version,
        "counts": row_counts,
        "sizes": sizes,
        "indexes_sql": INDEXES.strip(),
        "all_indexes": [{"name": a, "def": b} for a, b in idx_rows],
        "queries": [
            {"id": q["id"], "title": q["title"], "source": q["source"], "sql": q["sql"].strip()}
            for q in QUERIES
        ],
        "samples": samples,
        "before": before,
        "after": after,
        "params": params,
    }
    OUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print("saved", OUT, flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
