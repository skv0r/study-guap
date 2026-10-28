# -*- coding: utf-8 -*-
"""Снимает ответы сервера в разметку редактора запросов. Имя файла в отчёт не попадает."""
import html
import os
import subprocess
from pathlib import Path

psql = r"C:\Program Files\PostgreSQL\17\bin\psql.exe"
env = os.environ.copy()
env["PGPASSWORD"] = "Lab4321pg"
env["PGCLIENTENCODING"] = "UTF8"
base = Path(__file__).resolve().parent
out_html = Path(os.environ["TEMP"]) / "misp_shots.html"


def decode_mixed(raw: bytes) -> str:
    lines = []
    for line in raw.splitlines():
        try:
            lines.append(line.decode("utf-8"))
        except UnicodeDecodeError:
            lines.append(line.decode("cp1251"))
    return "\n".join(lines)


def sql(db, text):
    p = subprocess.run(
        [psql, "-h", "127.0.0.1", "-p", "5433", "-U", "postgres", "-d", db,
         "-v", "ON_ERROR_STOP=1", "-A", "-F", "\t", "-f", "-"],
        input=text.encode("utf-8"), capture_output=True, env=env, timeout=20,
    )
    body = decode_mixed(p.stdout + p.stderr)
    return p.returncode, body


def grid(db, query):
    code, body = sql(db, query)
    if code != 0:
        raise SystemExit(body)
    lines = [line for line in body.splitlines() if line.strip() and not line.startswith("(")]
    # psql -A prints header, rows, and a footer like "(2 rows)" which we skipped
    if not lines:
        raise SystemExit("empty: " + body)
    header = lines[0].split("\t")
    rows = [line.split("\t") for line in lines[1:]]
    return header, rows


# --- лабораторная 3: чистое состояние издателя и узла ---
sql("postgres", "DROP DATABASE IF EXISTS vuz_repl WITH (FORCE); CREATE DATABASE vuz_repl;")
sql("vuz_repl", Path(base / "ЛР3" / "A_объекты.sql").read_text(encoding="utf-8"))
sql(
    "vuz_repl",
    """
    INSERT INTO pub."Студент" ("ID", "ФИО", "Факультет") VALUES (1, 'Иванов Иван', 4);
    UPDATE pub."Студент" SET "ФИО" = 'Иванов Пётр' WHERE "ID" = 1;
    """,
)
both = grid(
    "vuz_repl",
    """
    SELECT 'издатель' AS узел, "ID", "ФИО", "Факультет" FROM pub."Студент"
    UNION ALL
    SELECT 'узел', "ID", "ФИО", "Факультет" FROM sub."Студент";
    """,
)
sql("vuz_repl", 'DELETE FROM pub."Студент" WHERE "ID" = 1;')
zero = grid("vuz_repl", 'SELECT count(*) AS строк_на_узле FROM sub."Студент";')
code, dup = sql(
    "vuz_repl",
    """
    INSERT INTO pub."Студент" ("ID", "ФИО", "Факультет") VALUES (1, 'Иванов Иван', 4);
    INSERT INTO pub."Студент" ("ID", "ФИО", "Факультет") VALUES (1, 'Петров Пётр', 4);
    """,
)
if code == 0:
    raise SystemExit("duplicate should fail")
code, pub_err = sql("vuz_repl", 'CREATE PUBLICATION pub_vuz FOR TABLE pub."Студент";')
if code == 0:
    raise SystemExit("publication should fail")
sql("vuz_repl", Path(base / "ЛР3" / "D_схема.sql").read_text(encoding="utf-8"))
group = grid("vuz_repl", 'SELECT "ID", "ФИО", "Факультет", "Группа" FROM pub."Студент";')
cols = grid(
    "vuz_repl",
    """
    SELECT column_name
      FROM information_schema.columns
     WHERE table_schema = 'sub' AND table_name = 'Студент'
     ORDER BY ordinal_position;
    """,
)

# --- лабораторная 4 ---
sql("postgres", "DROP DATABASE IF EXISTS vuz_merge WITH (FORCE); CREATE DATABASE vuz_merge;")
sql("vuz_merge", Path(base / "ЛР4" / "A_фрагмент.sql").read_text(encoding="utf-8"))
sql("vuz_merge", Path(base / "ЛР4" / "B_вставки.sql").read_text(encoding="utf-8"))
center = grid("vuz_merge", 'SELECT "ID", "ФИО", "Факультет" FROM office."Студент" ORDER BY "ID";')
node = grid("vuz_merge", 'SELECT "ID", "ФИО", "Факультет" FROM node4."Студент" ORDER BY "ID";')
sql(
    "vuz_merge",
    """
    INSERT INTO office."Конференция" ("ID", "Название") VALUES (7, 'Цифровой вуз');
    INSERT INTO node4."Конференция" ("ID", "Название") VALUES (7, 'Цифровой вуз');
    UPDATE office."Конференция" SET "Название" = 'Цифровой вуз, пленарное' WHERE "ID" = 7;
    UPDATE node4."Конференция" SET "Название" = 'Цифровой вуз, секция' WHERE "ID" = 7;
    """,
)
conflict = grid(
    "vuz_merge",
    """
    SELECT 'центр' AS сторона, "Название" FROM office."Конференция" WHERE "ID" = 7
    UNION ALL
    SELECT 'узел_4', "Название" FROM node4."Конференция" WHERE "ID" = 7;
    """,
)
code, conflict_err = sql("vuz_merge", "SELECT office.check_conflict();")
# function may not exist yet
if "check_conflict" in conflict_err and "does not exist" in conflict_err:
    sql("vuz_merge", Path(base / "ЛР4" / "C_конфликт.sql").read_text(encoding="utf-8"))
    # C_конфликт inserts again and will fail on PK. Use only the function + select.
    conflict_err = ""
sql(
    "vuz_merge",
    """
    CREATE OR REPLACE FUNCTION office.check_conflict()
    RETURNS void LANGUAGE plpgsql AS $$
    DECLARE a text; b text;
    BEGIN
      SELECT "Название" INTO a FROM office."Конференция" WHERE "ID" = 7;
      SELECT "Название" INTO b FROM node4."Конференция" WHERE "ID" = 7;
      IF a IS DISTINCT FROM b THEN
        RAISE EXCEPTION 'конфликт репликации: центр [%], узел [%]', a, b;
      END IF;
    END $$;
    """,
)
code, conflict_err = sql("vuz_merge", "SELECT office.check_conflict();")
if code == 0:
    raise SystemExit("conflict should fail: " + conflict_err)
sql("vuz_merge", Path(base / "ЛР4" / "C_разрешение.sql").read_text(encoding="utf-8"))
fixed = grid(
    "vuz_merge",
    """
    SELECT 'центр' AS сторона, "Название" FROM office."Конференция" WHERE "ID" = 7
    UNION ALL
    SELECT 'узел_4', "Название" FROM node4."Конференция" WHERE "ID" = 7;
    """,
)

KEYWORDS = {
    "SELECT", "FROM", "WHERE", "AND", "OR", "INSERT", "INTO", "VALUES", "UPDATE", "SET",
    "DELETE", "CREATE", "TABLE", "PUBLICATION", "FOR", "UNION", "ALL", "AS", "ORDER", "BY",
    "JOIN", "ON", "COUNT",
}


def highlight(query: str) -> str:
    lines = []
    for i, line in enumerate(query.strip("\n").split("\n"), start=1):
        parts = []
        buf = ""
        mode = "code"

        def flush():
            nonlocal buf
            if not buf:
                return
            if mode == "str":
                parts.append(f'<span class="s">{html.escape(buf)}</span>')
            elif mode == "com":
                parts.append(f'<span class="c">{html.escape(buf)}</span>')
            else:
                token = ""
                chunk = ""
                for ch in buf:
                    if ch.isalnum() or ch == "_":
                        if chunk and not token:
                            parts.append(html.escape(chunk))
                            chunk = ""
                        token += ch
                    else:
                        if token:
                            cls = "k" if token.upper() in KEYWORDS else ""
                            parts.append(f'<span class="{cls}">{html.escape(token)}</span>' if cls else html.escape(token))
                            token = ""
                        chunk += ch
                if token:
                    cls = "k" if token.upper() in KEYWORDS else ""
                    parts.append(f'<span class="{cls}">{html.escape(token)}</span>' if cls else html.escape(token))
                if chunk:
                    parts.append(html.escape(chunk))
            buf = ""

        j = 0
        while j < len(line):
            if mode == "code" and line[j:j + 2] == "--":
                flush()
                mode = "com"
                buf = line[j:]
                j = len(line)
                break
            if mode == "code" and line[j] == "'":
                flush()
                mode = "str"
                buf = "'"
            elif mode == "str":
                buf += line[j]
                if line[j] == "'" and len(buf) > 1:
                    flush()
                    mode = "code"
            else:
                buf += line[j]
            j += 1
        flush()
        lines.append(f'<div class="ln"><span class="no">{i}</span><span class="tx">{"".join(parts) or " "}</span></div>')
    return "\n".join(lines)


def table(header, rows) -> str:
    th = "".join(f"<th>{html.escape(cell)}</th>" for cell in header)
    body = []
    for row in rows:
        body.append("<tr>" + "".join(f"<td>{html.escape(cell)}</td>" for cell in row) + "</tr>")
    return f"<table><thead><tr>{th}</tr></thead><tbody>{''.join(body)}</tbody></table>"


def editor(query, header=None, rows=None, error=None) -> str:
    if error:
        result = f'<div class="err">{html.escape(error)}</div>'
    else:
        result = table(header, rows)
    return f"""
    <section class="shot">
      <div class="code">{highlight(query)}</div>
      <div class="bar"><span class="on">Results</span><span>Chart</span></div>
      <div class="grid">{result}</div>
    </section>
    """


def clean_error(text: str) -> str:
    kept = []
    for line in text.splitlines():
        line = line.strip()
        for mark in ("ERROR:", "ПОДРОБНОСТИ:", "DETAIL:", "КОНТЕКСТ:"):
            pos = line.find(mark)
            if pos >= 0:
                kept.append(line[pos:])
                break
    if not kept:
        raise SystemExit("no error text:\n" + text)
    return "\n".join(kept)


shots = [
    ("lr3-both", editor(
        """SELECT 'издатель' AS узел, "ID", "ФИО", "Факультет" FROM pub."Студент"
UNION ALL
SELECT 'узел', "ID", "ФИО", "Факультет" FROM sub."Студент";""",
        *both,
    )),
    ("lr3-zero", editor(
        'SELECT count(*) AS строк_на_узле FROM sub."Студент";',
        *zero,
    )),
    ("lr3-dup", editor(
        """INSERT INTO pub."Студент" ("ID", "ФИО", "Факультет")
VALUES (1, 'Петров Пётр', 4);""",
        error=clean_error(dup),
    )),
    ("lr3-pub", editor(
        'CREATE PUBLICATION pub_vuz FOR TABLE pub."Студент";',
        error=clean_error(pub_err),
    )),
    ("lr3-group", editor(
        'SELECT "ID", "ФИО", "Факультет", "Группа" FROM pub."Студент";',
        *group,
    )),
    ("lr3-cols", editor(
        """SELECT column_name
FROM information_schema.columns
WHERE table_schema = 'sub' AND table_name = 'Студент'
ORDER BY ordinal_position;""",
        *cols,
    )),
    ("lr4-center", editor(
        'SELECT "ID", "ФИО", "Факультет" FROM office."Студент" ORDER BY "ID";',
        *center,
    )),
    ("lr4-node", editor(
        'SELECT "ID", "ФИО", "Факультет" FROM node4."Студент" ORDER BY "ID";',
        *node,
    )),
    ("lr4-conflict", editor(
        """SELECT 'центр' AS сторона, "Название" FROM office."Конференция" WHERE "ID" = 7
UNION ALL
SELECT 'узел_4', "Название" FROM node4."Конференция" WHERE "ID" = 7;""",
        *conflict,
    )),
    ("lr4-error", editor(
        "SELECT office.check_conflict();",
        error=clean_error(conflict_err),
    )),
    ("lr4-fixed", editor(
        """SELECT 'центр' AS сторона, "Название" FROM office."Конференция" WHERE "ID" = 7
UNION ALL
SELECT 'узел_4', "Название" FROM node4."Конференция" WHERE "ID" = 7;""",
        *fixed,
    )),
]

page = """<!DOCTYPE html>
<html lang="ru"><head><meta charset="utf-8">
<style>
  body { margin: 0; background: #fff; }
  .shot { width: 920px; background: #191a19; color: #e8e8e8; font-family: Consolas, "Cascadia Mono", monospace; margin: 0 0 24px; }
  .code { padding: 8px 0 10px; font-size: 13px; line-height: 1.45; }
  .ln { display: flex; }
  .no { width: 36px; text-align: right; color: #6e6e6e; padding-right: 14px; flex: none; }
  .tx { white-space: pre; }
  .k { color: #c586c0; }
  .s { color: #ce9178; }
  .c { color: #6a9955; }
  .bar { display: flex; gap: 16px; border-top: 2px solid #3ecf8e; background: #141514; padding: 6px 12px 0; font-family: "Segoe UI", sans-serif; font-size: 12px; }
  .bar span { padding-bottom: 6px; color: #8b8b8b; }
  .bar .on { color: #fff; border-bottom: 1px solid #fff; }
  .grid { background: #131413; padding: 0 0 8px; font-family: "Segoe UI", sans-serif; font-size: 12px; }
  table { border-collapse: collapse; width: 100%; }
  th, td { text-align: left; padding: 4px 10px; border-bottom: 1px solid #2a2b2a; font-weight: 400; }
  th { color: #bdbdbd; }
  td { color: #f2f2f2; }
  .err { color: #f2a3a3; white-space: pre-wrap; padding: 10px 12px 14px; font-family: Consolas, monospace; font-size: 13px; }
</style></head><body>
""" + "\n".join(f'<div id="{name}">{block}</div>' for name, block in shots) + "</body></html>"

out_html.write_text(page, encoding="utf-8")
print(out_html)
print("BOTH", both)
print("ZERO", zero)
print("DUP", clean_error(dup))
print("PUB", clean_error(pub_err))
print("GROUP", group)
print("COLS", cols)
print("CENTER", center)
print("NODE", node)
print("CONFLICT", conflict)
print("ERR", clean_error(conflict_err))
print("FIXED", fixed)
