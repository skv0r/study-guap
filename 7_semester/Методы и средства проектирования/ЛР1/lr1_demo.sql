-- Показать преподавателю время и план до индексов и после.
-- Вставь файл целиком в SQL Editor и нажми Run.
-- Вкладка Results покажет одну таблицу: строки, время до, время после, начало плана.
-- После выполнения пять индексов остаются. Чтобы снова показать состояние «до»,
-- выполни lr1_otkat.sql, затем этот файл ещё раз.

SET statement_timeout = '180s';

DROP INDEX IF EXISTS "Студент_ФИО";
DROP INDEX IF EXISTS "Студент_ID_группы";
DROP INDEX IF EXISTS "Выступление_ID_программы";
DROP INDEX IF EXISTS "Выступление_студент_ID_студента";
DROP INDEX IF EXISTS "Группа_Номер_группы";

ANALYZE;

DROP TABLE IF EXISTS lab_queries;
CREATE TEMP TABLE lab_queries (
  ord integer PRIMARY KEY,
  title text NOT NULL,
  body text NOT NULL
);

INSERT INTO lab_queries (ord, title, body) VALUES
(1, 'Поиск по ФИО', $q$
SELECT "Студент"."ФИО", "Группа"."Номер_группы"
FROM "Студент"
JOIN "Группа" ON "Группа"."ID_группы" = "Студент"."ID_группы"
WHERE "Студент"."ФИО" = 'Алексеев Иван Сергеевич'
$q$),
(2, 'Запрос (а)', $q$
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
$q$),
(3, 'Запрос (б)', $q$
SELECT DISTINCT "Выступление"."Название_темы"
FROM "Выступление"
JOIN "Выступление_студент" ON "Выступление_студент"."ID_выступления" = "Выступление"."ID_выступления"
JOIN "Студент" ON "Студент"."ID_студента" = "Выступление_студент"."ID_студента"
JOIN "Группа" ON "Группа"."ID_группы" = "Студент"."ID_группы"
WHERE "Группа"."Номер_группы" = '4226'
$q$),
(4, 'Запрос (в)', $q$
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
$q$),
(5, 'Сортировка по ФИО', $q$
SELECT "Студент"."ФИО", "Группа"."Номер_группы"
FROM "Студент"
JOIN "Группа" ON "Группа"."ID_группы" = "Студент"."ID_группы"
ORDER BY "Студент"."ФИО" ASC
$q$),
(6, 'BETWEEN', $q$
SELECT "Факультет"."Номер_факультета", "Факультет"."Название"
FROM "Факультет"
JOIN "Вуз" ON "Вуз"."ID_вуза" = "Факультет"."ID_вуза"
WHERE "Вуз"."Название" = 'ГУАП'
  AND "Факультет"."Номер_факультета" BETWEEN 2 AND 4
$q$),
(7, 'Среднее число участников', $q$
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
$q$),
(8, 'Запрос (ж)', $q$
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
$q$);

DROP TABLE IF EXISTS lab_times;
CREATE TEMP TABLE lab_times (
  ord integer NOT NULL,
  title text NOT NULL,
  phase text NOT NULL,
  rows_n integer,
  execution_ms numeric,
  plan_root text
);

CREATE OR REPLACE FUNCTION pg_temp.lab_measure(p_phase text)
RETURNS void
LANGUAGE plpgsql
AS $fn$
DECLARE
  q record;
  line record;
  n integer;
  root text;
  scan_line text;
  index_line text;
  ms numeric;
  nrows integer;
  plan_line text;
BEGIN
  FOR q IN SELECT ord, title, body FROM lab_queries ORDER BY ord LOOP
    n := 0;
    root := NULL;
    scan_line := NULL;
    index_line := NULL;
    ms := NULL;
    nrows := NULL;
    FOR line IN EXECUTE 'EXPLAIN (ANALYZE, BUFFERS) ' || q.body LOOP
      n := n + 1;
      plan_line := line."QUERY PLAN";
      IF n = 1 THEN
        root := plan_line;
        nrows := substring(plan_line from 'actual time=[0-9.]+[.][.][0-9.]+ rows=([0-9]+)')::integer;
      END IF;
      IF index_line IS NULL AND plan_line ~ 'Index (Only )?Scan using "(Студент_ФИО|Студент_ID_группы|Выступление_ID_программы|Выступление_студент_ID_студента|Группа_Номер_группы)"' THEN
        index_line := btrim(plan_line);
      END IF;
      IF scan_line IS NULL AND plan_line ~ 'Seq Scan on' THEN
        scan_line := btrim(plan_line);
      END IF;
      IF plan_line LIKE 'Execution Time:%' THEN
        ms := substring(plan_line from '([0-9]+[.][0-9]+|[0-9]+)')::numeric;
      END IF;
    END LOOP;
    INSERT INTO lab_times (ord, title, phase, rows_n, execution_ms, plan_root)
    VALUES (q.ord, q.title, p_phase, nrows, ms, coalesce(index_line, scan_line, root));
  END LOOP;
END;
$fn$;

DO $$ BEGIN PERFORM pg_temp.lab_measure('before'); END $$;

CREATE INDEX "Студент_ФИО" ON "Студент" ("ФИО");
CREATE INDEX "Студент_ID_группы" ON "Студент" ("ID_группы");
CREATE INDEX "Выступление_ID_программы" ON "Выступление" ("ID_программы");
CREATE INDEX "Выступление_студент_ID_студента" ON "Выступление_студент" ("ID_студента");
CREATE INDEX "Группа_Номер_группы" ON "Группа" ("Номер_группы");

DO $$ BEGIN PERFORM pg_temp.lab_measure('after'); END $$;

SELECT
  b.title AS "Запрос",
  b.rows_n AS "Строк",
  b.execution_ms AS "До, мс",
  a.execution_ms AS "После, мс",
  round(b.execution_ms - a.execution_ms, 3) AS "Разница, мс",
  left(regexp_replace(b.plan_root, '^[\s>-]+', ''), 160) AS "План до",
  left(regexp_replace(a.plan_root, '^[\s>-]+', ''), 160) AS "План после"
FROM lab_times b
JOIN lab_times a ON a.ord = b.ord AND a.phase = 'after'
WHERE b.phase = 'before'
ORDER BY b.ord;
