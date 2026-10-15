-- Лабораторная 2. Транзакции, уровни изоляции и блокировки. Вариант 5, PostgreSQL.
-- База уже наполнена. Номера строк ниже — из выполненного опыта:
-- студент 100009 Соколов Игорь Павлович, студент 1 Алексеев Иван Сергеевич,
-- выступление 12010 «Черновик темы для проверки изоляции».
-- Два сеанса — два отдельных соединения. Команды одного сеанса идут подряд.

-- подготовка
INSERT INTO "Студент" ("ID_группы", "ФИО")
SELECT g."ID_группы", 'Соколов Игорь Павлович'
FROM "Группа" g
JOIN "Факультет" f ON f."ID_факультета" = g."ID_факультета"
JOIN "Вуз" v ON v."ID_вуза" = f."ID_вуза"
WHERE v."Название" = 'ГУАП'
  AND g."Номер_группы" = '4226'
  AND NOT EXISTS (
    SELECT 1 FROM "Студент" WHERE "ФИО" = 'Соколов Игорь Павлович'
  );

INSERT INTO "Выступление" ("ID_программы", "Название_темы")
SELECT p."ID_программы", 'Черновик темы для проверки изоляции'
FROM "Конференция" k
JOIN "Программа_конференции" p ON p."ID_конференции" = k."ID_конференции"
WHERE k."Название" = 'Информатика'
  AND NOT EXISTS (
    SELECT 1 FROM "Выступление"
    WHERE "Название_темы" = 'Черновик темы для проверки изоляции'
  );

-- 1. Грязное чтение
-- сеанс A
BEGIN ISOLATION LEVEL READ COMMITTED;
UPDATE "Выступление"
SET "Название_темы" = 'НЕЗАФИКСИРОВАННАЯ ПРАВКА ТЕМЫ'
WHERE "ID_выступления" = 12010;
SELECT "Название_темы"
FROM "Выступление"
WHERE "ID_выступления" = 12010;

-- сеанс B, пока сеанс A не зафиксировал правку
BEGIN ISOLATION LEVEL READ UNCOMMITTED;
SHOW transaction_isolation;
SELECT "Название_темы"
FROM "Выступление"
WHERE "ID_выступления" = 12010;

-- сеанс A
ROLLBACK;

-- сеанс B
SELECT "Название_темы"
FROM "Выступление"
WHERE "ID_выступления" = 12010;
ROLLBACK;

-- 2. Неповторяемое чтение
-- сеанс A
BEGIN ISOLATION LEVEL READ COMMITTED;
SELECT "Название_темы"
FROM "Выступление"
WHERE "ID_выступления" = 12010;

-- сеанс B
BEGIN;
UPDATE "Выступление"
SET "Название_темы" = 'Черновик темы после фиксации правки'
WHERE "ID_выступления" = 12010;
COMMIT;

-- сеанс A, второе чтение той же строки
SELECT "Название_темы"
FROM "Выступление"
WHERE "ID_выступления" = 12010;
ROLLBACK;

-- тот же опыт на снимке: у сеанса A уровень REPEATABLE READ,
-- оба чтения возвращают тему, которая была на BEGIN.
-- после опыта вернуть исходную тему:
UPDATE "Выступление"
SET "Название_темы" = 'Черновик темы для проверки изоляции'
WHERE "ID_выступления" = 12010;

-- 3. Фантомные строки
-- сеанс A
BEGIN ISOLATION LEVEL READ COMMITTED;
SELECT count(*)
FROM "Студент" s
JOIN "Группа" g ON g."ID_группы" = s."ID_группы"
WHERE g."Номер_группы" = '4226';

-- сеанс B
BEGIN;
INSERT INTO "Студент" ("ID_группы", "ФИО")
SELECT g."ID_группы", 'Фантомный Студент Сеанса'
FROM "Группа" g
WHERE g."Номер_группы" = '4226';
COMMIT;

-- сеанс A, второй подсчёт
SELECT count(*)
FROM "Студент" s
JOIN "Группа" g ON g."ID_группы" = s."ID_группы"
WHERE g."Номер_группы" = '4226';
COMMIT;

-- убрать вставленную строку
DELETE FROM "Студент" WHERE "ФИО" = 'Фантомный Студент Сеанса';

-- тот же подсчёт на REPEATABLE READ и на SERIALIZABLE:
-- оба раза 3, новая строка в снимок не входит.

-- 4. Блокировки
-- сеанс-держатель
BEGIN;
UPDATE "Студент"
SET "ФИО" = "ФИО"
WHERE "ID_студента" = 100009;

-- сеанс-ожидающий, останавливается на FOR SHARE
BEGIN;
SELECT "ФИО"
FROM "Студент"
WHERE "ID_студента" = 100009
FOR SHARE;

-- отчёт о блокировках, третье соединение
SELECT l.pid,
       c.application_name,
       l.locktype,
       l.mode,
       l.granted,
       CASE WHEN l.relation IS NULL THEN NULL ELSE l.relation::regclass::text END AS relation,
       c.wait_event_type,
       c.wait_event,
       left(c.query, 180) AS query
FROM pg_locks l
JOIN pg_stat_activity c ON c.pid = l.pid
WHERE c.query ILIKE '%100009%'
   OR c.application_name IN ('lock_holder', 'lock_waiter')
ORDER BY l.pid, l.granted DESC, l.locktype, l.mode;

-- SELECT pg_blocking_pids(<pid ожидающего>);
-- затем ROLLBACK в обоих сеансах.

-- 5. Взаимоблокировка
-- сеанс 1
BEGIN;
SET deadlock_timeout = '1s';
UPDATE "Студент" SET "ФИО" = "ФИО" WHERE "ID_студента" = 100009;
-- сеанс 2, до второго UPDATE сеанса 1
BEGIN;
SET deadlock_timeout = '1s';
UPDATE "Студент" SET "ФИО" = "ФИО" WHERE "ID_студента" = 1;
-- сеанс 1
UPDATE "Студент" SET "ФИО" = "ФИО" WHERE "ID_студента" = 1;
-- сеанс 2
UPDATE "Студент" SET "ФИО" = "ФИО" WHERE "ID_студента" = 100009;
-- один сеанс получает «обнаружена взаимоблокировка», второй выполняет COMMIT.
