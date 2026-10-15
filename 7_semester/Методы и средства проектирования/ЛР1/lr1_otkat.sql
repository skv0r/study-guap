-- Откат индексов лабораторной 1. Данные не удаляются.
-- После этого запросы снова идут без пяти дополнительных индексов,
-- и lr1_demo.sql можно выполнить заново: сначала «до», потом «после».

DROP INDEX IF EXISTS "Студент_ФИО";
DROP INDEX IF EXISTS "Студент_ID_группы";
DROP INDEX IF EXISTS "Выступление_ID_программы";
DROP INDEX IF EXISTS "Выступление_студент_ID_студента";
DROP INDEX IF EXISTS "Группа_Номер_группы";

SELECT indexname AS "Оставшиеся индексы"
FROM pg_indexes
WHERE schemaname = 'public'
ORDER BY tablename, indexname;
