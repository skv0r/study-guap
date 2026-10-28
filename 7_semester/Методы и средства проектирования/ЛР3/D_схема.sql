-- Блок изменения схемы. Столбец появляется у издателя.
-- Копия на узле его не получает: в узле остаются прежние три столбца.

ALTER TABLE pub."Студент" ADD COLUMN "Группа" varchar(20);

UPDATE pub."Студент"
   SET "Группа" = '4321'
 WHERE "ID" = 1;

SELECT "ID", "ФИО", "Факультет", "Группа" FROM pub."Студент";

SELECT column_name
  FROM information_schema.columns
 WHERE table_schema = 'sub' AND table_name = 'Студент'
 ORDER BY ordinal_position;
