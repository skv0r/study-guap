-- Блок C, разрешение. На узле оставляется значение центра, сверка проходит.

UPDATE node4."Конференция"
   SET "Название" = (SELECT "Название" FROM office."Конференция" WHERE "ID" = 7)
 WHERE "ID" = 7;

SELECT office.check_conflict();

SELECT 'центр' AS сторона, "Название" FROM office."Конференция" WHERE "ID" = 7
UNION ALL
SELECT 'узел_4', "Название" FROM node4."Конференция" WHERE "ID" = 7;
