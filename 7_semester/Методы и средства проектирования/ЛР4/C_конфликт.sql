-- Блок C. Одна и та же конференция изменена в центре и на узле.
-- Сверка поднимает конфликт. После выравнивания названия сверка проходит.

INSERT INTO office."Конференция" ("ID", "Название") VALUES (7, 'Цифровой вуз');
INSERT INTO node4."Конференция" ("ID", "Название") VALUES (7, 'Цифровой вуз');

UPDATE office."Конференция" SET "Название" = 'Цифровой вуз, пленарное' WHERE "ID" = 7;
UPDATE node4."Конференция" SET "Название" = 'Цифровой вуз, секция' WHERE "ID" = 7;

SELECT 'центр' AS сторона, "Название" FROM office."Конференция" WHERE "ID" = 7
UNION ALL
SELECT 'узел_4', "Название" FROM node4."Конференция" WHERE "ID" = 7;

CREATE OR REPLACE FUNCTION office.check_conflict()
RETURNS void
LANGUAGE plpgsql
AS $$
DECLARE
  a text;
  b text;
BEGIN
  SELECT "Название" INTO a FROM office."Конференция" WHERE "ID" = 7;
  SELECT "Название" INTO b FROM node4."Конференция" WHERE "ID" = 7;
  IF a IS DISTINCT FROM b THEN
    RAISE EXCEPTION 'конфликт репликации: центр [%], узел [%]', a, b;
  END IF;
END;
$$;

SELECT office.check_conflict();
