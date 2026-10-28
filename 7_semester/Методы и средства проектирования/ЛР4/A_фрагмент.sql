-- Блок A. Центральная база и узел факультета 4.
-- На узел попадают только строки факультета 4. Таблица конференции общая: на ней будет конфликт.

DROP SCHEMA IF EXISTS node4 CASCADE;
DROP SCHEMA IF EXISTS office CASCADE;

CREATE SCHEMA office;
CREATE SCHEMA node4;

CREATE TABLE office."Студент" (
  "ID" integer PRIMARY KEY,
  "ФИО" varchar(100) NOT NULL,
  "Факультет" smallint NOT NULL
);

CREATE TABLE node4."Студент" (
  LIKE office."Студент" INCLUDING ALL
);

CREATE TABLE office."Конференция" (
  "ID" integer PRIMARY KEY,
  "Название" varchar(150) NOT NULL
);

CREATE TABLE node4."Конференция" (
  LIKE office."Конференция" INCLUDING ALL
);

CREATE OR REPLACE FUNCTION office.copy_student()
RETURNS trigger
LANGUAGE plpgsql
AS $$
BEGIN
  IF NEW."Факультет" = 4 THEN
    INSERT INTO node4."Студент" ("ID", "ФИО", "Факультет")
    VALUES (NEW."ID", NEW."ФИО", NEW."Факультет");
  END IF;
  RETURN NEW;
END;
$$;

CREATE TRIGGER frag
AFTER INSERT ON office."Студент"
FOR EACH ROW EXECUTE FUNCTION office.copy_student();
