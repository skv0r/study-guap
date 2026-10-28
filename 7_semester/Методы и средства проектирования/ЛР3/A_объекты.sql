-- Блок A. Объекты издателя и узла в одной базе.
-- В Supabase этот блок вставляется целиком и выполняется отдельно от следующих.

DROP SCHEMA IF EXISTS sub CASCADE;
DROP SCHEMA IF EXISTS pub CASCADE;

CREATE SCHEMA pub;
CREATE SCHEMA sub;

CREATE TABLE pub."Студент" (
  "ID" integer PRIMARY KEY,
  "ФИО" varchar(100) NOT NULL,
  "Факультет" smallint NOT NULL
);

CREATE TABLE sub."Студент" (
  LIKE pub."Студент" INCLUDING ALL
);

CREATE OR REPLACE FUNCTION pub.copy_student()
RETURNS trigger
LANGUAGE plpgsql
AS $$
BEGIN
  IF TG_OP = 'INSERT' THEN
    INSERT INTO sub."Студент" VALUES (NEW."ID", NEW."ФИО", NEW."Факультет");
  ELSIF TG_OP = 'UPDATE' THEN
    UPDATE sub."Студент"
       SET "ФИО" = NEW."ФИО", "Факультет" = NEW."Факультет"
     WHERE "ID" = NEW."ID";
  ELSIF TG_OP = 'DELETE' THEN
    DELETE FROM sub."Студент" WHERE "ID" = OLD."ID";
  END IF;
  RETURN COALESCE(NEW, OLD);
END;
$$;

CREATE TRIGGER repl
AFTER INSERT OR UPDATE OR DELETE ON pub."Студент"
FOR EACH ROW EXECUTE FUNCTION pub.copy_student();

CREATE PUBLICATION pub_vuz FOR TABLE pub."Студент";
