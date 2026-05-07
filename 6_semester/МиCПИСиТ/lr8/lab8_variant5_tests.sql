-- ============================================================
-- ПУНКТ 1 / РИСУНОК 1
-- Вызов pr_cursor_iso_conference_stats + результат tmp_cursor_conf_stats
-- ============================================================

call pr_cursor_iso_conference_stats();

select
  t.id_konferencii,
  t.nazvanie_konferencii,
  t.kolichestvo_uchastiy
from tmp_cursor_conf_stats t
order by t.nazvanie_konferencii;


-- ============================================================
-- ПУНКТ 2 / РИСУНОК 2
-- Вызов pr_cursor_snapshot_like + журнал шагов tmp_cursor_action_log
-- ============================================================

call pr_cursor_snapshot_like();

select
  l.step_no,
  l.action_text,
  l.action_time
from tmp_cursor_action_log l
order by l.step_no;


-- ============================================================
-- ПУНКТ 3 / РИСУНОК 3
-- Исходные данные tmp_cursor_workset перед обновлением
-- ============================================================

-- Создаем рабочий набор с нуля, не меняя основную таблицу "Студент".
drop table if exists tmp_cursor_workset;

create temp table tmp_cursor_workset (
  "ID_студента" integer primary key,
  "ФИО" varchar(150) not null,
  "Номер_группы" varchar(20) not null,
  "Стипендия" numeric(12,2) not null
) on commit preserve rows;

insert into tmp_cursor_workset ("ID_студента", "ФИО", "Номер_группы", "Стипендия")
select
  s."ID_студента",
  s."ФИО",
  g."Номер_группы",
  1000::numeric(12,2) as "Стипендия"
from "Студент" s
join "Группа" g on g."ID_группы" = s."ID_группы";

select
  w."ID_студента",
  w."ФИО",
  w."Номер_группы",
  w."Стипендия"
from tmp_cursor_workset w
where w."Номер_группы" = '4226'
order by w."ФИО";


-- ============================================================
-- ПУНКТ 4 / РИСУНОК 4
-- UPDATE WHERE CURRENT OF (повышение стипендии в рабочем наборе)
-- ============================================================

call pr_cursor_update_scholarship_where_current_of('4226', 15);

select
  w."ID_студента",
  w."ФИО",
  w."Номер_группы",
  w."Стипендия"
from tmp_cursor_workset w
where w."Номер_группы" = '4226'
order by w."ФИО";


-- ============================================================
-- ПУНКТ 5 / РИСУНОК 5
-- Запуск DELETE WHERE CURRENT OF (демонстрация команды)
-- ============================================================

-- Для наглядности добавляем тестового студента без участия в "Информатика".
insert into "Студент" ("ID_группы", "ФИО")
select g."ID_группы", 'Тест ЛР8 Удаление'
from "Группа" g
where g."Номер_группы" = '4551'
order by g."ID_группы"
limit 1
on conflict do nothing;

select
  'До вызова процедуры' as "Этап",
  s."ID_студента",
  s."ФИО"
from "Студент" s
where s."ФИО" = 'Тест ЛР8 Удаление';

call pr_cursor_delete_inactive_where_current_of('Информатика');

select
  'После вызова процедуры' as "Этап",
  case
    when exists (
      select 1
      from "Студент" s
      where s."ФИО" = 'Тест ЛР8 Удаление'
    ) then 'Тестовый студент остался'
    else 'Тестовый студент удалён'
  end as "Статус";


-- ============================================================
-- ПУНКТ 6 / РИСУНОК 6
-- Состав таблицы "Студент" после демонстрации удаления
-- ============================================================

select
  s."ID_студента",
  s."ФИО",
  g."Номер_группы"
from "Студент" s
join "Группа" g on g."ID_группы" = s."ID_группы"
order by s."ФИО";
