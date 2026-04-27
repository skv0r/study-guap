-- ж) Студенты четвёртого факультета, не выступавших на конференциях (ни в одном выступлении)
-- В варианте 5 данные лр2 привязаны к 4-му факультету через группы 4550/4551; плюс отбор по номеру факультета 4.
-- Так запрос не «глушится», если в таблице «Вуз» другое написание или номера факультетов отличаются.

select distinct
  s."ФИО" as "ФИО"
from "Студент" as s
join "Группа" as g on g."ID_группы" = s."ID_группы"
join "Факультет" as f on f."ID_факультета" = g."ID_факультета"
where (
  f."Номер_факультета" = 4
  or g."Номер_группы" in ('4550', '4551')
)
  and not exists (
    select 1
    from "Выступление_студент" as vs
    where vs."ID_студента" = s."ID_студента"
  )
order by s."ФИО";


-- з) Студенты, выступившие на всех конференциях (по каждой конференции есть хотя бы одно участие)

select distinct
  s."ФИО" as "ФИО"
from "Студент" as s
where not exists (
  select 1
  from "Конференция" as k
  where not exists (
    select 1
    from "Выступление_студент" as vs
    join "Выступление" as v on v."ID_выступления" = vs."ID_выступления"
    join "Программа_конференции" as pc on pc."ID_программы" = v."ID_программы"
    where vs."ID_студента" = s."ID_студента"
      and pc."ID_конференции" = k."ID_конференции"
  )
)
order by s."ФИО";


-- и) Пары студентов, всегда выступающие вместе (множества выступлений совпадают; в паре — разные студенты)

select distinct
  s1."ФИО" as "Студент_1",
  s2."ФИО" as "Студент_2"
from "Студент" as s1
join "Студент" as s2 on s2."ID_студента" > s1."ID_студента"
where exists (
  select 1
  from "Выступление_студент" as vs0
  where vs0."ID_студента" = s1."ID_студента"
)
and exists (
  select 1
  from "Выступление_студент" as vs0b
  where vs0b."ID_студента" = s2."ID_студента"
)
and not exists (
  select 1
  from "Выступление_студент" as vs1
  where vs1."ID_студента" = s1."ID_студента"
    and not exists (
      select 1
      from "Выступление_студент" as vs1b
      where vs1b."ID_выступления" = vs1."ID_выступления"
        and vs1b."ID_студента" = s2."ID_студента"
    )
)
and not exists (
  select 1
  from "Выступление_студент" as vs2
  where vs2."ID_студента" = s2."ID_студента"
    and not exists (
      select 1
      from "Выступление_студент" as vs2b
      where vs2b."ID_выступления" = vs2."ID_выступления"
        and vs2b."ID_студента" = s1."ID_студента"
    )
)
order by s1."ФИО", s2."ФИО";


-- Подзапросы в операторах манипулирования данными (временная таблица)

drop table if exists tmp_student_participation_stats;

create temp table tmp_student_participation_stats (
  "ID_студента" integer not null primary key,
  "ФИО" varchar(150) not null,
  "Участий" integer not null
);

-- Студенты с числом участий ниже среднего по всем, у кого есть хотя бы одно участие
insert into tmp_student_participation_stats ("ID_студента", "ФИО", "Участий")
select
  s."ID_студента",
  s."ФИО",
  cnt.n
from (
  select
    vs."ID_студента",
    count(*)::integer as n
  from "Выступление_студент" as vs
  group by vs."ID_студента"
) as cnt
join "Студент" as s on s."ID_студента" = cnt."ID_студента"
where cnt.n < (
  select avg(c)::numeric
  from (
    select count(*)::numeric as c
    from "Выступление_студент" as vs2
    group by vs2."ID_студента"
  ) as t
);

-- студенты с факультета, по которому нет участников конференции «Информатика»
update tmp_student_participation_stats as t
set "ФИО" = t."ФИО" || ' (факультет без выступлений на «Информатике»)'
where t."ID_студента" in (
  select s."ID_студента"
  from "Студент" as s
  join "Группа" as g on g."ID_группы" = s."ID_группы"
  where g."ID_факультета" in (
    select f."ID_факультета"
    from "Факультет" as f
    join "Вуз" as vz on vz."ID_вуза" = f."ID_вуза"
    where vz."Название" = 'ГУАП'
      and not exists (
        select 1
        from "Студент" as st2
        join "Группа" as g2 on g2."ID_группы" = st2."ID_группы"
        join "Выступление_студент" as vs on vs."ID_студента" = st2."ID_студента"
        join "Выступление" as v on v."ID_выступления" = vs."ID_выступления"
        join "Программа_конференции" as pc on pc."ID_программы" = v."ID_программы"
        join "Конференция" as k on k."ID_конференции" = pc."ID_конференции"
        where g2."ID_факультета" = f."ID_факультета"
          and k."Название" = 'Информатика'
      )
  )
);

-- Удаление строк по студентам, не выступавшим на «Робототехника и ИИ»
delete from tmp_student_participation_stats as t
where t."ID_студента" in (
  select s."ID_студента"
  from "Студент" as s
  where not exists (
    select 1
    from "Выступление_студент" as vs
    join "Выступление" as v on v."ID_выступления" = vs."ID_выступления"
    join "Программа_конференции" as pc on pc."ID_программы" = v."ID_программы"
    join "Конференция" as k on k."ID_конференции" = pc."ID_конференции"
    where vs."ID_студента" = s."ID_студента"
      and k."Название" = 'Робототехника и ИИ'
  )
);

select *
from tmp_student_participation_stats
order by "Участий", "ФИО";

drop table if exists tmp_student_participation_stats;


-- теоретико-множественное пересечение A ∩ B через EXISTS
-- A: ФИО, выступавшие на «Информатика»; B: ФИО на «Робототехника и ИИ»
-- (в lr4 для мультимножеств использовался intersect all и группировка; при уникальных ФИО в A/B эквивалентно intersect)

with "A" as (
  select distinct s."ФИО"
  from "Студент" as s
  join "Выступление_студент" as vs on vs."ID_студента" = s."ID_студента"
  join "Выступление" as v on v."ID_выступления" = vs."ID_выступления"
  join "Программа_конференции" as pc on pc."ID_программы" = v."ID_программы"
  join "Конференция" as k on k."ID_конференции" = pc."ID_конференции"
  where k."Название" = 'Информатика'
),
"B" as (
  select distinct s."ФИО"
  from "Студент" as s
  join "Выступление_студент" as vs on vs."ID_студента" = s."ID_студента"
  join "Выступление" as v on v."ID_выступления" = vs."ID_выступления"
  join "Программа_конференции" as pc on pc."ID_программы" = v."ID_программы"
  join "Конференция" as k on k."ID_конференции" = pc."ID_конференции"
  where k."Название" = 'Робототехника и ИИ'
)
select distinct
  a."ФИО" as "ФИО"
from "A" as a
where exists (
  select 1
  from "B" as b
  where b."ФИО" = a."ФИО"
);

-- Разность A \ B через NOT EXISTS (аналог except; в lr4 не была основной демонстрацией)

-- select distinct a."ФИО" from "A" as a
-- where not exists (select 1 from "B" as b where b."ФИО" = a."ФИО");


-- Различие intersect / except и exists / not exists при наличии NULL в сравниваемом столбце
-- Искусственные множества кодов: A и B содержат NULL и общее непустое значение

with "A" as (
select * from (values (null::varchar), ('ИИСиТ')) as a(code)
),
  "B" as (
select * from (values (null::varchar), ('ИРТ')) as b(code)
)
  select a.code as "intersect"
  from "A" as a
  intersect
  select b.code
  from "B" as b;


with "A" as (
select * from (values (null::varchar), ('ИИСиТ')) as a(code)
),
  "B" as (
select * from (values (null::varchar), ('ИРТ')) as b(code)
)
  select a.code as "exists_пересечение"
  from "A" as a
  where exists (select 1 from "B" as b where b.code = a.code);


with "A" as (
select * from (values (null::varchar), ('ИИСиТ')) as a(code)
),
  "B" as (
select * from (values (null::varchar), ('ИРТ')) as b(code)
)
  select a.code as "except"
  from "A" as a
  except
  select b.code
  from "B" as b;


with "A" as (
select * from (values (null::varchar), ('ИИСиТ')) as a(code)
),
  "B" as (
select * from (values (null::varchar), ('ИРТ')) as b(code)
)
  select a.code as "not_exists_разность"
  from "A" as a
  where not exists (select 1 from "B" as b where b.code = a.code);
