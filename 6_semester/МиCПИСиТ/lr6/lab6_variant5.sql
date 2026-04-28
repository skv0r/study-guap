-- хранимые процедуры и функции
-- 1) вставка с пополнением справочников


create or replace procedure pr_add_student_with_group(
  p_fio text,
  p_nomer_grupy text,
  p_nomer_fakulteta smallint,
  p_vuz text default 'ГУАП'
)
language plpgsql
as $$
declare
  v_id_fakulteta integer;
  v_id_grupy integer;
begin
  if btrim(p_fio) = '' then
    raise exception 'ФИО не может быть пустым';
  end if;

  -- факультет вуза
  select f."ID_факультета"
  into v_id_fakulteta
  from "Факультет" f
  join "Вуз" vz on vz."ID_вуза" = f."ID_вуза"
  where btrim(vz."Название") = btrim(p_vuz)
    and f."Номер_факультета" = p_nomer_fakulteta;

  if v_id_fakulteta is null then
    raise exception 'Факультет с номером % для вуза «%» не найден', p_nomer_fakulteta, p_vuz;
  end if;

  -- группа; при отсутствии создаём
  select g."ID_группы"
  into v_id_grupy
  from "Группа" g
  where g."ID_факультета" = v_id_fakulteta
    and g."Номер_группы" = btrim(p_nomer_grupy);

  if v_id_grupy is null then
    insert into "Группа" ("ID_факультета", "Номер_группы")
    values (v_id_fakulteta, btrim(p_nomer_grupy))
    returning "ID_группы" into v_id_grupy;
  end if;

  -- не дублировать студента с тем же ФИО в той же группе
  if exists (
    select *
    from "Студент" s
    where s."ID_группы" = v_id_grupy
      and s."ФИО" = btrim(p_fio)
  ) then
    raise notice 'Студент «%» уже есть в группе %', btrim(p_fio), btrim(p_nomer_grupy);
    return;
  end if;

  insert into "Студент" ("ID_группы", "ФИО")
  values (v_id_grupy, btrim(p_fio));
end;
$$;


-- call pr_add_student_with_group('Пробный Студент ЛР6','9999',1::smallint,'ГУАП');


-- 2) удаление с очисткой справочника «Группа»


create or replace procedure pr_delete_student_cleanup_group(
  p_fio text
)
language plpgsql
as $$
declare
  v_id_student integer;
  v_id_grupy integer;
begin
  select s."ID_студента", s."ID_группы"
  into v_id_student, v_id_grupy
  from "Студент" s
  where s."ФИО" = btrim(p_fio);

  if v_id_student is null then
    raise notice 'Студент «%» не найден', p_fio;
    return;
  end if;

  delete from "Выступление_студент" vs
  where vs."ID_студента" = v_id_student;

  delete from "Студент" s
  where s."ID_студента" = v_id_student;

  if not exists (
    select *
    from "Студент" s2
    where s2."ID_группы" = v_id_grupy
  ) then
    delete from "Группа" g
    where g."ID_группы" = v_id_grupy;
  end if;
end;
$$;



-- call pr_delete_student_cleanup_group('Пробный Студент ЛР6');



-- 3) каскадное удаление группы


create or replace procedure pr_delete_group_with_students(
  p_nomer_grupy text,
  p_nomer_fakulteta smallint,
  p_vuz text default 'ГУАП'
)
language plpgsql
as $$
declare
  v_id_grupy integer;
begin
  select g."ID_группы"
  into v_id_grupy
  from "Группа" g
  join "Факультет" f on f."ID_факультета" = g."ID_факультета"
  join "Вуз" vz on vz."ID_вуза" = f."ID_вуза"
  where btrim(vz."Название") = btrim(p_vuz)
    and f."Номер_факультета" = p_nomer_fakulteta
    and g."Номер_группы" = btrim(p_nomer_grupy);

  if v_id_grupy is null then
    raise notice 'Группа не найдена';
    return;
  end if;

  delete from "Выступление_студент" vs
  using "Студент" s
  where vs."ID_студента" = s."ID_студента"
    and s."ID_группы" = v_id_grupy;

  delete from "Студент" s
  where s."ID_группы" = v_id_grupy;

  delete from "Группа" g
  where g."ID_группы" = v_id_grupy;
end;
$$;



-- call pr_delete_group_with_students('9999', 1::smallint, 'ГУАП');



-- 4) скалярная функция — среднее число участий на студента (все студенты вуза, в т.ч. с нулём участий)


create or replace function fn_avg_participations_per_student(
  p_vuz text default 'ГУАП'
)
returns numeric(12,2)
language plpgsql
stable
as $$
declare
  v_avg numeric(12,2);
begin
  select round(avg(coalesce(t.cnt, 0::bigint))::numeric, 2)
  into v_avg
  from (
    select count(vs."ID_выступления")::bigint as cnt
    from "Студент" s
    join "Группа" g on g."ID_группы" = s."ID_группы"
    join "Факультет" f on f."ID_факультета" = g."ID_факультета"
    join "Вуз" vz on vz."ID_вуза" = f."ID_вуза"
    left join "Выступление_студент" vs on vs."ID_студента" = s."ID_студента"
    where btrim(vz."Название") = btrim(p_vuz)
    group by s."ID_студента"
  ) t;

  return coalesce(v_avg, 0::numeric(12,2));
end;
$$;


select fn_avg_participations_per_student('ГУАП') as "Среднее_число_участий_на_студента";


-- 5) статистика во временной таблице по факультетам


create or replace procedure pr_fill_faculty_stats()
language plpgsql
as $$
begin
  drop table if exists tmp_faculty_stats;

  create temp table tmp_faculty_stats (
    nomer_fakulteta smallint not null,
    nazvanie_fakulteta varchar(200),
    grupp int not null,
    studentov int not null,
    uchastiy_v_vystupleniyah bigint not null,
    srednyaya_dlina_temy numeric(10, 2)
  ) on commit preserve rows;

  insert into tmp_faculty_stats (
    nomer_fakulteta,
    nazvanie_fakulteta,
    grupp,
    studentov,
    uchastiy_v_vystupleniyah,
    srednyaya_dlina_temy
  )
  select
    f."Номер_факультета",
    f."Название",

    (
      select count(*)::int
      from "Группа" g
      where g."ID_факультета" = f."ID_факультета"
    ),

    (
      select count(*)::int
      from "Студент" s
      join "Группа" g on g."ID_группы" = s."ID_группы"
      where g."ID_факультета" = f."ID_факультета"
    ),

    coalesce(
      (
        select count(*)
        from "Выступление_студент" vs
        join "Студент" s on s."ID_студента" = vs."ID_студента"
        join "Группа" g on g."ID_группы" = s."ID_группы"
        where g."ID_факультета" = f."ID_факультета"
      ),
      0
    ),

    coalesce(
      (
        select round(avg(char_length(v."Название_темы"))::numeric, 2)
        from "Выступление_студент" vs
        join "Студент" s on s."ID_студента" = vs."ID_студента"
        join "Группа" g on g."ID_группы" = s."ID_группы"
        join "Выступление" v on v."ID_выступления" = vs."ID_выступления"
        where g."ID_факультета" = f."ID_факультета"
      ),
      0::numeric
    )

  from "Факультет" f
  join "Вуз" vz on vz."ID_вуза" = f."ID_вуза"
  where btrim(vz."Название") = 'ГУАП'
  order by f."Номер_факультета";
end;
$$;


call pr_fill_faculty_stats();

select *
from tmp_faculty_stats
order by nomer_fakulteta;



-- 6) управляющие конструкции - цикл по конференциям, заполнение отчёта


create or replace procedure pr_demo_conference_participation_loop(
  p_vuz text default 'ГУАП'
)
language plpgsql
as $$
declare
  r record;
  v_cnt bigint;
begin
  drop table if exists tmp_conf_loop;

  create temp table tmp_conf_loop (
    konferenciya varchar(200) not null,
    uchastiy bigint not null
  ) on commit preserve rows;

  for r in
    select k."Название" as nazvanie
    from "Конференция" k
    order by k."Название"
  loop
    select count(*)::bigint
    into v_cnt
    from "Выступление_студент" vs
    join "Студент" s on s."ID_студента" = vs."ID_студента"
    join "Группа" g on g."ID_группы" = s."ID_группы"
    join "Факультет" f on f."ID_факультета" = g."ID_факультета"
    join "Вуз" vz on vz."ID_вуза" = f."ID_вуза"
    join "Выступление" v on v."ID_выступления" = vs."ID_выступления"
    join "Программа_конференции" pc on pc."ID_программы" = v."ID_программы"
    join "Конференция" k2 on k2."ID_конференции" = pc."ID_конференции"
    where btrim(vz."Название") = btrim(p_vuz)
      and k2."Название" = r.nazvanie;

    insert into tmp_conf_loop (konferenciya, uchastiy)
    values (r.nazvanie, coalesce(v_cnt, 0));
  end loop;
end;
$$;


call pr_demo_conference_participation_loop('ГУАП');

select *
from tmp_conf_loop
order by konferenciya;



-- 7) табличная функция — студенты конференции (ФИО и номер группы)


create or replace function fn_students_by_conference(
  p_konferenciya text
)
returns table (
  fio varchar(150),
  nomer_grupy varchar(20)
)
language sql
stable
as $$
  select distinct
    s."ФИО",
    g."Номер_группы"
  from "Студент" s
  join "Группа" g on g."ID_группы" = s."ID_группы"
  join "Выступление_студент" vs on vs."ID_студента" = s."ID_студента"
  join "Выступление" v on v."ID_выступления" = vs."ID_выступления"
  join "Программа_конференции" pc on pc."ID_программы" = v."ID_программы"
  join "Конференция" k on k."ID_конференции" = pc."ID_конференции"
  where k."Название" = btrim(p_konferenciya)
  order by s."ФИО";
$$;


select *
from fn_students_by_conference('Информатика')
limit 20;


-- 8) ситуация с CASE — категории активности студентов по числу участий


create or replace procedure pr_fill_student_activity_case(
  p_vuz text default 'ГУАП'
)
language plpgsql
as $$
begin
  drop table if exists tmp_student_activity_case;

  create temp table tmp_student_activity_case (
    fio varchar(150) not null,
    nomer_grupy varchar(20) not null,
    kolichestvo_uchastiy bigint not null,
    uroven_aktivnosti varchar(30) not null
  ) on commit preserve rows;

  insert into tmp_student_activity_case (
    fio,
    nomer_grupy,
    kolichestvo_uchastiy,
    uroven_aktivnosti
  )
  select
    s."ФИО",
    g."Номер_группы",
    count(vs."ID_выступления")::bigint as kolichestvo_uchastiy,
    case
      when count(vs."ID_выступления") = 0 then 'Нет участий'
      when count(vs."ID_выступления") between 1 and 2 then 'Низкая'
      when count(vs."ID_выступления") between 3 and 4 then 'Средняя'
      else 'Высокая'
    end as uroven_aktivnosti
  from "Студент" s
  join "Группа" g on g."ID_группы" = s."ID_группы"
  join "Факультет" f on f."ID_факультета" = g."ID_факультета"
  join "Вуз" vz on vz."ID_вуза" = f."ID_вуза"
  left join "Выступление_студент" vs on vs."ID_студента" = s."ID_студента"
  where btrim(vz."Название") = btrim(p_vuz)
  group by s."ID_студента", s."ФИО", g."Номер_группы"
  order by s."ФИО";
end;
$$;


call pr_fill_student_activity_case('ГУАП');

select *
from tmp_student_activity_case
order by fio;
