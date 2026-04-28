-- 5) формирование статистики во временной таблице
-- для каждого района:
-- количество квартир,
-- количество квартир в продаже,
-- количество однокомнатных квартир в продаже,
-- количество продаж,
-- количество агентов,
-- средняя цена квартиры,
-- средняя цена за квадратный метр у квартир в продаже

create or replace procedure pr_fill_district_stats()
language plpgsql
as $$
begin
  drop table if exists tmp_district_stats;

  create temp table tmp_district_stats (
    city_name varchar,
    district_name varchar,
    apartments_count int,
    apartments_for_sale_count int,
    one_room_for_sale_count int,
    sales_count int,
    agents_count int,
    avg_apartment_price numeric(12,2),
    avg_price_per_meter_sale numeric(12,2)
  ) on commit preserve rows;

  insert into tmp_district_stats(
    city_name,
    district_name,
    apartments_count,
    apartments_for_sale_count,
    one_room_for_sale_count,
    sales_count,
    agents_count,
    avg_apartment_price,
    avg_price_per_meter_sale
  )
  select
    c.name as city_name,
    d.name as district_name,

    (
      select count(*)
      from apartment a
      where a.district_id = d.district_id
    ) as apartments_count,

    (
      select count(*)
      from apartment a
      where a.district_id = d.district_id
        and a.is_for_sale = true
    ) as apartments_for_sale_count,

    (
      select count(*)
      from apartment a
      where a.district_id = d.district_id
        and a.is_for_sale = true
        and a.rooms_count = 1
    ) as one_room_for_sale_count,

    (
      select count(*)
      from sale s
      join apartment a
        on a.apartment_id = s.apartment_id
      where a.district_id = d.district_id
    ) as sales_count,

    (
      select count(distinct s.agent_id)
      from sale s
      join apartment a
        on a.apartment_id = s.apartment_id
      where a.district_id = d.district_id
    ) as agents_count,

    (
      select round(avg(a.evaluated_price), 2)
      from apartment a
      where a.district_id = d.district_id
    ) as avg_apartment_price,

    (
      select round(avg(a.evaluated_price / a.area), 2)
      from apartment a
      where a.district_id = d.district_id
        and a.is_for_sale = true
    ) as avg_price_per_meter_sale

  from district d
  join city c
    on c.city_id = d.city_id
  order by c.name, d.name;
end;
$$;


-- пример вызова процедуры статистики
call pr_fill_district_stats();

select *
from tmp_district_stats
order by city_name, district_name;