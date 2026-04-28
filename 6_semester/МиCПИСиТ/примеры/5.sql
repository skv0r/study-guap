--ж) улицы, продолжительность которых ограничивается только одним районом;
select distinct c.name as "Город", s.name as "Улица"
from street s
join city c
  on s.city_id = c.city_id
join street_district_city sdc
  on sdc.street_id = s.street_id
 and sdc.city_id = s.city_id
where not exists ( -- для улицы не должно существовать второй связи с другим районом
  select *
  from street_district_city sdc2
  where sdc2.street_id = sdc.street_id
    and sdc2.city_id = sdc.city_id
    and sdc2.district_id <> sdc.district_id
)


--з) районы, в которых не продаются однокомнатные квартиры;
select c.name as "Город", d.name as "Район"
from district d
join city c
  on d.city_id = c.city_id
where not exists ( -- исключаем однокомнатные квартиры
  select *
  from apartment a
  where a.district_id = d.district_id
    and a.is_for_sale = true
    and a.rooms_count = 1
)


--и) районы, в которых продаются квартиры всех строительных серий;
select c.name as "Город", d.name as "Район"
from district d
join city c
  on d.city_id = c.city_id
where not exists ( -- не должно существовать серий, которых нет среди продаваемых квартир района
  select *
  from building_series bs
  where not exists (
    select *
    from apartment a
    where a.district_id = d.district_id
      and a.is_for_sale = true
      and a.series_id = bs.series_id
  )
)


-- использование подзапросов в операторах манипулирования данными
drop table if exists tmp_sale_apartment_stats;

create temp table tmp_sale_apartment_stats (
  apartment_id uuid primary key,
  district_name varchar,
  price_per_meter numeric(12,2)
);

-- добавление продаваемых квартир, у которых цена за м2 ниже средней по БД
insert into tmp_sale_apartment_stats(apartment_id, district_name, price_per_meter)
select a.apartment_id,
       d.name,
       a.evaluated_price / a.area
from apartment a
join district d
  on a.district_id = d.district_id
where a.is_for_sale = true
  and (a.evaluated_price / a.area) < (
    select avg(a2.evaluated_price / a2.area)	-- среднее цена за м2 по бд
    from apartment a2
    where a2.is_for_sale = true
  );

-- отметка квартир, которые находятся в районах без продаваемых однокомнатных квартир
update tmp_sale_apartment_stats
set district_name = district_name || ' (без 1-комн. в продаже)'
where apartment_id in ( 						-- Получение квартир, где надо поставить отметку
  select a.apartment_id
  from apartment a
  where a.district_id in (						-- Получение районов, без одномкомнаятных квартир
    select d.district_id
    from district d
    where not exists (							-- Проверка, что в районе нет однокомнатных квартир
      select 1
      from apartment a1
      where a1.district_id = d.district_id
        and a1.is_for_sale = true
        and a1.rooms_count = 1
    )
  )
);

-- удаление квартир без указанной строительной серии
delete from tmp_sale_apartment_stats
where apartment_id in (
  select a.apartment_id
  from apartment a
  where a.series_id is null
);


select *
from tmp_sale_apartment_stats;

drop table if exists tmp_sale_apartment_stats;

-- заросы из лр4 чере exists/not exists
with A as (  -- Московский
  select s.name
  from apartment a
  join street s
    on s.street_id = a.street_id
  join city c
    on c.city_id = a.city_id
   and c.name = 'Санкт-Петербург'
  join district d
    on d.district_id = a.district_id
  where d.name = 'Московский'
),
B as (  -- Центральный
  select s.name
  from apartment a
  join street s
    on s.street_id = a.street_id
  join city c
    on c.city_id = a.city_id
   and c.name = 'Санкт-Петербург'
  join district d
    on d.district_id = a.district_id
  where d.name = 'Центральный'
)
select distinct a.name as "Улица"
from A a

-- intersect
--where exists (
--  select 1
--  from B b
--  where b.name = a.name
--);

-- except
where not exists (
  select 1
  from B b
  where b.name = a.name
);




-- различие между intersect / except и exists / not exists при наличии null-значений
-- A = строительные серии квартир Санкт-Петербурга
-- B = строительные серии двухкомнатных квартир

with A as (
  select distinct bs.code as series_code
  from apartment a
  left join building_series bs
    on bs.series_id = a.series_id
  join city c
    on c.city_id = a.city_id
  where c.name = 'Санкт-Петербург'
),
B as (
  select distinct bs.code as series_code
  from apartment a
  left join building_series bs
    on bs.series_id = a.series_id
  where a.rooms_count = 2
)

-- intersect считает null обычным значением
select series_code
from A
intersect
select series_code
from B



with A as (
  select distinct bs.code as series_code
  from apartment a
  left join building_series bs
    on bs.series_id = a.series_id
  join city c
    on c.city_id = a.city_id
  where c.name = 'Санкт-Петербург'
),
B as (
  select distinct bs.code as series_code
  from apartment a
  left join building_series bs
    on bs.series_id = a.series_id
  where a.rooms_count = 2
)

-- exists использует "=" , поэтому null = null не дает true
select a.series_code
from A a
where exists (
  select 1
  from B b
  where b.series_code = a.series_code
)


with A as (
  select distinct bs.code as series_code
  from apartment a
  left join building_series bs
    on bs.series_id = a.series_id
  join city c
    on c.city_id = a.city_id
  where c.name = 'Санкт-Петербург'
),
B as (
  select distinct bs.code as series_code
  from apartment a
  left join building_series bs
    on bs.series_id = a.series_id
  where a.rooms_count = 2
)

-- except удаляет null, если он есть в обеих выборках
select series_code
from A
except
select series_code
from B




with A as (
  select distinct bs.code as series_code
  from apartment a
  left join building_series bs
    on bs.series_id = a.series_id
  join city c
    on c.city_id = a.city_id
  where c.name = 'Санкт-Петербург'
),
B as (
  select distinct bs.code as series_code
  from apartment a
  left join building_series bs
    on bs.series_id = a.series_id
  where a.rooms_count = 2
)

-- not exists оставляет null, так как сравнение с null возвращает unknown
select a.series_code
from A a
where not exists (
  select 1
  from B b
  where b.series_code = a.series_code
)













