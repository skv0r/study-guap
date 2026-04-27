-- 8) табличная функция
-- продаваемые квартиры заданного района

create or replace function fn_sale_apartments_by_district(
  p_city_name varchar,
  p_district_name varchar
)
returns table (
  apartment_id uuid,
  full_address text,
  rooms_count smallint,
  area numeric(8,2),
  evaluated_price numeric(12,2),
  price_per_meter numeric(12,2),
  series_code varchar
)
language sql
as $$
  select
    a.apartment_id,
    c.name || ', ' || d.name || ', ' || s.name || ', д. ' || a.house_num || ', кв. ' || a.apt_num as full_address,
    a.rooms_count,
    a.area,
    a.evaluated_price,
    round(a.evaluated_price / a.area, 2) as price_per_meter,
    bs.code as series_code
  from apartment a
  join city c
    on c.city_id = a.city_id
  join district d
    on d.district_id = a.district_id
  join street s
    on s.street_id = a.street_id
  left join building_series bs
    on bs.series_id = a.series_id
  where c.name = p_city_name
    and d.name = p_district_name
    and a.is_for_sale = true
  order by a.evaluated_price desc;
$$;


-- запрос с табличной функцией
select *
from fn_sale_apartments_by_district('Санкт-Петербург', 'Московский');