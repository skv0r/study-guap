-- 7) скалярная функция
-- цена за квадратный метр для квартиры

create or replace function fn_price_per_meter(
  p_apartment_id uuid
)
returns numeric(12,2)
language plpgsql
as $$
declare
  v_price_per_meter numeric(12,2);
begin
  select round(evaluated_price / area, 2)
  into v_price_per_meter
  from apartment
  where apartment_id = p_apartment_id;

  return v_price_per_meter;
end;
$$;


-- запрос со скалярной функцией
-- квартиры Санкт-Петербурга, у которых цена за м2 выше средней по городу
select
  c.name as "Город",
  d.name as "Район",
  s.name as "Улица",
  a.house_num as "Дом",
  a.apt_num as "Квартира",
  fn_price_per_meter(a.apartment_id) as "Цена за м2"
from apartment a
join city c
  on c.city_id = a.city_id
join district d
  on d.district_id = a.district_id
join street s
  on s.street_id = a.street_id
where c.name = 'Санкт-Петербург'
  and a.is_for_sale = true
  and fn_price_per_meter(a.apartment_id) > (
    select avg(fn_price_per_meter(a2.apartment_id))
    from apartment a2
    join city c2
      on c2.city_id = a2.city_id
    where c2.name = 'Санкт-Петербург'
      and a2.is_for_sale = true
  )
order by "Цена за м2" desc;