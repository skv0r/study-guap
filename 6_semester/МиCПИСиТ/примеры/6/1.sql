-- лр-6
-- хранимые процедуры и функции

-- 1) вставка с пополнением справочников
-- если города / района / улицы / связи улица-район-город / серии дома нет,
-- процедура создаёт их автоматически, после чего добавляет квартиру

create or replace procedure pr_add_apartment(
  p_city_name varchar,
  p_district_name varchar,
  p_street_name varchar,
  p_house_num varchar,
  p_apt_num varchar,
  p_rooms_count smallint,
  p_area numeric(8,2),
  p_evaluated_price numeric(12,2),
  p_is_for_sale boolean default true,
  p_series_code varchar default null
)
language plpgsql
as $$
declare
  v_city_id uuid;
  v_district_id uuid;
  v_street_id uuid;
  v_series_id uuid;
begin

  -- город
  select city_id
  into v_city_id
  from city
  where name = p_city_name;

  if v_city_id is null then
    insert into city(name)
    values (p_city_name)
    returning city_id into v_city_id;
  end if;

  -- район
  select district_id
  into v_district_id
  from district
  where city_id = v_city_id
    and name = p_district_name;

  if v_district_id is null then
    insert into district(city_id, name)
    values (v_city_id, p_district_name)
    returning district_id into v_district_id;
  end if;

  -- улица
  select street_id
  into v_street_id
  from street
  where city_id = v_city_id
    and name = p_street_name;

  if v_street_id is null then
    insert into street(city_id, name)
    values (v_city_id, p_street_name)
    returning street_id into v_street_id;
  end if; 

  -- связка улица-район-город
  if not exists (
    select *
    from street_district_city
    where street_id = v_street_id
      and district_id = v_district_id
      and city_id = v_city_id
  ) then
    insert into street_district_city(street_id, district_id, city_id)
    values (v_street_id, v_district_id, v_city_id);
  end if;

  -- строительная серия
  if p_series_code is not null then
    select series_id
    into v_series_id
    from building_series
    where code = p_series_code;

    if v_series_id is null then
      insert into building_series(code)
      values (p_series_code)
      returning series_id into v_series_id;
    end if;
  else
    v_series_id := null;
  end if;

  -- проверка, что квартира уже зарегистрирована
  if exists (
    select *
    from apartment
    where city_id = v_city_id
      and district_id = v_district_id
      and street_id = v_street_id
      and house_num = p_house_num
      and apt_num = p_apt_num
  ) then
    raise notice 'Квартира по адресу уже существует';
    return;
  end if;

  -- добавлене квартиры
  insert into apartment(
    city_id,
    district_id,
    street_id,
    house_num,
    apt_num,
    rooms_count,
    area,
    evaluated_price,
    is_for_sale,
    series_id
  )
  values (
    v_city_id,
    v_district_id,
    v_street_id,
    p_house_num,
    p_apt_num,
    p_rooms_count,
    p_area,
    p_evaluated_price,
    p_is_for_sale,
    v_series_id
  );
end;
$$;


-- пример вызова процедуры вставки
call pr_add_apartment(
  'Казань',					--  p_city_name varchar,
  'Вахитовский',			--  p_district_name varchar,
  'Улица Баумана',			--  p_street_name varchar,
  '1',						--  p_house_num varchar,
  '15',						--  p_apt_num varchar,
  1::smallint,				--  p_rooms_count smallint,
  40.00,					--  p_area numeric(8,2),
  5100000.00,				--  p_evaluated_price numeric(12,2),
  true,						--  p_is_for_sale boolean default true,
  'П-44'					--  p_series_code varchar default null							
);		


select
  c.name as "Город",
  d.name as "Район",
  s.name as "Улица",
  a.house_num as "Дом",
  a.apt_num as "Квартира",
  a.rooms_count as "Комнат",
  a.area as "Площадь",
  a.evaluated_price as "Цена",
  bs.code as "Серия"
from apartment a
join city c
  on c.city_id = a.city_id
join district d
  on d.district_id = a.district_id
join street s
  on s.street_id = a.street_id
left join building_series bs
  on bs.series_id = a.series_id
where c.name = 'Казань'
  and d.name = 'Вахитовский'
  and s.name = 'Улица Баумана'
  and a.house_num = '1'
  and a.apt_num = '15';


