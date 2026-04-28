-- 2) удаление с очисткой справочников
-- удаляется квартира; если после удаления не осталось квартир по связи
-- улица-район-город, удаляется связь.
-- если улица / район / город / серия больше нигде не используются,
-- они также удаляются

create or replace procedure pr_delete_apartment_with_cleanup(
  p_city_name varchar,
  p_district_name varchar,
  p_street_name varchar,
  p_house_num varchar,
  p_apt_num varchar
)
language plpgsql
as $$
declare
  v_apartment_id uuid;
  v_city_id uuid;
  v_district_id uuid;
  v_street_id uuid;
  v_series_id uuid;
begin
  -- заполнение всех переменных
  select
    a.apartment_id,
    a.city_id,
    a.district_id,
    a.street_id,
    a.series_id
  into
    v_apartment_id,
    v_city_id,
    v_district_id,
    v_street_id,
    v_series_id
  from apartment a
  join city c
    on c.city_id = a.city_id
  join district d
    on d.district_id = a.district_id
  join street s
    on s.street_id = a.street_id
  where c.name = p_city_name
    and d.name = p_district_name
    and s.name = p_street_name
    and a.house_num = p_house_num
    and a.apt_num = p_apt_num;

  if v_apartment_id is null then
    raise notice 'Квартира не найдена';
    return;
  end if;

  -- сначала удаляем продажи, т.к. есть ограничение restrict
  delete from sale
  where apartment_id = v_apartment_id;

  delete from apartment
  where apartment_id = v_apartment_id;

  -- если квартир с такой связкой город-район-улица больше нет, то эту связку надо удалить
  if not exists (
    select *
    from apartment
    where city_id = v_city_id
      and district_id = v_district_id
      and street_id = v_street_id
  ) then
    delete from street_district_city
    where city_id = v_city_id
      and district_id = v_district_id
      and street_id = v_street_id;
  end if;
  
 -- если картир с такой улицей больше нет, 
  if not exists (
    select *
    from apartment
    where city_id = v_city_id
      and street_id = v_street_id
  )
  and not exists (
    select *
    from street_district_city
    where city_id = v_city_id
      and street_id = v_street_id
  ) then
    delete from street
    where street_id = v_street_id;
  end if;

  if not exists (
    select *
    from apartment
    where city_id = v_city_id
      and district_id = v_district_id
  )
  and not exists (
    select *
    from street_district_city
    where city_id = v_city_id
      and district_id = v_district_id
  ) then
    delete from district
    where district_id = v_district_id;
  end if;

  if v_series_id is not null
     and not exists (
       select *
       from apartment
       where series_id = v_series_id
     ) then
    delete from building_series
    where series_id = v_series_id;
  end if;

  if not exists (
    select * from apartment where city_id = v_city_id
  )
  and not exists (
    select * from district where city_id = v_city_id
  )
  and not exists (
    select * from street where city_id = v_city_id
  ) then
    delete from city
    where city_id = v_city_id;
  end if;
end;
$$;


-- пример вызова процедуры удаления с очисткой справочников
call pr_delete_apartment_with_cleanup(
  'Казань',
  'Вахитовский',
  'Улица Баумана',
  '1',
  '15'
);

select
  c.name as "Город",
  d.name as "Район",
  s.name as "Улица",
  a.house_num,
  a.apt_num
from apartment a
join city c
  on c.city_id = a.city_id
join district d
  on d.district_id = a.district_id
join street s
  on s.street_id = a.street_id
where c.name = 'Казань';