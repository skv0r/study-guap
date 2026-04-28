-- 3) каскадное удаление
-- при удалении района сначала удаляются продажи квартир этого района,
-- затем квартиры, затем связи улица-район-город, после чего сам район

create or replace procedure pr_delete_district_cascade(
  p_city_name varchar,
  p_district_name varchar
)
language plpgsql
as $$
declare
  v_city_id uuid;
  v_district_id uuid;
begin
  select c.city_id, d.district_id
  into v_city_id, v_district_id
  from city c
  join district d
    on d.city_id = c.city_id
  where c.name = p_city_name
    and d.name = p_district_name;

  if v_district_id is null then
    raise notice 'Район не найден';
    return;
  end if;

  delete from sale s
  using apartment a
  where s.apartment_id = a.apartment_id
    and a.district_id = v_district_id;

  delete from apartment
  where district_id = v_district_id;

  delete from street_district_city
  where district_id = v_district_id;

  delete from district
  where district_id = v_district_id;

  delete from street st
  where st.city_id = v_city_id
    and not exists (
      select *
      from street_district_city sdc
      where sdc.street_id = st.street_id
    )
    and not exists (
      select *
      from apartment a
      where a.street_id = st.street_id
    );

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


-- тестовые данные для каскадного удаления
call pr_add_apartment(
  'Самара',
  'Ленинский',
  'Молодогвардейская улица',
  '5',
  '8',
  2::smallint,
  54.00,
  6300000.00,
  true,
  'К-7'
);

call pr_add_apartment(
  'Самара',
  'Ленинский',
  'Молодогвардейская улица',
  '5',
  '9',
  1::smallint,
  38.00,
  4700000.00,
  true,
  'П-44'
);

select
  c.name as "Город",
  d.name as "Район",
  count(*) as "Квартир в районе"
from apartment a
join city c
  on c.city_id = a.city_id
join district d
  on d.district_id = a.district_id
where c.name = 'Самара'
  and d.name = 'Ленинский'
group by c.name, d.name;

call pr_delete_district_cascade('Самара', 'Ленинский');

select
  c.name as "Город",
  d.name as "Район"
from city c
left join district d
  on d.city_id = c.city_id
where c.name = 'Самара';