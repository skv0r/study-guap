-- Вставка данных

-- Города
insert into city(name) values
  ('Санкт-Петербург'),
  ('Москва');

-- Серии домов (строительные серии)
insert into building_series(code) values
  ('П-44'),
  ('К-7'),
  ('Лен-60'),
  ('И-209');

-- Районы
insert into district(city_id, name)
select city.city_id, dname
from city
join (values
  ('Санкт-Петербург', 'Московский'),
  ('Санкт-Петербург', 'Центральный'),
  ('Санкт-Петербург', 'Петроградский'),
  ('Москва', 'Тверской'),
  ('Москва', 'Пресненский')
) t(city_name, dname) on t.city_name = city.name;

-- Улицы
insert into street(city_id, name)
select city.city_id, sname
from city
join (values
  ('Санкт-Петербург', 'Пулковское шоссе'),
  ('Санкт-Петербург', 'Лиговский проспект'),
  ('Санкт-Петербург', 'Невский проспект'),
  ('Санкт-Петербург', 'Каменноостровский проспект'),
  ('Москва', 'Тверская улица'),
  ('Москва', 'Зоологическая улица')
) t(city_name, sname) on t.city_name = city.name;

-- Связь улица–район–город 
-- СПБ: Пулковское шоссе только в Московском
insert into street_district_city(street_id, district_id, city_id)
select street.street_id, district.district_id, city.city_id
from city
join street on street.city_id = city.city_id and street.name = 'Пулковское шоссе'
join district on district.city_id = city.city_id and district.name = 'Московский'
where city.name = 'Санкт-Петербург';

-- СПБ: Лиговский проспект в Московском и Центральном
insert into street_district_city(street_id, district_id, city_id)
select street.street_id, district.district_id, city.city_id
from city
join street on street.city_id = city.city_id and street.name = 'Лиговский проспект'
join district on district.city_id = city.city_id and district.name = 'Московский'
where city.name = 'Санкт-Петербург';

insert into street_district_city(street_id, district_id, city_id)
select street.street_id, district.district_id, city.city_id
from city
join street on street.city_id = city.city_id and street.name = 'Лиговский проспект'
join district on district.city_id = city.city_id and district.name = 'Центральный'
where city.name = 'Санкт-Петербург';

-- СПБ: Невский проспект в Центральном (верная связь)
insert into street_district_city(street_id, district_id, city_id)
select street.street_id, district.district_id, city.city_id
from city
join street on street.city_id = city.city_id and street.name = 'Невский проспект'
join district on district.city_id = city.city_id and district.name = 'Центральный'
where city.name = 'Санкт-Петербург';

-- СПБ: Невский проспект в Московском (ошибочная связь, будет удалена)
insert into street_district_city(street_id, district_id, city_id)
select street.street_id, district.district_id, city.city_id
from city
join street on street.city_id = city.city_id and street.name = 'Невский проспект'
join district on district.city_id = city.city_id and district.name = 'Московский'
where city.name = 'Санкт-Петербург';

-- СПБ: Каменноостровский проспект в Петроградском
insert into street_district_city(street_id, district_id, city_id)
select street.street_id, district.district_id, city.city_id
from city
join street on street.city_id = city.city_id and street.name = 'Каменноостровский проспект'
join district on district.city_id = city.city_id and district.name = 'Петроградский'
where city.name = 'Санкт-Петербург';

-- Москва: Тверская улица в Тверском
insert into street_district_city(street_id, district_id, city_id)
select street.street_id, district.district_id, city.city_id
from city
join street on street.city_id = city.city_id and street.name = 'Тверская улица'
join district on district.city_id = city.city_id and district.name = 'Тверской'
where city.name = 'Москва';

-- Москва: Зоологическая улица в Пресненском
insert into street_district_city(street_id, district_id, city_id)
select street.street_id, district.district_id, city.city_id
from city
join street on street.city_id = city.city_id and street.name = 'Зоологическая улица'
join district on district.city_id = city.city_id and district.name = 'Пресненский'
where city.name = 'Москва';

-- Квартиры
-- СПБ, Московский, Пулковское шоссе, 1-комн., продаётся
insert into apartment(
  city_id, district_id, street_id, house_num, apt_num,
  rooms_count, area, evaluated_price, is_for_sale, series_id
)
select city.city_id, district.district_id, street.street_id, '10','1',
       1, 35.00, 4200000.00, true, building_series.series_id
from city
join district on district.city_id = city.city_id and district.name = 'Московский'
join street on street.city_id = city.city_id and street.name = 'Пулковское шоссе'
join building_series on building_series.code = 'П-44'
where city.name = 'Санкт-Петербург';

-- СПБ, Московский, Пулковское шоссе, 1-комн., НЕ продаётся 
insert into apartment(
  city_id, district_id, street_id, house_num, apt_num,
  rooms_count, area, evaluated_price, is_for_sale, series_id
)
select city.city_id, district.district_id, street.street_id, '10','2',
       1, 30.00, 3600000.00, false, building_series.series_id
from city
join district on district.city_id = city.city_id and district.name = 'Московский'
join street on street.city_id = city.city_id and street.name = 'Пулковское шоссе'
join building_series on building_series.code = 'П-44'
where city.name = 'Санкт-Петербург';

-- СПБ, Московский, Лиговский проспект, 2-комн., 60 м2, продаётся 
insert into apartment(
  city_id, district_id, street_id, house_num, apt_num,
  rooms_count, area, evaluated_price, is_for_sale, series_id
)
select city.city_id, district.district_id, street.street_id, '100','12',
       2, 60.00, 7200000.00, true, building_series.series_id
from city
join district on district.city_id = city.city_id and district.name = 'Московский'
join street on street.city_id = city.city_id and street.name = 'Лиговский проспект'
join building_series on building_series.code = 'К-7'
where city.name = 'Санкт-Петербург';

-- СПБ, Центральный, Лиговский проспект, 3-комн., 60 м2, продаётся 
insert into apartment(
  city_id, district_id, street_id, house_num, apt_num,
  rooms_count, area, evaluated_price, is_for_sale, series_id
)
select city.city_id, district.district_id, street.street_id, '80','8',
       3, 60.00, 9000000.00, true, building_series.series_id
from city
join district on district.city_id = city.city_id and district.name = 'Центральный'
join street on street.city_id = city.city_id and street.name = 'Лиговский проспект'
join building_series on building_series.code = 'Лен-60'
where city.name = 'Санкт-Петербург';

-- СПБ, Центральный, Невский проспект, 1-комн., продаётся
insert into apartment(
  city_id, district_id, street_id, house_num, apt_num,
  rooms_count, area, evaluated_price, is_for_sale, series_id
)
select city.city_id, district.district_id, street.street_id, '1','10',
       1, 28.00, 5000000.00, true, building_series.series_id
from city
join district on district.city_id = city.city_id and district.name = 'Центральный'
join street on street.city_id = city.city_id and street.name = 'Невский проспект'
join building_series on building_series.code = 'И-209'
where city.name = 'Санкт-Петербург';

-- СПБ, Петроградский, Каменноостровский проспект, 2-комн., продаётся
insert into apartment(
  city_id, district_id, street_id, house_num, apt_num,
  rooms_count, area, evaluated_price, is_for_sale, series_id
)
select city.city_id, district.district_id, street.street_id, '5','15',
       2, 45.00, 4500000.00, true, null
from city
join district on district.city_id = city.city_id and district.name = 'Петроградский'
join street on street.city_id = city.city_id and street.name = 'Каменноостровский проспект'
where city.name = 'Санкт-Петербург';

-- СПБ, Московский, Лиговский проспект, 1-комн., продаётся
insert into apartment(
  city_id, district_id, street_id, house_num, apt_num,
  rooms_count, area, evaluated_price, is_for_sale, series_id
)
select city.city_id, district.district_id, street.street_id, '102','3',
       1, 33.00, 4400000.00, true, building_series.series_id
from city
join district on district.city_id = city.city_id and district.name = 'Московский'
join street on street.city_id = city.city_id and street.name = 'Лиговский проспект'
join building_series on building_series.code = 'И-209'
where city.name = 'Санкт-Петербург';

-- СПБ, Невский пр-т в Московском (неверно), 2-комн., продаётся
-- Будет перенесена в Центральный (корректировка)
insert into apartment(
  city_id, district_id, street_id, house_num, apt_num,
  rooms_count, area, evaluated_price, is_for_sale, series_id
)
select city.city_id, district.district_id, street.street_id, '3','9',
       2, 55.00, 5500000.00, true, building_series.series_id
from city
join district on district.city_id = city.city_id and district.name = 'Московский'      -- ошибка
join street on street.city_id = city.city_id and street.name = 'Невский проспект'
join building_series on building_series.code = 'Лен-60'
where city.name = 'Санкт-Петербург';

-- СПБ, Московский, Пулковское шоссе, 2-комн., продаётся,
insert into apartment(
  city_id, district_id, street_id, house_num, apt_num,
  rooms_count, area, evaluated_price, is_for_sale, series_id
)
select city.city_id, district.district_id, street.street_id, '12','1',
       2, 50.00, 6000000.00, true, building_series.series_id
from city
join district on district.city_id = city.city_id and district.name = 'Московский'
join street on street.city_id = city.city_id and street.name = 'Пулковское шоссе'
join building_series on building_series.code = 'Лен-60'
where city.name = 'Санкт-Петербург';

-- Москва, Тверской, Тверская улица, 1-комн., продаётся
insert into apartment(
  city_id, district_id, street_id, house_num, apt_num,
  rooms_count, area, evaluated_price, is_for_sale, series_id
)
select city.city_id, district.district_id, street.street_id, '1','1',
       1, 25.00, 6500000.00, true, building_series.series_id
from city
join district on district.city_id = city.city_id and district.name = 'Тверской'
join street on street.city_id = city.city_id and street.name = 'Тверская улица'
join building_series on building_series.code = 'К-7'
where city.name = 'Москва';

-- Комнаты 
insert into room(apartment_id, area, description)
select apartment.apartment_id, 20.00, 'Жилая'
from apartment
join city on city.city_id = apartment.city_id and city.name = 'Санкт-Петербург'
join district on district.district_id = apartment.district_id and district.name = 'Московский'
join street on street.street_id = apartment.street_id and street.name = 'Пулковское шоссе'
where apartment.house_num = '10' and apartment.apt_num = '1';

insert into room(apartment_id, area, description)
select apartment.apartment_id, 18.00, 'Спальня'
from apartment
join city on city.city_id = apartment.city_id and city.name = 'Санкт-Петербург'
join district on district.district_id = apartment.district_id and district.name = 'Московский'
join street on street.street_id = apartment.street_id and street.name = 'Лиговский проспект'
where apartment.house_num = '100' and apartment.apt_num = '12';

insert into room(apartment_id, area, description)
select apartment.apartment_id, 16.00, 'Гостиная'
from apartment
join city on city.city_id = apartment.city_id and city.name = 'Санкт-Петербург'
join district on district.district_id = apartment.district_id and district.name = 'Московский'
join street on street.street_id = apartment.street_id and street.name = 'Лиговский проспект'
where apartment.house_num = '100' and apartment.apt_num = '12';

insert into room(apartment_id, area, description)
select apartment.apartment_id, 15.00, 'Комната 1'
from apartment
join city on city.city_id = apartment.city_id and city.name = 'Санкт-Петербург'
join district on district.district_id = apartment.district_id and district.name = 'Центральный'
join street on street.street_id = apartment.street_id and street.name = 'Лиговский проспект'
where apartment.house_num = '80' and apartment.apt_num = '8';

insert into room(apartment_id, area, description)
select apartment.apartment_id, 14.00, 'Комната 2'
from apartment
join city on city.city_id = apartment.city_id and city.name = 'Санкт-Петербург'
join district on district.district_id = apartment.district_id and district.name = 'Центральный'
join street on street.street_id = apartment.street_id and street.name = 'Лиговский проспект'
where apartment.house_num = '80' and apartment.apt_num = '8';

insert into room(apartment_id, area, description)
select apartment.apartment_id, 13.00, 'Комната 3'
from apartment
join city on city.city_id = apartment.city_id and city.name = 'Санкт-Петербург'
join district on district.district_id = apartment.district_id and district.name = 'Центральный'
join street on street.street_id = apartment.street_id and street.name = 'Лиговский проспект'
where apartment.house_num = '80' and apartment.apt_num = '8';

insert into agent(full_name) values
  ('Иванов И.И.'),
  ('Претров П.П.'),  -- опечатка (будет исправлена на Петров П.П.)
  ('Сидорова А.А.');

-- Продажи
insert into sale(sale_date, agent_id, apartment_id, sale_price)
select date '2025-12-01', agent.agent_id, apartment.apartment_id, 3700000.00
from agent, apartment
join city on city.city_id = apartment.city_id and city.name = 'Санкт-Петербург'
join district on district.district_id = apartment.district_id and district.name = 'Московский'
join street on street.street_id = apartment.street_id and street.name = 'Пулковское шоссе'
where apartment.house_num = '10' and apartment.apt_num = '2' and agent.full_name = 'Иванов И.И.'
limit 1;

insert into sale(sale_date, agent_id, apartment_id, sale_price)
select date '2026-01-15', agent.agent_id, apartment.apartment_id, 8800000.00
from agent, apartment
join city on city.city_id = apartment.city_id and city.name = 'Санкт-Петербург'
join district on district.district_id = apartment.district_id and district.name = 'Центральный'
join street on street.street_id = apartment.street_id and street.name = 'Лиговский проспект'
where apartment.house_num = '80' and apartment.apt_num = '8' and agent.full_name = 'Сидорова А.А.'
limit 1;

insert into sale(sale_date, agent_id, apartment_id, sale_price)
select date '2026-02-10', agent.agent_id, apartment.apartment_id, 0.00
from agent, apartment
join city on city.city_id = apartment.city_id and city.name = 'Санкт-Петербург'
join district on district.district_id = apartment.district_id and district.name = 'Московский'
join street on street.street_id = apartment.street_id and street.name = 'Лиговский проспект'
where apartment.house_num = '100' and apartment.apt_num = '12' and agent.full_name = 'Претров П.П.'
limit 1;

-- Исправления неправильных данных

-- 1) Исправление опечатки в имени агента: 'Претров' -> 'Петров'
update agent
set full_name = 'Петров П.П.'
where full_name = 'Претров П.П.';

-- 2) Исправление нулевой цены продажи до корректной 
update sale
set sale_price = 7200000.00
from apartment
join city on city.city_id = apartment.city_id and city.name = 'Санкт-Петербург'
join district on district.district_id = apartment.district_id and district.name = 'Московский'
join street on street.street_id = apartment.street_id and street.name = 'Лиговский проспект'
where sale.apartment_id = apartment.apartment_id
  and apartment.house_num = '100' and apartment.apt_num = '12';

-- 3) Квартира ошибочно привязана к Невскому проспекту в Московский.
update apartment
set district_id = district_new.district_id
from city, district district_old, district district_new, street
where city.name = 'Санкт-Петербург'
  and district_new.city_id = city.city_id
  and district_new.name = 'Центральный'
  and street.name = 'Невский проспект'
  and street.street_id = apartment.street_id
  and district_old.district_id = apartment.district_id
  and district_old.name = 'Московский'
  and apartment.house_num = '3'
  and apartment.apt_num = '9';

-- Удаление ошибочной связи улица–район (Невский проспект в Московском)
delete from street_district_city
using city, street, district
where street_district_city.city_id = city.city_id
and street_district_city.street_id = street.street_id
and street_district_city.district_id = district.district_id
and city.name = 'Санкт-Петербург'
and street.name = 'Невский проспект'
and district.name = 'Московский';
