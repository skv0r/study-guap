--г) средняя цена однокомнатной квартиры в городе;
select avg(a.evaluated_price)
from apartment a join city c on a.city_id = c.city_id 
where a.rooms_count = 1
  and c.name like 'Санкт-Петербург';


--д) районы, в которых продается наибольшее число объектов недвижимости;
with district_count as ( -- Подзапрос для вычисления количества квартир а районе
  select d.name as district_name, count(a.apartment_id) as count
  from apartment a join district d 
    on a.district_id = d.district_id 
  group by d.name
)
select dc.district_name as "Район", dc.count  as "Всего квартир"
from district_count dc
where dc.count = (select max(dc.count) from district_count dc) -- Получение записей с максимальным зачением dc.count


--е) районы, в которых минимальна стоимость квадратного метра;
with avg_prices as (
  select d.name as "district_name", avg(a.evaluated_price / a.area) as "avg_price"
  from apartment a join district d 
    on a.district_id = d.district_id 
  group by d.name
) select ap.district_name as "Район", ap.avg_price  as "Всего квартир"
from avg_prices ap 
where ap.avg_price = (select min(ap.avg_price) from avg_prices ap)


-- демонстрация всех агрегатных функций
-- статистика агента по продажам
select ag.full_name as "Агент",
  sum(s.sale_price) as "Сумма продаж",
   min(s.sale_price) as "Мин. сделка",
   max(s.sale_price) as "Макс. сделка",
   avg(s.sale_price) as "Средний чек",
   count(*) as "Сделок"
from sale s
join agent ag on ag.agent_id = s.agent_id
group by ag.full_name
order by "Сумма продаж" desc;

-- теоретико-множественные операции и мультимножества

-- Улицы квартир СПб в Московском и Центральном районах
with A as (  -- Московский
  select s.name
  from apartment a
  join street s   on s.street_id = a.street_id
  join city   c   on c.city_id = a.city_id and c.name = 'Санкт-Петербург'
  join district d on d.district_id = a.district_id
  where d.name = 'Московский'
),
B as (  -- Центральный
  select s.name
  from apartment a
  join street s   on s.street_id = a.street_id
  join city   c   on c.city_id = a.city_id and c.name = 'Санкт-Петербург'
  join district d on d.district_id = a.district_id
  where d.name = 'Центральный'
)
select name as "Улица", count(*) as "Количество квартир"
from (
  select * from A
--  union all
--  union
--  except all
--  except
  intersect all
--  intersect
  select * from B
) t
group by name;
