-- а) перечень однокомнатных квартир, продаваемых в Московском районе;
select a.* 
from apartment a  join district d on a.district_id = d.district_id 
where d.name like 'Московский'
	and a.rooms_count = 1
	and a.is_for_sale;

-- б) квартиры, находящиеся на одной улице, но в различных районах;
select
	a1.apartment_id as "Квартира 1",
	d1.name as "Район 1",
	a2.apartment_id as "Квартира 2",
	d2.name  as "Район 2",
	s.name as "Улица"
from apartment a1 join apartment a2 
	on a1.district_id < a2.district_id 
		and a1.apartment_id != a2.apartment_id
		and a1.street_id = a2.street_id
	join street s on s.street_id = a1.street_id 
	join district d1 on d1.district_id = a1.district_id
	join district d2 on d2.district_id = a2.district_id;

-- в) двух- и трехкомнатные квартиры, имеющие одинаковую площадь;
select 
	a1.apartment_id as "Квартира 1",
	a1.area as "Площадь 1",
	a1.rooms_count as "Комнат в кв 1",	
	a2.apartment_id as "Квартира 2",
	a2.area as "Площадь 2",
	a2.rooms_count as "Комнат в кв 2"
from apartment a1 join apartment a2 
	on a1.apartment_id < a2.apartment_id
		and a1.area  = a2.area
		and a1.rooms_count  in (2, 3)
		and a2.rooms_count in (2, 3)
		and a1.rooms_count != a2.rooms_count;


-- distinct районы в которых продаются квартиры с площадью между 35 и 60 кв.м.
select distinct d.name
from apartment a join district d on a.district_id = d.district_id 
where a.area between 35 and 60;

-- is null (известные строительные серии и отсортированные по цене)
select a.apartment_id, bs.code, a.evaluated_price, a.is_for_sale 
from apartment a join building_series bs on a.series_id = bs.series_id 
where a.series_id is not null
	and a.is_for_sale 
order by a.evaluated_price;



