-- 4) вычисление и возврат значения агрегатной функции
-- средняя цена однокомнатной квартиры в заданном городе

create or replace function fn_avg_one_room_price(p_city_name varchar)
returns numeric(12,2)
language plpgsql
as $$
declare
  v_avg numeric(12,2);
begin
  select round(avg(a.evaluated_price), 2)
  into v_avg
  from apartment a
  join city c on c.city_id = a.city_id
  where c.name = p_city_name and a.rooms_count = 1;
  return v_avg;
end;
$$;


select fn_avg_one_room_price('Санкт-Петербург');