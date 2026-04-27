-- 6) Пример ХП, демонстрирующего управляющие конструкции
-- ХП формирует временный отчет по разным ценовым диапазонам продаваемых квартир заданного города.

create or replace procedure fill_price_segments_proc(p_city_name text)
language plpgsql as $$
declare
    v_total_count int;
    v_step int := 1;
    v_max_price numeric(12,2) := 5000000.00;		-- текущее значение
begin
    -- удаляем старую временную таблицу, если она существует
    drop table if exists tmp_price_segments;

    -- создаём новую временную таблицу
    create temp table tmp_price_segments (
        step_no int,
        max_price numeric(12,2),
        apartments_count int,
        segment_name varchar
    );

    -- проверяем наличие продаваемых квартир в указанном городе
    select count(*)
    into v_total_count
    from apartment a
    join city c on c.city_id = a.city_id
    where c.name = p_city_name
      and a.is_for_sale = true;

    if v_total_count = 0 then
        raise notice 'В городе "%" нет квартир, выставленных на продажу', p_city_name;
    else
        -- формируем 3 ценовых сегмента
        while v_step <= 3 loop
            insert into tmp_price_segments (step_no, max_price, apartments_count, segment_name)
            select
                v_step,
                v_max_price,
                count(*),
                case
                    when v_step = 1 then 'До 5 млн'
                    when v_step = 2 then 'До 7 млн'
                    else 'До 9 млн'
                end
            from apartment a
            join city c on c.city_id = a.city_id
            where c.name = p_city_name
              and a.is_for_sale = true
              and a.evaluated_price <= v_max_price;

            v_step := v_step + 1;
            v_max_price := v_max_price + 2000000.00;
        end loop;
    end if;
end;
$$;


call fill_price_segments_proc('Санкт-Петербург');


select * from tmp_price_segments order by step_no;