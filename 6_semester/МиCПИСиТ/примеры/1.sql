create extension if not exists pgcrypto;

drop table if exists sale cascade;
drop table if exists room cascade;
drop table if exists apartment cascade;
drop table if exists street_district_city cascade;
drop table if exists street cascade;
drop table if exists district cascade;
drop table if exists agent cascade;
drop table if exists building_series cascade;
drop table if exists city cascade;

create table city (
  city_id uuid primary key default gen_random_uuid(),
  name varchar not null
);

create table district (
  district_id uuid primary key default gen_random_uuid(),
  city_id uuid not null references city(city_id) on delete cascade,
  name varchar not null,

  -- В одном городе район с таким названием должен быть один
  constraint uq_district_city_name unique (city_id, name),

  -- Нужно для составного FK (district_id, city_id)
  constraint uq_district_id_city unique (district_id, city_id)
);


create table street (
  street_id uuid primary key default gen_random_uuid(),
  city_id uuid not null references city(city_id) on delete cascade,
  name varchar not null,

  -- В одном городе улица с таким названием должна быть одна
  constraint uq_street_city_name unique (city_id, name),

  -- Нужно для составного FK (street_id, city_id)
  constraint uq_street_id_city unique (street_id, city_id)
);


create table building_series (
  series_id uuid primary key default gen_random_uuid(),
  code varchar not null unique
);


create table street_district_city (
  street_id uuid not null,
  district_id uuid not null,
  city_id uuid not null,

  primary key (street_id, district_id, city_id),

  foreign key (street_id, city_id)
    references street(street_id, city_id)
    on delete restrict
    on update cascade,

  foreign key (district_id, city_id)
    references district(district_id, city_id)
    on delete restrict
    on update cascade
);


create table apartment (
  apartment_id uuid primary key default gen_random_uuid(),

  city_id uuid not null,
  district_id uuid not null,
  street_id uuid not null,
  house_num varchar not null,
  apt_num varchar not null,

  rooms_count smallint not null check (rooms_count > 0),
  area numeric(8,2) not null check (area > 0),

  evaluated_price numeric(12,2) not null check (evaluated_price >= 0),
  is_for_sale boolean not null default true,
  series_id uuid references building_series(series_id) on delete set null,

  -- FK, проверяющий возможность комбинации город–район–улица
  constraint fk_apartment_address
    foreign key (street_id, district_id, city_id)
    references street_district_city(street_id, district_id, city_id)
    on update cascade
    on delete restrict,

  -- Запрет одинаковых квартир по адресу
  constraint uq_apartment_full_address
    unique (city_id, district_id, street_id, house_num, apt_num)
);


create table room (
  room_id uuid primary key default gen_random_uuid(),
  apartment_id uuid not null references apartment(apartment_id) on delete cascade,
  area numeric(8,2) not null check (area > 0),
  description varchar
);


create table agent (
  agent_id uuid primary key default gen_random_uuid(),
  full_name varchar not null
);


create table sale (
  sale_id uuid primary key default gen_random_uuid(),
  sale_date date not null default current_date,
  agent_id uuid not null references agent(agent_id) on delete restrict,
  apartment_id uuid not null references apartment(apartment_id) on delete restrict,
  sale_price numeric(12,2) not null check (sale_price >= 0)
);

