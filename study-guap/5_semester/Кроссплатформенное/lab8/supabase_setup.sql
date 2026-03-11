-- ===============================================
-- SQL скрипт для настройки Supabase
-- Проект: F1 Drivers Management App
-- ===============================================

-- 1. Создание таблицы drivers
CREATE TABLE IF NOT EXISTS drivers (
    id SERIAL PRIMARY KEY,
    full_name TEXT NOT NULL,
    driver_number INTEGER NOT NULL,
    first_name TEXT,
    last_name TEXT,
    team_name TEXT,
    team_colour TEXT,
    name_acronym TEXT,
    country_code TEXT,
    broadcast_name TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT TIMEZONE('utc'::text, NOW()) NOT NULL
);

-- 2. Отключение Row Level Security (для разработки и тестирования)
ALTER TABLE drivers DISABLE ROW LEVEL SECURITY;

-- АЛЬТЕРНАТИВНО: Можно включить RLS и создать политики для публичного доступа
-- ALTER TABLE drivers ENABLE ROW LEVEL SECURITY;
-- 
-- CREATE POLICY "Allow public read access" ON drivers
--     FOR SELECT USING (true);
-- 
-- CREATE POLICY "Allow public insert access" ON drivers
--     FOR INSERT WITH CHECK (true);
-- 
-- CREATE POLICY "Allow public update access" ON drivers
--     FOR UPDATE USING (true);
-- 
-- CREATE POLICY "Allow public delete access" ON drivers
--     FOR DELETE USING (true);

-- 3. Создание индекса для быстрого поиска по номеру гонщика
CREATE INDEX IF NOT EXISTS idx_driver_number ON drivers(driver_number);

-- 4. Добавление тестовых данных (гонщики F1 2024)
INSERT INTO drivers (full_name, driver_number, first_name, last_name, team_name, team_colour, name_acronym, country_code, broadcast_name)
VALUES 
    ('Max Verstappen', 1, 'Max', 'Verstappen', 'Red Bull Racing', '3671C6', 'VER', 'NED', 'M VERSTAPPEN'),
    ('Sergio Perez', 11, 'Sergio', 'Perez', 'Red Bull Racing', '3671C6', 'PER', 'MEX', 'S PEREZ'),
    ('Lewis Hamilton', 44, 'Lewis', 'Hamilton', 'Mercedes', '27F4D2', 'HAM', 'GBR', 'L HAMILTON'),
    ('George Russell', 63, 'George', 'Russell', 'Mercedes', '27F4D2', 'RUS', 'GBR', 'G RUSSELL'),
    ('Charles Leclerc', 16, 'Charles', 'Leclerc', 'Ferrari', 'E80020', 'LEC', 'MON', 'C LECLERC'),
    ('Carlos Sainz', 55, 'Carlos', 'Sainz', 'Ferrari', 'E80020', 'SAI', 'ESP', 'C SAINZ'),
    ('Lando Norris', 4, 'Lando', 'Norris', 'McLaren', 'FF8000', 'NOR', 'GBR', 'L NORRIS'),
    ('Oscar Piastri', 81, 'Oscar', 'Piastri', 'McLaren', 'FF8000', 'PIA', 'AUS', 'O PIASTRI'),
    ('Fernando Alonso', 14, 'Fernando', 'Alonso', 'Aston Martin', '229971', 'ALO', 'ESP', 'F ALONSO'),
    ('Lance Stroll', 18, 'Lance', 'Stroll', 'Aston Martin', '229971', 'STR', 'CAN', 'L STROLL'),
    ('Pierre Gasly', 10, 'Pierre', 'Gasly', 'Alpine', 'FF87BC', 'GAS', 'FRA', 'P GASLY'),
    ('Esteban Ocon', 31, 'Esteban', 'Ocon', 'Alpine', 'FF87BC', 'OCO', 'FRA', 'E OCON'),
    ('Alexander Albon', 23, 'Alexander', 'Albon', 'Williams', '64C4FF', 'ALB', 'THA', 'A ALBON'),
    ('Logan Sargeant', 2, 'Logan', 'Sargeant', 'Williams', '64C4FF', 'SAR', 'USA', 'L SARGEANT'),
    ('Valtteri Bottas', 77, 'Valtteri', 'Bottas', 'Alfa Romeo', 'C92D4B', 'BOT', 'FIN', 'V BOTTAS'),
    ('Zhou Guanyu', 24, 'Zhou', 'Guanyu', 'Alfa Romeo', 'C92D4B', 'ZHO', 'CHN', 'G ZHOU'),
    ('Kevin Magnussen', 20, 'Kevin', 'Magnussen', 'Haas F1 Team', 'B6BABD', 'MAG', 'DEN', 'K MAGNUSSEN'),
    ('Nico Hulkenberg', 27, 'Nico', 'Hulkenberg', 'Haas F1 Team', 'B6BABD', 'HUL', 'GER', 'N HULKENBERG'),
    ('Yuki Tsunoda', 22, 'Yuki', 'Tsunoda', 'AlphaTauri', '5E8FAA', 'TSU', 'JPN', 'Y TSUNODA'),
    ('Daniel Ricciardo', 3, 'Daniel', 'Ricciardo', 'AlphaTauri', '5E8FAA', 'RIC', 'AUS', 'D RICCIARDO')
ON CONFLICT DO NOTHING;

-- 5. Проверка данных
SELECT COUNT(*) as total_drivers FROM drivers;

-- 6. Вывод всех гонщиков
SELECT id, full_name, driver_number, team_name FROM drivers ORDER BY driver_number;

