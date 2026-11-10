# ЛР5. Веб-программирование

Источник: методичка «ЛР5 Установка PostgreSQL и подключение Prisma ORM».

Тема: «Установка PostgreSQL и подключение Prisma ORM».

Сдача: код в репозитории pairwise, ветка `lr5/postgres-prisma`, без слияния в main. Отчёт Word. Преподаватель: ассистент В.В. Жукалин. Студент Г.В. Буренков, группа 4321.

Продолжение PairWise из ЛР4. Имена UI — PairWise; роль БД `study_app`, база `study_space`, Prisma ORM 6.19.2.

Код пишет студент. Агент собирает только отчёт по готовым доказательствам.

Обязательный результат:

- PostgreSQL установлен, служба `postgresql-x64-17`, порт 5432;
- роль `study_app` без суперпользователя, база `study_space` с владельцем `study_app`;
- `prisma` и `@prisma/client` 6.19.2, schema с временной моделью `SetupCheck`;
- `DATABASE_URL` только в локальном `.env`, `.env.example` без секрета;
- `app/lib/prisma.ts`, `scripts/check-prisma.ts`, успешные validate/generate и TypeScript-запрос;
- предметные таблицы ещё не созданы; миграции не выполняются;
- `npm run lint` и `npm run build` без ошибок.

Содержание отчёта — пункты раздела 7 методички, включая все контрольные вопросы. Раздел 8 методички списком не переносить.

Критерий готовности: соединение `study_space | study_app | 5432`, lint/build ок.
