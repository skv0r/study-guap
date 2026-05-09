---
name: cc-1c-overview
description: Ориентирует агента в наборе навыков 1С:Предприятие 8.3 из cc-1c-skills (EPF/ERF, формы, СКД, метаданные, CFE, базы, веб, web-test). Использовать при разработке на 1С, XML-выгрузках конфигуратора, внешних обработках и отчётах, расширениях, публикации или тестировании веб-клиента.
---

# Набор cc-1c-skills в этом проекте

В каталоге `.cursor/skills/` лежит порт [cc-1c-skills](https://github.com/Nikolay-Shirokov/cc-1c-skills) (ветка **port-cursor**, рантайм PowerShell). Каждый подкаталог — отдельный навык со своим `SKILL.md` и скриптами.

## Как выбирать навык

- Описывай задачу естественным языком — агент сопоставит её с полем `description` в `SKILL.md` нужного навыка.
- Точный вызов: слеш-команды из заголовков навыков (например `/epf-init`, `/db-load-cf`, `/web-test`).

## Требования среды

- **Windows**, PowerShell 5.1+ — основной рантайм скриптов.
- **1С:Предприятие 8.3** — сборка/разбор EPF/ERF, работа с ИБ и конфигуратором.
- **Node.js 18+** — навык `/web-test` (браузерная автоматизация).
- Для Python-альтернативы см. ветку **port-cursor-py** upstream и зависимости `lxml`, `psutil` в документации репозитория.

## Обновление навыков

1. В каталоге `tools/cc-1c-skills` выполнить `git pull` (ветка **port-cursor**).
2. Повторно скопировать всё из `tools/cc-1c-skills/.cursor/skills/` в `.cursor/skills/` корня проекта с заменой файлов.

Альтернатива: клонировать [полный репозиторий](https://github.com/Nikolay-Shirokov/cc-1c-skills) (ветка **main**), затем `python scripts/switch.py cursor --project-dir "<корень проекта>"` — скрипт сам положит навыки в `.cursor/skills/`.

Полные гайды и спецификации XML — в [документации upstream](https://github.com/Nikolay-Shirokov/cc-1c-skills/tree/main/docs).
