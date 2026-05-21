# ЛР8 — LOF (обнаружение аномалий на ECG5000)

## Файлы

| Файл | Назначение |
|------|------------|
| `lr8_lof.lgp` | Готовый пакет (открывать из папки `тест` или `ЛР8`) |
| `LOF_template.lgp` | Исходный шаблон преподавателя |
| `ecg_train.txt`, `ecg_test.txt` | Данные |
| `sample_by_object.tsv` | Эталонные `SAMPLE` / `IsTestSet` (из `export_reference.txt`) |
| `export_reference.txt` | Итоговый датасет от преподавателя/одногруппника |
| `Методичка.pdf` | Задание |

Сборка пакета:  
`.cursor/skills/loginom-assistant/scripts/build_lr8_project.py`

Проверка метрик в Python (эталон):  
`.cursor/skills/loginom-assistant/scripts/validate_lr8_lof.py`

Выгрузка и сравнение с `export_reference.txt`:

```bash
python .cursor/skills/loginom-assistant/scripts/export_lr8_dataset.py
python .cursor/skills/loginom-assistant/scripts/compare_lr8_export.py ЛР8/export_reference.txt ЛР8/export_out.txt
```

## Сценарий 1 — «Подготовка выборок» (узлы Loginom)

В `lr8_lof.lgp` цепочка **без Python** для разбиения:

1. Импорт `ecg_train.txt` / `ecg_test.txt` → **IsTestSet** (`false` / `true`) → **Объединение** (сначала train-файл, затем test-файл).
2. **Калькулятор** `OBJECT` = `"obj"+str(RowNum()+1)`.
3. **Фильтр** `CLASS = 1` → только нормальные записи.
4. **Разбиение на множества** 70/30, seed **42**, метод **Последовательный** (`smSequence`) → обучающая часть нормальных.
5. Ветка «не CLASS=1» + 30% нормальных → **Объединение** «остаток + аномалии».
6. **Разбиение** 50/50, seed **42**, **Стратифицированный** по `CLASS` → valid / test.
7. **Калькуляторы** на ветках: только `SAMPLE` = `train` / `valid` / `test`.
8. Три ветки → **Объединение** (с `SAMPLE`) → **IsTestSet (из SAMPLE)** (`IsTestSet = SAMPLE<>"train"`) → публичный узел **Датасет ECG (публичный)**.

Ожидаемые объёмы: **3233** записи, train **2043** (только CLASS=1), valid **595**, test **595** (с аномалиями CLASS 3,4,5).

> **Важно:** фильтр `CLASS = 1` стоит **только** на ветке 70% нормальных, не на всём датасете до первого разбиения.

**Loginom 7.3.1:** пересборка:

```bash
python .cursor/skills/loginom-assistant/scripts/build_lr8_project.py
```

После первого запуска в Loginom откройте узлы **Разбиение** и нажмите «Выполнить» / обновите статистику (число строк подтянется автоматически).

**Выгрузка в `Выход-скрипта.txt`:** в **Объединении** сопоставляют только `CLASS`, при необходимости `SAMPLE` и `VAR2`–`VAR141` — **не** `OBJECT` и **не** `IsTestSet` (иначе дубликаты `OBJECT.1`, `IsTestSet.1` и лишние столбцы). `IsTestSet` задаётся одним калькулятором после финального объединения. Пересборка: `build_lr8_project.py` или `patch_lr8_fix_unions.py`.

> Разбиение узлами Loginom (seed=42) может **не совпасть построчно** с `export_reference.txt` одногруппника (~50% совпадений меток `SAMPLE` при чистом sklearn). Для **точного** совпадения с эталоном используйте `sample_by_object.tsv` (см. `export_lr8_dataset.py`). Для ЛР и LOF достаточно методичных долей и наличия аномалий в valid/test.

## Сценарий 2 — «Построение модели» (нативные модули Loginom)

Цепочка как у одногруппника (`мареев.lgp`), с исправлениями:

1. **Ссылка на узел** → публичный **Датасет ECG (публичный)**.
2. **meta-scaling** → z-нормализация.
3. **neighbors.LOF Novelty** + **model.fitter** (k=15, c=0.05, модель `C:\model\k15_c05`).
4. Скоринг: **meta-scaled ALL** → `is_valid` / `is_test` → **model.fitter** (порт «Тестовая выборка»).
5. **Замена** `outlier_label` и **CLASS** (0/1) → **classification metrics** (β=2).

Пересборка сценария 2:

```bash
python .cursor/skills/loginom-assistant/scripts/rebuild_lr8_unit1_native.py
```

Отдельно: `merge_lr8_unit1_from_mareev.py`, `fix_lr8_unit1_metrics.py`, `patch_lr8_unit1_wiring.py`.

Запасной вариант (те же метрики, без kits): `build_lr8_unit1_python_metrics.py`.

**Порядок запуска:** сначала **Подготовка выборок**, затем **Построение модели**.

### Ожидаемые метрики (k=15, c=0.05)

| Выборка | recall | fn (ориентир) |
|---------|--------|----------------|
| valid | ~0.9 | ~10 |
| test | ~0.97 | ~4 |

### Ориентиры по valid (Fbeta, β=2)

| k | c | Fbeta (ориентир) |
|---|-----|------------------|
| 15 | 0.02 | ~0.84 |
| 15 | 0.05 | **~0.97** |
| 20 | 0.05 | ~0.97 |
| 25 | 0.05 | ~0.97 |

### Test для k=15, c=0.05

| Метрика | Значение |
|---------|----------|
| Precision | ~0.87 |
| Recall | ~0.97 |
| F1 | ~0.92 |
| Fbeta | ~0.97 |

## Ссылки на библиотеки

Пакет ссылается на:

- `../ЛР4/libs/python_kits/python_kits/loginom_sklearn_kit.lgp`
- `../ЛР4/libs/python_kits/python_kits/loginom_sklearn_meta.lgp`
- `../ЛР4/libs/silver_kit/silver_kit/loginom_silver_kit.lgp`

Открывайте `lr8_lof.lgp` из корня репозитория `тест`, чтобы пути к ЛР4 и ЛР8 были корректны.

## Если пакет не открывается

Как в ЛР6/ЛР7: не используется устаревший `PackageIndex.bin` из шаблона. Пересоберите:

```bash
python .cursor/skills/loginom-assistant/scripts/build_lr8_project.py
```
