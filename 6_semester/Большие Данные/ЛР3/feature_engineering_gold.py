# -*- coding: utf-8 -*-
"""
Лабораторная работа №3. Построение и отбор признаков.
Скрипт для данных по золоту (GoldUSD.csv).
Совместим с Loginom: при запуске в узле Python данные берутся из InputTable.
При запуске из командной строки — из CSV-файла.
"""

import os
import numpy as np
import pandas as pd
from typing import Optional, List, Tuple

# Попытка импорта модулей Loginom (доступны только при запуске в платформе)
try:
    import builtin_data
    from builtin_data import InputTable, OutputTable
    from builtin_pandas_utils import to_data_frame, prepare_compatible_table, fill_table
    IN_LOGINOM = True
except ImportError:
    IN_LOGINOM = False

# =============================================================================
# Конфигурация
# =============================================================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(SCRIPT_DIR, "GoldUSD.csv")
TARGET_COLUMN = "Close"
K_BEST = 20
DROP_NA_THRESHOLD = 0.5
OUTLIER_STD_THRESHOLD = 3.0
# =============================================================================


def load_data(df: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    """Загрузка данных: из Loginom или из CSV."""
    if df is not None:
        return df.copy()
    if IN_LOGINOM:
        return to_data_frame(InputTable).copy()
    return pd.read_csv(CSV_PATH).copy()


def parse_date(df: pd.DataFrame) -> pd.DataFrame:
    """Парсинг даты и создание признаков времени."""
    df = df.copy()
    df["Date"] = pd.to_datetime(df["Date"], format="%d-%m-%y", errors="coerce")
    df["Year"] = df["Date"].dt.year
    df["Month"] = df["Date"].dt.month
    df["DayOfWeek"] = df["Date"].dt.dayofweek
    return df


def fill_missing(df: pd.DataFrame, numeric_cols: List[str]) -> pd.DataFrame:
    """Заполнение пропусков: числовые — медианой."""
    df = df.copy()
    for col in numeric_cols:
        if col in df.columns and df[col].isna().any():
            median_val = df[col].median()
            df[col] = df[col].fillna(median_val)
    return df


def drop_high_na_columns(
    df: pd.DataFrame,
    threshold: float = 0.5,
    target_column: str = "Close",
) -> pd.DataFrame:
    """Удаление столбцов с долей пропусков выше порога."""
    na_ratios = df.isna().mean()
    to_drop = [c for c, r in na_ratios.items() if r > threshold and c != target_column]
    if to_drop:
        df = df.drop(columns=to_drop)
    return df


def handle_outliers(
    df: pd.DataFrame,
    numeric_cols: List[str],
    n_std: float = 3.0,
    target_column: str = "Close",
) -> pd.DataFrame:
    """Ограничение выбросов (clip): не удаляем строки, а обрезаем экстремумы."""
    df = df.copy()
    for col in numeric_cols:
        if col not in df.columns or col == target_column:
            continue
        mean_val = df[col].mean()
        std_val = df[col].std()
        if std_val > 0:
            lower = mean_val - n_std * std_val
            upper = mean_val + n_std * std_val
            df[col] = df[col].clip(lower=lower, upper=upper)
    return df


def engineer_features(df: pd.DataFrame) -> pd.DataFrame:
    """Построение признаков для OHLC-данных."""
    df = df.copy()
    if "High" in df.columns and "Low" in df.columns:
        df["OHLC_Range"] = df["High"] - df["Low"]
    if "Close" in df.columns and "Open" in df.columns:
        df["Price_Change"] = df["Close"] - df["Open"]
    if "Open" in df.columns and df["Open"].replace(0, np.nan).notna().any():
        df["Price_Change_Pct"] = (df["Close"] - df["Open"]) / df["Open"].replace(0, np.nan)
    
    # Скользящие средние (5 и 20 периодов)
    if "Close" in df.columns:
        df["MA5"] = df["Close"].rolling(5, min_periods=1).mean()
        df["MA20"] = df["Close"].rolling(20, min_periods=1).mean()
    
    # Объём: логарифм (Volume часто имеет много нулей)
    if "Volume" in df.columns:
        df["Log_Volume"] = np.log1p(df["Volume"].fillna(0))
    
    return df


def normalize_minmax(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    """Нормализация min-max: диапазон [0, 1]."""
    df = df.copy()
    for col in cols:
        if col not in df.columns:
            continue
        min_val = df[col].min()
        max_val = df[col].max()
        if max_val > min_val:
            df[col] = (df[col] - min_val) / (max_val - min_val)
    return df


def standardize(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    """Стандартизация: среднее=0, std=1."""
    df = df.copy()
    for col in cols:
        if col not in df.columns:
            continue
        mean_val = df[col].mean()
        std_val = df[col].std()
        if std_val > 0:
            df[col] = (df[col] - mean_val) / std_val
    return df


def correlation_table(X: pd.DataFrame, y: pd.Series) -> pd.DataFrame:
    """Таблица корреляций Пирсона: признак — коэффициент (отсортировано по |corr|)."""
    rows = []
    y_std = (y - y.mean()).std(ddof=0)
    if y_std == 0:
        return pd.DataFrame(columns=["Признак", "Коэффициент_корреляции"])
    
    for col in X.columns:
        x = X[col]
        if x.std(ddof=0) == 0:
            continue
        try:
            corr = x.corr(y)
        except Exception:
            corr = np.nan
        if pd.notna(corr):
            rows.append((col, float(corr)))
    
    corr_df = pd.DataFrame(rows, columns=["Признак", "Коэффициент_корреляции"])
    corr_df["|corr|"] = corr_df["Коэффициент_корреляции"].abs()
    corr_df = corr_df.sort_values("|corr|", ascending=False).reset_index(drop=True)
    corr_df = corr_df.drop(columns=["|corr|"])
    return corr_df


def select_top_features(corr_df: pd.DataFrame, k: int) -> List[str]:
    """Выбор топ-K признаков по модулю корреляции."""
    if corr_df.empty:
        return []
    return corr_df["Признак"].head(k).tolist()


def run(df_input: Optional[pd.DataFrame] = None) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Основной пайплайн построения и отбора признаков.
    Возвращает: (processed_df, corr_table)
    """
    df = load_data(df_input)

    # Явно приводим стандартные числовые столбцы к типу float
    for col in ["Open", "High", "Low", "Close", "Volume"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # 1. Парсинг даты
    if "Date" in df.columns:
        df = parse_date(df)
    
    # Числовые столбцы (исключая Date)
    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()

    # Определяем целевой столбец: сначала пытаемся использовать TARGET_COLUMN,
    # если его нет, пробуем подобрать автоматически.
    target_col = TARGET_COLUMN
    if target_col not in df.columns:
        # Популярные имена целевой переменной
        preferred_names = ("close", "saleprice", "target", "y")
        candidates = [c for c in df.columns if c.lower() in preferred_names]
        if candidates:
            target_col = candidates[0]
        else:
            # Последний числовой столбец как запасной вариант
            num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            if num_cols:
                target_col = num_cols[-1]

    if target_col not in df.columns:
        raise ValueError(
            f"Целевой столбец не найден. Ожидался '{TARGET_COLUMN}', "
            f"но ни один из кандидатов не обнаружен. Колонки набора данных: {list(df.columns)}"
        )

    # 2. Удаление столбцов с большим числом пропусков
    df = drop_high_na_columns(df, DROP_NA_THRESHOLD, target_column=target_col)
    numeric_cols = [c for c in numeric_cols if c in df.columns]
    
    # 3. Заполнение пропусков
    df = fill_missing(df, numeric_cols)
    
    # 4. Построение признаков
    df = engineer_features(df)
    
    # Обновляем список числовых столбцов
    feature_cols = [
        c for c in df.select_dtypes(include=[np.number]).columns if c != target_col
    ]
    
    # 5. Обработка выбросов (ограничение, не удаление)
    df = handle_outliers(df, feature_cols, OUTLIER_STD_THRESHOLD, target_column=target_col)
    
    # 6. Удаляем строки с NaN (после построения признаков могли появиться)
    df = df.dropna(subset=[target_col])
    
    # 7. Матрица признаков и целевая переменная
    X = df[feature_cols].copy()
    y = df[target_col]
    
    # 8. Корреляционный анализ
    corr_table = correlation_table(X, y)
    
    # Топ-K признаков
    selected = select_top_features(corr_table, K_BEST)
    
    print("=" * 60)
    print("Лабораторная работа №3. Построение и отбор признаков")
    print("=" * 60)
    print(f"Строк: {len(df)}, Признаков: {len(feature_cols)}")
    print(f"Топ-{K_BEST} признаков по корреляции с '{target_col}':")
    for i, row in corr_table.head(K_BEST).iterrows():
        print(f"  {row['Признак']}: {row['Коэффициент_корреляции']:.4f}")
    print("=" * 60)
    
    # Обработанный датасет (с Date и целевой переменной)
    processed_df = df.copy()
    
    return processed_df, corr_table


# =============================================================================
# Точка входа
# =============================================================================
if __name__ == "__main__":
    processed_df, corr_table = run()
    
    if IN_LOGINOM:
        # Вывод в Loginom: таблица корреляций (Признак — Коэффициент_корреляции)
        try:
            prepare_compatible_table(OutputTable, corr_table, with_index=False)
            fill_table(OutputTable, corr_table, with_index=False)
        except Exception:
            pass
    else:
        # Локальный запуск: сохранение результатов
        corr_table.to_csv("correlation_results.csv", index=False)
        print("Результаты сохранены в correlation_results.csv")
