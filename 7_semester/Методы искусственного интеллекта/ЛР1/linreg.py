# -*- coding: utf-8 -*-
"""Линейная регрессия времени доставки: линейная выборка 20 % и МНК."""
from __future__ import annotations

import json
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

BASE = Path(__file__).resolve().parent
DATA = BASE / "data" / "courier_delivery.csv"
FIG = BASE / "figures"
OUT = BASE / "data"


def linear_indices(n: int, fraction: float = 0.2) -> list[int]:
    """Отбор строк как в LinearSamplingRowFilter: доля по всей таблице.

    Первая и последняя строки входят в отобранную часть. Остальные берутся,
    когда накопленная доля переходит через целое.
    """
    weight = fraction
    left = int(round(n * fraction))
    acc = 0.0
    chosen: list[int] = []
    for i in range(n):
        prev = acc
        acc += weight
        if i == 0 or i == n - 1:
            left -= 1
            chosen.append(i)
            continue
        if left == 1:
            continue
        if int(prev) < int(acc):
            left -= 1
            chosen.append(i)
    if left != 0:
        raise RuntimeError(f"линейная выборка не сошлась, остаток {left}")
    return chosen


def round_half_up(value: float) -> int:
    return int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def metrics(y: np.ndarray, pred: np.ndarray) -> dict:
    err = pred - y
    resid = y - pred
    rss = float(np.sum(resid**2))
    tss = float(np.sum((y - y.mean()) ** 2))
    r2 = 1.0 - rss / tss
    n = int(len(y))
    adj = 1.0 - (1.0 - r2) * (n - 1) / (n - 2)
    return {
        "n": n,
        "R2": r2,
        "adj_R2": adj,
        "MAE": float(np.mean(np.abs(resid))),
        "MSE": float(np.mean(resid**2)),
        "RMSE": float(np.sqrt(np.mean(resid**2))),
        "MAPE_pct": float(np.mean(np.abs(resid) / np.abs(y)) * 100),
        "mean_pred_minus_fact": float(np.mean(err)),
    }


def draw_scheme(path: Path) -> None:
    labels = [
        "Чтение таблицы заказов",
        "Разбиение: 20 %, линейная выборка",
        "Цвет: синий — обучение, оранжевый — тест",
        "Оценка прямой по 20 обучающим строкам",
        "Прогноз на 80 тестовых строках",
        "Показатели качества и диаграмма рассеяния",
        "Округление прогноза до целой минуты",
        "Таблица результата без дробного прогноза",
    ]
    fig, ax = plt.subplots(figsize=(7.2, 9.0))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, len(labels) + 0.4)
    ax.axis("off")
    fig.patch.set_facecolor("white")
    for i, text in enumerate(labels):
        y = len(labels) - i - 0.35
        box = FancyBboxPatch(
            (1.1, y - 0.38),
            7.8,
            0.76,
            boxstyle="square,pad=0",
            linewidth=1.0,
            edgecolor="black",
            facecolor="white",
        )
        ax.add_patch(box)
        ax.text(5.0, y, text, ha="center", va="center", fontsize=12, color="black")
        if i < len(labels) - 1:
            ax.add_patch(
                FancyArrowPatch(
                    (5.0, y - 0.42),
                    (5.0, y - 0.62),
                    arrowstyle="-|>",
                    mutation_scale=12,
                    linewidth=1.0,
                    color="black",
                )
            )
    fig.tight_layout()
    fig.savefig(path, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def draw_scatter(df: pd.DataFrame, train_idx: list[int], b0: float, b1: float, path: Path) -> None:
    train = df.iloc[train_idx]
    test = df.drop(df.index[train_idx])
    fig, ax = plt.subplots(figsize=(7.4, 5.2))
    ax.scatter(
        test["distance_km"],
        test["time_min"],
        s=28,
        c="#c46b1a",
        label="Тест, 80 заказов",
        zorder=2,
    )
    ax.scatter(
        train["distance_km"],
        train["time_min"],
        s=42,
        c="#1f4e79",
        label="Обучение, 20 заказов",
        zorder=3,
    )
    xs = np.linspace(df["distance_km"].min(), df["distance_km"].max(), 100)
    ax.plot(xs, b0 + b1 * xs, color="black", linewidth=1.2, label="Прямая регрессии", zorder=1)
    ax.set_xlabel("Расстояние, км", fontsize=12)
    ax.set_ylabel("Время доставки, мин", fontsize=12)
    ax.legend(frameon=False, fontsize=11)
    ax.grid(True, linewidth=0.4, color="#cccccc")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    fig.savefig(path, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def main() -> None:
    plt.rcParams["font.family"] = "Times New Roman"
    df = pd.read_csv(DATA)
    n = len(df)
    idx = linear_indices(n, 0.2)
    train = df.iloc[idx]
    test = df.drop(df.index[idx]).copy()
    x = train["distance_km"].to_numpy(float)
    y = train["time_min"].to_numpy(float)
    design = np.column_stack([np.ones(len(x)), x])
    beta, *_ = np.linalg.lstsq(design, y, rcond=None)
    b0, b1 = float(beta[0]), float(beta[1])
    pred = b0 + b1 * test["distance_km"].to_numpy(float)
    actual = test["time_min"].to_numpy(float)
    scored = metrics(actual, pred)
    rounded = np.array([round_half_up(v) for v in pred], dtype=float)
    scored_rounded = metrics(actual, rounded)

    test_out = test.copy()
    test_out["prediction"] = pred
    test_out["prediction_min"] = rounded.astype(int)
    test_out.to_csv(OUT / "predictions.csv", index=False)

    result = {
        "n": n,
        "n_train": int(len(train)),
        "n_test": int(len(test)),
        "train_order_ids": train["order_id"].astype(int).tolist(),
        "intercept": b0,
        "slope": b1,
        "distance_min": float(df["distance_km"].min()),
        "distance_max": float(df["distance_km"].max()),
        "time_min": float(df["time_min"].min()),
        "time_max": float(df["time_min"].max()),
        "train_distance_min": float(train["distance_km"].min()),
        "train_distance_max": float(train["distance_km"].max()),
        "test_distance_min": float(test["distance_km"].min()),
        "test_distance_max": float(test["distance_km"].max()),
        "sorted_by_distance": bool(df["distance_km"].is_monotonic_increasing),
        "test": scored,
        "test_rounded": scored_rounded,
    }
    (OUT / "metrics.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    FIG.mkdir(exist_ok=True)
    draw_scheme(FIG / "scheme.png")
    draw_scatter(df, idx, b0, b1, FIG / "scatter.png")
    print(json.dumps({k: result[k] for k in ("intercept", "slope", "n_train", "n_test")}, ensure_ascii=False))
    print("test", {k: round(v, 5) if isinstance(v, float) else v for k, v in scored.items()})


if __name__ == "__main__":
    main()
