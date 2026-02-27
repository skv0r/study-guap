#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Графики для ЛР1 (БСВ z ~ R[0,1]) для варианта 1 (LCG).

Скрипт использует генератор из `lab.py` и строит через matplotlib:
- Гистограмму выборки z (bins=K)
- Столбчатую диаграмму частот по K интервалам [0,1)
- Диаграмму рассеяния пар (z_i, z_{i+s}) (визуальная проверка независимости)
- Автокорреляцию corr(z_i, z_{i+lag}) по lag=1..max_lag

Запуск (из корня репо):
  python3 "6_semester/Моделирование систем/plot_lab1.py" --n 200000 --k 20 --s 1 --max-lag 30

Зависимости:
  pip install matplotlib numpy
"""

from __future__ import annotations

import argparse
import os
from pathlib import Path
from typing import List


def _ensure_outdir(path: str) -> Path:
    outdir = Path(path).expanduser().resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    return outdir


def main() -> int:
    ap = argparse.ArgumentParser(description="ЛР1: построение графиков (гистограмма/частоты/корреляции) для LCG")
    ap.add_argument("--n", type=int, default=200_000, help="объём выборки")
    ap.add_argument("--k", type=int, default=20, help="число интервалов/бинов на [0,1)")
    ap.add_argument("--s", type=int, default=1, help="шаг s для пар (z_i, z_{i+s})")
    ap.add_argument("--max-lag", type=int, default=30, help="максимальный лаг для графика автокорреляции")
    ap.add_argument("--seed", type=int, default=123456789, help="стартовое значение X0")
    ap.add_argument("--high-bits", type=int, default=24, help="сколько старших бит использовать для z")
    ap.add_argument(
        "--outdir",
        type=str,
        default=str(Path(__file__).with_name("plots")),
        help="директория для сохранения PNG",
    )
    ap.add_argument("--show", action="store_true", help="показать окна графиков (если среда поддерживает GUI)")
    args = ap.parse_args()

    try:
        import numpy as np
        import matplotlib.pyplot as plt
    except ModuleNotFoundError as e:
        missing = getattr(e, "name", "matplotlib/numpy")
        print(f"Не найдена зависимость: {missing}")
        print("Установите зависимости и повторите запуск:")
        print("  python3 -m pip install matplotlib numpy")
        return 2

    # Важно: `lab.py` лежит рядом, поэтому импорт должен работать при запуске этого файла.
    import lab  # type: ignore

    outdir = _ensure_outdir(args.outdir)

    xs: List[float] = lab.lcg_sample(
        seed=args.seed,
        n=args.n,
        params=lab.LCGParams(),
        high_bits=args.high_bits,
    )
    x = np.asarray(xs, dtype=float)

    # --- 1) Гистограмма ---
    plt.figure(figsize=(10, 5))
    plt.hist(x, bins=args.k, range=(0.0, 1.0), density=True, edgecolor="black", alpha=0.75)
    plt.title(f"Гистограмма выборки z (n={args.n}, bins={args.k})")
    plt.xlabel("z")
    plt.ylabel("Плотность (нормировано)")
    plt.grid(True, alpha=0.25)
    hist_path = outdir / "histogram.png"
    plt.tight_layout()
    plt.savefig(hist_path, dpi=160)

    # --- 2) Частоты по интервалам (bar) ---
    counts, freqs = lab.frequency_test(xs, args.k)
    centers = np.arange(args.k)
    plt.figure(figsize=(10, 5))
    plt.bar(centers, freqs, width=0.9, edgecolor="black", alpha=0.8)
    plt.axhline(1.0 / args.k, color="red", linestyle="--", linewidth=1.5, label="теория = 1/K")
    plt.title(f"Относительные частоты по K интервалам (K={args.k})")
    plt.xlabel("Номер интервала")
    plt.ylabel("Частота")
    plt.grid(True, axis="y", alpha=0.25)
    plt.legend()
    freq_path = outdir / "frequencies_bar.png"
    plt.tight_layout()
    plt.savefig(freq_path, dpi=160)

    # --- 3) Диаграмма рассеяния (z_i, z_{i+s}) ---
    s = args.s
    # Чтобы scatter не был слишком тяжёлым, ограничим количество точек
    max_points = min(30_000, max(0, args.n - s))
    a = x[:max_points]
    b = x[s : s + max_points]
    plt.figure(figsize=(6.5, 6.5))
    plt.scatter(a, b, s=3, alpha=0.35)
    plt.title(f"Точечный график пар (z_i, z_(i+{s}))")
    plt.xlabel("z_i")
    plt.ylabel(f"z_(i+{s})")
    plt.xlim(0.0, 1.0)
    plt.ylim(0.0, 1.0)
    plt.grid(True, alpha=0.2)
    scatter_path = outdir / f"scatter_pairs_s{s}.png"
    plt.tight_layout()
    plt.savefig(scatter_path, dpi=160)

    # --- 4) Автокорреляция по лагам ---
    max_lag = max(1, int(args.max_lag))
    lags = np.arange(1, max_lag + 1)
    rs = np.array([lab.correlation(xs, int(lag)) for lag in lags], dtype=float)
    plt.figure(figsize=(10, 5))
    plt.plot(lags, rs, marker="o", linewidth=1.5)
    plt.axhline(0.0, color="black", linewidth=1.0)
    plt.title("Автокорреляция corr(z_i, z_(i+lag))")
    plt.xlabel("lag")
    plt.ylabel("corr")
    plt.grid(True, alpha=0.25)
    ac_path = outdir / "autocorr.png"
    plt.tight_layout()
    plt.savefig(ac_path, dpi=160)

    print("Графики сохранены в:", str(outdir))
    print("-", hist_path.name)
    print("-", freq_path.name)
    print("-", scatter_path.name)
    print("-", ac_path.name)

    if args.show:
        plt.show()
    else:
        # чтобы не держать ресурсы при пакетном запуске
        plt.close("all")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())


