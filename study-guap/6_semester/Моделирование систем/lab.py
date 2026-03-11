#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Лабораторная работа №1. Моделирование базовой случайной величины (БСВ) z ~ R[0,1].

Вариант 1: линейный конгруэнтный метод (LCG):
    X_{k+1} = (a * X_k + c) mod m

Рекомендуемые константы (из задания):
    a = 1664525
    c = 1013904223
    m = 2^32

Требование: использовать только старшие разряды (уменьшает проблемы низших битов LCG).

В отчёте/выводе выполняются пункты 1–4 (по смыслу методички):
1) Построить датчик БСВ (генератор z in [0,1)).
2) Получить выборку объёма n.
3) Проверить равномерность: частотный тест на K интервалах, гистограмма частот.
4) Проверить независимость: оценка линейной корреляции между z_i и z_{i+s}.

Также оцениваются мат. ожидание и дисперсия и сравниваются с теорией:
    M(z) = 0.5
    D(z) = 1/12
"""

from __future__ import annotations

import argparse
import math
from dataclasses import dataclass
from typing import Iterable, List, Tuple


THEORETICAL_MEAN = 0.5
THEORETICAL_VAR = 1.0 / 12.0


@dataclass(frozen=True)
class LCGParams:
    a: int = 1664525
    c: int = 1013904223
    m: int = 2**32


def lcg_next(x: int, p: LCGParams) -> int:
    # Делаем вычисления в 32-битной арифметике, как для m=2^32
    return (p.a * x + p.c) & (p.m - 1)


def lcg_sample(
    *,
    seed: int,
    n: int,
    params: LCGParams,
    high_bits: int = 24,
) -> List[float]:
    """
    Генерирует n значений БСВ z in [0,1) из LCG.

    high_bits: сколько старших бит использовать (<= 32).
    Например, high_bits=24: берём биты [31..8] и делим на 2^24.
    """
    if n <= 0:
        raise ValueError("n must be > 0")
    if not (1 <= high_bits <= 32):
        raise ValueError("high_bits must be in [1, 32]")

    x = seed & (params.m - 1)
    shift = 32 - high_bits
    denom = float(1 << high_bits)

    out: List[float] = []
    for _ in range(n):
        x = lcg_next(x, params)
        # используем старшие разряды
        z_int = x >> shift
        z = z_int / denom  # [0, 1)
        out.append(z)
    return out


def mean(xs: Iterable[float]) -> float:
    xs = list(xs)
    if not xs:
        raise ValueError("empty sample")
    return sum(xs) / len(xs)


def variance_population(xs: Iterable[float]) -> float:
    """Оценка дисперсии с делением на n (как для сравнения с теоретической D)."""
    xs = list(xs)
    n = len(xs)
    if n == 0:
        raise ValueError("empty sample")
    m = sum(xs) / n
    return sum((x - m) ** 2 for x in xs) / n


def variance_unbiased(xs: Iterable[float]) -> float:
    """Несмещённая оценка дисперсии (деление на n-1)."""
    xs = list(xs)
    n = len(xs)
    if n < 2:
        raise ValueError("need at least 2 samples")
    m = sum(xs) / n
    return sum((x - m) ** 2 for x in xs) / (n - 1)


def frequency_test(xs: List[float], k: int) -> Tuple[List[int], List[float]]:
    """
    Частотный тест равномерности на [0,1) по k равным интервалам.
    Возвращает counts и relative frequencies.
    """
    if k <= 0:
        raise ValueError("k must be > 0")
    counts = [0] * k
    for x in xs:
        if x < 0.0 or x >= 1.0:
            # теоретически не должно быть, но защитимся
            idx = min(max(int(x * k), 0), k - 1)
        else:
            idx = int(x * k)
            if idx == k:
                idx = k - 1
        counts[idx] += 1
    n = len(xs)
    freqs = [c / n for c in counts]
    return counts, freqs


def ascii_histogram(freqs: List[float], width: int = 50) -> str:
    """Простая текстовая гистограмма по относительным частотам."""
    if not freqs:
        return ""
    mx = max(freqs)
    if mx <= 0:
        mx = 1.0
    lines: List[str] = []
    for i, f in enumerate(freqs):
        bar_len = int(round((f / mx) * width))
        bar = "#" * bar_len
        lines.append(f"{i:02d}: {f:0.6f} |{bar}")
    return "\n".join(lines)


def correlation(xs: List[float], s: int) -> float:
    """
    Оценка линейной корреляции r между (x_0..x_{n-s-1}) и (x_s..x_{n-1}).
    """
    n = len(xs)
    if s <= 0:
        raise ValueError("s must be >= 1")
    if n - s < 2:
        raise ValueError("n must be large enough for given s")

    x1 = xs[: n - s]
    x2 = xs[s:]
    m1 = mean(x1)
    m2 = mean(x2)

    num = sum((a - m1) * (b - m2) for a, b in zip(x1, x2))
    den1 = math.sqrt(sum((a - m1) ** 2 for a in x1))
    den2 = math.sqrt(sum((b - m2) ** 2 for b in x2))
    denom = den1 * den2
    if denom == 0:
        return 0.0
    return num / denom


def fmt_diff(estimate: float, theoretical: float) -> str:
    return f"оценка={estimate:.8f}, теория={theoretical:.8f}, |Δ|={abs(estimate - theoretical):.8e}"


def main() -> int:
    ap = argparse.ArgumentParser(
        description="ЛР1: датчик БСВ R[0,1] (вариант 1, LCG) + тесты равномерности и независимости"
    )
    ap.add_argument("--n", type=int, default=100_000, help="объём выборки")
    ap.add_argument("--k", type=int, default=10, help="число интервалов для частотного теста")
    ap.add_argument("--s", type=int, default=1, help="шаг s для корреляционного теста (>=1)")
    ap.add_argument("--seed", type=int, default=123456789, help="стартовое значение X0")
    ap.add_argument(
        "--high-bits",
        type=int,
        default=24,
        help="сколько старших бит использовать для z (рекомендуется 24..31)",
    )
    args = ap.parse_args()

    params = LCGParams()

    print("Лабораторная работа №1 — Моделирование базовой случайной величины (БСВ) z ~ R[0,1]")
    print("Вариант 1: линейный конгруэнтный метод (LCG)")
    print(f"Параметры: a={params.a}, c={params.c}, m=2^32, seed(X0)={args.seed}, high_bits={args.high_bits}")
    print(f"Настройки тестов: n={args.n}, K={args.k}, s={args.s}")
    print()

    # 1–2) Датчик + выборка
    xs = lcg_sample(seed=args.seed, n=args.n, params=params, high_bits=args.high_bits)

    # 3) Мат. ожидание и дисперсия
    m_hat = mean(xs)
    d_hat_n = variance_population(xs)
    d_hat_n1 = variance_unbiased(xs)

    print("Проверка основных свойств БСВ (M и D):")
    print(f"- M(z): {fmt_diff(m_hat, THEORETICAL_MEAN)}")
    print(f"- D(z) (деление на n): {fmt_diff(d_hat_n, THEORETICAL_VAR)}")
    print(f"- D(z) (несмещ., деление на n-1): оценка={d_hat_n1:.8f} (для справки)")
    print()

    # 4) Частотный тест (равномерность)
    counts, freqs = frequency_test(xs, args.k)
    print("Частотный тест равномерности на [0,1):")
    print("- counts:", counts)
    print("- freqs :", [round(f, 6) for f in freqs])
    print("Гистограмма относительных частот (ASCII):")
    print(ascii_histogram(freqs))
    print()

    # 5) Тест независимости (корреляция)
    r = correlation(xs, args.s)
    print("Тест статистической независимости (линейная корреляция):")
    print(f"- corr(z_i, z_(i+{args.s})) = {r:.8f}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

