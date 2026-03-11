

from __future__ import annotations

import argparse
import os
from pathlib import Path
from typing import Iterable, List, Sequence


def _ensure_outdir(path: str) -> Path:
    outdir = Path(path).expanduser().resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    return outdir


def _parse_int_list(s: str) -> List[int]:
    # пример: "2,5,10" -> [2,5,10]
    parts = [p.strip() for p in s.split(",") if p.strip()]
    out: List[int] = []
    for p in parts:
        out.append(int(p))
    return out


def corr_vs_T(xs: Sequence[float], s: int, Ts: Iterable[int]):
    """
    Оценка корреляции r(T) = corr(z_1..z_T, z_{1+s}..z_{T}) для заданного шага s,
    как функция T.

    Реализовано эффективно через накапливаемые суммы по парам (x_i, y_i) = (z_i, z_{i+s}).
    Для фиксированного T число пар равно n_pairs = T - s.
    """
    if s <= 0:
        raise ValueError("s must be >= 1")
    n = len(xs)
    if n <= s + 2:
        raise ValueError("n is too small for this s")

    # running sums for pairs (x_i, y_i), i = 0..(t-1)
    sum_x = 0.0
    sum_y = 0.0
    sum_x2 = 0.0
    sum_y2 = 0.0
    sum_xy = 0.0
    t_pairs = 0  # number of accumulated pairs

    # We will generate answers in increasing T; require Ts increasing for efficiency.
    Ts_sorted = list(Ts)
    if Ts_sorted != sorted(Ts_sorted):
        raise ValueError("Ts must be sorted ascending")

    results = []
    next_pair_idx = 0  # next i to add pair (xs[i], xs[i+s])

    for T in Ts_sorted:
        if T > n:
            break
        if T <= s + 1:
            results.append(0.0)
            continue

        target_pairs = T - s  # number of pairs for this T
        while t_pairs < target_pairs and (next_pair_idx + s) < n:
            x = float(xs[next_pair_idx])
            y = float(xs[next_pair_idx + s])
            next_pair_idx += 1
            t_pairs += 1
            sum_x += x
            sum_y += y
            sum_x2 += x * x
            sum_y2 += y * y
            sum_xy += x * y

        # Pearson correlation using running sums:
        # r = (t*sum_xy - sum_x*sum_y) / sqrt((t*sum_x2 - sum_x^2)*(t*sum_y2 - sum_y^2))
        t = float(t_pairs)
        num = t * sum_xy - sum_x * sum_y
        den_x = t * sum_x2 - sum_x * sum_x
        den_y = t * sum_y2 - sum_y * sum_y
        denom = (den_x * den_y) ** 0.5 if den_x > 0 and den_y > 0 else 0.0
        r = num / denom if denom != 0.0 else 0.0
        results.append(r)

    return Ts_sorted[: len(results)], results


def main() -> int:
    ap = argparse.ArgumentParser(description="ЛР1: построение графиков (гистограмма/частоты/корреляции) для LCG")
    ap.add_argument("--n", type=int, default=200_000, help="объём выборки")
    ap.add_argument("--k", type=int, default=20, help="число интервалов/бинов на [0,1)")
    ap.add_argument("--s", type=int, default=1, help="шаг s для пар (z_i, z_{i+s})")
    ap.add_argument("--max-lag", type=int, default=30, help="максимальный лаг для графика автокорреляции")
    ap.add_argument(
        "--s-list",
        type=str,
        default="2,5,10",
        help="список s для графика R(T) (пример: 2,5,10)",
    )
    ap.add_argument("--t-max", type=int, default=2000, help="максимальный T для графика R(T)")
    ap.add_argument("--t-step", type=int, default=5, help="шаг по T для графика R(T)")
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

    # --- 1) Сравнение теоретических и экспериментальных M и D ---
    m_hat = lab.mean(xs)
    d_hat = lab.variance_population(xs)
    m_theor = lab.THEORETICAL_MEAN
    d_theor = lab.THEORETICAL_VAR

    plt.figure(figsize=(8, 5))
    labels = ["M(z)", "D(z)"]
    x_pos = np.arange(len(labels))
    exp_vals = [m_hat, d_hat]
    theor_vals = [m_theor, d_theor]

    width = 0.35
    plt.bar(x_pos - width / 2, theor_vals, width, label="Теоретическое значение")
    plt.bar(x_pos + width / 2, exp_vals, width, label="Экспериментальная оценка")

    # немного увеличиваем «зум», чтобы различия были видны
    all_vals = theor_vals + exp_vals
    y_min = min(all_vals)
    y_max = max(all_vals)
    margin = 0.2 * (y_max - y_min) if y_max > y_min else 0.02
    plt.ylim(y_min - margin, y_max + margin)

    # подпишем точные значения над столбцами
    for i, (t_val, e_val) in enumerate(zip(theor_vals, exp_vals)):
        plt.text(i - width / 2, t_val, f"{t_val:.4f}", ha="center", va="bottom", fontsize=9, color="navy")
        plt.text(i + width / 2, e_val, f"{e_val:.4f}", ha="center", va="bottom", fontsize=9, color="darkorange")

    plt.xticks(x_pos, labels)
    plt.ylabel("Значение")
    plt.title("Сравнение теоретических и экспериментальных значений M и D")
    plt.grid(True, axis="y", alpha=0.25)
    plt.legend()
    md_path = outdir / "md_compare.png"
    plt.tight_layout()
    plt.savefig(md_path, dpi=160)

    # --- 2) Гистограмма (как в отчётах: относительные частоты по интервалам) ---
    # Важно: делим [0,1] на K равных интервалов => ширина интервала = 1/K.
    # Высота столбца = (число попаданий в интервал) / n.
    plt.figure(figsize=(10, 5))
    edges = np.linspace(0.0, 1.0, args.k + 1)
    weights = np.ones_like(x) / float(len(x))  # чтобы получить относительные частоты
    plt.hist(x, bins=edges, weights=weights, edgecolor="black", alpha=0.85)
    plt.axhline(1.0 / args.k, color="red", linestyle="--", linewidth=1.5, label="Теоретическая частота (1/K)")
    plt.title("Гистограмма распределения БСВ")
    plt.xlabel("Интервал")
    plt.ylabel("Относительная частота")
    plt.grid(True, alpha=0.25)
    plt.legend()
    hist_path = outdir / "histogram.png"
    plt.tight_layout()
    plt.savefig(hist_path, dpi=160)

    # --- 3) Частоты по интервалам (bar) ---
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

    # --- 4) Диаграмма рассеяния (z_i, z_{i+s}) ---
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

    # --- 5) Автокорреляция по лагам ---
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

    # --- 6) R_hat(T) от T для нескольких s ---
    s_list = [v for v in _parse_int_list(args.s_list) if v >= 1]
    t_max = max(10, int(args.t_max))
    t_step = max(1, int(args.t_step))
    t_max = min(t_max, args.n)
    Ts = list(range(1, t_max + 1, t_step))

    plt.figure(figsize=(10, 5))
    for s_val in s_list:
        if s_val >= args.n - 2:
            continue
        TT, RR = corr_vs_T(xs, s_val, Ts)
        plt.plot(TT, RR, linewidth=1.5, label=f"s={s_val}")
    plt.axhline(0.0, color="black", linewidth=1.0)
    plt.title("Графики зависимости R от T")
    plt.xlabel("T")
    plt.ylabel("R(T)")
    plt.grid(True, alpha=0.25)
    plt.legend()
    rt_path = outdir / "corr_vs_T.png"
    plt.tight_layout()
    plt.savefig(rt_path, dpi=160)

    print("Графики сохранены в:", str(outdir))
    print("-", md_path.name)
    print("-", hist_path.name)
    print("-", freq_path.name)
    print("-", scatter_path.name)
    print("-", ac_path.name)
    print("-", rt_path.name)

    if args.show:
        plt.show()
    else:
        # чтобы не держать ресурсы при пакетном запуске
        plt.close("all")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())


