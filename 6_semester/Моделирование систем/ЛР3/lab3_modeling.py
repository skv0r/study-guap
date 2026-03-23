from __future__ import annotations

from dataclasses import dataclass
from math import factorial, gamma, pi, sqrt
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np


N_SAMPLES = 50000
RNG_SEED = 42


@dataclass
class DistResult:
    name: str
    sample: np.ndarray
    m_theoretical: float
    d_theoretical: float
    m_empirical: float
    d_empirical: float


def pdf_exponential(x: np.ndarray, lam: float) -> np.ndarray:
    y = np.zeros_like(x)
    mask = x >= 0
    y[mask] = lam * np.exp(-lam * x[mask])
    return y


def pdf_uniform(x: np.ndarray, a: float, b: float) -> np.ndarray:
    y = np.zeros_like(x)
    mask = (x >= a) & (x <= b)
    y[mask] = 1.0 / (b - a)
    return y


def pdf_erlang(x: np.ndarray, k: int, lam: float) -> np.ndarray:
    y = np.zeros_like(x)
    mask = x >= 0
    y[mask] = (lam**k) * (x[mask] ** (k - 1)) * np.exp(-lam * x[mask]) / factorial(k - 1)
    return y


def pdf_normal(x: np.ndarray, m: float, sigma: float) -> np.ndarray:
    return np.exp(-((x - m) ** 2) / (2 * sigma**2)) / (sigma * sqrt(2 * pi))


def pdf_beta(x: np.ndarray, alpha: int, beta: int) -> np.ndarray:
    y = np.zeros_like(x)
    mask = (x >= 0) & (x <= 1)
    beta_func = gamma(alpha) * gamma(beta) / gamma(alpha + beta)
    y[mask] = (x[mask] ** (alpha - 1)) * ((1 - x[mask]) ** (beta - 1)) / beta_func
    return y


def sample_erlang_from_uniform(rng: np.random.Generator, n: int, k: int, lam: float) -> np.ndarray:
    u = rng.random((n, k))
    return -np.log(u).sum(axis=1) / lam


def sample_normal_box_muller(rng: np.random.Generator, n: int, m: float, sigma: float) -> np.ndarray:
    n2 = n if n % 2 == 0 else n + 1
    u1 = rng.random(n2 // 2)
    u2 = rng.random(n2 // 2)
    r = np.sqrt(-2.0 * np.log(u1))
    theta = 2.0 * np.pi * u2
    z1 = r * np.cos(theta)
    z2 = r * np.sin(theta)
    z = np.concatenate([z1, z2])[:n]
    return m + sigma * z


def sample_beta_integer_shape(rng: np.random.Generator, n: int, alpha: int, beta: int) -> np.ndarray:
    # Для целых alpha и beta: Beta(alpha, beta) = G1 / (G1 + G2),
    # где G1 и G2 — гамма-распределения (суммы экспоненциальных СВ)
    g1 = -np.log(rng.random((n, alpha))).sum(axis=1)
    g2 = -np.log(rng.random((n, beta))).sum(axis=1)
    return g1 / (g1 + g2)


def evaluate_distribution(name: str, sample: np.ndarray, m_theoretical: float, d_theoretical: float) -> DistResult:
    return DistResult(
        name=name,
        sample=sample,
        m_theoretical=m_theoretical,
        d_theoretical=d_theoretical,
        m_empirical=float(np.mean(sample)),
        d_empirical=float(np.var(sample)),
    )


def save_hist_with_pdf(
    result: DistResult,
    x_grid: np.ndarray,
    pdf_y: np.ndarray,
    file_path: Path,
    x_label: str = "x",
) -> None:
    fig, ax = plt.subplots(figsize=(9.5, 5.5))
    ax.hist(result.sample, bins=50, density=True, alpha=0.65, color="#4C72B0", edgecolor="white")
    ax.plot(x_grid, pdf_y, color="#C44E52", linewidth=2.2)
    ax.set_title(f"{result.name}: эмпирическая гистограмма и теоретическая плотность")
    ax.set_xlabel(x_label)
    ax.set_ylabel("Плотность вероятности")
    ax.grid(axis="y", linestyle="--", alpha=0.45)
    fig.tight_layout()
    fig.savefig(file_path, dpi=200)
    plt.close(fig)


def save_summary_table(results: list[DistResult], file_path: Path) -> None:
    rows = []
    for r in results:
        rows.append(
            [
                r.name,
                f"{r.m_theoretical:.6f}",
                f"{r.m_empirical:.6f}",
                f"{abs(r.m_theoretical - r.m_empirical):.6f}",
                f"{r.d_theoretical:.6f}",
                f"{r.d_empirical:.6f}",
                f"{abs(r.d_theoretical - r.d_empirical):.6f}",
            ]
        )

    fig, ax = plt.subplots(figsize=(13, 4.6))
    ax.axis("off")
    ax.set_title("Сравнение теоретических и эмпирических характеристик (ЛР3, вариант 1)", pad=12)
    table = ax.table(
        cellText=rows,
        colLabels=["Распределение", "M теор", "M эмп", "|ΔM|", "D теор", "D эмп", "|ΔD|"],
        cellLoc="center",
        colLoc="center",
        loc="center",
    )
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1, 1.45)
    fig.tight_layout()
    fig.savefig(file_path, dpi=220)
    plt.close(fig)


def build_report(results: list[DistResult], output_path: Path) -> None:
    exp_r, uni_r, erl_r, norm_r, beta_r = results
    report = f"""ЛАБОРАТОРНАЯ РАБОТА 3
Моделирование непрерывной случайной величины
Вариант 1

Цель
Целью лабораторной работы является реализация генераторов непрерывных случайных величин на основе базовой случайной величины и проверка соответствия эмпирических характеристик теоретическим значениям для различных законов распределения.

Задание и исходные данные
В соответствии с методическими указаниями требуется построить пять генераторов непрерывных случайных величин и исследовать их статистические свойства. В работе реализованы экспоненциальное распределение, равномерное распределение, распределение Эрланга, нормальное распределение и распределение, заданное вариантом 1. Для варианта 1 задано бета распределение, которое было смоделировано с параметрами alpha равно 2 и beta равно 5 на интервале от 0 до 1.

Ход выполнения задания
На первом этапе были реализованы формулы моделирования для всех требуемых распределений. Экспоненциальное, равномерное и Эрланговское распределения построены по формулам методических указаний через преобразование базовой случайной величины. Нормальное распределение реализовано методом Бокса Мюллера, а для бета распределения применено представление через отношение двух гамма величин, каждая из которых получалась как сумма экспоненциальных случайных величин.

На втором этапе для каждого закона распределения была сформирована выборка большого объема, после чего вычислены эмпирические оценки математического ожидания и дисперсии. Теоретические значения M и D рассчитаны по известным аналитическим формулам для соответствующих распределений и использованы для прямого сравнения с результатами моделирования.

На заключительном этапе построены гистограммы распределений и выполнен визуальный анализ соответствия формы эмпирических плотностей теоретическим кривым. Дополнительно подготовлена сводная таблица, где для всех пяти распределений представлены теоретические и эмпирические значения M и D, а также абсолютные расхождения.

3.1 Построение генераторов непрерывных случайных величин
Для каждого вида распределения использован собственный способ преобразования значений базовой случайной величины. Экспоненциальная величина получалась через логарифмическое преобразование, равномерная через линейное масштабирование на заданный интервал, Эрланговская как сумма независимых экспоненциальных величин фиксированного порядка, нормальная через преобразование пары независимых равномерных величин по методу Бокса Мюллера, а бета распределение через отношение двух гамма величин с целыми параметрами. Такой подход обеспечивает единый принцип моделирования и позволяет корректно получить все пять требуемых распределений в рамках одной программы.

3.2 Гистограммы распределений непрерывной случайной величины
Для каждой выборки построена гистограмма эмпирической плотности вероятности и наложена соответствующая теоретическая кривая. По результатам визуального сравнения видно, что формы гистограмм согласуются с ожидаемыми законами: экспоненциальная плотность убывает на положительной полуоси, равномерная остается практически постоянной на заданном интервале, распределение Эрланга имеет характерную правостороннюю асимметрию, нормальная плотность симметрична относительно среднего значения, а бета распределение сосредоточено в области малых значений на интервале от 0 до 1. Небольшие локальные колебания столбцов объясняются случайной природой конечной выборки.

На рисунке 1 изображена гистограмма и теоретическая плотность экспоненциального распределения.
Рисунок 1 - Экспоненциальное распределение
На рисунке 2 изображена гистограмма и теоретическая плотность равномерного распределения.
Рисунок 2 - Равномерное распределение
На рисунке 3 изображена гистограмма и теоретическая плотность распределения Эрланга.
Рисунок 3 - Распределение Эрланга
На рисунке 4 изображена гистограмма и теоретическая плотность нормального распределения.
Рисунок 4 - Нормальное распределение
На рисунке 5 изображена гистограмма и теоретическая плотность бета распределения для варианта 1.
Рисунок 5 - Бета распределение

3.3 Сравнение эмпирических и теоретических значений M и D
В ходе расчетов получено, что для экспоненциального распределения M теоретическое равно {exp_r.m_theoretical:.6f}, M эмпирическое равно {exp_r.m_empirical:.6f}, D теоретическое равно {exp_r.d_theoretical:.6f}, D эмпирическое равно {exp_r.d_empirical:.6f}. Для равномерного распределения M теоретическое равно {uni_r.m_theoretical:.6f}, M эмпирическое равно {uni_r.m_empirical:.6f}, D теоретическое равно {uni_r.d_theoretical:.6f}, D эмпирическое равно {uni_r.d_empirical:.6f}. Для распределения Эрланга M теоретическое равно {erl_r.m_theoretical:.6f}, M эмпирическое равно {erl_r.m_empirical:.6f}, D теоретическое равно {erl_r.d_theoretical:.6f}, D эмпирическое равно {erl_r.d_empirical:.6f}. Для нормального распределения M теоретическое равно {norm_r.m_theoretical:.6f}, M эмпирическое равно {norm_r.m_empirical:.6f}, D теоретическое равно {norm_r.d_theoretical:.6f}, D эмпирическое равно {norm_r.d_empirical:.6f}. Для бета распределения M теоретическое равно {beta_r.m_theoretical:.6f}, M эмпирическое равно {beta_r.m_empirical:.6f}, D теоретическое равно {beta_r.d_theoretical:.6f}, D эмпирическое равно {beta_r.d_empirical:.6f}.

Анализ показывает, что по всем пяти законам эмпирические значения близки к теоретическим. Наибольшие отклонения остаются малыми относительно масштаба соответствующих дисперсий, что подтверждает корректность программной реализации генераторов и достаточный объем выборки для устойчивой оценки характеристик.

На рисунке 6 изображена сводная таблица сравнения теоретических и эмпирических значений M и D для всех распределений.
Рисунок 6 - Сводные результаты сравнения M и D

Вывод
В лабораторной работе реализованы пять генераторов непрерывных случайных величин, включая распределение, заданное вариантом 1, и для каждого генератора выполнен статистический анализ результатов моделирования. Построенные гистограммы соответствуют теоретическим формам плотностей, что подтверждает корректность выбранных методов генерации.

Сравнение теоретических и эмпирических характеристик показало хорошее совпадение математических ожиданий и дисперсий по всем рассмотренным распределениям. Наблюдаемые расхождения находятся в пределах статистически ожидаемых отклонений для случайной выборки конечного объема.

Таким образом, поставленная цель работы достигнута, а разработанная программа может использоваться для дальнейших задач имитационного моделирования непрерывных случайных величин.
"""
    output_path.write_text(report, encoding="utf-8")


def main() -> None:
    out_dir = Path("plots")
    out_dir.mkdir(parents=True, exist_ok=True)
    rng = np.random.default_rng(RNG_SEED)

    # 1) Экспоненциальное распределение
    lam_exp = 1.5
    exp_sample = -np.log(rng.random(N_SAMPLES)) / lam_exp
    exp_result = evaluate_distribution(
        "Экспоненциальное",
        exp_sample,
        m_theoretical=1.0 / lam_exp,
        d_theoretical=1.0 / (lam_exp**2),
    )
    x_exp = np.linspace(0, np.quantile(exp_sample, 0.995), 400)
    save_hist_with_pdf(exp_result, x_exp, pdf_exponential(x_exp, lam_exp), out_dir / "1_exponential.png")

    # 2) Равномерное распределение
    a_uni, b_uni = -5.0, 7.0
    uni_sample = a_uni + (b_uni - a_uni) * rng.random(N_SAMPLES)
    uni_result = evaluate_distribution(
        "Равномерное",
        uni_sample,
        m_theoretical=(a_uni + b_uni) / 2.0,
        d_theoretical=((b_uni - a_uni) ** 2) / 12.0,
    )
    x_uni = np.linspace(a_uni - 1, b_uni + 1, 400)
    save_hist_with_pdf(uni_result, x_uni, pdf_uniform(x_uni, a_uni, b_uni), out_dir / "2_uniform.png")

    # 3) Распределение Эрланга
    k_erl, lam_erl = 4, 1.1
    erl_sample = sample_erlang_from_uniform(rng, N_SAMPLES, k_erl, lam_erl)
    erl_result = evaluate_distribution(
        "Эрланга",
        erl_sample,
        m_theoretical=k_erl / lam_erl,
        d_theoretical=k_erl / (lam_erl**2),
    )
    x_erl = np.linspace(0, np.quantile(erl_sample, 0.995), 450)
    save_hist_with_pdf(erl_result, x_erl, pdf_erlang(x_erl, k_erl, lam_erl), out_dir / "3_erlang.png")

    # 4) Нормальное распределение
    m_norm, sigma_norm = 1.5, 2.0
    norm_sample = sample_normal_box_muller(rng, N_SAMPLES, m_norm, sigma_norm)
    norm_result = evaluate_distribution(
        "Нормальное",
        norm_sample,
        m_theoretical=m_norm,
        d_theoretical=sigma_norm**2,
    )
    x_norm = np.linspace(np.quantile(norm_sample, 0.002), np.quantile(norm_sample, 0.998), 450)
    save_hist_with_pdf(norm_result, x_norm, pdf_normal(x_norm, m_norm, sigma_norm), out_dir / "4_normal.png")

    # 5) Вариант 1: Бета-распределение
    alpha_b, beta_b = 2, 5
    beta_sample = sample_beta_integer_shape(rng, N_SAMPLES, alpha_b, beta_b)
    beta_result = evaluate_distribution(
        "Бета",
        beta_sample,
        m_theoretical=alpha_b / (alpha_b + beta_b),
        d_theoretical=(alpha_b * beta_b) / (((alpha_b + beta_b) ** 2) * (alpha_b + beta_b + 1)),
    )
    x_beta = np.linspace(0, 1, 450)
    save_hist_with_pdf(beta_result, x_beta, pdf_beta(x_beta, alpha_b, beta_b), out_dir / "5_beta_variant1.png")

    results = [exp_result, uni_result, erl_result, norm_result, beta_result]
    save_summary_table(results, out_dir / "6_summary_md.png")
    build_report(results, Path("report_lab3_variant1_gost.txt"))

    print("ЛР3 (вариант 1) выполнена.")
    print(f"N = {N_SAMPLES}, seed = {RNG_SEED}")
    print("Отчет: report_lab3_variant1_gost.txt")
    print("Графики: plots/1_exponential.png, plots/2_uniform.png, plots/3_erlang.png, plots/4_normal.png, plots/5_beta_variant1.png, plots/6_summary_md.png")


if __name__ == "__main__":
    main()
