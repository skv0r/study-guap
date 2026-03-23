from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np


# Вариант 1 (таблица 3.2 из методички)
X_VALUES = np.array([-73.4, -70.7, -51.5, -43.9, 13.3, 73.0, 73.8], dtype=float)
P_VALUES = np.array([0.241, 0.023, 0.166, 0.078, 0.272, 0.192, 0.028], dtype=float)
N_SAMPLES = 500
RNG_SEED = 42


def generate_discrete_sample(x: np.ndarray, p: np.ndarray, n: int, rng: np.random.Generator) -> np.ndarray:
    """Генерация дискретной СВ методом обратной функции распределения."""
    cumulative_prob = np.cumsum(p)
    u = rng.random(n)
    indices = np.searchsorted(cumulative_prob, u, side="right")
    return x[indices]


def compute_theoretical_stats(x: np.ndarray, p: np.ndarray) -> tuple[float, float]:
    m = float(np.sum(x * p))
    d = float(np.sum((x**2) * p) - m**2)
    return m, d


def compute_empirical_probabilities(x: np.ndarray, sample: np.ndarray) -> np.ndarray:
    counts = np.array([(sample == xi).sum() for xi in x], dtype=float)
    return counts / sample.size


def save_histograms(x: np.ndarray, p_theoretical: np.ndarray, p_empirical: np.ndarray, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    positions = np.arange(len(x))

    fig, axes = plt.subplots(2, 1, figsize=(11, 8), sharex=True)

    axes[0].bar(positions, p_empirical, color="#4C72B0")
    axes[0].set_title("Эмпирическое распределение вероятностей")
    axes[0].set_ylabel("p*")
    axes[0].grid(axis="y", linestyle="--", alpha=0.5)

    axes[1].bar(positions, p_theoretical, color="#55A868")
    axes[1].set_title("Теоретическое распределение вероятностей")
    axes[1].set_ylabel("p")
    axes[1].set_xlabel("Значения xj")
    axes[1].grid(axis="y", linestyle="--", alpha=0.5)
    axes[1].set_xticks(positions)
    axes[1].set_xticklabels([f"{val:.1f}" for val in x], rotation=0)

    fig.tight_layout()
    fig.savefig(output_dir / "hist_compare_variant1.png", dpi=200)
    plt.close(fig)

    # Дополнительный сравнительный столбчатый график (удобно для отчета)
    fig2, ax2 = plt.subplots(figsize=(11, 4.5))
    width = 0.38
    ax2.bar(positions - width / 2, p_theoretical, width=width, label="Теоретические", color="#55A868")
    ax2.bar(positions + width / 2, p_empirical, width=width, label="Эмпирические", color="#4C72B0")
    ax2.set_title("Сравнение теоретических и эмпирических вероятностей")
    ax2.set_xlabel("Значения xj")
    ax2.set_ylabel("Вероятность")
    ax2.set_xticks(positions)
    ax2.set_xticklabels([f"{val:.1f}" for val in x], rotation=0)
    ax2.grid(axis="y", linestyle="--", alpha=0.5)
    ax2.legend()
    fig2.tight_layout()
    fig2.savefig(output_dir / "bar_compare_variant1.png", dpi=200)
    plt.close(fig2)


def save_results_figure(
    m_theoretical: float,
    d_theoretical: float,
    m_empirical: float,
    d_empirical: float,
    output_dir: Path,
) -> None:
    """Сохраняет картинку с итоговыми числовыми результатами для отчета."""
    fig, ax = plt.subplots(figsize=(10, 3.8))
    ax.axis("off")
    ax.set_title("Итоговые результаты ЛР2 (вариант 1)", pad=12)

    rows = [
        ["M(x) теоретическое", f"{m_theoretical:.6f}"],
        ["D(x) теоретическое", f"{d_theoretical:.6f}"],
        ["M*(x) эмпирическое", f"{m_empirical:.6f}"],
        ["D*(x) эмпирическое", f"{d_empirical:.6f}"],
        ["|M - M*|", f"{abs(m_theoretical - m_empirical):.6f}"],
        ["|D - D*|", f"{abs(d_theoretical - d_empirical):.6f}"],
    ]

    table = ax.table(
        cellText=rows,
        colLabels=["Показатель", "Значение"],
        cellLoc="left",
        colLoc="center",
        loc="center",
    )
    table.auto_set_font_size(False)
    table.set_fontsize(11)
    table.scale(1, 1.4)

    fig.tight_layout()
    fig.savefig(output_dir / "results_variant1.png", dpi=220)
    plt.close(fig)


def write_report_files(
    sample: np.ndarray,
    x: np.ndarray,
    p_theoretical: np.ndarray,
    p_empirical: np.ndarray,
    m_theoretical: float,
    d_theoretical: float,
    m_empirical: float,
    d_empirical: float,
) -> None:
    first_30 = ", ".join(f"{v:.1f}" for v in sample[:30])
    probabilities_table = "\n".join(
        f"| {xj:.1f} | {pt:.3f} | {pe:.3f} |"
        for xj, pt, pe in zip(x, p_theoretical, p_empirical)
    )

    report_text = f"""# Лабораторная работа №2
## Моделирование дискретной случайной величины (вариант 1)

### Ход работы
1. Заданы значения дискретной случайной величины xj и соответствующие вероятности pj (вариант 1).
2. Построены накопленные вероятности tj = sum(p1..pj), по которым реализован датчик методом обратной функции распределения.
3. Сгенерирована выборка из N = {sample.size} значений.
4. Рассчитаны:
   - теоретические характеристики: M(x), D(x);
   - эмпирические оценки: M*(x), D*(x).
5. Построены гистограммы эмпирического и теоретического распределений.

### Первые 30 значений выборки
{first_30}

### Теоретические и эмпирические характеристики
- M(x) теоретическое = {m_theoretical:.6f}
- D(x) теоретическое = {d_theoretical:.6f}
- M*(x) эмпирическое = {m_empirical:.6f}
- D*(x) эмпирическое = {d_empirical:.6f}
- |M - M*| = {abs(m_theoretical - m_empirical):.6f}
- |D - D*| = {abs(d_theoretical - d_empirical):.6f}

### Сравнение вероятностей
| xj | pj (теор.) | p*j (эмп.) |
|---:|-----------:|-----------:|
{probabilities_table}

### Вывод
Сгенерированный датчик корректно моделирует дискретную случайную величину варианта 1.
Эмпирические оценки математического ожидания и дисперсии близки к теоретическим значениям, а форма эмпирической гистограммы соответствует теоретическому распределению.
Наблюдаемые различия объясняются конечным объёмом выборки (N = {sample.size}) и носят случайный характер.
"""

    Path("lab2_report.md").write_text(report_text, encoding="utf-8")
    Path("first30_variant1.txt").write_text(first_30 + "\n", encoding="utf-8")

    txt_report = f"""ЛАБОРАТОРНАЯ РАБОТА №2
МОДЕЛИРОВАНИЕ ДИСКРЕТНОЙ СЛУЧАЙНОЙ ВЕЛИЧИНЫ
Вариант 1

1) Листинг программной реализации датчика заданной дискретной СВ
Файл программы: lab2_modeling.py

2) Первые 30 значений xi
{first_30}

3) Результаты эмпирических и теоретических значений M и D
M(x) теоретическое  = {m_theoretical:.6f}
D(x) теоретическое  = {d_theoretical:.6f}
M*(x) эмпирическое  = {m_empirical:.6f}
D*(x) эмпирическое  = {d_empirical:.6f}
|M - M*|            = {abs(m_theoretical - m_empirical):.6f}
|D - D*|            = {abs(d_theoretical - d_empirical):.6f}

4) Гистограммы распределения эмпирических и теоретических вероятностей дискретной СВ
Сформированные файлы графиков:
- plots/hist_compare_variant1.png
- plots/bar_compare_variant1.png
- plots/results_variant1.png

5) Выводы по результатам сравнительной оценки распределений
Датчик корректно моделирует дискретную случайную величину варианта 1.
Эмпирические оценки математического ожидания и дисперсии близки к теоретическим.
Форма эмпирического распределения согласуется с теоретическим распределением.
Различия обусловлены конечным объемом выборки N = {sample.size}.
"""
    Path("report_lab2_variant1.txt").write_text(txt_report, encoding="utf-8")


def main() -> None:
    if not np.isclose(P_VALUES.sum(), 1.0):
        raise ValueError("Сумма вероятностей pj должна быть равна 1.")

    rng = np.random.default_rng(RNG_SEED)
    sample = generate_discrete_sample(X_VALUES, P_VALUES, N_SAMPLES, rng)

    m_theoretical, d_theoretical = compute_theoretical_stats(X_VALUES, P_VALUES)
    m_empirical = float(np.mean(sample))
    d_empirical = float(np.var(sample))
    p_empirical = compute_empirical_probabilities(X_VALUES, sample)

    plots_dir = Path("plots")
    save_histograms(X_VALUES, P_VALUES, p_empirical, plots_dir)
    save_results_figure(m_theoretical, d_theoretical, m_empirical, d_empirical, plots_dir)
    write_report_files(
        sample,
        X_VALUES,
        P_VALUES,
        p_empirical,
        m_theoretical,
        d_theoretical,
        m_empirical,
        d_empirical,
    )

    print("ЛР2 (вариант 1) выполнена.")
    print(f"N = {N_SAMPLES}, seed = {RNG_SEED}")
    print(f"M_theoretical = {m_theoretical:.6f}")
    print(f"D_theoretical = {d_theoretical:.6f}")
    print(f"M_empirical   = {m_empirical:.6f}")
    print(f"D_empirical   = {d_empirical:.6f}")
    print("Файлы отчета: lab2_report.md, report_lab2_variant1.txt, first30_variant1.txt")
    print("Графики: plots/hist_compare_variant1.png, plots/bar_compare_variant1.png, plots/results_variant1.png")


if __name__ == "__main__":
    main()
