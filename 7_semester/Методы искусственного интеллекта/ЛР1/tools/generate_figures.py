#!/usr/bin/env python3
"""Figures for MII lab 1: scatter, scorer table, KNIME-style workflow."""
from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Circle

LAB = Path(__file__).resolve().parents[1]
FIG = LAB / "figures"
DATA = LAB / "data" / "courier_delivery.csv"
FIG.mkdir(exist_ok=True)

GOLD = "#F7C01A"
NAVY = "#1F3A5F"
GREY = "#5C6570"


def load():
    rows = np.genfromtxt(DATA, delimiter=",", names=True, dtype=None, encoding="utf-8")
    x = np.array([r["distance_km"] for r in rows], dtype=float)
    y = np.array([r["time_min"] for r in rows], dtype=float)
    return x, y


def scatter():
    x, y = load()
    n_train = 20
    b0, b1 = 6.7917, 3.4624
    xs = np.linspace(x.min(), x.max(), 200)
    fig, ax = plt.subplots(figsize=(8.2, 5.2), dpi=140)
    ax.scatter(x[n_train:], y[n_train:], s=28, c="#4C78A8", alpha=0.85, label="Тест (80%)", zorder=3)
    ax.scatter(x[:n_train], y[:n_train], s=36, c="#F58518", marker="D", label="Обучение (20%)", zorder=4)
    ax.plot(xs, b0 + b1 * xs, color="#E45756", lw=2.0, label=r"$\widehat{time}=6{,}792+3{,}462\cdot distance$")
    ax.set_xlabel("distance_km, км")
    ax.set_ylabel("time_min, мин")
    ax.set_title("Время доставки от расстояния")
    ax.grid(True, ls=":", alpha=0.5)
    ax.legend(frameon=True, fontsize=9)
    fig.tight_layout()
    fig.savefig(FIG / "scatter.png", bbox_inches="tight")
    plt.close(fig)


def scorer_table():
    rows = [
        ("R²", "0,915"),
        ("Mean absolute error", "2,472"),
        ("Mean squared error", "9,482"),
        ("Root mean squared error", "3,079"),
        ("Mean signed difference", "2,054"),
        ("Mean absolute percentage error", "7,196 %"),
        ("Adjusted R²", "0,914"),
    ]
    fig, ax = plt.subplots(figsize=(7.2, 3.6), dpi=140)
    ax.set_axis_off()
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.add_patch(FancyBboxPatch((0.02, 0.02), 0.96, 0.96, boxstyle="round,pad=0.01",
                                facecolor="white", edgecolor="#C8CCD0", lw=1.2))
    ax.text(0.05, 0.90, "Numeric Scorer — Statistics", fontsize=12, fontweight="bold", color=NAVY)
    y = 0.78
    for name, val in rows:
        ax.text(0.08, y, name, fontsize=10, color="#222")
        ax.text(0.92, y, val, fontsize=10, ha="right", fontfamily="monospace", color="#111")
        y -= 0.10
    fig.tight_layout()
    fig.savefig(FIG / "numeric_scorer.png", bbox_inches="tight")
    plt.close(fig)


def _node(ax, x, y, title, w=1.55, h=0.72):
    box = FancyBboxPatch((x - w / 2, y - h / 2), w, h, boxstyle="round,pad=0.04,rounding_size=0.08",
                         facecolor=GOLD, edgecolor="#B8860B", lw=1.3, zorder=3)
    ax.add_patch(box)
    ax.text(x, y + 0.06, title, ha="center", va="center", fontsize=7.2, fontweight="bold", zorder=4)
    # KNIME-style ports
    ax.add_patch(Circle((x - w / 2, y), 0.055, facecolor="white", edgecolor="#333", lw=0.8, zorder=5))
    ax.add_patch(Circle((x + w / 2, y), 0.055, facecolor="white", edgecolor="#333", lw=0.8, zorder=5))
    return (x + w / 2, y), (x - w / 2, y)


def workflow():
    fig, ax = plt.subplots(figsize=(12.0, 4.2), dpi=150)
    ax.set_xlim(0.2, 11.8)
    ax.set_ylim(0.15, 3.35)
    ax.axis("off")

    def node(x, y, title, w=1.62):
        h = 0.78
        box = FancyBboxPatch(
            (x - w / 2, y - h / 2), w, h,
            boxstyle="round,pad=0.03,rounding_size=0.08",
            facecolor=GOLD, edgecolor="#B8860B", lw=1.3, zorder=3,
        )
        ax.add_patch(box)
        ax.text(x, y, title, ha="center", va="center", fontsize=7.4, fontweight="bold", zorder=4)
        ax.add_patch(Circle((x - w / 2, y), 0.05, facecolor="white", edgecolor="#333", lw=0.8, zorder=5))
        ax.add_patch(Circle((x + w / 2, y), 0.05, facecolor="white", edgecolor="#333", lw=0.8, zorder=5))
        return x - w / 2, x + w / 2, y

    def arrow(x0, y0, x1, y1):
        ax.add_patch(FancyArrowPatch(
            (x0, y0), (x1, y1), arrowstyle="-|>", mutation_scale=11,
            lw=1.35, color="#3A3A3A", zorder=2, connectionstyle="arc3,rad=0",
        ))

    a0, a1, ay = node(1.05, 1.7, "CSV Reader")
    b0, b1, by = node(2.95, 1.7, "Color Manager")
    c0, c1, cy = node(4.90, 1.7, "Partitioning")
    d0, d1, dy = node(7.05, 2.55, "Linear Regression\nLearner", w=1.85)
    e0, e1, ey = node(7.05, 0.90, "Regression\nPredictor", w=1.75)
    f0, f1, fy = node(9.45, 0.90, "Numeric Scorer")
    g0, g1, gy = node(9.45, 2.55, "Scatter Plot")

    arrow(a1, ay, b0, by)
    arrow(b1, by, c0, cy)
    arrow(c1, cy + 0.12, d0, dy)
    arrow(c1, cy - 0.12, e0, ey)
    arrow(d1, dy, e1 - 0.02, ey + 0.35)  # model down to predictor
    arrow(e1, ey, f0, fy)
    arrow(e1, ey + 0.18, g0, gy)
    ax.text(5.85, 2.20, "обуч. 20%", fontsize=7.2, color=GREY)
    ax.text(5.85, 1.12, "тест 80%", fontsize=7.2, color=GREY)
    ax.set_title("Модель линейной регрессии (узлы KNIME по методичке)", fontsize=11, pad=6)
    fig.tight_layout()
    fig.savefig(FIG / "workflow.png", bbox_inches="tight")
    plt.close(fig)


def main():
    scatter()
    scorer_table()
    workflow()
    print("figures", list(FIG.glob("*.png")))


if __name__ == "__main__":
    main()
