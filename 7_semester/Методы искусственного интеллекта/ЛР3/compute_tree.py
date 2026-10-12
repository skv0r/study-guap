# -*- coding: utf-8 -*-
"""Данные IT-тикетов и дерево решений. В текст отчёта не входит."""
from __future__ import annotations

import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
import numpy as np
import pandas as pd
from sklearn.metrics import (
    accuracy_score,
    confusion_matrix,
    f1_score,
    precision_recall_fscore_support,
)
from sklearn.model_selection import train_test_split
from sklearn.tree import DecisionTreeClassifier, export_text

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
FIG = BASE / "figures"

CLASSES = ["сеть", "доступ", "ПО"]
FEATURES = ["users_affected", "downtime_min", "latency_ms", "error_rate_pct"]
FEATURE_RU = {
    "users_affected": "Пользователи",
    "downtime_min": "Простой, мин",
    "latency_ms": "Задержка, мс",
    "error_rate_pct": "Ошибки, %",
}
COLORS = {"сеть": "#1f77b4", "доступ": "#2ca02c", "ПО": "#d62728"}
SEED = 29


def generate(rng: np.random.Generator) -> pd.DataFrame:
    """Три типа IT-инцидентов с шумом и пересечением признаков."""
    specs = {
        # сеть: много пользователей, высокая задержка, умеренные ошибки
        "сеть": dict(
            users=(55, 10),
            downtime=(40, 8),
            latency=(450, 70),
            errors=(3.8, 1.2),
        ),
        # доступ: мало пользователей, короткий простой, низкая задержка и ошибки
        "доступ": dict(
            users=(5, 2.2),
            downtime=(10, 3.5),
            latency=(80, 25),
            errors=(0.9, 0.45),
        ),
        # ПО: средний охват, длинный простой, много ошибок приложения
        "ПО": dict(
            users=(18, 6),
            downtime=(60, 12),
            latency=(160, 40),
            errors=(14.0, 3.0),
        ),
    }
    boundary = {
        "сеть": dict(latency=(250, 35), n=5),
        "доступ": dict(users=(14, 3), n=5),
        "ПО": dict(errors=(6.0, 1.0), n=5),
    }
    limits = {
        "users_affected": (1, 90, 0),
        "downtime_min": (2, 120, 0),
        "latency_ms": (20, 800, 0),
        "error_rate_pct": (0.1, 25.0, 1),
    }
    rows = []
    for name in CLASSES:
        s = specs[name]
        n = 30
        users = rng.normal(s["users"][0], s["users"][1], n)
        downtime = rng.normal(s["downtime"][0], s["downtime"][1], n)
        latency = rng.normal(s["latency"][0], s["latency"][1], n)
        errors = rng.normal(s["errors"][0], s["errors"][1], n)
        b = boundary[name]
        idx = rng.choice(n, size=b["n"], replace=False)
        if "latency" in b:
            latency[idx] = rng.normal(b["latency"][0], b["latency"][1], b["n"])
        if "users" in b:
            users[idx] = rng.normal(b["users"][0], b["users"][1], b["n"])
        if "errors" in b:
            errors[idx] = rng.normal(b["errors"][0], b["errors"][1], b["n"])
        for i in range(n):
            raw = {
                "users_affected": users[i],
                "downtime_min": downtime[i],
                "latency_ms": latency[i],
                "error_rate_pct": errors[i],
            }
            item = {"ticket_type": name}
            for key, (lo, hi, nd) in limits.items():
                item[key] = float(np.round(np.clip(raw[key], lo, hi), nd))
            rows.append(item)
    df = pd.DataFrame(rows)
    df = df.iloc[rng.permutation(len(df))].reset_index(drop=True)
    df.insert(0, "ticket_id", np.arange(1, len(df) + 1))
    return df


def rules_from_tree(clf: DecisionTreeClassifier) -> list[str]:
    text = export_text(
        clf,
        feature_names=[FEATURE_RU[f] for f in FEATURES],
        class_names=list(clf.classes_),
        decimals=1,
        show_weights=False,
    )
    return [line.rstrip() for line in text.splitlines() if line.strip()]


def logical_rules(clf: DecisionTreeClassifier) -> list[str]:
    """Краткие если-то правила по листьям дерева."""
    tree = clf.tree_
    labels = list(clf.classes_)
    phrases: list[str] = []

    def recurse(node: int, path: list[str]) -> None:
        if tree.feature[node] == -2:
            counts = tree.value[node][0]
            label = labels[int(np.argmax(counts))]
            cond = " и ".join(path) if path else "без условий"
            phrases.append(f"Если {cond}, то тип «{label}».")
            return
        name = FEATURES[tree.feature[node]]
        ru = FEATURE_RU[name]
        thr = tree.threshold[node]
        thr_s = f"{thr:.1f}".replace(".", ",")
        recurse(tree.children_left[node], path + [f"{ru} ≤ {thr_s}"])
        recurse(tree.children_right[node], path + [f"{ru} > {thr_s}"])

    recurse(0, [])
    return phrases


def draw_scheme(path: Path) -> None:
    plt.rcParams["font.family"] = "Times New Roman"
    labels = [
        "Чтение таблицы IT-тикетов",
        "Цвет класса для диаграммы",
        "Разбиение 70 % / 30 % по классу",
        "Обучение дерева решений",
        "Прогноз типа инцидента на тесте",
        "Матрица ошибок и метрики качества",
        "Правила дерева и диаграмма рассеяния",
    ]
    n = len(labels)
    box_h = 0.72
    gap = 0.42
    top = 0.15
    fig_h = top + n * box_h + (n - 1) * gap + 0.12
    fig, ax = plt.subplots(figsize=(7.6, fig_h))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, fig_h)
    ax.axis("off")
    y = fig_h - top - box_h
    for i, text in enumerate(labels):
        ax.add_patch(plt.Rectangle((0.55, y), 8.9, box_h, fill=False, edgecolor="black", linewidth=1.0))
        ax.text(5.0, y + box_h / 2, text, ha="center", va="center", fontsize=12, color="black")
        if i < n - 1:
            ax.annotate(
                "",
                xy=(5.0, y - gap + 0.02),
                xytext=(5.0, y - 0.01),
                arrowprops=dict(arrowstyle="-|>", color="black", lw=1.05, mutation_scale=11),
            )
        y -= box_h + gap
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=160, bbox_inches="tight", facecolor="white", pad_inches=0.12)
    plt.close()


def _thr(value: float) -> str:
    return f"{value:.1f}".replace(".", ",")


def draw_tree(clf: DecisionTreeClassifier, path: Path) -> None:
    """Читаемая схема дерева: узлы-условия и цветные листья классов."""
    from matplotlib.patches import FancyBboxPatch

    plt.rcParams["font.family"] = "Times New Roman"
    leaf_fc = {
        "доступ": "#d9ead3",
        "ПО": "#f4cccc",
        "сеть": "#cfe2f3",
    }
    tree = clf.tree_
    labels = list(clf.classes_)
    root_thr = tree.threshold[0]
    right = tree.children_right[0]
    right_thr = tree.threshold[right]
    left_leaf = labels[int(np.argmax(tree.value[tree.children_left[0]][0]))]
    soft_leaf = labels[int(np.argmax(tree.value[tree.children_left[right]][0]))]
    net_leaf = labels[int(np.argmax(tree.value[tree.children_right[right]][0]))]

    fig, ax = plt.subplots(figsize=(10.8, 6.4))
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")
    fig.patch.set_facecolor("white")

    def node(x, y, w, h, text, *, fc="#ffffff", size=11, bold=False):
        ax.add_patch(
            FancyBboxPatch(
                (x - w / 2, y - h / 2),
                w,
                h,
                boxstyle="round,pad=0.02,rounding_size=0.8",
                facecolor=fc,
                edgecolor="black",
                linewidth=1.15,
            )
        )
        ax.text(
            x,
            y,
            text,
            ha="center",
            va="center",
            fontsize=size,
            fontweight="bold" if bold else "normal",
            color="black",
            linespacing=1.25,
        )

    def link(x1, y1, x2, y2, label, *, side="left"):
        ax.plot([x1, x2], [y1, y2], color="black", lw=1.05, zorder=1)
        mx, my = (x1 + x2) / 2, (y1 + y2) / 2
        dx = -3.2 if side == "left" else 3.2
        ax.text(mx + dx, my + 1.2, label, ha="center", va="center", fontsize=10, color="#333333")

    root_feat = FEATURE_RU[FEATURES[tree.feature[0]]]
    right_feat = FEATURE_RU[FEATURES[tree.feature[right]]]

    node(50, 88, 38, 13, f"{root_feat}\n≤ {_thr(root_thr)} ?", fc="#fff2cc", size=13, bold=True)
    node(72, 55, 38, 13, f"{right_feat}\n≤ {_thr(right_thr)} ?", fc="#fff2cc", size=13, bold=True)
    node(22, 18, 28, 12, f"Класс:\n{left_leaf}", fc=leaf_fc[left_leaf], size=13, bold=True)
    node(58, 18, 28, 12, f"Класс:\n{soft_leaf}", fc=leaf_fc[soft_leaf], size=13, bold=True)
    node(88, 18, 28, 12, f"Класс:\n{net_leaf}", fc=leaf_fc[net_leaf], size=13, bold=True)

    link(50, 81.5, 22, 24, "да", side="left")
    link(50, 81.5, 72, 61.5, "нет", side="right")
    link(72, 48.5, 58, 24, "да", side="left")
    link(72, 48.5, 88, 24, "нет", side="right")

    ax.text(
        50,
        4.0,
        "Жёлтые узлы — условия разбиения; зелёный / красный / синий — итоговый класс",
        ha="center",
        va="center",
        fontsize=10,
        color="#444444",
    )
    fig.savefig(path, dpi=180, bbox_inches="tight", facecolor="white", pad_inches=0.15)
    plt.close()


def draw_scatter(df: pd.DataFrame, error_ids: list[int], path: Path) -> None:
    plt.rcParams["font.family"] = "Times New Roman"
    fig, ax = plt.subplots(figsize=(7.5, 5.7))
    for name in CLASSES:
        part = df[df["ticket_type"] == name]
        ax.scatter(
            part["latency_ms"],
            part["error_rate_pct"],
            c=COLORS[name],
            s=42,
            label=name,
            zorder=3,
            edgecolors="white",
            linewidths=0.4,
        )
    err = df[df["ticket_id"].isin(error_ids)]
    if len(err):
        ax.scatter(
            err["latency_ms"],
            err["error_rate_pct"],
            facecolors="none",
            edgecolors="black",
            s=130,
            linewidths=1.15,
            label="ошибка на тесте",
            zorder=4,
        )

    def comma_tick(value, _pos):
        return f"{value:g}".replace(".", ",")

    ax.xaxis.set_major_formatter(FuncFormatter(comma_tick))
    ax.yaxis.set_major_formatter(FuncFormatter(comma_tick))
    ax.set_xlabel("Задержка, мс", fontsize=12)
    ax.set_ylabel("Доля ошибок, %", fontsize=12)
    ax.tick_params(labelsize=11)
    ax.grid(True, color="#d0d0d0", linewidth=0.6)
    ax.set_axisbelow(True)
    ax.legend(frameon=True, fontsize=11)
    fig.tight_layout()
    fig.savefig(path, dpi=160, bbox_inches="tight", facecolor="white")
    plt.close()


def main() -> None:
    DATA.mkdir(parents=True, exist_ok=True)
    FIG.mkdir(parents=True, exist_ok=True)
    rng = np.random.default_rng(SEED)
    df = generate(rng)
    df.to_csv(DATA / "tickets.csv", index=False, encoding="utf-8")

    X = df[FEATURES].to_numpy(dtype=float)
    y = df["ticket_type"].to_numpy()
    ids = df["ticket_id"].to_numpy()

    X_train, X_test, y_train, y_test, id_train, id_test = train_test_split(
        X,
        y,
        ids,
        test_size=0.30,
        random_state=SEED,
        stratify=y,
    )
    # Порядок классов в дереве — clf.classes_, не порядок списка CLASSES.

    clf = DecisionTreeClassifier(
        criterion="gini",
        max_depth=3,
        min_samples_leaf=3,
        random_state=SEED,
    )
    clf.fit(X_train, y_train)
    y_pred = clf.predict(X_test)

    cm = confusion_matrix(y_test, y_pred, labels=CLASSES)
    precision, recall, f1, support = precision_recall_fscore_support(
        y_test, y_pred, labels=CLASSES, zero_division=0
    )
    acc = float(accuracy_score(y_test, y_pred))
    correct = int((y_pred == y_test).sum())
    n_test = int(len(y_test))
    error_ids = [int(i) for i, yt, yp in zip(id_test, y_test, y_pred) if yt != yp]

    per_class = []
    for i, name in enumerate(CLASSES):
        per_class.append(
            {
                "class": name,
                "precision": float(precision[i]),
                "recall": float(recall[i]),
                "f1": float(f1[i]),
                "support": int(support[i]),
            }
        )

    predictions = []
    for tid, yt, yp, row in zip(id_test, y_test, y_pred, X_test):
        predictions.append(
            {
                "ticket_id": int(tid),
                "ticket_type": str(yt),
                "predicted": str(yp),
                "correct": bool(yt == yp),
                "features": {f: float(v) for f, v in zip(FEATURES, row)},
            }
        )

    metrics = {
        "n": int(len(df)),
        "n_train": int(len(y_train)),
        "n_test": n_test,
        "correct": correct,
        "accuracy": acc,
        "confusion": cm.tolist(),
        "per_class": per_class,
        "macro_precision": float(np.mean(precision)),
        "macro_recall": float(np.mean(recall)),
        "macro_f1": float(np.mean(f1)),
        "train_ids": sorted(int(i) for i in id_train),
        "test_ids": sorted(int(i) for i in id_test),
        "error_ids": error_ids,
        "predictions": predictions,
        "rules": logical_rules(clf),
        "export_text": rules_from_tree(clf),
        "max_depth": 3,
        "criterion": "gini",
        "feature_importance": {
            f: float(v) for f, v in zip(FEATURES, clf.feature_importances_)
        },
        "ranges": {
            name: {
                f: [
                    float(df.loc[df["ticket_type"] == name, f].min()),
                    float(df.loc[df["ticket_type"] == name, f].max()),
                ]
                for f in FEATURES
            }
            for name in CLASSES
        },
    }
    (DATA / "metrics.json").write_text(
        json.dumps(metrics, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    draw_scheme(FIG / "scheme.png")
    draw_tree(clf, FIG / "tree.png")
    draw_scatter(df, error_ids, FIG / "scatter.png")

    print("n", metrics["n"], "train", metrics["n_train"], "test", metrics["n_test"])
    print("accuracy", round(acc, 4), "correct", correct, "/", n_test)
    print("errors", error_ids)
    print("rules:")
    for r in metrics["rules"]:
        print(" ", r)


if __name__ == "__main__":
    main()
