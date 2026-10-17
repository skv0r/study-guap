# -*- coding: utf-8 -*-
"""Опорные векторы по измерениям мяча и сравнение четырёх методов. В текст отчёта не входит."""
from __future__ import annotations

import json
import shutil
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from sklearn.svm import SVC

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
FIG = BASE / "figures"
ROOT = BASE.parent
FEATURES = ["elasticity", "size", "mass", "pressure", "bounce"]
FEATURE_RU = {
    "elasticity": "Коэффициент упругости",
    "size": "Окружность, см",
    "mass": "Масса, г",
    "pressure": "Давление, бар",
    "bounce": "Высота отскока, см",
}
CLASSES = ["футбольный", "баскетбольный", "волейбольный"]
COLORS = {"футбольный": "#1f77b4", "баскетбольный": "#ff7f0e", "волейбольный": "#2ca02c"}
SIGMA = 0.5
PENALTY = 1.0
SHORT = {
    "elasticity": "упругость",
    "size": "окружность",
    "mass": "масса",
    "pressure": "давление",
    "bounce": "отскок",
}


def fmt_num(feature: str, value: float) -> str:
    number = float(value)
    if feature in ("elasticity", "pressure"):
        return f"{number:.2f}".replace(".", ",")
    if feature == "size":
        return f"{number:.1f}".replace(".", ",")
    return str(int(round(number)))


def metrics(y_true, y_pred) -> dict:
    matrix = [[0, 0, 0] for _ in CLASSES]
    index = {name: i for i, name in enumerate(CLASSES)}
    for actual, pred in zip(y_true, y_pred):
        matrix[index[actual]][index[pred]] += 1
    per = []
    for i, name in enumerate(CLASSES):
        tp = matrix[i][i]
        fp = sum(matrix[r][i] for r in range(3)) - tp
        fn = sum(matrix[i]) - tp
        precision = tp / (tp + fp) if tp + fp else 0.0
        recall = tp / (tp + fn) if tp + fn else 0.0
        f1 = 2 * precision * recall / (precision + recall) if precision + recall else 0.0
        per.append(
            {
                "class": name,
                "tp": tp,
                "fp": fp,
                "fn": fn,
                "precision": precision,
                "recall": recall,
                "f1": f1,
                "support": tp + fn,
            }
        )
    correct = sum(matrix[i][i] for i in range(3))
    return {
        "correct": correct,
        "accuracy": correct / len(y_true),
        "confusion": matrix,
        "per_class": per,
        "macro_precision": sum(item["precision"] for item in per) / 3,
        "macro_recall": sum(item["recall"] for item in per) / 3,
        "macro_f1": sum(item["f1"] for item in per) / 3,
    }


def row_diff(ball: str, features: dict, typical: dict) -> list[dict]:
    diffs = []
    for feature, usual in zip(FEATURES, typical[ball]):
        value = float(features[feature])
        if abs(value - float(usual)) > 1e-9:
            diffs.append(
                {
                    "feature": feature,
                    "usual": float(usual),
                    "value": value,
                    "phrase": (
                        f"{SHORT[feature]} {fmt_num(feature, value)} "
                        f"вместо {fmt_num(feature, usual)}"
                    ),
                }
            )
    return diffs


def main() -> None:
    DATA.mkdir(exist_ok=True)
    FIG.mkdir(exist_ok=True)
    for name in ("balls.csv", "split.json", "codes.json"):
        shutil.copyfile(ROOT / "ЛР3" / "data" / name, DATA / name)
    split = json.loads((DATA / "split.json").read_text(encoding="utf-8"))
    codes = json.loads((DATA / "codes.json").read_text(encoding="utf-8"))
    typical = codes["typical"]
    df = pd.read_csv(DATA / "balls.csv")
    train = df[df["id"].isin(split["train_ids"])].sort_values("id")
    test = df[df["id"].isin(split["test_ids"])].sort_values("id")
    n_total = int(codes.get("n", len(df)))
    n_train = int(codes.get("n_train", split.get("n_train", 105)))
    n_test = int(codes.get("n_test", split.get("n_test", 45)))
    if len(train) != n_train or len(test) != n_test:
        raise RuntimeError(f"разбиение не {n_train}/{n_test}")
    x_train_raw = train[FEATURES].to_numpy(dtype=float)
    y_train = train["ball"].to_numpy()
    x_test_raw = test[FEATURES].to_numpy(dtype=float)
    # Масштаб как у kNN: иначе масса в граммах подавляет упругость в ядре RBF.
    mu = x_train_raw.mean(axis=0)
    sd = x_train_raw.std(axis=0, ddof=0)
    sd = np.where(sd < 1e-9, 1.0, sd)
    x_train = (x_train_raw - mu) / sd
    x_test = (x_test_raw - mu) / sd
    gamma = 1.0 / (2.0 * SIGMA * SIGMA)
    model = SVC(kernel="rbf", C=PENALTY, gamma=gamma, probability=True)
    model.fit(x_train, y_train)
    pred = model.predict(x_test)
    proba = model.predict_proba(x_test)
    class_order = list(model.classes_)
    scored = metrics(test["ball"].tolist(), pred.tolist())
    if scored["correct"] >= n_test:
        raise RuntimeError("SVM дал 100 % на тесте — данные слишком простые")
    if scored["correct"] / n_test < 0.80:
        raise RuntimeError(f"SVM ниже 80 %: {scored['correct']}/{n_test}")
    rows = []
    for i, (_, row) in enumerate(test.iterrows()):
        features = {feature: float(row[feature]) for feature in FEATURES}
        diffs = row_diff(row["ball"], features, typical)
        probs = {name: float(proba[i, class_order.index(name)]) for name in CLASSES}
        rows.append(
            {
                "id": int(row["id"]),
                "ball": row["ball"],
                "pred": str(pred[i]),
                "correct": str(pred[i]) == row["ball"],
                "probs": probs,
                "features": features,
                "diffs": diffs,
            }
        )
    sv_rows = []
    for vector, label in zip(model.support_vectors_, y_train[model.support_]):
        original = vector * sd + mu
        sv_rows.append(
            {
                "class": str(label),
                "features": {feature: float(value) for feature, value in zip(FEATURES, original)},
            }
        )
    counts = {name: int((y_train[model.support_] == name).sum()) for name in CLASSES}
    knn = json.loads((ROOT / "ЛР2" / "data" / "metrics.json").read_text(encoding="utf-8"))
    tree = json.loads((ROOT / "ЛР3" / "data" / "metrics.json").read_text(encoding="utf-8"))
    bayes = json.loads((ROOT / "ЛР4" / "data" / "metrics.json").read_text(encoding="utf-8"))
    for label, source in (("ЛР2", knn), ("ЛР3", tree), ("ЛР4", bayes)):
        if sorted(source.get("test_ids", [])) != sorted(split["test_ids"]):
            raise RuntimeError(f"тест {label} не совпал")
    comparison = []
    for lab, title, source in (
        ("2", "k ближайших соседей", knn),
        ("3", "дерево решений", tree),
        ("4", "наивный Байес", bayes),
        ("5", "опорные векторы", {**scored}),
    ):
        comparison.append(
            {
                "lab": lab,
                "method": title,
                "correct": source["correct"],
                "n_test": n_test,
                "accuracy": source["accuracy"],
                "precision": source["macro_precision"],
                "recall": source["macro_recall"],
                "f1": source["macro_f1"],
            }
        )
    best = max(item["accuracy"] for item in comparison)
    winners = [item["method"] for item in comparison if abs(item["accuracy"] - best) < 1e-12]
    example = next(
        (row for row in rows if row["correct"] and not row["diffs"]),
        None,
    )
    if example is None:
        example = next(row for row in rows if row["correct"])
    error_ids = [row["id"] for row in rows if not row["correct"]]
    payload = {
        "n": n_total,
        "n_train": n_train,
        "n_test": n_test,
        "n_train_per": int(codes.get("n_train_per", split.get("n_train_per", 35))),
        "n_test_per": int(codes.get("n_test_per", split.get("n_test_per", 15))),
        "sigma": SIGMA,
        "penalty": PENALTY,
        "gamma": gamma,
        "scaled": True,
        "mu": {feature: float(mu[i]) for i, feature in enumerate(FEATURES)},
        "sd": {feature: float(sd[i]) for i, feature in enumerate(FEATURES)},
        "n_support": counts,
        "n_support_total": int(model.n_support_.sum()),
        "support_vectors": sv_rows,
        "predictions": rows,
        "example": example,
        "comparison": comparison,
        "best_accuracy": best,
        "winners": winners,
        "train_ids": split["train_ids"],
        "test_ids": split["test_ids"],
        "error_ids": error_ids,
        "typical": typical,
        "seed": split["seed"],
        "knn_k": knn.get("k"),
        "tree_depth": tree.get("actual_depth"),
        **scored,
    }
    (DATA / "metrics.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    plt.rcParams["font.family"] = "Times New Roman"
    fig, ax = plt.subplots(figsize=(7.6, 5.8))
    for name in CLASSES:
        part = df[df["ball"] == name]
        ax.scatter(
            part["size"],
            part["elasticity"],
            s=70,
            c=COLORS[name],
            label=name,
            alpha=0.9,
            edgecolors="white",
            linewidths=0.6,
            zorder=3,
        )
    err = df[df["id"].isin(error_ids)]
    if len(err):
        ax.scatter(
            err["size"],
            err["elasticity"],
            facecolors="none",
            edgecolors="black",
            s=180,
            linewidths=1.3,
            label="ошибка на тесте",
            zorder=4,
        )
    ax.set_xlabel("Окружность, см", fontsize=12)
    ax.set_ylabel("Коэффициент упругости", fontsize=12)
    ax.tick_params(labelsize=11)
    ax.grid(True, color="#d0d0d0", linewidth=0.6)
    ax.set_axisbelow(True)
    ax.legend(frameon=True, fontsize=11, loc="best")
    fig.tight_layout()
    fig.savefig(FIG / "scatter.png", dpi=160, bbox_inches="tight", facecolor="white")
    plt.close(fig)

    labels = [
        "Чтение таблицы измерений мяча",
        "Разбиение 70 % / 30 % по классу",
        "Обучение SVM с ядром RBF",
        "Прогноз на тесте",
        "Матрица ошибок и сравнение методов",
        "Диаграмма рассеяния",
    ]
    n = len(labels)
    box_h = 0.72
    gap = 0.38
    top = 0.15
    fig_h = top + n * box_h + (n - 1) * gap + 0.12
    fig, ax = plt.subplots(figsize=(8.4, fig_h))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, fig_h)
    ax.axis("off")
    y = fig_h - top - box_h
    for i, text in enumerate(labels):
        ax.add_patch(plt.Rectangle((0.35, y), 9.3, box_h, fill=False, edgecolor="black", linewidth=1.0))
        ax.text(5.0, y + box_h / 2, text, ha="center", va="center", fontsize=12, color="black")
        if i < n - 1:
            ax.annotate(
                "",
                xy=(5.0, y - gap + 0.02),
                xytext=(5.0, y - 0.01),
                arrowprops=dict(arrowstyle="-|>", color="black", lw=1.05, mutation_scale=11),
            )
        y -= box_h + gap
    fig.savefig(FIG / "scheme.png", dpi=160, bbox_inches="tight", facecolor="white", pad_inches=0.12)
    plt.close(fig)
    print("svm", scored["correct"], "supports", counts, "winners", winners)
    print("comparison", [(item["method"], item["correct"]) for item in comparison])


if __name__ == "__main__":
    main()
