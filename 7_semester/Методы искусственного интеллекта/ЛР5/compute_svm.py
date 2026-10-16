# -*- coding: utf-8 -*-
"""Опорные векторы по мячу и сравнение четырёх методов. В текст отчёта не входит."""
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
FEATURES = ["net_share", "play_zone", "hold", "attack", "ground"]
FEATURE_RU = {
    "net_share": "Динамика владения",
    "play_zone": "Игровая зона",
    "hold": "Удержание контроля",
    "attack": "Вектор атаки",
    "ground": "Контакт с землёй",
}
CLASSES = ["футбольный", "баскетбольный", "волейбольный"]
COLORS = {"футбольный": "#1f77b4", "баскетбольный": "#ff7f0e", "волейбольный": "#2ca02c"}
SIGMA = 0.5
PENALTY = 1.0
TYPICAL = {
    "футбольный": [0, 0, 1, 0, 2],
    "баскетбольный": [1, 2, 2, 1, 0],
    "волейбольный": [2, 1, 0, 1, 0],
}


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


def main() -> None:
    DATA.mkdir(exist_ok=True)
    FIG.mkdir(exist_ok=True)
    shutil.copyfile(ROOT / "ЛР3" / "data" / "balls.csv", DATA / "balls.csv")
    split = json.loads((ROOT / "ЛР3" / "data" / "split.json").read_text(encoding="utf-8"))
    df = pd.read_csv(DATA / "balls.csv")
    train = df[df["id"].isin(split["train_ids"])].sort_values("id")
    test = df[df["id"].isin(split["test_ids"])].sort_values("id")
    x_train = train[FEATURES].to_numpy(dtype=float)
    y_train = train["ball"].to_numpy()
    x_test = test[FEATURES].to_numpy(dtype=float)
    gamma = 1.0 / (2.0 * SIGMA * SIGMA)
    model = SVC(kernel="rbf", C=PENALTY, gamma=gamma, probability=True)
    model.fit(x_train, y_train)
    pred = model.predict(x_test)
    proba = model.predict_proba(x_test)
    class_order = list(model.classes_)
    scored = metrics(test["ball"].tolist(), pred.tolist())
    rows = []
    for i, (_, row) in enumerate(test.iterrows()):
        usual = TYPICAL[row["ball"]]
        diffs = []
        for feature, usual_value in zip(FEATURES, usual):
            if int(row[feature]) != int(usual_value):
                diffs.append({"feature": feature, "usual": int(usual_value), "value": int(row[feature])})
        probs = {name: float(proba[i, class_order.index(name)]) for name in CLASSES}
        rows.append(
            {
                "id": int(row["id"]),
                "ball": row["ball"],
                "pred": str(pred[i]),
                "correct": str(pred[i]) == row["ball"],
                "probs": probs,
                "features": {feature: int(row[feature]) for feature in FEATURES},
                "diffs": diffs,
            }
        )
    sv_rows = []
    for vector, label in zip(model.support_vectors_, y_train[model.support_]):
        sv_rows.append(
            {
                "class": str(label),
                "features": {feature: float(value) for feature, value in zip(FEATURES, vector)},
            }
        )
    counts = {name: int((y_train[model.support_] == name).sum()) for name in CLASSES}
    knn = json.loads((ROOT / "ЛР2" / "data" / "metrics.json").read_text(encoding="utf-8"))
    tree = json.loads((ROOT / "ЛР3" / "data" / "metrics.json").read_text(encoding="utf-8"))
    bayes = json.loads((ROOT / "ЛР4" / "data" / "metrics.json").read_text(encoding="utf-8"))
    comparison = []
    for lab, title, source in (
        ("2", "k ближайших соседей", knn),
        ("3", "дерево решений", tree),
        ("4", "наивный Байес", bayes),
        ("5", "опорные векторы", scored),
    ):
        comparison.append(
            {
                "lab": lab,
                "method": title,
                "correct": source["correct"],
                "n_test": 27,
                "accuracy": source["accuracy"],
                "precision": source["macro_precision"],
                "recall": source["macro_recall"],
                "f1": source["macro_f1"],
            }
        )
    best = max(item["accuracy"] for item in comparison)
    winners = [item["method"] for item in comparison if item["accuracy"] == best]
    payload = {
        "sigma": SIGMA,
        "penalty": PENALTY,
        "gamma": gamma,
        "n_support": counts,
        "n_support_total": int(model.n_support_.sum()),
        "support_vectors": sv_rows,
        "predictions": rows,
        "example": next(row for row in rows if row["id"] == 1),
        "comparison": comparison,
        "best_accuracy": best,
        "winners": winners,
        "train_ids": split["train_ids"],
        "test_ids": split["test_ids"],
        **scored,
    }
    (DATA / "metrics.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    fig, ax = plt.subplots(figsize=(7.2, 4.6))
    for name in CLASSES:
        part = df[df["ball"] == name]
        ax.scatter(part["play_zone"], part["attack"], s=46, c=COLORS[name], label=name, alpha=0.75, edgecolors="white", linewidths=0.4)
    ax.set_xticks([0, 1, 2])
    ax.set_yticks([0, 1, 2])
    ax.set_xlabel("Игровая зона")
    ax.set_ylabel("Вектор атаки")
    ax.legend(frameon=False)
    ax.set_xlim(-0.4, 2.4)
    ax.set_ylim(-0.4, 2.4)
    fig.tight_layout()
    fig.savefig(FIG / "scatter.png", dpi=140)
    plt.close(fig)
    fig, ax = plt.subplots(figsize=(8.2, 2.2))
    ax.set_xlim(0, 8.2)
    ax.set_ylim(0, 2.2)
    ax.axis("off")
    for x, text in ((0.2, "90 розыгрышей"), (2.2, "63 на обучение"), (4.2, "ядро RBF"), (6.2, "27 на проверку")):
        ax.add_patch(plt.Rectangle((x, 0.7), 1.7, 0.8, fill=False, linewidth=1.2))
        ax.text(x + 0.85, 1.1, text, ha="center", va="center", fontsize=9)
    for x in (1.9, 3.9, 5.9):
        ax.annotate("", xy=(x + 0.3, 1.1), xytext=(x, 1.1), arrowprops={"arrowstyle": "->"})
    fig.tight_layout()
    fig.savefig(FIG / "scheme.png", dpi=140)
    plt.close(fig)
    print("svm", scored["correct"], "supports", counts, "winners", winners)


if __name__ == "__main__":
    main()
