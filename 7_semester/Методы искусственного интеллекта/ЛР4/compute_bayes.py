# -*- coding: utf-8 -*-
"""Наивный Байес по мячу. В текст отчёта не входит."""
from __future__ import annotations

import json
import math
import shutil
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
FIG = BASE / "figures"
SRC = BASE.parent / "ЛР3"
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
MIN_SD = 1e-4
TYPICAL = {
    "футбольный": [0, 0, 1, 0, 2],
    "баскетбольный": [1, 2, 2, 1, 0],
    "волейбольный": [2, 1, 0, 1, 0],
}


def load():
    DATA.mkdir(exist_ok=True)
    FIG.mkdir(exist_ok=True)
    shutil.copyfile(SRC / "data" / "balls.csv", DATA / "balls.csv")
    split = json.loads((SRC / "data" / "split.json").read_text(encoding="utf-8"))
    df = pd.read_csv(DATA / "balls.csv")
    train = df[df["id"].isin(split["train_ids"])].copy()
    test = df[df["id"].isin(split["test_ids"])].copy()
    if len(train) != 63 or len(test) != 27:
        raise RuntimeError("разбиение не 63/27")
    return df, train, test, split


def fit(train: pd.DataFrame) -> dict:
    model = {"prior": {}, "stats": []}
    n = len(train)
    for name in CLASSES:
        part = train[train["ball"] == name]
        model["prior"][name] = len(part) / n
        for feature in FEATURES:
            values = part[feature].to_numpy(dtype=float)
            mean = float(values.mean())
            sd = float(values.std(ddof=1)) if len(values) > 1 else 0.0
            used = sd if sd >= MIN_SD else MIN_SD
            model["stats"].append(
                {
                    "class": name,
                    "feature": feature,
                    "n": int(len(values)),
                    "mean": mean,
                    "sd": sd,
                    "sd_used": used,
                    "floored": sd < MIN_SD,
                }
            )
    return model


def log_gaussian(x: float, mean: float, sd: float) -> float:
    z = (x - mean) / sd
    return -math.log(sd) - 0.5 * math.log(2 * math.pi) - 0.5 * z * z


def predict_row(row, model: dict) -> dict:
    log_posts = {}
    for name in CLASSES:
        log_p = math.log(model["prior"][name])
        for stat in model["stats"]:
            if stat["class"] != name:
                continue
            log_p += log_gaussian(float(row[stat["feature"]]), stat["mean"], stat["sd_used"])
        log_posts[name] = log_p
    top = max(log_posts.values())
    weights = {name: math.exp(log_posts[name] - top) if log_posts[name] - top > -700 else 0.0 for name in CLASSES}
    total = sum(weights.values())
    probs = {name: weights[name] / total for name in CLASSES}
    pred = max(CLASSES, key=lambda name: (probs[name], -CLASSES.index(name)))
    return {"pred": pred, "probs": probs}


def metrics(y_true, y_pred) -> dict:
    labels = CLASSES
    matrix = [[0, 0, 0] for _ in labels]
    index = {name: i for i, name in enumerate(labels)}
    for actual, pred in zip(y_true, y_pred):
        matrix[index[actual]][index[pred]] += 1
    per = []
    for i, name in enumerate(labels):
        tp = matrix[i][i]
        fp = sum(matrix[r][i] for r in range(3)) - tp
        fn = sum(matrix[i]) - tp
        tn = 27 - tp - fp - fn
        precision = tp / (tp + fp) if tp + fp else 0.0
        recall = tp / (tp + fn) if tp + fn else 0.0
        f1 = 2 * precision * recall / (precision + recall) if precision + recall else 0.0
        per.append(
            {
                "class": name,
                "tp": tp,
                "fp": fp,
                "tn": tn,
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


def scatter(df: pd.DataFrame) -> None:
    fig, ax = plt.subplots(figsize=(7.2, 4.6))
    for name in CLASSES:
        part = df[df["ball"] == name]
        ax.scatter(
            part["play_zone"],
            part["attack"],
            s=46,
            c=COLORS[name],
            label=name,
            alpha=0.75,
            edgecolors="white",
            linewidths=0.4,
        )
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


def scheme() -> None:
    fig, ax = plt.subplots(figsize=(8.2, 2.2))
    ax.set_xlim(0, 8.2)
    ax.set_ylim(0, 2.2)
    ax.axis("off")
    boxes = [
        (0.2, "90 розыгрышей"),
        (2.2, "63 на обучение"),
        (4.2, "средние и отклонения"),
        (6.2, "27 на проверку"),
    ]
    for x, text in boxes:
        ax.add_patch(plt.Rectangle((x, 0.7), 1.7, 0.8, fill=False, linewidth=1.2))
        ax.text(x + 0.85, 1.1, text, ha="center", va="center", fontsize=9)
    for x in (1.9, 3.9, 5.9):
        ax.annotate("", xy=(x + 0.3, 1.1), xytext=(x, 1.1), arrowprops={"arrowstyle": "->"})
    fig.tight_layout()
    fig.savefig(FIG / "scheme.png", dpi=140)
    plt.close(fig)


def main() -> None:
    df, train, test, split = load()
    model = fit(train)
    rows = []
    for _, row in test.sort_values("id").iterrows():
        result = predict_row(row, model)
        usual = TYPICAL[row["ball"]]
        diffs = []
        for feature, usual_value in zip(FEATURES, usual):
            if int(row[feature]) != int(usual_value):
                diffs.append({"feature": feature, "usual": int(usual_value), "value": int(row[feature])})
        rows.append(
            {
                "id": int(row["id"]),
                "ball": row["ball"],
                "pred": result["pred"],
                "correct": result["pred"] == row["ball"],
                "probs": result["probs"],
                "features": {feature: int(row[feature]) for feature in FEATURES},
                "diffs": diffs,
            }
        )
    scored = metrics([row["ball"] for row in rows], [row["pred"] for row in rows])
    example = next(row for row in rows if row["id"] == 1)
    variant = next(row for row in rows if row["diffs"])
    payload = {
        "n": 90,
        "n_train": 63,
        "n_test": 27,
        "min_sd": MIN_SD,
        "priors": model["prior"],
        "stats": model["stats"],
        "example": example,
        "variant_example": variant,
        "predictions": rows,
        "train_ids": split["train_ids"],
        "test_ids": split["test_ids"],
        **scored,
    }
    (DATA / "metrics.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    scatter(df)
    scheme()
    print("correct", scored["correct"], "example", example["id"], example["probs"])


if __name__ == "__main__":
    main()
