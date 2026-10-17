# -*- coding: utf-8 -*-
"""Наивный Байес по измерениям мяча. В текст отчёта не входит."""
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
MIN_SD = 1e-4
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


def load():
    DATA.mkdir(exist_ok=True)
    FIG.mkdir(exist_ok=True)
    for name in ("balls.csv", "split.json", "codes.json"):
        shutil.copyfile(SRC / "data" / name, DATA / name)
    split = json.loads((DATA / "split.json").read_text(encoding="utf-8"))
    codes = json.loads((DATA / "codes.json").read_text(encoding="utf-8"))
    df = pd.read_csv(DATA / "balls.csv")
    train = df[df["id"].isin(split["train_ids"])].copy()
    test = df[df["id"].isin(split["test_ids"])].copy()
    n_train = int(codes.get("n_train", split.get("n_train", 105)))
    n_test = int(codes.get("n_test", split.get("n_test", 45)))
    if len(train) != n_train or len(test) != n_test:
        raise RuntimeError(f"разбиение не {n_train}/{n_test}")
    if list(codes.get("features", FEATURES)) != FEATURES:
        raise RuntimeError("признаки не совпали")
    return df, train, test, split, codes


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
        tn = len(y_true) - tp - fp - fn
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


def scatter(df: pd.DataFrame, error_ids: list[int]) -> None:
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


def scheme() -> None:
    plt.rcParams["font.family"] = "Times New Roman"
    labels = [
        "Чтение таблицы измерений мяча",
        "Разбиение 70 % / 30 % по классу",
        "Средние и отклонения признаков по классу",
        "Наивный Байес на тесте",
        "Матрица ошибок и метрики качества",
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


def main() -> None:
    df, train, test, split, codes = load()
    typical = codes["typical"]
    model = fit(train)
    rows = []
    for _, row in test.sort_values("id").iterrows():
        result = predict_row(row, model)
        features = {feature: float(row[feature]) for feature in FEATURES}
        diffs = row_diff(row["ball"], features, typical)
        rows.append(
            {
                "id": int(row["id"]),
                "ball": row["ball"],
                "pred": result["pred"],
                "correct": result["pred"] == row["ball"],
                "probs": result["probs"],
                "features": features,
                "diffs": diffs,
            }
        )
    scored = metrics([row["ball"] for row in rows], [row["pred"] for row in rows])
    n_total = int(codes.get("n", len(df)))
    n_train = int(codes.get("n_train", len(train)))
    n_test = int(codes.get("n_test", len(test)))
    if scored["correct"] >= n_test:
        raise RuntimeError("на тесте нужны ошибки, не 100 %")
    if scored["correct"] / n_test < 0.80:
        raise RuntimeError(f"Байес ниже 80 %: {scored['correct']}/{n_test}")
    example = next(
        (row for row in rows if row["correct"] and not row["diffs"]),
        None,
    )
    if example is None:
        example = next(row for row in rows if row["correct"])
    variant = next((row for row in rows if row["diffs"]), rows[0])
    error_ids = [row["id"] for row in rows if not row["correct"]]
    payload = {
        "n": n_total,
        "n_train": n_train,
        "n_test": n_test,
        "n_train_per": int(codes.get("n_train_per", split.get("n_train_per", 35))),
        "n_test_per": int(codes.get("n_test_per", split.get("n_test_per", 15))),
        "min_sd": MIN_SD,
        "priors": model["prior"],
        "stats": model["stats"],
        "example": example,
        "variant_example": variant,
        "predictions": rows,
        "train_ids": split["train_ids"],
        "test_ids": split["test_ids"],
        "error_ids": error_ids,
        "typical": typical,
        "seed": split["seed"],
        **scored,
    }
    (DATA / "metrics.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    scatter(df, error_ids)
    scheme()
    print("correct", scored["correct"], "example", example["id"])
    print("errors", error_ids)
    print("floored", sum(1 for s in model["stats"] if s["floored"]))


if __name__ == "__main__":
    main()
