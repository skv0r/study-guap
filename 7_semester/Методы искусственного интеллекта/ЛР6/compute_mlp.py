# -*- coding: utf-8 -*-
"""Обучение многослойного перцептрона на общем тесте мячей. В отчёт не попадает."""
from __future__ import annotations

import json
import shutil
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from sklearn.metrics import confusion_matrix, precision_recall_fscore_support
from sklearn.neural_network import MLPClassifier
from sklearn.preprocessing import StandardScaler

BASE = Path(__file__).resolve().parent
SRC = BASE.parent / "ЛР5" / "data"
DATA = BASE / "data"
FIG = BASE / "figures"
FEATURES = ["elasticity", "size", "mass", "pressure", "bounce"]
CLASSES = ["футбольный", "баскетбольный", "волейбольный"]
USUAL = {
    "футбольный": [0.75, 70.0, 420.0, 0.55, 125.0],
    "баскетбольный": [0.82, 72.0, 480.0, 0.5, 135.0],
    "волейбольный": [0.68, 69.0, 380.0, 0.42, 118.0],
}
NAMES = {
    "elasticity": "упругость",
    "size": "окружность",
    "mass": "масса",
    "pressure": "давление",
    "bounce": "отскок",
}


def phrase(feature: str, value: float, usual: float) -> str:
    if feature in ("elasticity", "pressure"):
        left, right = f"{value:.2f}", f"{usual:.2f}"
    elif feature == "size":
        left, right = f"{value:.1f}", f"{usual:.1f}"
    else:
        left, right = str(int(round(value))), str(int(round(usual)))
    return f"{NAMES[feature]} {left} вместо {right}".replace(".", ",")


def main() -> None:
    DATA.mkdir(exist_ok=True)
    FIG.mkdir(exist_ok=True)
    shutil.copy2(SRC / "balls.csv", DATA / "balls.csv")
    split = json.loads((SRC / "split.json").read_text(encoding="utf-8"))
    (DATA / "split.json").write_text(json.dumps(split, ensure_ascii=False, indent=2), encoding="utf-8")

    df = pd.read_csv(DATA / "balls.csv")
    train = df[df["id"].isin(split["train_ids"])].copy()
    test = df[df["id"].isin(split["test_ids"])].copy()
    if len(train) != 105 or len(test) != 45:
        raise SystemExit(f"split {len(train)} {len(test)}")

    scaler = StandardScaler()
    x_train = scaler.fit_transform(train[FEATURES].to_numpy(float))
    x_test = scaler.transform(test[FEATURES].to_numpy(float))
    y_train = train["ball"].to_numpy()
    y_test = test["ball"].to_numpy()

    model = MLPClassifier(
        hidden_layer_sizes=(10,),
        activation="relu",
        solver="lbfgs",
        max_iter=500,
        random_state=300,
        learning_rate_init=0.01,
    )
    model.fit(x_train, y_train)
    pred = model.predict(x_test)
    proba = model.predict_proba(x_test)
    labels = list(model.classes_)
    matrix = confusion_matrix(y_test, pred, labels=CLASSES)
    correct = int((pred == y_test).sum())
    precision, recall, f1, _support = precision_recall_fscore_support(
        y_test, pred, labels=CLASSES, average=None, zero_division=0
    )
    macro_p, macro_r, macro_f, _ = precision_recall_fscore_support(
        y_test, pred, average="macro", zero_division=0
    )

    rows = []
    for index, row in enumerate(test.itertuples(index=False)):
        values = [float(getattr(row, feature)) for feature in FEATURES]
        usual = USUAL[row.ball]
        diffs = []
        for feature, value, center in zip(FEATURES, values, usual):
            if abs(value - center) > 1e-9:
                diffs.append({"feature": feature, "usual": center, "value": value, "phrase": phrase(feature, value, center)})
        probs = {label: float(proba[index, labels.index(label)]) for label in CLASSES}
        rows.append(
            {
                "id": int(row.id),
                "ball": row.ball,
                "pred": str(pred[index]),
                "correct": bool(pred[index] == row.ball),
                "probs": probs,
                "features": dict(zip(FEATURES, values)),
                "diffs": diffs,
            }
        )
    correct_rows = [item for item in rows if item["correct"]]
    example = min(correct_rows, key=lambda item: max(item["probs"].values()))
    errors = [item for item in rows if not item["correct"]]
    prev = json.loads((SRC / "metrics.json").read_text(encoding="utf-8"))
    comparison = list(prev["comparison"])
    comparison.append(
        {
            "lab": "6",
            "method": "нейронная сеть",
            "correct": correct,
            "n_test": 45,
            "accuracy": correct / 45,
            "precision": float(macro_p),
            "recall": float(macro_r),
            "f1": float(macro_f),
        }
    )

    hidden_w = model.coefs_[0]
    out_w = model.coefs_[1]
    payload = {
        "n": 150,
        "n_train": 105,
        "n_test": 45,
        "n_train_per": 35,
        "n_test_per": 15,
        "hidden": 10,
        "max_iter": 500,
        "n_iter": int(model.n_iter_),
        "loss": float(model.loss_),
        "inputs": 5,
        "outputs": 3,
        "activation": "relu",
        "solver": "lbfgs",
        "seed": 300,
        "mu": {feature: float(scaler.mean_[i]) for i, feature in enumerate(FEATURES)},
        "scale": {feature: float(scaler.scale_[i]) for i, feature in enumerate(FEATURES)},
        "correct": correct,
        "accuracy": correct / 45,
        "confusion": matrix.tolist(),
        "classes": CLASSES,
        "per_class": [
            {
                "class": name,
                "precision": float(precision[i]),
                "recall": float(recall[i]),
                "f1": float(f1[i]),
                "support": 15,
            }
            for i, name in enumerate(CLASSES)
        ],
        "macro_precision": float(macro_p),
        "macro_recall": float(macro_r),
        "macro_f1": float(macro_f),
        "example": example,
        "error_ids": [item["id"] for item in errors],
        "errors": [
            {
                "id": item["id"],
                "ball": item["ball"],
                "pred": item["pred"],
                "features": item["features"],
                "probs": item["probs"],
            }
            for item in errors
        ],
        "knn_k": prev.get("knn_k", 5),
        "tree_depth": prev.get("tree_depth", 3),
        "comparison": comparison,
        "hidden_bias": [float(v) for v in model.intercepts_[0]],
        "output_bias": [float(v) for v in model.intercepts_[1]],
        "first_hidden_weights": [float(v) for v in hidden_w[:, 0]],
        "output_weight_abs_mean": float(np.mean(np.abs(out_w))),
        "typical": USUAL,
    }
    (DATA / "metrics.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    fig, ax = plt.subplots(figsize=(8.2, 3.2))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 4)
    ax.axis("off")
    ax.add_patch(plt.Rectangle((0.4, 1.2), 2.2, 1.6, fill=False, linewidth=1.2))
    ax.text(1.5, 2.0, "5 входов\nупругость, окружность,\nмасса, давление, отскок", ha="center", va="center", fontsize=8)
    ax.add_patch(plt.Rectangle((3.8, 0.7), 2.4, 2.6, fill=False, linewidth=1.2))
    ax.text(5.0, 2.0, "скрытый слой\n10 нейронов\nReLU", ha="center", va="center", fontsize=9)
    ax.add_patch(plt.Rectangle((7.4, 1.2), 2.2, 1.6, fill=False, linewidth=1.2))
    ax.text(8.5, 2.0, "3 выхода\nфутбол, баскетбол,\nволейбол", ha="center", va="center", fontsize=8)
    ax.annotate("", xy=(3.8, 2.0), xytext=(2.6, 2.0), arrowprops={"arrowstyle": "->"})
    ax.annotate("", xy=(7.4, 2.0), xytext=(6.2, 2.0), arrowprops={"arrowstyle": "->"})
    fig.tight_layout()
    fig.savefig(FIG / "scheme.png", dpi=140)
    plt.close(fig)
    print("correct", correct, "iters", model.n_iter_, "loss", round(model.loss_, 4))
    print("example", example["id"], {k: round(v, 4) for k, v in example["probs"].items()})
    print("errors", [(item["id"], item["ball"], item["pred"]) for item in errors])


if __name__ == "__main__":
    main()
