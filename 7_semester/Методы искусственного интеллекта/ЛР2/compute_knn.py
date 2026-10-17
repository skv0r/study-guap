# -*- coding: utf-8 -*-
"""k ближайших соседей по измерениям мяча. В отчёт не входит."""
from __future__ import annotations

import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
FIG = BASE / "figures"

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
K_CANDIDATES = [1, 3, 5, 7, 9, 11, 13, 15]
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


def scale_fit(x: np.ndarray) -> tuple[np.ndarray, np.ndarray]:
    mu = x.mean(axis=0)
    sd = x.std(axis=0, ddof=0)
    sd = np.where(sd == 0, 1.0, sd)
    return mu, sd


def scale_apply(x: np.ndarray, mu: np.ndarray, sd: np.ndarray) -> np.ndarray:
    return (x - mu) / sd


def predict_one(z_train: np.ndarray, y_train: np.ndarray, z: np.ndarray, k: int):
    dist = np.sqrt(((z_train - z) ** 2).sum(axis=1))
    nn = np.argsort(dist, kind="mergesort")[:k]
    labels = y_train[nn]
    values, counts = np.unique(labels, return_counts=True)
    best = int(counts.max())
    candidates = values[counts == best]
    tie = len(candidates) > 1
    if not tie:
        pred = str(candidates[0])
    else:
        allowed = set(candidates.tolist())
        pred = str(next(y_train[i] for i in nn if y_train[i] in allowed))
    return pred, nn, dist, tie


def loo_correct(x: np.ndarray, y: np.ndarray, k: int) -> int:
    n = len(y)
    correct = 0
    for i in range(n):
        mask = np.ones(n, dtype=bool)
        mask[i] = False
        mu, sd = scale_fit(x[mask])
        z_train = scale_apply(x[mask], mu, sd)
        z = scale_apply(x[i : i + 1], mu, sd)[0]
        pred, _, _, _ = predict_one(z_train, y[mask], z, k)
        correct += int(pred == y[i])
    return correct


def confusion(y_true: np.ndarray, y_pred: list[str]) -> np.ndarray:
    index = {name: i for i, name in enumerate(CLASSES)}
    cm = np.zeros((3, 3), dtype=int)
    for true, pred in zip(y_true, y_pred):
        cm[index[true], index[pred]] += 1
    return cm


def class_metrics(cm: np.ndarray) -> list[dict]:
    rows = []
    n = int(cm.sum())
    for i, name in enumerate(CLASSES):
        tp = int(cm[i, i])
        fp = int(cm[:, i].sum() - tp)
        fn = int(cm[i, :].sum() - tp)
        tn = n - tp - fp - fn
        precision = tp / (tp + fp) if tp + fp else 0.0
        recall = tp / (tp + fn) if tp + fn else 0.0
        f1 = 2 * precision * recall / (precision + recall) if precision + recall else 0.0
        rows.append(
            {
                "class": name,
                "tp": tp,
                "fp": fp,
                "tn": tn,
                "fn": fn,
                "precision": precision,
                "recall": recall,
                "f1": f1,
                "support": int(cm[i, :].sum()),
            }
        )
    return rows


def draw_scheme(path: Path) -> None:
    plt.rcParams["font.family"] = "Times New Roman"
    labels = [
        "Чтение таблицы измерений мяча",
        "Цвет класса для диаграммы",
        "Разбиение 70 % / 30 % по классу",
        "Общий масштаб пяти признаков по обучению",
        "Выбор нечётного k по обучающей выборке",
        "Классификация k ближайшими соседями",
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
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=160, bbox_inches="tight", facecolor="white", pad_inches=0.12)
    plt.close()


def draw_scatter(df: pd.DataFrame, error_ids: list[int], path: Path) -> None:
    plt.rcParams["font.family"] = "Times New Roman"
    fig, ax = plt.subplots(figsize=(7.6, 5.8))
    for name in CLASSES:
        part = df[df["ball"] == name]
        ax.scatter(
            part["size"],
            part["elasticity"],
            c=COLORS[name],
            s=70,
            label=name,
            zorder=3,
            edgecolors="white",
            linewidths=0.6,
            alpha=0.9,
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
    fig.savefig(path, dpi=160, bbox_inches="tight", facecolor="white")
    plt.close()


def neighbor_record(train_ids, y_train, x_train, nn, dist) -> list[dict]:
    rows = []
    for pos in nn:
        rows.append(
            {
                "id": int(train_ids[pos]),
                "ball": str(y_train[pos]),
                "distance": float(dist[pos]),
                "features": {feature: float(x_train[pos, j]) for j, feature in enumerate(FEATURES)},
            }
        )
    return rows


def row_diff(ball: str, features: dict, typical: dict) -> dict:
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
    return {"diffs": diffs, "variant_name": ""}


def main() -> None:
    requests = DATA / "requests.csv"
    if not requests.exists():
        raise RuntimeError("requests.csv должен остаться")
    df = pd.read_csv(DATA / "balls.csv")
    split = json.loads((DATA / "split.json").read_text(encoding="utf-8"))
    codes = json.loads((DATA / "codes.json").read_text(encoding="utf-8"))
    typical = codes["typical"]
    if list(codes.get("features", FEATURES)) != FEATURES:
        raise RuntimeError("признаки в codes.json не совпали")
    n_total = int(codes.get("n", len(df)))
    n_train = int(codes.get("n_train", split.get("n_train", 105)))
    n_test = int(codes.get("n_test", split.get("n_test", 45)))
    n_train_per = int(codes.get("n_train_per", split.get("n_train_per", 35)))
    n_test_per = int(codes.get("n_test_per", split.get("n_test_per", 15)))
    if len(df) != n_total:
        raise RuntimeError(f"в таблице не {n_total} строк")
    for feature in FEATURES:
        if feature not in df.columns:
            raise RuntimeError(f"нет столбца {feature}")
    X = df[FEATURES].to_numpy(dtype=float)
    y = df["ball"].to_numpy()
    ids = df["id"].to_numpy()
    seed = int(split["seed"])
    X_train, X_test, y_train, y_test, id_train, id_test = train_test_split(
        X, y, ids, test_size=0.30, random_state=seed, stratify=y
    )
    train_ids = [int(i) for i in id_train]
    test_ids = [int(i) for i in id_test]
    if sorted(train_ids) != split["train_ids"] or sorted(test_ids) != split["test_ids"]:
        raise RuntimeError("разбиение не совпало с деревом решений")
    if sorted(pd.Series(y_train).value_counts().tolist()) != [n_train_per] * 3:
        raise RuntimeError(f"в обучении не по {n_train_per}")
    if sorted(pd.Series(y_test).value_counts().tolist()) != [n_test_per] * 3:
        raise RuntimeError(f"в тесте не по {n_test_per}")
    if len(train_ids) != n_train or len(test_ids) != n_test:
        raise RuntimeError("объём обучения/теста не совпал")

    loo = []
    for k in K_CANDIDATES:
        ok = loo_correct(X_train, y_train, k)
        loo.append({"k": k, "correct": ok, "n": n_train, "accuracy": ok / n_train})
    best_correct = max(row["correct"] for row in loo)
    winners = [row["k"] for row in loo if row["correct"] == best_correct]
    k_star = min(winners)
    if k_star % 2 == 0 or any(k % 2 == 0 for k in winners):
        raise RuntimeError("k должно быть нечётным")

    mu, sd = scale_fit(X_train)
    z_train = scale_apply(X_train, mu, sd)
    z_test = scale_apply(X_test, mu, sd)

    by_id = {int(row.id): row for row in df.itertuples(index=False)}
    predictions = []
    for i in np.argsort(id_test):
        pred, nn, dist, tie = predict_one(z_train, y_train, z_test[i], k_star)
        votes = {name: int((y_train[nn] == name).sum()) for name in CLASSES}
        features = {feature: float(X_test[i, j]) for j, feature in enumerate(FEATURES)}
        described = row_diff(str(y_test[i]), features, typical)
        predictions.append(
            {
                "id": int(id_test[i]),
                "ball": str(y_test[i]),
                "pred": pred,
                "correct": bool(pred == y_test[i]),
                "tie": bool(tie),
                "votes": votes,
                "features": features,
                "neighbors": neighbor_record(id_train, y_train, X_train, nn, dist),
                "diffs": described["diffs"],
                "variant_name": described["variant_name"],
            }
        )
    order = np.argsort(id_test)
    y_true_ordered = y_test[order]
    cm = confusion(y_true_ordered, [item["pred"] for item in predictions])
    per_class = class_metrics(cm)
    correct = int(sum(item["correct"] for item in predictions))
    n_ties = int(sum(item["tie"] for item in predictions))
    example = next(
        (
            item
            for item in predictions
            if item["correct"] and not item["tie"] and not item["diffs"]
        ),
        None,
    )
    if example is None:
        example = next(item for item in predictions if item["correct"] and not item["tie"])

    mass_i = FEATURES.index("mass")
    elast_i = FEATURES.index("elasticity")
    if correct >= n_test:
        raise RuntimeError("kNN дал 100 % на тесте — данные слишком простые")
    if correct / n_test < 0.80:
        raise RuntimeError(f"kNN ниже 80 %: {correct}/{n_test}")
    metrics = {
        "n": n_total,
        "n_train": n_train,
        "n_test": n_test,
        "n_train_per": n_train_per,
        "n_test_per": n_test_per,
        "k": k_star,
        "k_tied": winners,
        "loo": loo,
        "correct": correct,
        "errors": n_test - correct,
        "ties_on_test": n_ties,
        "accuracy": correct / n_test,
        "confusion": cm.tolist(),
        "per_class": per_class,
        "macro_precision": float(np.mean([row["precision"] for row in per_class])),
        "macro_recall": float(np.mean([row["recall"] for row in per_class])),
        "macro_f1": float(np.mean([row["f1"] for row in per_class])),
        "mu": {feature: float(mu[i]) for i, feature in enumerate(FEATURES)},
        "sd": {feature: float(sd[i]) for i, feature in enumerate(FEATURES)},
        "train_ids": sorted(train_ids),
        "test_ids": sorted(test_ids),
        "error_ids": [item["id"] for item in predictions if not item["correct"]],
        "predictions": predictions,
        "example_id": int(example["id"]),
        "typical": typical,
        "scale_note": {
            "mass_sd_g": float(sd[mass_i]),
            "elasticity_sd": float(sd[elast_i]),
            "mass_10g_scaled": float(10.0 / sd[mass_i]),
            "elasticity_001_scaled": float(0.01 / sd[elast_i]),
        },
        "classes_present": sorted(df["ball"].unique().tolist()),
        "seed": seed,
    }
    if not requests.exists():
        raise RuntimeError("requests.csv должен остаться")
    assert by_id[example["id"]].ball == example["ball"]
    (DATA / "metrics.json").write_text(json.dumps(metrics, ensure_ascii=False, indent=2), encoding="utf-8")
    FIG.mkdir(parents=True, exist_ok=True)
    draw_scheme(FIG / "scheme.png")
    draw_scatter(df, metrics["error_ids"], FIG / "scatter.png")
    print("k", k_star, "tied", winners, "correct", correct, "ties", n_ties)
    print("errors", metrics["error_ids"])
    print("loo", [(row["k"], row["correct"]) for row in loo])
    print("sd", metrics["sd"])
    for item in predictions:
        if not item["correct"]:
            print(item["id"], item["ball"], item["pred"], item["votes"], item["features"])


if __name__ == "__main__":
    main()
