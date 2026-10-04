# -*- coding: utf-8 -*-
"""Расчёт k ближайших соседей для заявок мастерской. В отчёт не входит."""
from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"

CLASSES = ["прокол", "тормоза", "трансмиссия"]
FEATURES = ["pressure_bar", "pad_mm", "chain_pct", "lever_mm"]
K_CANDIDATES = [1, 3, 5, 7, 9, 11, 13, 15]


def generate(rng: np.random.Generator) -> pd.DataFrame:
    """Три класса с шумом и пограничными заявками, у которых свой симптом слабый."""
    specs = {
        "прокол": dict(pressure=(1.95, 0.42), pad=(1.95, 0.42), chain=(0.55, 0.16), lever=(13.0, 3.2)),
        "тормоза": dict(pressure=(2.70, 0.40), pad=(1.15, 0.36), chain=(0.58, 0.16), lever=(18.5, 3.6)),
        "трансмиссия": dict(pressure=(2.65, 0.40), pad=(1.80, 0.42), chain=(0.82, 0.18), lever=(13.5, 3.2)),
    }
    # Часть заявок со слабым главным симптомом: класс есть, а измерения уже похожи на соседний.
    boundary = {
        "прокол": dict(pressure=(2.55, 0.16), n=7),
        "тормоза": dict(pad=(1.70, 0.14), lever=(14.0, 1.5), n=7),
        "трансмиссия": dict(chain=(0.60, 0.07), n=7),
    }
    limits = {
        "pressure_bar": (0.70, 4.20, 2),
        "pad_mm": (0.20, 3.40, 2),
        "chain_pct": (0.10, 1.70, 2),
        "lever_mm": (6.0, 32.0, 1),
    }
    rows = []
    for name in CLASSES:
        spec = specs[name]
        n = 30
        pressure = rng.normal(spec["pressure"][0], spec["pressure"][1], n)
        pad = rng.normal(spec["pad"][0], spec["pad"][1], n)
        chain = rng.normal(spec["chain"][0], spec["chain"][1], n)
        lever = rng.normal(spec["lever"][0], spec["lever"][1], n)
        b = boundary[name]
        idx = rng.choice(n, size=b["n"], replace=False)
        if "pressure" in b:
            pressure[idx] = rng.normal(b["pressure"][0], b["pressure"][1], b["n"])
        if "pad" in b:
            pad[idx] = rng.normal(b["pad"][0], b["pad"][1], b["n"])
        if "lever" in b:
            lever[idx] = rng.normal(b["lever"][0], b["lever"][1], b["n"])
        if "chain" in b:
            chain[idx] = rng.normal(b["chain"][0], b["chain"][1], b["n"])
        for i in range(n):
            raw = {
                "pressure_bar": pressure[i],
                "pad_mm": pad[i],
                "chain_pct": chain[i],
                "lever_mm": lever[i],
            }
            item = {"fault": name}
            for key, (lo, hi, nd) in limits.items():
                item[key] = float(np.round(np.clip(raw[key], lo, hi), nd))
            rows.append(item)
    df = pd.DataFrame(rows)
    order = rng.permutation(len(df))
    df = df.iloc[order].reset_index(drop=True)
    df.insert(0, "request_id", np.arange(1, len(df) + 1))
    return df


def scale_fit(x: np.ndarray) -> tuple[np.ndarray, np.ndarray]:
    mu = x.mean(axis=0)
    sd = x.std(axis=0, ddof=0)
    sd = np.where(sd == 0, 1.0, sd)
    return mu, sd


def scale_apply(x: np.ndarray, mu: np.ndarray, sd: np.ndarray) -> np.ndarray:
    return (x - mu) / sd


def predict_one(z_train: np.ndarray, y_train: np.ndarray, z: np.ndarray, k: int) -> tuple[str, np.ndarray, bool]:
    dist = np.sqrt(((z_train - z) ** 2).sum(axis=1))
    nn = np.argsort(dist, kind="mergesort")[:k]
    labels = y_train[nn]
    values, counts = np.unique(labels, return_counts=True)
    best = counts.max()
    candidates = values[counts == best]
    tie = len(candidates) > 1
    if not tie:
        pred = str(candidates[0])
    else:
        pred = None
        allowed = set(candidates.tolist())
        for i in nn:
            if y_train[i] in allowed:
                pred = str(y_train[i])
                break
    return pred, nn, tie


def loo_accuracy(x: np.ndarray, y: np.ndarray, k: int) -> float:
    n = len(y)
    correct = 0
    for i in range(n):
        mask = np.ones(n, dtype=bool)
        mask[i] = False
        mu, sd = scale_fit(x[mask])
        z_train = scale_apply(x[mask], mu, sd)
        z = scale_apply(x[i : i + 1], mu, sd)[0]
        pred, _, _ = predict_one(z_train, y[mask], z, k)
        correct += int(pred == y[i])
    return correct / n


def confusion(y_true: np.ndarray, y_pred: list[str]) -> np.ndarray:
    index = {name: i for i, name in enumerate(CLASSES)}
    cm = np.zeros((3, 3), dtype=int)
    for t, p in zip(y_true, y_pred):
        cm[index[t], index[p]] += 1
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


def stratified_split(df: pd.DataFrame, rng: np.random.Generator, test_per_class: int = 9):
    train_idx = []
    test_idx = []
    for name in CLASSES:
        idx = df.index[df["fault"] == name].to_numpy()
        take = rng.choice(idx, size=test_per_class, replace=False)
        test_idx.extend(take.tolist())
        train_idx.extend([i for i in idx if i not in set(take.tolist())])
    train_idx.sort()
    test_idx.sort()
    return df.loc[train_idx].reset_index(drop=True), df.loc[test_idx].reset_index(drop=True)


def main() -> None:
    rng = np.random.default_rng(4321)
    df = generate(rng)
    split_rng = np.random.default_rng(42)
    train, test = stratified_split(df, split_rng)
    x_train = train[FEATURES].to_numpy(float)
    y_train = train["fault"].to_numpy()
    x_test = test[FEATURES].to_numpy(float)
    y_test = test["fault"].to_numpy()

    loo = []
    for k in K_CANDIDATES:
        acc = loo_accuracy(x_train, y_train, k)
        loo.append({"k": k, "correct": int(round(acc * len(y_train))), "n": int(len(y_train)), "accuracy": acc})

    best = max(loo, key=lambda row: (row["accuracy"], -row["k"]))
    # При равной доле верных берём меньшее k: граница меньше размывается.
    top = max(row["accuracy"] for row in loo)
    chosen = min(row["k"] for row in loo if row["accuracy"] == top)

    mu, sd = scale_fit(x_train)
    z_train = scale_apply(x_train, mu, sd)
    z_test = scale_apply(x_test, mu, sd)

    preds = []
    details = []
    ties = 0
    for i in range(len(test)):
        pred, nn, tie = predict_one(z_train, y_train, z_test[i], chosen)
        ties += int(tie)
        preds.append(pred)
        neighbors = []
        dist = np.sqrt(((z_train - z_test[i]) ** 2).sum(axis=1))
        for j in nn:
            neighbors.append(
                {
                    "request_id": int(train.iloc[j]["request_id"]),
                    "fault": str(y_train[j]),
                    "distance": float(dist[j]),
                }
            )
        details.append(
            {
                "request_id": int(test.iloc[i]["request_id"]),
                "fault": str(y_test[i]),
                "pred": pred,
                "correct": pred == y_test[i],
                "tie": tie,
                "pressure_bar": float(test.iloc[i]["pressure_bar"]),
                "pad_mm": float(test.iloc[i]["pad_mm"]),
                "chain_pct": float(test.iloc[i]["chain_pct"]),
                "lever_mm": float(test.iloc[i]["lever_mm"]),
                "neighbors": neighbors,
            }
        )

    cm = confusion(y_test, preds)
    per_class = class_metrics(cm)
    correct = int(cm.trace())
    n_test = int(cm.sum())
    accuracy = correct / n_test
    macro_p = float(np.mean([r["precision"] for r in per_class]))
    macro_r = float(np.mean([r["recall"] for r in per_class]))
    macro_f = float(np.mean([r["f1"] for r in per_class]))

    # Центроиды классов в исходных единицах, по всей выборке — для текста про пересечение.
    centroids = {}
    for name in CLASSES:
        part = df.loc[df["fault"] == name, FEATURES]
        centroids[name] = {col: float(part[col].mean()) for col in FEATURES}

    payload = {
        "n": int(len(df)),
        "n_train": int(len(train)),
        "n_test": n_test,
        "per_class_total": 30,
        "per_class_train": 21,
        "per_class_test": 9,
        "k": chosen,
        "loo": loo,
        "ties": ties,
        "mu": {col: float(mu[i]) for i, col in enumerate(FEATURES)},
        "sd": {col: float(sd[i]) for i, col in enumerate(FEATURES)},
        "confusion": cm.tolist(),
        "correct": correct,
        "errors": n_test - correct,
        "accuracy": accuracy,
        "macro_precision": macro_p,
        "macro_recall": macro_r,
        "macro_f1": macro_f,
        "per_class": per_class,
        "train_ids": [int(v) for v in train["request_id"]],
        "test_ids": [int(v) for v in test["request_id"]],
        "predictions": details,
        "centroids": centroids,
        "ranges": {
            col: {"min": float(df[col].min()), "max": float(df[col].max())} for col in FEATURES
        },
        "class_counts": {name: int((df["fault"] == name).sum()) for name in CLASSES},
    }
    DATA.mkdir(parents=True, exist_ok=True)
    df.to_csv(DATA / "requests.csv", index=False)
    (DATA / "metrics.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    print("k", chosen)
    print("loo")
    for row in loo:
        print(f"  k={row['k']} {row['correct']}/{row['n']} {row['accuracy']:.4f}")
    print("test", correct, "/", n_test, f"{accuracy:.4f}", "ties", ties)
    print("cm")
    print(cm)
    for row in per_class:
        print(row["class"], f"P={row['precision']:.3f}", f"R={row['recall']:.3f}", f"F={row['f1']:.3f}", "tp", row["tp"])
    print("macro", f"{macro_p:.3f}", f"{macro_r:.3f}", f"{macro_f:.3f}")
    print("errors:")
    for d in details:
        if not d["correct"]:
            votes = {}
            for nb in d["neighbors"]:
                votes[nb["fault"]] = votes.get(nb["fault"], 0) + 1
            print(
                d["request_id"],
                d["fault"],
                "->",
                d["pred"],
                votes,
                "p",
                d["pressure_bar"],
                "pad",
                d["pad_mm"],
                "chain",
                d["chain_pct"],
                "lever",
                d["lever_mm"],
            )


if __name__ == "__main__":
    main()
