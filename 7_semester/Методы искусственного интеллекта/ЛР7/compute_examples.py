# -*- coding: utf-8 -*-
"""Короткий ручной пример трёх алгоритмов. В отчёт этот файл не попадает."""
from __future__ import annotations

import json
import math
from pathlib import Path

import matplotlib.pyplot as plt

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
FIG = BASE / "figures"

OBJECTS = [
    {"name": "В1", "ball": "волейбольный", "size": 68, "bounce": 116},
    {"name": "В2", "ball": "волейбольный", "size": 69, "bounce": 119},
    {"name": "Ф1", "ball": "футбольный", "size": 70, "bounce": 126},
    {"name": "Ф2", "ball": "футбольный", "size": 72, "bounce": 127},
    {"name": "Б1", "ball": "баскетбольный", "size": 73, "bounce": 136},
    {"name": "Б2", "ball": "баскетбольный", "size": 75, "bounce": 140},
]
QUERY = {"name": "Н", "size": 71, "bounce": 125}


def knn() -> dict:
    rows = []
    for obj in OBJECTS:
        d2 = (QUERY["size"] - obj["size"]) ** 2 + (QUERY["bounce"] - obj["bounce"]) ** 2
        rows.append({**obj, "d2": d2, "d": math.sqrt(d2)})
    ranked = sorted(rows, key=lambda item: (item["d2"], item["name"]))
    for place, item in enumerate(ranked, start=1):
        item["rank"] = place
    top3 = ranked[:3]
    votes: dict[str, int] = {}
    for item in top3:
        votes[item["ball"]] = votes.get(item["ball"], 0) + 1
    winner = max(votes, key=votes.get)
    top5 = ranked[:5]
    votes5: dict[str, int] = {}
    for item in top5:
        votes5[item["ball"]] = votes5.get(item["ball"], 0) + 1
    if winner != "футбольный" or votes != {"футбольный": 2, "волейбольный": 1}:
        raise RuntimeError(votes)
    if votes5["футбольный"] != 2 or votes5["волейбольный"] != 2:
        raise RuntimeError(votes5)
    return {"query": QUERY, "ranked": ranked, "k": 3, "votes": votes, "winner": winner, "votes5": votes5}


def bayes() -> dict:
    groups = {
        "волейбольный": [67, 68, 69],
        "футбольный": [69, 70, 71],
        "баскетбольный": [72, 73, 74],
    }
    x = 71
    classes = []
    for name, values in groups.items():
        mu = sum(values) / len(values)
        var = sum((value - mu) ** 2 for value in values) / (len(values) - 1)
        weight = math.exp(-((x - mu) ** 2) / (2 * var))
        classes.append({"ball": name, "values": values, "mu": mu, "var": var, "weight": weight})
    total = sum(item["weight"] for item in classes)
    for item in classes:
        item["prob"] = item["weight"] / total
    boundary = 71.5
    tied = math.exp(-((boundary - 70) ** 2) / 2)
    tied_b = math.exp(-((boundary - 73) ** 2) / 2)
    if abs(tied - tied_b) > 1e-12:
        raise RuntimeError("граница 71,5 не симметрична")
    winner = max(classes, key=lambda item: item["prob"])["ball"]
    if winner != "футбольный":
        raise RuntimeError(winner)
    if abs(sum(item["prob"] for item in classes) - 1) > 1e-12:
        raise RuntimeError("вероятности не нормированы")
    return {"x": x, "classes": classes, "total": total, "boundary": boundary, "tied_weight": tied, "winner": winner}


def perceptron() -> dict:
    points = [
        {"name": "П1", "size": 68, "bounce": 116, "y": 0, "ball": "волейбольный"},
        {"name": "П2", "size": 71, "bounce": 126, "y": 1, "ball": "футбольный"},
        {"name": "П3", "size": 69, "bounce": 118, "y": 0, "ball": "волейбольный"},
        {"name": "П4", "size": 72, "bounce": 128, "y": 1, "ball": "футбольный"},
    ]
    for item in points:
        item["x1"] = item["size"] - 70
        item["x2"] = (item["bounce"] - 120) / 2
        if item["x2"] != int(item["x2"]):
            raise RuntimeError(item)
        item["x2"] = int(item["x2"])
    w1 = w2 = b = 0
    steps = []
    for item in points:
        z = w1 * item["x1"] + w2 * item["x2"] + b
        yhat = 1 if z >= 0 else 0
        err = item["y"] - yhat
        w1 += err * item["x1"]
        w2 += err * item["x2"]
        b += err
        steps.append({**item, "z": z, "yhat": yhat, "w1": w1, "w2": w2, "b": b})
    if (w1, w2, b) != (2, 2, -1):
        raise RuntimeError((w1, w2, b))
    checks = []
    for size, bounce, expect in ((70, 122, 1), (68, 120, 0)):
        x1 = size - 70
        x2 = (bounce - 120) / 2
        z = w1 * x1 + w2 * x2 + b
        yhat = 1 if z >= 0 else 0
        if yhat != expect:
            raise RuntimeError((size, bounce, z, yhat))
        checks.append({"size": size, "bounce": bounce, "x1": x1, "x2": int(x2), "z": z, "yhat": yhat})
    return {"points": points, "steps": steps, "w1": w1, "w2": w2, "b": b, "checks": checks}


def draw(knn_rows: list[dict]) -> None:
    FIG.mkdir(exist_ok=True)
    colors = {"волейбольный": "#2c7bb6", "футбольный": "#1a7f37", "баскетбольный": "#d95f02"}
    fig, ax = plt.subplots(figsize=(7.2, 4.6))
    for ball, color in colors.items():
        pts = [row for row in knn_rows if row["ball"] == ball]
        ax.scatter(
            [row["size"] for row in pts],
            [row["bounce"] for row in pts],
            s=60,
            color=color,
            label=ball,
            zorder=3,
        )
        for row in pts:
            ax.annotate(row["name"], (row["size"] + 0.12, row["bounce"] + 0.4), fontsize=9)
    ax.scatter([QUERY["size"]], [QUERY["bounce"]], s=120, marker="*", color="black", label="новое Н", zorder=4)
    ax.annotate("Н", (QUERY["size"] + 0.12, QUERY["bounce"] + 0.4), fontsize=9)
    ax.set_xlabel("окружность, см")
    ax.set_ylabel("отскок, см")
    ax.legend(frameon=False)
    ax.set_xlim(67, 76)
    ax.set_ylim(112, 144)
    fig.tight_layout()
    fig.savefig(FIG / "neighbors.png", dpi=140)
    plt.close(fig)


def main() -> None:
    DATA.mkdir(exist_ok=True)
    payload = {"knn": knn(), "bayes": bayes(), "perceptron": perceptron()}
    draw(payload["knn"]["ranked"])
    (DATA / "examples.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print("knn", payload["knn"]["winner"], payload["knn"]["votes"])
    print("bayes", [(c["ball"], round(c["prob"], 4)) for c in payload["bayes"]["classes"]])
    print("perceptron", payload["perceptron"]["w1"], payload["perceptron"]["w2"], payload["perceptron"]["b"])


if __name__ == "__main__":
    main()
