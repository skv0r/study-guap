# -*- coding: utf-8 -*-
"""Наблюдения о мяче и дерево решений. В текст отчёта не входит."""
from __future__ import annotations

import json
import shutil
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.patches import FancyBboxPatch
from sklearn.metrics import confusion_matrix, precision_recall_fscore_support
from sklearn.model_selection import train_test_split
from sklearn.tree import DecisionTreeClassifier

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
FIG = BASE / "figures"
LR2_DATA = BASE.parent / "ЛР2" / "data"

FEATURES = ["net_share", "play_zone", "hold", "attack", "ground"]
FEATURE_RU = {
    "net_share": "Динамика владения",
    "play_zone": "Игровая зона",
    "hold": "Удержание контроля",
    "attack": "Вектор атаки",
    "ground": "Контакт с землёй",
}
SHORT = {
    "net_share": "динамика владения",
    "play_zone": "зона",
    "hold": "удержание",
    "attack": "вектор",
    "ground": "земля",
}
CLASSES = ["футбольный", "баскетбольный", "волейбольный"]
COLORS = {"футбольный": "#1f77b4", "баскетбольный": "#ff7f0e", "волейбольный": "#2ca02c"}
LEAF_FC = {"футбольный": "#cfe2f3", "баскетбольный": "#fce5cd", "волейбольный": "#d9ead3"}
TYPICAL = {
    "футбольный": [0, 0, 1, 0, 1],
    "баскетбольный": [0, 2, 2, 1, 1],
    "волейбольный": [1, 1, 0, 1, 0],
}
VARIANT = {
    "футбольный": {
        "feature": "attack",
        "value": 1,
        "name": "футбольный навес",
        "phrase": "вектор 1 вместо 0",
    },
    "волейбольный": {
        "feature": "ground",
        "value": 1,
        "name": "волейбольный приём, мяч коснулся пола",
        "phrase": "земля 1 вместо 0",
    },
    "баскетбольный": {
        "feature": "attack",
        "value": 0,
        "name": "баскетбольный пас по площадке без броска вверх",
        "phrase": "вектор 0 вместо 1",
    },
}
CODEBOOK = [
    ["Динамика владения", "0", "пас и ведение, удара через сетку нет", "футбольный, баскетбольный"],
    ["Динамика владения", "1", "розыгрыш ударами через сетку", "волейбольный"],
    ["Игровая зона", "0", "ворота на уровне поля", "футбольный"],
    ["Игровая зона", "1", "площадка соперника за сеткой", "волейбольный"],
    ["Игровая зона", "2", "кольцо над площадкой", "баскетбольный"],
    ["Удержание контроля", "0", "краткий удар, мяч не ведут", "волейбольный"],
    ["Удержание контроля", "1", "ведение ногой", "футбольный"],
    ["Удержание контроля", "2", "ведение рукой с отскоком и захватом", "баскетбольный"],
    ["Вектор атаки", "0", "горизонтальный", "футбольный"],
    ["Вектор атаки", "1", "вертикальный", "баскетбольный, волейбольный"],
    ["Контакт с землёй", "0", "нет, касание пола завершает розыгрыш", "волейбольный"],
    ["Контакт с землёй", "1", "да, мяч идёт по земле или отскок обязателен", "футбольный, баскетбольный"],
]
SEED = 24
MAX_DEPTH = 3
MIN_LEAF = 3


def generate(rng: np.random.Generator) -> pd.DataFrame:
    rows = []
    for name in CLASSES:
        typical = TYPICAL[name]
        spec = VARIANT[name]
        changed = typical.copy()
        changed[FEATURES.index(spec["feature"])] = spec["value"]
        for _ in range(27):
            item = {"ball": name}
            for feature, value in zip(FEATURES, typical):
                item[feature] = int(value)
            rows.append(item)
        for _ in range(3):
            item = {"ball": name}
            for feature, value in zip(FEATURES, changed):
                item[feature] = int(value)
            rows.append(item)
    df = pd.DataFrame(rows)
    df = df.iloc[rng.permutation(len(df))].reset_index(drop=True)
    df.insert(0, "id", np.arange(1, len(df) + 1))
    for feature in FEATURES:
        df[feature] = df[feature].astype(int)
    return df


def audit(df: pd.DataFrame) -> list[dict]:
    if len(df) != 90 or sorted(df["ball"].value_counts().tolist()) != [30, 30, 30]:
        raise RuntimeError("нужно 90 строк, по 30 в классе")
    raw = df[FEATURES].to_numpy()
    if not set(int(v) for v in raw.ravel()) <= {0, 1, 2}:
        raise RuntimeError("код вне 0, 1, 2")
    found = []
    for name in CLASSES:
        spec = VARIANT[name]
        index = FEATURES.index(spec["feature"])
        n_typical = 0
        n_variant = 0
        for row in df[df["ball"] == name].itertuples(index=False):
            vec = [int(getattr(row, feature)) for feature in FEATURES]
            changed = [i for i, (got, exp) in enumerate(zip(vec, TYPICAL[name])) if got != exp]
            if not changed:
                n_typical += 1
                continue
            if changed != [index] or vec[index] != spec["value"]:
                raise RuntimeError(f"неожиданный код в строке {int(row.id)}: {vec}")
            n_variant += 1
            found.append(
                {
                    "id": int(row.id),
                    "ball": name,
                    "name": spec["name"],
                    "feature": spec["feature"],
                    "usual": int(TYPICAL[name][index]),
                    "value": int(spec["value"]),
                    "phrase": spec["phrase"],
                }
            )
        if n_typical != 27 or n_variant != 3:
            raise RuntimeError(f"{name}: обычных {n_typical}, особых {n_variant}")
    return found


def row_diff(ball: str, features: dict) -> dict:
    diffs = []
    for feature, usual in zip(FEATURES, TYPICAL[ball]):
        value = int(features[feature])
        if value != int(usual):
            diffs.append(
                {
                    "feature": feature,
                    "usual": int(usual),
                    "value": value,
                    "phrase": f"{SHORT[feature]} {value} вместо {usual}",
                }
            )
    spec = VARIANT[ball]
    variant_name = ""
    if len(diffs) == 1 and diffs[0]["feature"] == spec["feature"] and diffs[0]["value"] == spec["value"]:
        variant_name = spec["name"]
    return {"diffs": diffs, "variant_name": variant_name}


def path_to(tree, target: int) -> list[tuple[int, bool, float]]:
    found: list[list[tuple[int, bool, float]]] = []

    def dfs(node: int, acc: list[tuple[int, bool, float]]) -> bool:
        if node == target:
            found.append(list(acc))
            return True
        if tree.feature[node] == -2:
            return False
        feat = int(tree.feature[node])
        thr = float(tree.threshold[node])
        if dfs(int(tree.children_left[node]), acc + [(feat, True, thr)]):
            return True
        return dfs(int(tree.children_right[node]), acc + [(feat, False, thr)])

    if not dfs(0, []):
        raise RuntimeError("узел дерева не найден")
    return found[0]


def apply_path(X: np.ndarray, path: list[tuple[int, bool, float]]) -> np.ndarray:
    mask = np.ones(len(X), dtype=bool)
    for feat, is_left, thr in path:
        if is_left:
            mask &= X[:, feat] <= thr
        else:
            mask &= X[:, feat] > thr
    return mask


def side_codes(tree, X: np.ndarray, node: int, is_left: bool) -> list[int]:
    mask = apply_path(X, path_to(tree, node))
    feat = int(tree.feature[node])
    thr = float(tree.threshold[node])
    column = X[mask, feat]
    picked = column[column <= thr] if is_left else column[column > thr]
    values = sorted({int(v) for v in picked})
    if not values:
        raise RuntimeError("пустое ребро дерева")
    child = int(tree.children_left[node] if is_left else tree.children_right[node])
    if len(picked) != int(tree.n_node_samples[child]):
        raise RuntimeError("коды на ребре не совпали с листом")
    return values


def codes_text(feature: str, values: list[int]) -> str:
    shown = str(values[0]) if len(values) == 1 else " или ".join(str(v) for v in values)
    return f"{FEATURE_RU[feature]} = {shown}"


def edge_text(values: list[int]) -> str:
    if len(values) == 1:
        return f"= {values[0]}"
    return "= " + " или ".join(str(v) for v in values)


def logical_rules(clf: DecisionTreeClassifier, X: np.ndarray) -> tuple[list[str], list[dict]]:
    tree = clf.tree_
    labels = list(clf.classes_)
    phrases: list[str] = []
    splits: list[dict] = []

    def recurse(node: int, constraints: dict[str, list[int]]) -> None:
        if tree.feature[node] == -2:
            counts = tree.value[node][0]
            label = labels[int(np.argmax(counts))]
            ordered = [feature for feature in FEATURES if feature in constraints]
            cond = " и ".join(codes_text(feature, constraints[feature]) for feature in ordered) or "без условий"
            n_leaf = int(tree.n_node_samples[node])
            pure = int((counts > 0).sum()) == 1
            tail = (
                f"В листе {n_leaf} обучающих наблюдений одного класса."
                if pure
                else f"В листе {n_leaf} обучающих наблюдений, классы перемешаны."
            )
            phrases.append(f"Если {cond}, то класс «{label}». {tail}")
            return
        feature = FEATURES[int(tree.feature[node])]
        left_values = side_codes(tree, X, node, True)
        right_values = side_codes(tree, X, node, False)
        splits.append({"feature": feature, "left": left_values, "right": right_values})

        def narrow(values: list[int]) -> dict[str, list[int]]:
            nxt = {key: list(val) for key, val in constraints.items()}
            if feature in nxt:
                nxt[feature] = [value for value in nxt[feature] if value in values]
            else:
                nxt[feature] = list(values)
            return nxt

        recurse(int(tree.children_left[node]), narrow(left_values))
        recurse(int(tree.children_right[node]), narrow(right_values))

    recurse(0, {})
    return phrases, splits


def draw_scheme(path: Path) -> None:
    plt.rcParams["font.family"] = "Times New Roman"
    labels = [
        "Чтение таблицы розыгрышей",
        "Цвет класса для диаграммы",
        "Разбиение 70 % / 30 % по классу",
        "Дерево решений: Джини, глубина не больше 3",
        "Прогноз класса мяча на тесте",
        "Матрица ошибок и метрики качества",
        "Правила листьев и диаграмма рассеяния",
    ]
    n = len(labels)
    box_h = 0.72
    gap = 0.42
    top = 0.15
    fig_h = top + n * box_h + (n - 1) * gap + 0.12
    fig, ax = plt.subplots(figsize=(8.2, fig_h))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, fig_h)
    ax.axis("off")
    y = fig_h - top - box_h
    for i, text in enumerate(labels):
        ax.add_patch(plt.Rectangle((0.4, y), 9.2, box_h, fill=False, edgecolor="black", linewidth=1.0))
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


def draw_tree(clf: DecisionTreeClassifier, X: np.ndarray, path: Path) -> None:
    plt.rcParams["font.family"] = "Times New Roman"
    tree = clf.tree_
    labels = list(clf.classes_)

    def leaves_under(node: int) -> list[int]:
        if tree.feature[node] == -2:
            return [node]
        return leaves_under(int(tree.children_left[node])) + leaves_under(int(tree.children_right[node]))

    order = leaves_under(0)
    x_of_leaf = {node: i for i, node in enumerate(order)}
    depth_of: dict[int, int] = {}

    def set_depth(node: int, depth: int) -> None:
        depth_of[node] = depth
        if tree.feature[node] == -2:
            return
        set_depth(int(tree.children_left[node]), depth + 1)
        set_depth(int(tree.children_right[node]), depth + 1)

    set_depth(0, 0)
    max_depth = max(depth_of.values())
    n_leaves = len(order)

    def x_center(node: int) -> float:
        desc = leaves_under(node)
        return sum(x_of_leaf[item] for item in desc) / len(desc)

    def px(node: int) -> float:
        return 14 + (x_center(node) + 0.5) / n_leaves * 72

    def py(node: int) -> float:
        if max_depth == 0:
            return 50
        return 86 - depth_of[node] * (68 / max_depth)

    fig_w = max(10.2, 2.5 * n_leaves + 1.4)
    fig_h = 2.4 + max_depth * 2.15
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")
    fig.patch.set_facecolor("white")

    def box_width(node: int) -> float:
        same = [other for other, depth in depth_of.items() if depth == depth_of[node] and other != node]
        if not same:
            return 34
        dist = min(abs(px(node) - px(other)) for other in same)
        return min(34, dist * 0.82)

    def node_box(node: int, text: str, *, fc: str) -> None:
        w = box_width(node)
        h = 13.5 if tree.feature[node] != -2 else 12.5
        x, y = px(node), py(node)
        ax.add_patch(
            FancyBboxPatch(
                (x - w / 2, y - h / 2),
                w,
                h,
                boxstyle="round,pad=0.02,rounding_size=0.8",
                facecolor=fc,
                edgecolor="black",
                linewidth=1.15,
                zorder=2,
            )
        )
        ax.text(x, y, text, ha="center", va="center", fontsize=12, color="black", linespacing=1.25, zorder=3)

    def link(parent: int, child: int, label: str) -> None:
        x1, y1 = px(parent), py(parent)
        x2, y2 = px(child), py(child)
        h1 = 13.5 / 2
        h2 = (13.5 if tree.feature[child] != -2 else 12.5) / 2
        y1 -= h1
        y2 += h2
        ax.plot([x1, x2], [y1, y2], color="black", lw=1.05, zorder=1)
        mx, my = (x1 + x2) / 2, (y1 + y2) / 2
        dx = -4.6 if x2 <= x1 else 4.6
        ax.text(mx + dx, my + 1.2, label, ha="center", va="center", fontsize=12, color="#222222", zorder=3)

    def walk(node: int) -> None:
        if tree.feature[node] == -2:
            counts = tree.value[node][0]
            label = labels[int(np.argmax(counts))]
            title = f"Класс:\n{label}" if int((counts > 0).sum()) == 1 else f"Смешанный лист\n{label}"
            node_box(node, title, fc=LEAF_FC.get(label, "#ffffff"))
            return
        feature = FEATURES[int(tree.feature[node])]
        left = int(tree.children_left[node])
        right = int(tree.children_right[node])
        node_box(node, FEATURE_RU[feature], fc="#fff2cc")
        link(node, left, edge_text(side_codes(tree, X, node, True)))
        link(node, right, edge_text(side_codes(tree, X, node, False)))
        walk(left)
        walk(right)

    walk(0)
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=180, bbox_inches="tight", facecolor="white", pad_inches=0.18)
    plt.close()


def draw_scatter(df: pd.DataFrame, error_ids: list[int], path: Path) -> None:
    plt.rcParams["font.family"] = "Times New Roman"
    fig, ax = plt.subplots(figsize=(7.6, 5.8))
    for name in CLASSES:
        part = df[df["ball"] == name]
        grouped = part.groupby(["play_zone", "net_share"], sort=True)
        first = True
        for (zone, share), group in grouped:
            ax.scatter(
                [int(zone)],
                [int(share)],
                c=COLORS[name],
                s=90,
                label=name if first else None,
                zorder=3,
                edgecolors="white",
                linewidths=0.6,
            )
            ax.annotate(
                str(len(group)),
                (int(zone), int(share)),
                textcoords="offset points",
                xytext=(8, 8),
                fontsize=11,
                color="#222222",
            )
            first = False
    err = df[df["id"].isin(error_ids)]
    if len(err):
        ax.scatter(
            err["play_zone"],
            err["net_share"],
            facecolors="none",
            edgecolors="black",
            s=220,
            linewidths=1.3,
            label="ошибка на тесте",
            zorder=4,
        )
    ax.set_xticks([0, 1, 2])
    ax.set_yticks([0, 1, 2])
    ax.set_xticklabels(["0", "1", "2"])
    ax.set_yticklabels(["0", "1", "2"])
    ax.set_xlim(-0.45, 2.45)
    ax.set_ylim(-0.45, 2.45)
    ax.set_xlabel("Игровая зона", fontsize=12)
    ax.set_ylabel("Динамика владения", fontsize=12)
    ax.tick_params(labelsize=11)
    ax.grid(True, color="#d0d0d0", linewidth=0.6)
    ax.set_axisbelow(True)
    ax.legend(frameon=True, fontsize=11, loc="upper right")
    fig.tight_layout()
    fig.savefig(path, dpi=160, bbox_inches="tight", facecolor="white")
    plt.close()


def main() -> None:
    tickets = DATA / "tickets.csv"
    requests = LR2_DATA / "requests.csv"
    if not tickets.exists():
        raise RuntimeError("tickets.csv должен остаться")
    if not requests.exists():
        raise RuntimeError("requests.csv должен остаться")
    DATA.mkdir(parents=True, exist_ok=True)
    FIG.mkdir(parents=True, exist_ok=True)
    rng = np.random.default_rng(SEED)
    df = generate(rng)
    variants = audit(df)
    df.to_csv(DATA / "balls.csv", index=False, encoding="utf-8")

    X = df[FEATURES].to_numpy(dtype=float)
    y = df["ball"].to_numpy()
    ids = df["id"].to_numpy()
    X_train, X_test, y_train, y_test, id_train, id_test = train_test_split(
        X, y, ids, test_size=0.30, random_state=SEED, stratify=y
    )
    train_ids = [int(i) for i in id_train]
    test_ids = [int(i) for i in id_test]
    if sorted(pd.Series(y_train).value_counts().tolist()) != [21, 21, 21]:
        raise RuntimeError("в обучении не по 21")
    if sorted(pd.Series(y_test).value_counts().tolist()) != [9, 9, 9]:
        raise RuntimeError("в тесте не по 9")

    clf = DecisionTreeClassifier(
        criterion="gini",
        max_depth=MAX_DEPTH,
        min_samples_leaf=MIN_LEAF,
        random_state=SEED,
    )
    clf.fit(X_train, y_train)
    y_pred = clf.predict(X_test)
    correct = int((y_pred == y_test).sum())
    rules, splits = logical_rules(clf, X_train)

    cm = confusion_matrix(y_test, y_pred, labels=CLASSES)
    precision, recall, f1, support = precision_recall_fscore_support(
        y_test, y_pred, labels=CLASSES, zero_division=0
    )
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

    order = np.argsort(id_test)
    predictions = []
    for i in order:
        row = {feature: int(X_test[i, j]) for j, feature in enumerate(FEATURES)}
        described = row_diff(str(y_test[i]), row)
        predictions.append(
            {
                "id": int(id_test[i]),
                "ball": str(y_test[i]),
                "pred": str(y_pred[i]),
                "correct": bool(y_test[i] == y_pred[i]),
                "features": row,
                "diffs": described["diffs"],
                "variant_name": described["variant_name"],
            }
        )
    error_ids = [item["id"] for item in predictions if not item["correct"]]
    metrics = {
        "n": 90,
        "n_train": 63,
        "n_test": 27,
        "correct": correct,
        "accuracy": float(correct / 27),
        "confusion": cm.tolist(),
        "per_class": per_class,
        "macro_precision": float(np.mean(precision)),
        "macro_recall": float(np.mean(recall)),
        "macro_f1": float(np.mean(f1)),
        "train_ids": sorted(train_ids),
        "test_ids": sorted(test_ids),
        "error_ids": error_ids,
        "predictions": predictions,
        "rules": rules,
        "splits": splits,
        "max_depth_limit": MAX_DEPTH,
        "actual_depth": int(clf.get_depth()),
        "min_samples_leaf": MIN_LEAF,
        "criterion": "gini",
        "n_leaves": int(clf.get_n_leaves()),
        "typical": TYPICAL,
        "variants": variants,
        "codebook": CODEBOOK,
        "seed": SEED,
    }
    (DATA / "metrics.json").write_text(json.dumps(metrics, ensure_ascii=False, indent=2), encoding="utf-8")
    codes = {"typical": TYPICAL, "variant": VARIANT, "codebook": CODEBOOK, "features": FEATURES}
    (DATA / "codes.json").write_text(json.dumps(codes, ensure_ascii=False, indent=2), encoding="utf-8")
    split = {"seed": SEED, "train_ids": metrics["train_ids"], "test_ids": metrics["test_ids"]}
    (DATA / "split.json").write_text(json.dumps(split, ensure_ascii=False, indent=2), encoding="utf-8")

    draw_scheme(FIG / "scheme.png")
    draw_tree(clf, X_train, FIG / "tree.png")
    draw_scatter(df, error_ids, FIG / "scatter.png")

    LR2_DATA.mkdir(parents=True, exist_ok=True)
    shutil.copy2(DATA / "balls.csv", LR2_DATA / "balls.csv")
    shutil.copy2(DATA / "split.json", LR2_DATA / "split.json")
    shutil.copy2(DATA / "codes.json", LR2_DATA / "codes.json")
    if not tickets.exists() or not requests.exists():
        raise RuntimeError("tickets.csv или requests.csv пропали")

    print("correct", correct, "depth", metrics["actual_depth"], "leaves", metrics["n_leaves"])
    print("errors", error_ids)
    print("splits", splits)
    for rule in rules:
        print(rule)
    print("variants", [(item["id"], item["name"]) for item in variants])


if __name__ == "__main__":
    main()
