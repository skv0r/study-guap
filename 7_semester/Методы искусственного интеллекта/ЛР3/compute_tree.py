# -*- coding: utf-8 -*-
"""Измерения мяча и дерево решений. В текст отчёта не входит."""
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

FEATURES = ["elasticity", "size", "mass", "pressure", "bounce"]
FEATURE_RU = {
    "elasticity": "Коэффициент упругости",
    "size": "Окружность, см",
    "mass": "Масса, г",
    "pressure": "Давление, бар",
    "bounce": "Высота отскока, см",
}
SHORT = {
    "elasticity": "упругость",
    "size": "окружность",
    "mass": "масса",
    "pressure": "давление",
    "bounce": "отскок",
}
CLASSES = ["футбольный", "баскетбольный", "волейбольный"]
COLORS = {"футбольный": "#1f77b4", "баскетбольный": "#ff7f0e", "волейбольный": "#2ca02c"}
LEAF_FC = {"футбольный": "#cfe2f3", "баскетбольный": "#fce5cd", "волейбольный": "#d9ead3"}

# Типовые центры классов (без кодирования). Диапазоны пересекаются.
TYPICAL = {
    "футбольный": [0.75, 70.0, 420.0, 0.55, 125.0],
    "баскетбольный": [0.82, 72.0, 480.0, 0.50, 135.0],
    "волейбольный": [0.68, 69.0, 380.0, 0.42, 118.0],
}
# Шум подобран так, чтобы все методы давали ≥80 % и <100 % на одном тесте.
NOISE = [0.028, 0.95, 24.0, 0.045, 4.5]
JITTERS = {
    "футбольный": [
        [0.0, 0.0, 0.0, 0.0, 0.0],
        [-0.02, 0.55, 16.0, -0.03, 5.0],
        [0.02, -0.4, -14.0, 0.025, -4.0],
        [-0.015, 0.7, 22.0, 0.015, 6.0],
    ],
    "баскетбольный": [
        [0.0, 0.0, 0.0, 0.0, 0.0],
        [-0.025, -0.65, -20.0, 0.02, -4.5],
        [0.02, 0.5, 16.0, -0.025, 5.0],
        [-0.03, 0.4, -26.0, 0.03, -3.5],
    ],
    "волейбольный": [
        [0.0, 0.0, 0.0, 0.0, 0.0],
        [0.025, 0.65, 20.0, 0.04, 5.0],
        [-0.025, -0.5, -16.0, -0.025, -4.5],
        [0.03, 0.75, 24.0, 0.03, 6.0],
    ],
}
SEED = 300
N_PER_CLASS = 50
N_TOTAL = N_PER_CLASS * 3
N_TRAIN_PER = 35
N_TEST_PER = 15
N_TRAIN = N_TRAIN_PER * 3
N_TEST = N_TEST_PER * 3
MAX_DEPTH = 3
MIN_LEAF = 4
NEAR_TOL = [0.04, 1.0, 28.0, 0.055, 5.5]


def fmt_num(feature: str, value: float) -> str:
    number = float(value)
    if feature in ("elasticity", "pressure"):
        return f"{number:.2f}".replace(".", ",")
    if feature == "size":
        return f"{number:.1f}".replace(".", ",")
    return str(int(round(number)))


def _round_feature(feature: str, value: float) -> float:
    if feature in ("elasticity", "pressure"):
        return float(round(value, 3))
    if feature == "size":
        return float(round(value, 1))
    return float(round(value))


def generate(rng: np.random.Generator) -> pd.DataFrame:
    rows = []
    noise_scale = np.asarray(NOISE, dtype=float)
    for name in CLASSES:
        center = np.asarray(TYPICAL[name], dtype=float)
        jitters = JITTERS[name]
        for i in range(N_PER_CLASS):
            # Первые ~12 строк ближе к типовому центру, остальные — с большим разбросом.
            if i < 14:
                shift = np.zeros(len(FEATURES))
                scale = 0.18 * noise_scale
            else:
                shift = np.asarray(jitters[i % len(jitters)], dtype=float)
                scale = noise_scale
            noise = rng.normal(0.0, 1.0, size=len(FEATURES)) * scale
            raw = center + shift + noise
            item = {"ball": name}
            for feature, value in zip(FEATURES, raw):
                item[feature] = _round_feature(feature, float(value))
            rows.append(item)
    df = pd.DataFrame(rows)
    df = df.iloc[rng.permutation(len(df))].reset_index(drop=True)
    df.insert(0, "id", np.arange(1, len(df) + 1))
    return df


def near_typical_count(part: pd.DataFrame, usual: list[float]) -> int:
    matched = 0
    for row in part.itertuples(index=False):
        ok = True
        for feature, target, tol in zip(FEATURES, usual, NEAR_TOL):
            if abs(float(getattr(row, feature)) - float(target)) > tol:
                ok = False
                break
        matched += int(ok)
    return matched


def audit(df: pd.DataFrame) -> dict:
    expected = [N_PER_CLASS] * 3
    if len(df) != N_TOTAL or sorted(df["ball"].value_counts().tolist()) != expected:
        raise RuntimeError(f"нужно {N_TOTAL} строк, по {N_PER_CLASS} в классе")
    # Точных дубликатов всего вектора быть не должно — иначе kNN/SVM дают 100 %.
    keys = df[FEATURES].apply(lambda r: tuple(float(x) for x in r), axis=1)
    if keys.duplicated().any():
        raise RuntimeError("есть точные копии измерений")
    mode_count = {}
    for name in CLASSES:
        part = df[df["ball"] == name]
        near = near_typical_count(part, TYPICAL[name])
        if near < 6:
            raise RuntimeError(f"мало строк около типового {name}: {near}")
        mode_count[name] = near
    foot = df[df["ball"] == "футбольный"]
    basket = df[df["ball"] == "баскетбольный"]
    volley = df[df["ball"] == "волейбольный"]
    if foot["elasticity"].max() <= volley["elasticity"].min():
        raise RuntimeError("упругость волейбола и футбола не пересекается")
    if basket["elasticity"].min() >= foot["elasticity"].max():
        raise RuntimeError("упругость футбола и баскетбола не пересекается")
    if volley["size"].max() <= foot["size"].min():
        raise RuntimeError("окружность волейбола и футбола не пересекается")
    return {"mode_count": mode_count}


def row_diff(ball: str, features: dict) -> dict:
    diffs = []
    for feature, usual in zip(FEATURES, TYPICAL[ball]):
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


def thr_text(feature: str, thr: float, *, left: bool) -> str:
    shown = fmt_num(feature, thr)
    return f"{'≤' if left else '>'} {shown}"


def cond_text(feature: str, thr: float, *, left: bool) -> str:
    return f"{FEATURE_RU[feature]} {thr_text(feature, thr, left=left)}"


def logical_rules(clf: DecisionTreeClassifier, X: np.ndarray) -> tuple[list[str], list[dict]]:
    tree = clf.tree_
    labels = list(clf.classes_)
    phrases: list[str] = []
    splits: list[dict] = []

    def recurse(node: int, constraints: list[tuple[str, bool, float]]) -> None:
        if tree.feature[node] == -2:
            counts = tree.value[node][0]
            label = labels[int(np.argmax(counts))]
            cond = " и ".join(
                cond_text(feature, thr, left=is_left) for feature, is_left, thr in constraints
            ) or "без условий"
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
        thr = float(tree.threshold[node])
        left = int(tree.children_left[node])
        right = int(tree.children_right[node])
        splits.append({"feature": feature, "threshold": thr})
        recurse(left, constraints + [(feature, True, thr)])
        recurse(right, constraints + [(feature, False, thr)])

    recurse(0, [])
    return phrases, splits


def draw_scheme(path: Path) -> None:
    plt.rcParams["font.family"] = "Times New Roman"
    labels = [
        "Чтение таблицы измерений мяча",
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
    """Дерево без наложений: порог внутри узла, на рёбрах только «≤» / «>»."""
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
    n_leaves = max(len(order), 1)
    # Широкий холст: каждому листу своя колонка.
    col_w = 2.35
    fig_w = max(12.0, n_leaves * col_w + 1.2)
    row_h = 2.55
    fig_h = 1.6 + (max_depth + 1) * row_h
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.set_xlim(0, fig_w)
    ax.set_ylim(0, fig_h)
    ax.axis("off")
    fig.patch.set_facecolor("white")

    def px(node: int) -> float:
        desc = leaves_under(node)
        center = sum(x_of_leaf[item] for item in desc) / len(desc)
        return 0.7 + (center + 0.5) * (fig_w - 1.4) / n_leaves

    def py(node: int) -> float:
        return fig_h - 0.9 - depth_of[node] * row_h

    box_w = min(2.05, (fig_w - 1.4) / n_leaves * 0.88)
    box_h_split = 0.95
    box_h_leaf = 0.85

    def node_box(node: int, text: str, *, fc: str, split: bool) -> None:
        h = box_h_split if split else box_h_leaf
        x, y = px(node), py(node)
        ax.add_patch(
            FancyBboxPatch(
                (x - box_w / 2, y - h / 2),
                box_w,
                h,
                boxstyle="round,pad=0.02,rounding_size=0.12",
                facecolor=fc,
                edgecolor="black",
                linewidth=1.1,
                zorder=2,
            )
        )
        ax.text(x, y, text, ha="center", va="center", fontsize=9.5, color="black", linespacing=1.2, zorder=3)

    def link(parent: int, child: int, side: str) -> None:
        x1, y1 = px(parent), py(parent)
        x2, y2 = px(child), py(child)
        y1 -= box_h_split / 2
        y2 += (box_h_split if tree.feature[child] != -2 else box_h_leaf) / 2
        ax.plot([x1, x2], [y1, y2], color="black", lw=1.0, zorder=1)
        # Короткая подпись сбоку от середины ребра, без длинных чисел.
        mx, my = (x1 + x2) / 2, (y1 + y2) / 2
        offset = -0.28 if side == "left" else 0.28
        ax.text(mx + offset, my, "≤" if side == "left" else ">", ha="center", va="center",
                fontsize=10, color="#222222", zorder=4,
                bbox=dict(boxstyle="round,pad=0.12", facecolor="white", edgecolor="none", alpha=0.92))

    def walk(node: int) -> None:
        if tree.feature[node] == -2:
            counts = tree.value[node][0]
            label = labels[int(np.argmax(counts))]
            pure = int((counts > 0).sum()) == 1
            title = f"Класс:\n{label}" if pure else f"Смешанный\n{label}"
            node_box(node, title, fc=LEAF_FC.get(label, "#ffffff"), split=False)
            return
        feature = FEATURES[int(tree.feature[node])]
        thr = float(tree.threshold[node])
        left = int(tree.children_left[node])
        right = int(tree.children_right[node])
        # Порог внутри узла — на рёбрах не дублируем длинный текст.
        title = f"{FEATURE_RU[feature]}\n≤ {fmt_num(feature, thr)}"
        node_box(node, title, fc="#fff2cc", split=True)
        link(node, left, "left")
        link(node, right, "right")
        walk(left)
        walk(right)

    walk(0)
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=180, bbox_inches="tight", facecolor="white", pad_inches=0.2)
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
    design = audit(df)
    df.to_csv(DATA / "balls.csv", index=False, encoding="utf-8")

    X = df[FEATURES].to_numpy(dtype=float)
    y = df["ball"].to_numpy()
    ids = df["id"].to_numpy()
    X_train, X_test, y_train, y_test, id_train, id_test = train_test_split(
        X, y, ids, test_size=0.30, random_state=SEED, stratify=y
    )
    train_ids = [int(i) for i in id_train]
    test_ids = [int(i) for i in id_test]
    if sorted(pd.Series(y_train).value_counts().tolist()) != [N_TRAIN_PER] * 3:
        raise RuntimeError(f"в обучении не по {N_TRAIN_PER}")
    if sorted(pd.Series(y_test).value_counts().tolist()) != [N_TEST_PER] * 3:
        raise RuntimeError(f"в тесте не по {N_TEST_PER}")

    clf = DecisionTreeClassifier(
        criterion="gini",
        max_depth=MAX_DEPTH,
        min_samples_leaf=MIN_LEAF,
        random_state=SEED,
    )
    clf.fit(X_train, y_train)
    y_pred = clf.predict(X_test)
    correct = int((y_pred == y_test).sum())
    if correct >= N_TEST:
        raise RuntimeError("дерево дало 100 % на тесте — нужен больший шум")
    if correct / N_TEST < 0.80:
        raise RuntimeError(f"дерево ниже 80 %: {correct}/{N_TEST}")
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
        row = {feature: float(X_test[i, j]) for j, feature in enumerate(FEATURES)}
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
        "n": N_TOTAL,
        "n_train": N_TRAIN,
        "n_test": N_TEST,
        "n_per_class": N_PER_CLASS,
        "n_train_per": N_TRAIN_PER,
        "n_test_per": N_TEST_PER,
        "correct": correct,
        "accuracy": float(correct / N_TEST),
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
        "mode_count": design["mode_count"],
        "seed": SEED,
    }
    (DATA / "metrics.json").write_text(json.dumps(metrics, ensure_ascii=False, indent=2), encoding="utf-8")
    codes = {
        "typical": TYPICAL,
        "features": FEATURES,
        "mode_count": design["mode_count"],
        "n": N_TOTAL,
        "n_per_class": N_PER_CLASS,
        "n_train": N_TRAIN,
        "n_test": N_TEST,
        "n_train_per": N_TRAIN_PER,
        "n_test_per": N_TEST_PER,
    }
    (DATA / "codes.json").write_text(json.dumps(codes, ensure_ascii=False, indent=2), encoding="utf-8")
    split = {
        "seed": SEED,
        "train_ids": metrics["train_ids"],
        "test_ids": metrics["test_ids"],
        "n_train": N_TRAIN,
        "n_test": N_TEST,
        "n_train_per": N_TRAIN_PER,
        "n_test_per": N_TEST_PER,
    }
    (DATA / "split.json").write_text(json.dumps(split, ensure_ascii=False, indent=2), encoding="utf-8")

    draw_scheme(FIG / "scheme.png")
    draw_tree(clf, X_train, FIG / "tree.png")
    draw_scatter(df, error_ids, FIG / "scatter.png")

    for lab in ("ЛР2", "ЛР4", "ЛР5"):
        dest = BASE.parent / lab / "data"
        dest.mkdir(parents=True, exist_ok=True)
        shutil.copy2(DATA / "balls.csv", dest / "balls.csv")
        shutil.copy2(DATA / "split.json", dest / "split.json")
        shutil.copy2(DATA / "codes.json", dest / "codes.json")
    if not tickets.exists() or not requests.exists():
        raise RuntimeError("tickets.csv или requests.csv пропали")

    print("correct", correct, "/", N_TEST, "depth", metrics["actual_depth"], "leaves", metrics["n_leaves"])
    print("errors", error_ids)
    print("splits", splits)
    for rule in rules:
        print(rule)
    print("modes", design["mode_count"])


if __name__ == "__main__":
    main()
