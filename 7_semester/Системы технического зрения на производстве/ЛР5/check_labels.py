# -*- coding: utf-8 -*-
"""Проверка разметки, вариант 1: мелкие рамки. В текст отчёта не входит."""

import json
import shutil
from pathlib import Path

import cv2
import numpy as np

ROOT = Path(__file__).resolve().parent
IMG = ROOT / "data" / "images"
LBL = ROOT / "data" / "labels"
RES = ROOT / "results"
NAMES = ["gear", "washer", "nut"]
MIN_AREA = 0.001
DIRTY = ROOT / "data" / "labels_dirty"


def imread(path: Path) -> np.ndarray:
    return cv2.imdecode(np.fromfile(path, dtype=np.uint8), cv2.IMREAD_COLOR)


def imwrite(path: Path, image: np.ndarray) -> None:
    ok, encoded = cv2.imencode(".jpg", image)
    encoded.tofile(path)


def check_label_line(line, n_classes):
    parts = line.split()
    if len(parts) != 5:
        return "ошибка: должно быть 5 чисел"
    cls = int(parts[0])
    vals = [float(value) for value in parts[1:]]
    if not 0 <= cls < n_classes:
        return "ошибка: неизвестный класс"
    if any(value < 0 or value > 1 for value in vals):
        return "ошибка: значение вне [0, 1]"
    return "ok"


def audit_small(label_dir: Path):
    problems = []
    for path in sorted(label_dir.glob("*.txt")):
        rows = [line.split() for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]
        for row in rows:
            _, _, _, bw, bh = map(float, row[:5])
            if bw * bh < MIN_AREA:
                problems.append((path.name, "слишком малая рамка", round(bw * bh, 6)))
    return problems


def draw(name: str, label_dir: Path) -> np.ndarray:
    image = imread(IMG / f"{name}.jpg")
    height, width = image.shape[:2]
    for line in (label_dir / f"{name}.txt").read_text(encoding="utf-8").splitlines():
        if check_label_line(line, len(NAMES)) != "ok":
            continue
        cls, xc, yc, bw, bh = line.split()
        xc, bw = float(xc) * width, float(bw) * width
        yc, bh = float(yc) * height, float(bh) * height
        x1, y1 = int(xc - bw / 2), int(yc - bh / 2)
        x2, y2 = int(xc + bw / 2), int(yc + bh / 2)
        cv2.rectangle(image, (x1, y1), (x2, y2), (40, 170, 40), 2)
        cv2.putText(image, NAMES[int(cls)], (x1, max(16, y1 - 6)), cv2.FONT_HERSHEY_SIMPLEX, 0.55, (40, 170, 40), 2, cv2.LINE_AA)
    return image


def main():
    if DIRTY.exists():
        shutil.rmtree(DIRTY)
    shutil.copytree(LBL, DIRTY)
    target = DIRTY / "video03_0001.txt"
    target.write_text(target.read_text(encoding="utf-8") + "2 0.500000 0.500000 0.012000 0.012000\n", encoding="utf-8")
    dirty = audit_small(DIRTY)
    clean = audit_small(LBL)
    shown = ["video01_0001", "video02_0005", "video04_0002"]
    panels = [draw(name, LBL) for name in shown]
    imwrite(RES / "marked.jpg", np.hstack(panels))
    bad = draw("video03_0001", DIRTY)
    imwrite(RES / "small_box.jpg", bad)
    counts = {name: 0 for name in NAMES}
    for path in LBL.glob("*.txt"):
        for line in path.read_text(encoding="utf-8").splitlines():
            counts[NAMES[int(line.split()[0])]] += 1
    report = {
        "dirty": [{"file": item[0], "text": item[1], "area": item[2]} for item in dirty],
        "clean": clean,
        "counts": counts,
        "shown": shown,
        "min_area": MIN_AREA,
    }
    (RES / "audit.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
