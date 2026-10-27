# -*- coding: utf-8 -*-
"""Набор кадров и разметка YOLO для ЛР5. В текст отчёта не входит."""

import json
import random
import shutil
from pathlib import Path

import cv2
import numpy as np

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "data" / "src"
COG = ROOT.parent / "ЛР4" / "data" / "cog.jpg"
RES = ROOT / "results"
SET = ROOT / "dataset"
NAMES = ["gear", "washer", "nut"]
W, H = 640, 480
N = 50


def imread(path: Path) -> np.ndarray:
    image = cv2.imdecode(np.fromfile(path, dtype=np.uint8), cv2.IMREAD_COLOR)
    if image is None:
        raise SystemExit(path)
    return image


def imwrite(path: Path, image: np.ndarray) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    ext = ".png" if path.suffix.lower() == ".png" else ".jpg"
    ok, encoded = cv2.imencode(ext, image)
    if not ok:
        raise SystemExit(path)
    encoded.tofile(path)


def largest_mask(binary: np.ndarray) -> np.ndarray:
    contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    mask = np.zeros(binary.shape, np.uint8)
    contour = max(contours, key=cv2.contourArea)
    cv2.drawContours(mask, [contour], -1, 255, thickness=-1)
    return mask


def fit(sprite: np.ndarray, side: int = 220) -> np.ndarray:
    height, width = sprite.shape[:2]
    scale = side / max(height, width)
    if scale >= 1:
        return sprite
    return cv2.resize(sprite, (int(width * scale), int(height * scale)), interpolation=cv2.INTER_AREA)


def crop_alpha(image: np.ndarray, mask: np.ndarray) -> np.ndarray:
    ys, xs = np.where(mask > 0)
    y1, y2, x1, x2 = ys.min(), ys.max() + 1, xs.min(), xs.max() + 1
    bgr = image[y1:y2, x1:x2]
    alpha = mask[y1:y2, x1:x2]
    bgra = cv2.cvtColor(bgr, cv2.COLOR_BGR2BGRA)
    bgra[:, :, 3] = alpha
    return bgra


def sprite_gear() -> np.ndarray:
    image = imread(COG)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    binary = (gray > 125).astype(np.uint8) * 255
    binary = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, np.ones((5, 5), np.uint8))
    contours, hierarchy = cv2.findContours(binary, cv2.RETR_CCOMP, cv2.CHAIN_APPROX_SIMPLE)
    outer = max(contours, key=cv2.contourArea)
    mask = np.zeros(gray.shape, np.uint8)
    cv2.drawContours(mask, [outer], -1, 255, thickness=-1)
    if hierarchy is not None:
        for contour, info in zip(contours, hierarchy[0]):
            if info[3] != -1 and cv2.contourArea(contour) > 400:
                cv2.drawContours(mask, [contour], -1, 0, thickness=-1)
    return fit(crop_alpha(image, mask))


def sprite_washer() -> np.ndarray:
    image = imread(SRC / "Plain_washer1.jpg")
    hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
    metal = ((hsv[:, :, 1] < 70) & (hsv[:, :, 2] > 70)).astype(np.uint8) * 255
    metal = cv2.morphologyEx(metal, cv2.MORPH_OPEN, np.ones((3, 3), np.uint8))
    metal = cv2.morphologyEx(metal, cv2.MORPH_CLOSE, np.ones((3, 3), np.uint8))
    return fit(crop_alpha(image, metal))


def sprite_nut() -> np.ndarray:
    image = imread(SRC / "crops" / "hole4.jpg")
    height, width = image.shape[:2]
    radius = int(0.33 * min(height, width))
    center = (width // 2, height // 2)
    points = []
    for step in range(6):
        angle = np.deg2rad(30 + 60 * step)
        points.append([center[0] + radius * np.cos(angle), center[1] + radius * np.sin(angle)])
    mask = np.zeros((height, width), np.uint8)
    cv2.fillPoly(mask, [np.round(points).astype(np.int32)], 255)
    return fit(crop_alpha(image, mask))


def transform(sprite: np.ndarray, scale: float, angle: float) -> np.ndarray:
    height, width = sprite.shape[:2]
    side = max(height, width)
    canvas = np.zeros((side, side, 4), np.uint8)
    y0, x0 = (side - height) // 2, (side - width) // 2
    canvas[y0 : y0 + height, x0 : x0 + width] = sprite
    matrix = cv2.getRotationMatrix2D((side / 2, side / 2), angle, scale)
    bound = int(np.ceil(side * scale * 1.5))
    matrix[0, 2] += (bound - side) / 2
    matrix[1, 2] += (bound - side) / 2
    return cv2.warpAffine(canvas, matrix, (bound, bound), flags=cv2.INTER_LINEAR, borderValue=(0, 0, 0, 0))


def paste(canvas: np.ndarray, sprite: np.ndarray, left: int, top: int):
    height, width = sprite.shape[:2]
    x1, y1 = max(0, left), max(0, top)
    x2, y2 = min(W, left + width), min(H, top + height)
    if x2 <= x1 or y2 <= y1:
        return None
    cut = sprite[y1 - top : y2 - top, x1 - left : x2 - left]
    alpha = cut[:, :, 3:4].astype(np.float32) / 255
    if alpha.max() < 0.2:
        return None
    view = canvas[y1:y2, x1:x2]
    view[:] = (cut[:, :, :3] * alpha + view * (1 - alpha)).astype(np.uint8)
    ys, xs = np.where(cut[:, :, 3] > 20)
    if len(xs) == 0:
        return None
    return x1 + int(xs.min()), y1 + int(ys.min()), x1 + int(xs.max()), y1 + int(ys.max())


def background(rng: random.Random) -> np.ndarray:
    color = np.array([rng.randint(35, 90), rng.randint(40, 100), rng.randint(45, 110)], np.float32)
    noise = rng.random() * 18
    field = np.zeros((H, W, 3), np.float32)
    field[:] = color
    field += np.random.default_rng(rng.randrange(10**6)).normal(0, noise, field.shape)
    yy = np.linspace(0, 1, H, dtype=np.float32)[:, None]
    field *= (0.82 + 0.28 * yy)[:, :, None]
    return np.clip(field, 0, 255).astype(np.uint8)


def yolo_line(cls: int, box) -> str:
    x1, y1, x2, y2 = box
    xc = ((x1 + x2) / 2) / W
    yc = ((y1 + y2) / 2) / H
    bw = (x2 - x1) / W
    bh = (y2 - y1) / H
    return f"{cls} {xc:.6f} {yc:.6f} {bw:.6f} {bh:.6f}"


def place_object(canvas, sprite, rng, scale):
    turned = transform(sprite, scale, rng.uniform(0, 360))
    left = rng.randint(8, W - turned.shape[1] - 8)
    top = rng.randint(8, H - turned.shape[0] - 8)
    box = paste(canvas, turned, left, top)
    return box


def main():
    if RES.exists():
        shutil.rmtree(RES)
    if SET.exists():
        shutil.rmtree(SET)
    RES.mkdir(parents=True)
    sprites = [sprite_gear(), sprite_washer(), sprite_nut()]
    for name, sprite in zip(NAMES, sprites):
        imwrite(RES / f"sprite_{name}.png", sprite)

    rng = random.Random(7)
    records = []
    flat_img = ROOT / "data" / "images"
    flat_lbl = ROOT / "data" / "labels"
    flat_img.mkdir(parents=True, exist_ok=True)
    flat_lbl.mkdir(parents=True, exist_ok=True)
    for index in range(N):
        scene = index // 10 + 1
        frame = index % 10 + 1
        name = f"video{scene:02d}_{frame:04d}"
        canvas = background(rng)
        gain = rng.uniform(0.75, 1.2)
        canvas = np.clip(canvas.astype(np.float32) * gain, 0, 255).astype(np.uint8)
        boxes = []
        if index % 8 == 0:
            nut_box = place_object(canvas, sprites[2], rng, rng.uniform(0.42, 0.62))
            gear = transform(sprites[0], rng.uniform(0.34, 0.48), rng.uniform(0, 360))
            if nut_box:
                left = nut_box[0] - gear.shape[1] // 3
                top = nut_box[1] - gear.shape[0] // 5
                gear_box = paste(canvas, gear, left, top)
                if gear_box:
                    boxes.append((0, gear_box))
                boxes.append((2, nut_box))
            washer_box = place_object(canvas, sprites[1], rng, rng.uniform(0.45, 0.7))
            if washer_box:
                boxes.append((1, washer_box))
        else:
            for cls in rng.sample([0, 1, 2], rng.choice([1, 2, 2, 3])):
                scale = rng.uniform(0.28, 0.5) if cls == 0 else rng.uniform(0.4, 0.75)
                box = place_object(canvas, sprites[cls], rng, scale)
                if box and (box[2] - box[0]) > 24 and (box[3] - box[1]) > 24:
                    boxes.append((cls, box))
        if not boxes:
            box = place_object(canvas, sprites[1], rng, 0.6)
            boxes.append((1, box))
        imwrite(flat_img / f"{name}.jpg", canvas)
        lines = [yolo_line(cls, box) for cls, box in boxes]
        (flat_lbl / f"{name}.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")
        records.append({"name": name, "scene": f"video{scene:02d}", "labels": lines, "boxes": boxes})

    files = sorted(flat_img.glob("*.jpg"))
    order = files[:]
    random.Random(42).shuffle(order)
    n_train, n_valid = int(0.7 * len(order)), int(0.2 * len(order))
    roles = {}
    for path in order[:n_train]:
        roles[path.stem] = "train"
    for path in order[n_train : n_train + n_valid]:
        roles[path.stem] = "valid"
    for path in order[n_train + n_valid :]:
        roles[path.stem] = "test"
    counts = {part: {name: 0 for name in NAMES} for part in ("train", "valid", "test")}
    images_n = {part: 0 for part in counts}
    for path in files:
        part = roles[path.stem]
        images_n[part] += 1
        target_img = SET / part / "images"
        target_lbl = SET / part / "labels"
        target_img.mkdir(parents=True, exist_ok=True)
        target_lbl.mkdir(parents=True, exist_ok=True)
        shutil.copy(path, target_img / path.name)
        label = flat_lbl / f"{path.stem}.txt"
        shutil.copy(label, target_lbl / label.name)
        for line in label.read_text(encoding="utf-8").splitlines():
            counts[part][NAMES[int(line.split()[0])]] += 1
    yaml = (
        "path: .\ntrain: train/images\nval: valid/images\ntest: test/images\nnames:\n"
        "  0: gear\n  1: washer\n  2: nut\n"
    )
    (SET / "data.yaml").write_text(yaml, encoding="utf-8")
    summary = {
        "images": images_n,
        "boxes": counts,
        "total_images": N,
        "scenes": 5,
        "seed": 42,
    }
    (RES / "split.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
