# -*- coding: utf-8 -*-
"""Расчёт признаков ЛР4. В текст отчёта не входит."""

import json
from collections import deque
from pathlib import Path

import cv2
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data" / "lag.png"
RES = ROOT / "results"
WORK_SIDE = 960


def imread(path: Path) -> np.ndarray:
    data = np.fromfile(path, dtype=np.uint8)
    image = cv2.imdecode(data, cv2.IMREAD_COLOR)
    if image is None:
        raise SystemExit(f"Не прочитан кадр: {path}")
    return image


def imwrite(path: Path, image: np.ndarray) -> None:
    suffix = path.suffix.lower()
    ext = ".jpg" if suffix in {".jpg", ".jpeg"} else suffix
    ok, encoded = cv2.imencode(ext, image)
    if not ok:
        raise SystemExit(f"Не записан кадр: {path}")
    encoded.tofile(path)


def to_gray(bgr: np.ndarray) -> np.ndarray:
    rgb = cv2.cvtColor(bgr, cv2.COLOR_BGR2RGB).astype(np.float32)
    y = 0.2126 * rgb[:, :, 0] + 0.7152 * rgb[:, :, 1] + 0.0722 * rgb[:, :, 2]
    return np.clip(np.round(y), 0, 255).astype(np.uint8)


def close_color(current, neighbor, lo, up) -> bool:
    diff = neighbor.astype(np.int16) - current.astype(np.int16)
    return bool(np.all(diff >= -lo) and np.all(diff <= up))


def check_bottom_neighbor(field, img, row, col, n_rows, lo, up):
    if row + 1 >= n_rows:
        return None
    if field[row + 1, col] == 1:
        return None
    if close_color(img[row, col], img[row + 1, col], lo, up):
        field[row + 1, col] = 1
        return (row + 1, col)
    return None


def check_top_neighbor(field, img, row, col, lo, up):
    if row - 1 < 0:
        return None
    if field[row - 1, col] == 1:
        return None
    if close_color(img[row, col], img[row - 1, col], lo, up):
        field[row - 1, col] = 1
        return (row - 1, col)
    return None


def check_left_neighbor(field, img, row, col, lo, up):
    if col - 1 < 0:
        return None
    if field[row, col - 1] == 1:
        return None
    if close_color(img[row, col], img[row, col - 1], lo, up):
        field[row, col - 1] = 1
        return (row, col - 1)
    return None


def check_right_neighbor(field, img, row, col, n_cols, lo, up):
    if col + 1 >= n_cols:
        return None
    if field[row, col + 1] == 1:
        return None
    if close_color(img[row, col], img[row, col + 1], lo, up):
        field[row, col + 1] = 1
        return (row, col + 1)
    return None


def flood_own(img, seed_row, seed_col, lo, up):
    n_rows, n_cols = img.shape[:2]
    field = np.zeros((n_rows, n_cols), np.uint8)
    field[seed_row, seed_col] = 1
    queue = deque([(seed_row, seed_col)])
    lo_arr = np.array(lo, np.int16)
    up_arr = np.array(up, np.int16)
    while queue:
        row, col = queue.popleft()
        for nxt in (
            check_bottom_neighbor(field, img, row, col, n_rows, lo_arr, up_arr),
            check_top_neighbor(field, img, row, col, lo_arr, up_arr),
            check_left_neighbor(field, img, row, col, lo_arr, up_arr),
            check_right_neighbor(field, img, row, col, n_cols, lo_arr, up_arr),
        ):
            if nxt is not None:
                queue.append(nxt)
    return field


def largest_contour(binary):
    contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return None
    return max(contours, key=cv2.contourArea)


def choose_threshold(gray):
    height, width = gray.shape
    best = None
    for level in range(60, 210, 5):
        _, binary = cv2.threshold(gray, level, 255, cv2.THRESH_BINARY)
        contour = largest_contour(binary)
        if contour is None:
            continue
        area = float(cv2.contourArea(contour))
        x, y, w, h = cv2.boundingRect(contour)
        touches = x <= 1 or y <= 1 or x + w >= width - 1 or y + h >= height - 1
        fraction = area / float(height * width)
        if touches or fraction < 0.05 or fraction > 0.85:
            continue
        mask = np.zeros_like(gray)
        cv2.drawContours(mask, [contour], -1, 255, thickness=-1)
        inside = gray[mask == 255]
        outside = gray[mask == 0]
        if inside.size == 0 or outside.size == 0:
            continue
        contrast = float(inside.mean() - outside.mean())
        if contrast < 15:
            continue
        score = contrast * min(fraction, 0.55)
        if best is None or score > best[0]:
            best = (score, level, area, contrast)
    if best is None:
        raise SystemExit("Порог контура не найден")
    return best[1]


def draw_hog(gray_roi, path):
    roi = cv2.resize(gray_roi, (256, 256), interpolation=cv2.INTER_AREA)
    gx = cv2.Sobel(roi, cv2.CV_32F, 1, 0, ksize=1)
    gy = cv2.Sobel(roi, cv2.CV_32F, 0, 1, ksize=1)
    mag, ang = cv2.cartToPolar(gx, gy, angleInDegrees=True)
    cell = 16
    bins = 9
    fig, axes = plt.subplots(1, 2, figsize=(8.2, 4.1))
    axes[0].imshow(roi, cmap="gray")
    axes[0].set_title("Область объекта")
    axes[0].axis("off")
    axes[1].imshow(roi, cmap="gray")
    dominant = []
    for y in range(0, 256, cell):
        for x in range(0, 256, cell):
            hist = np.zeros(bins, np.float32)
            patch_mag = mag[y : y + cell, x : x + cell]
            patch_ang = ang[y : y + cell, x : x + cell] % 180
            idx = np.clip((patch_ang / 180 * bins).astype(np.int32), 0, bins - 1)
            for bin_id in range(bins):
                hist[bin_id] = patch_mag[idx == bin_id].sum()
            peak = int(hist.argmax())
            direction = (peak + 0.5) * (np.pi / bins)
            length = 6.5
            cx, cy = x + cell / 2, y + cell / 2
            dx, dy = length * np.cos(direction), length * np.sin(direction)
            axes[1].plot([cx - dx, cx + dx], [cy - dy, cy + dy], color="#1F4E79", linewidth=1.1)
            dominant.append(direction)
    axes[1].set_title("HOG, ячейка 16×16")
    axes[1].axis("off")
    fig.tight_layout()
    fig.savefig(path, dpi=140)
    plt.close(fig)
    return float(np.median(dominant) * 180 / np.pi)


def main():
    RES.mkdir(parents=True, exist_ok=True)
    source = imread(DATA)
    src_h, src_w = source.shape[:2]
    scale = WORK_SIDE / max(src_h, src_w)
    work = cv2.resize(source, (int(round(src_w * scale)), int(round(src_h * scale))), interpolation=cv2.INTER_AREA)
    height, width = work.shape[:2]
    gray = to_gray(work)
    imwrite(RES / "original.jpg", work)

    low_pair = (40, 110)
    high_pair = (90, 180)
    edges_low = cv2.Canny(gray, *low_pair)
    edges_high = cv2.Canny(gray, *high_pair)
    canny = np.hstack([edges_low, edges_high])
    imwrite(RES / "canny.png", canny)

    level = choose_threshold(gray)
    _, binary = cv2.threshold(gray, level, 255, cv2.THRESH_BINARY)
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (3, 3))
    binary = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel)
    contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    min_area = 300
    kept = [c for c in contours if cv2.contourArea(c) >= min_area]
    kept.sort(key=cv2.contourArea, reverse=True)
    main = kept[0]
    area = float(cv2.contourArea(main))
    perimeter = float(cv2.arcLength(main, True))
    x, y, w, h = cv2.boundingRect(main)
    hull = cv2.convexHull(main)
    solidity = area / float(cv2.contourArea(hull))

    drawn = work.copy()
    cv2.drawContours(drawn, [main], -1, (0, 180, 0), 2)
    cv2.rectangle(drawn, (x, y), (x + w, y + h), (0, 0, 220), 2)
    imwrite(RES / "contours.png", drawn)

    moments = cv2.moments(main)
    mask = np.zeros((height, width), np.uint8)
    cv2.drawContours(mask, [main], -1, 255, thickness=-1)
    inside = cv2.distanceTransform(mask, cv2.DIST_L2, 5)
    target = float(np.percentile(gray[mask > 0], 55))
    score = np.where((mask > 0) & (inside > 6), np.abs(gray.astype(np.float32) - target), 1e9)
    seed_row, seed_col = np.unravel_index(int(score.argmin()), score.shape)

    best = None
    for delta in (8, 12, 16, 20, 28, 36, 48):
        canvas = work.copy()
        probe = np.zeros((height + 2, width + 2), np.uint8)
        cv2.floodFill(
            canvas,
            probe,
            (int(seed_col), int(seed_row)),
            (0, 210, 255),
            (delta, delta, delta),
            (delta, delta, delta),
        )
        filled = probe[1:-1, 1:-1] > 0
        probe_area = int(filled.sum())
        outside = int(np.logical_and(filled, mask == 0).sum())
        if probe_area < 200 or outside > 0.02 * probe_area:
            continue
        if best is None or probe_area > best[0]:
            best = (probe_area, delta)
    if best is None:
        raise SystemExit("Допуск заливки не найден")
    _, delta = best
    lo = up = (delta, delta, delta)
    own = flood_own(work, int(seed_row), int(seed_col), lo, up)
    own_area = int(own.sum())

    cv_img = work.copy()
    cv_mask = np.zeros((height + 2, width + 2), np.uint8)
    cv2.floodFill(cv_img, cv_mask, (int(seed_col), int(seed_row)), (0, 210, 255), lo, up)
    cv_area = int((cv_mask[1:-1, 1:-1] > 0).sum())
    both = int(np.logical_and(own > 0, cv_mask[1:-1, 1:-1] > 0).sum())
    either = int(np.logical_or(own > 0, cv_mask[1:-1, 1:-1] > 0).sum())
    agree = both / either if either else 0
    inside = int(np.logical_and(own > 0, mask > 0).sum())
    union = int(np.logical_or(own > 0, mask > 0).sum())
    iou = inside / union if union else 0

    seed_view = work.copy()
    cv2.circle(seed_view, (int(seed_col), int(seed_row)), 8, (0, 0, 255), -1)
    own_view = work.copy()
    own_view[own > 0] = (0, 210, 255)
    pair = np.hstack([seed_view, own_view])
    imwrite(RES / "flood.png", pair)
    imwrite(RES / "flood_cv.png", cv_img)

    orb = cv2.ORB_create(nfeatures=120)
    side = int(np.ceil(np.hypot(width, height))) + 8
    y0 = (side - height) // 2
    x0 = (side - width) // 2
    square = np.zeros((side, side, 3), np.uint8)
    square[y0 : y0 + height, x0 : x0 + width] = work
    gray_square = to_gray(square)
    kp1, des1 = orb.detectAndCompute(gray_square, None)
    center = (side / 2, side / 2)
    matrix = cv2.getRotationMatrix2D(center, 25, 1.0)
    rotated = cv2.warpAffine(square, matrix, (side, side), flags=cv2.INTER_LINEAR, borderValue=(0, 0, 0))
    gray_rot = to_gray(rotated)
    kp2, des2 = orb.detectAndCompute(gray_rot, None)
    matcher = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)
    matches = matcher.match(des1, des2) if des1 is not None and des2 is not None else []
    src_pts = np.float32([kp1[m.queryIdx].pt for m in matches]).reshape(-1, 1, 2) if matches else np.zeros((0, 1, 2), np.float32)
    expected = cv2.transform(src_pts, matrix) if len(matches) else np.zeros((0, 1, 2), np.float32)
    dst_pts = np.float32([kp2[m.trainIdx].pt for m in matches]).reshape(-1, 1, 2) if matches else np.zeros((0, 1, 2), np.float32)
    errors = np.linalg.norm(expected - dst_pts, axis=2).ravel() if len(matches) else np.array([])
    inliers = int((errors < 12).sum()) if errors.size else 0

    vis1 = cv2.drawKeypoints(square, kp1, None, color=(0, 180, 0), flags=cv2.DRAW_MATCHES_FLAGS_DRAW_RICH_KEYPOINTS)
    vis2 = cv2.drawKeypoints(rotated, kp2, None, color=(0, 180, 0), flags=cv2.DRAW_MATCHES_FLAGS_DRAW_RICH_KEYPOINTS)
    imwrite(RES / "orb.png", np.hstack([vis1, vis2]))

    harris = cv2.cornerHarris(np.float32(gray), 2, 3, 0.04)
    local = cv2.dilate(harris, np.ones((7, 7), np.uint8))
    peak_mask = (harris == local) & (harris > 0.04 * float(harris.max()))
    peak_yx = np.column_stack(np.where(peak_mask))
    harris_view = work.copy()
    for row_i, col_i in peak_yx:
        cv2.circle(harris_view, (int(col_i), int(row_i)), 4, (0, 0, 255), 1)
    imwrite(RES / "harris.png", harris_view)
    harris_count = int(len(peak_yx))

    metal_pixels = work[mask > 0]
    hist = [np.bincount(metal_pixels[:, channel], minlength=256) for channel in range(3)]
    means = [float(metal_pixels[:, channel].mean()) for channel in range(3)]
    fig, axis = plt.subplots(figsize=(7.2, 3.6))
    colors = ("#3060C0", "#2E8B57", "#C0392B")
    names = ("B", "G", "R")
    for values, color, name in zip(hist, colors, names):
        axis.plot(values, color=color, label=name, linewidth=1.2)
    axis.set_xlim(0, 255)
    axis.set_xlabel("Уровень")
    axis.set_ylabel("Число пикселей")
    axis.legend()
    fig.tight_layout()
    fig.savefig(RES / "hist_rgb.png", dpi=140)
    plt.close(fig)

    hog_angle = draw_hog(gray[y : y + h, x : x + w], RES / "hog.png")

    metrics = {
        "src_w": src_w,
        "src_h": src_h,
        "width": width,
        "height": height,
        "scale": round(scale, 4),
        "canny_low": low_pair,
        "canny_high": high_pair,
        "edges_low": int((edges_low > 0).sum()),
        "edges_high": int((edges_high > 0).sum()),
        "threshold": int(level),
        "contours_total": len(contours),
        "contours_kept": len(kept),
        "area": round(area, 1),
        "perimeter": round(perimeter, 1),
        "bbox": [int(x), int(y), int(w), int(h)],
        "kept": [
            {
                "area": round(float(cv2.contourArea(c)), 1),
                "perimeter": round(float(cv2.arcLength(c, True)), 1),
                "bbox": [int(v) for v in cv2.boundingRect(c)],
            }
            for c in kept
        ],
        "solidity": round(solidity, 3),
        "seed": [int(seed_col), int(seed_row)],
        "delta": int(delta),
        "own_area": int(own_area),
        "cv_area": int(cv_area),
        "flood_iou_metal": round(float(iou), 3),
        "flood_agree": round(float(agree), 3),
        "orb_src": len(kp1),
        "orb_rot": len(kp2),
        "matches": len(matches),
        "inliers": inliers,
        "harris": harris_count,
        "metal_mean_bgr": [round(v, 1) for v in means],
        "hog_median_deg": round(hog_angle, 1),
    }
    (RES / "metrics.json").write_text(json.dumps(metrics, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(metrics, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
