# -*- coding: utf-8 -*-
"""Прогон предобработки ЛР3. В отчёт не входит."""

import json
import time
from pathlib import Path

import cv2
import matplotlib.pyplot as plt
import numpy as np

ROOT = Path(__file__).resolve().parent
RAW_VIDEO = Path(r"C:\Users\skvor\AppData\Local\Temp\stz_lr3_cam.avi")
FRAMES = ROOT / "data" / "frames"
PROCESSED = ROOT / "data" / "processed"
RESULTS = ROOT / "results"


def imread(path):
    data = np.fromfile(str(path), dtype=np.uint8)
    return cv2.imdecode(data, cv2.IMREAD_COLOR)


def imwrite(path, image):
    path = Path(path)
    ok, buf = cv2.imencode(path.suffix, image)
    if not ok:
        raise RuntimeError(f"imencode failed: {path}")
    buf.tofile(str(path))


def resize_with_padding(image, size=640, value=114):
    h, w = image.shape[:2]
    scale = min(size / w, size / h)
    nw, nh = int(round(w * scale)), int(round(h * scale))
    resized = cv2.resize(image, (nw, nh), interpolation=cv2.INTER_AREA)
    top = (size - nh) // 2
    bottom = size - nh - top
    left = (size - nw) // 2
    right = size - nw - left
    fill = value if image.ndim == 2 else (value, value, value)
    return cv2.copyMakeBorder(
        resized, top, bottom, left, right, cv2.BORDER_CONSTANT, value=fill
    )


def mean_time_ms(func, frame, repeats=30):
    times = []
    for _ in range(repeats):
        t0 = time.perf_counter()
        func(frame)
        times.append((time.perf_counter() - t0) * 1000)
    return float(np.mean(times)), float(np.std(times))


def stats(gray):
    return {
        "mean": round(float(gray.mean()), 2),
        "std": round(float(gray.std()), 2),
        "p5": int(np.percentile(gray, 5)),
        "p95": int(np.percentile(gray, 95)),
    }


def add_salt_pepper(image, amount=0.02, seed=7):
    rng = np.random.default_rng(seed)
    out = image.copy()
    n = int(amount * image.shape[0] * image.shape[1])
    ys = rng.integers(0, image.shape[0], n)
    xs = rng.integers(0, image.shape[1], n)
    half = n // 2
    out[ys[:half], xs[:half]] = 0
    out[ys[half:], xs[half:]] = 255
    return out


def panel(images, titles, path, cmap=None):
    n = len(images)
    fig, axes = plt.subplots(1, n, figsize=(3.2 * n, 3.4))
    if n == 1:
        axes = [axes]
    for ax, im, title in zip(axes, images, titles):
        if im.ndim == 2:
            ax.imshow(im, cmap="gray", vmin=0, vmax=255)
        else:
            ax.imshow(cv2.cvtColor(im, cv2.COLOR_BGR2RGB))
        ax.set_title(title, fontsize=10)
        ax.axis("off")
    fig.tight_layout()
    fig.savefig(path, dpi=140, bbox_inches="tight")
    plt.close(fig)


def main():
    for folder in (FRAMES, PROCESSED, RESULTS):
        folder.mkdir(parents=True, exist_ok=True)

    cap = cv2.VideoCapture(str(RAW_VIDEO))
    fps = cap.get(cv2.CAP_PROP_FPS)
    total = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
    height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
    step = max(total // 50, 1)
    saved_idx = []
    idx = 0
    while True:
        ok, frame = cap.read()
        if not ok:
            break
        if idx % step == 0:
            name = f"frame_{idx:06d}.jpg"
            imwrite(FRAMES / name, frame)
            saved_idx.append(idx)
        idx += 1
    cap.release()

    sample_name = f"frame_{saved_idx[len(saved_idx) // 2]:06d}.jpg"
    img = imread(FRAMES / sample_name)
    h, w = img.shape[:2]
    channels = img.shape[2]
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    hue, saturation, value = cv2.split(hsv)

    imwrite(RESULTS / "original.jpg", img)
    imwrite(RESULTS / "gray.jpg", gray)
    imwrite(RESULTS / "hue.jpg", hue)
    imwrite(RESULTS / "saturation.jpg", saturation)
    imwrite(RESULTS / "value.jpg", value)
    panel(
        [gray, hue, saturation, value],
        ["Оттенки серого", "Тон H", "Насыщенность S", "Яркость V"],
        RESULTS / "color_spaces.png",
    )

    fig, ax = plt.subplots(figsize=(7, 4))
    ax.hist(gray.ravel(), 256, [0, 256], color="#444444")
    ax.set_xlabel("Яркость")
    ax.set_ylabel("Количество пикселей")
    ax.grid(True, alpha=0.3)
    fig.tight_layout()
    fig.savefig(RESULTS / "hist_original.png", dpi=150, bbox_inches="tight")
    plt.close(fig)

    alpha, beta = 1.15, 10
    linear = cv2.convertScaleAbs(img, alpha=alpha, beta=beta)
    gamma = 0.85
    table = np.array([((i / 255.0) ** gamma) * 255 for i in np.arange(256)]).astype("uint8")
    gamma_img = cv2.LUT(img, table)
    lab = cv2.cvtColor(img, cv2.COLOR_BGR2LAB)
    l, a, b = cv2.split(lab)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    l2 = clahe.apply(l)
    clahe_img = cv2.cvtColor(cv2.merge([l2, a, b]), cv2.COLOR_LAB2BGR)

    panel(
        [img, linear, gamma_img, clahe_img],
        ["Исходный", "Линейная", "Гамма 0,85", "CLAHE"],
        RESULTS / "correction.png",
    )

    noisy = add_salt_pepper(img, 0.015)
    gaussian = cv2.GaussianBlur(noisy, (5, 5), 0)
    median = cv2.medianBlur(noisy, 5)
    bilateral = cv2.bilateralFilter(noisy, 9, 75, 75)
    panel(
        [noisy, gaussian, median, bilateral],
        ["Соль и перец", "Гауссов 5×5", "Медианный 5×5", "Билатеральный"],
        RESULTS / "filters_noise.png",
    )
    clean_g = cv2.GaussianBlur(img, (5, 5), 0)
    clean_m = cv2.medianBlur(img, 5)
    clean_b = cv2.bilateralFilter(img, 9, 75, 75)
    panel(
        [img, clean_g, clean_m, clean_b],
        ["Исходный", "Гауссов 5×5", "Медианный 5×5", "Билатеральный"],
        RESULTS / "filters_clean.png",
    )

    times = {
        "gaussian": mean_time_ms(lambda x: cv2.GaussianBlur(x, (5, 5), 0), img),
        "median": mean_time_ms(lambda x: cv2.medianBlur(x, 5), img),
        "bilateral": mean_time_ms(lambda x: cv2.bilateralFilter(x, 9, 75, 75), img),
        "clahe": mean_time_ms(
            lambda x: cv2.cvtColor(
                cv2.merge(
                    [
                        cv2.createCLAHE(2.0, (8, 8)).apply(
                            cv2.split(cv2.cvtColor(x, cv2.COLOR_BGR2LAB))[0]
                        ),
                        cv2.split(cv2.cvtColor(x, cv2.COLOR_BGR2LAB))[1],
                        cv2.split(cv2.cvtColor(x, cv2.COLOR_BGR2LAB))[2],
                    ]
                ),
                cv2.COLOR_LAB2BGR,
            ),
            img,
        ),
    }

    def pipeline(frame):
        lab_i = cv2.cvtColor(frame, cv2.COLOR_BGR2LAB)
        ll, aa, bb = cv2.split(lab_i)
        ll = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(ll)
        out = cv2.cvtColor(cv2.merge([ll, aa, bb]), cv2.COLOR_LAB2BGR)
        return resize_with_padding(out, 640, 114)

    pipe_mean, pipe_std = mean_time_ms(pipeline, img, repeats=30)
    prepared = pipeline(img)
    imwrite(RESULTS / "prepared_640.jpg", prepared)

    # три кадра с разным положением: начало, середина, конец
    pick = [saved_idx[0], saved_idx[len(saved_idx) // 2], saved_idx[-1]]
    pairs = []
    for n in pick:
        name = f"frame_{n:06d}.jpg"
        src = imread(FRAMES / name)
        dst = pipeline(src)
        imwrite(RESULTS / f"pair_src_{n:06d}.jpg", src)
        imwrite(RESULTS / f"pair_dst_{n:06d}.jpg", dst)
        pairs.append(name)

    for file in sorted(FRAMES.glob("*.jpg")):
        frame = imread(file)
        imwrite(PROCESSED / file.name, pipeline(frame))

    gray_lin = cv2.cvtColor(linear, cv2.COLOR_BGR2GRAY)
    gray_gam = cv2.cvtColor(gamma_img, cv2.COLOR_BGR2GRAY)
    gray_cl = cv2.cvtColor(clahe_img, cv2.COLOR_BGR2GRAY)

    def lap(g):
        return round(float(cv2.Laplacian(g, cv2.CV_64F).var()), 2)

    meta = {
        "video": str(RAW_VIDEO),
        "fps": fps,
        "total_frames": total,
        "duration_s": round(total / fps, 2),
        "width": width,
        "height": height,
        "step": step,
        "saved": len(saved_idx),
        "sample": sample_name,
        "channels": channels,
        "dtype": str(img.dtype),
        "gray": stats(gray),
        "linear": {**stats(gray_lin), "alpha": alpha, "beta": beta, "lap": lap(gray_lin)},
        "gamma": {**stats(gray_gam), "gamma": gamma, "lap": lap(gray_gam)},
        "clahe": {**stats(gray_cl), "clip": 2.0, "grid": 8, "lap": lap(gray_cl)},
        "original_lap": lap(gray),
        "times_ms": {k: {"mean": round(v[0], 3), "std": round(v[1], 3)} for k, v in times.items()},
        "pipeline_ms": {"mean": round(pipe_mean, 3), "std": round(pipe_std, 3)},
        "prepared_shape": list(prepared.shape),
        "pairs": pairs,
        "processed_count": len(list(PROCESSED.glob('*.jpg'))),
        "scale": {
            "nw": int(round(w * min(640 / w, 640 / h))),
            "nh": int(round(h * min(640 / w, 640 / h))),
        },
    }
    (RESULTS / "metrics.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(meta, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
