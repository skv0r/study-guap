# -*- coding: utf-8 -*-
"""Рисует пять диаграмм короткой модели. В текст отчёта не входит."""

from model import DIAGRAMS, check_balance
import sheet


def main():
    boxes, errors = check_balance()
    if errors:
        raise SystemExit("ICOM:\n" + "\n".join(errors))
    for spec in DIAGRAMS:
        path = sheet.render(spec)
        print(path.name, "h", round(spec["_height"], 1))
    print("diagrams", len(DIAGRAMS), "boxes", boxes)


if __name__ == "__main__":
    main()
