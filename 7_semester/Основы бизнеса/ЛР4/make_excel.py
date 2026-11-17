#!/usr/bin/env python3
"""Финансовая модель ПР4: формулы Excel, не готовые итоги."""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT = Path(__file__).resolve().parent / "ПР4_Расчет_мощность_Жар-короб.xlsx"

MONTHS = ["Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"]
CURRENT = [370, 350, 380, 390, 410, 430, 440, 450, 400, 380, 360, 440]
FORECAST = [500, 473, 513, 527, 553, 581, 594, 608, 540, 513, 486, 592]


def main():
    wb = Workbook()
    thin = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    fill_h = PatternFill("solid", fgColor="D9D9D9")
    font_h = Font(name="Times New Roman", bold=True, size=11)
    font_n = Font(name="Times New Roman", size=11)
    wrap = Alignment(wrap_text=True, vertical="center")

    # --- параметры ---
    wp = wb.active
    wp.title = "Параметры"
    wp.append(["Показатель", "Значение", "Пояснение"])
    rows = [
        ["Цена короба, ₽", 8000, "38 400 000 / 4 800 изделий"],
        ["Мощность одной смены, шт./мес", 480, "загрузка 400 / 480 = 83 %"],
        ["Ставка дисконтирования", 0.18, "18 % годовых"],
        ["Горизонт NPV, лет", 3, "по заданию"],
        ["Срок амортизации, лет", 5, "линейная"],
        ["Мощность сценария 1, шт./мес", 580, "экстенсивный"],
        ["Мощность сценария 2, шт./мес", 620, "интенсивный"],
        ["Мощность сценария 3, шт./мес", 560, "гибридный"],
        ["Капитальные затраты 1, ₽", 620000, "второй термопресс"],
        ["Капитальные затраты 2, ₽", 1850000, "раскройный комплекс ППЭ"],
        ["Капитальные затраты 3, ₽", 0, "аутсорсинг, оборудования нет"],
        ["Доп. операционные расходы 1, ₽/мес", 188000, "аренда 84 000 + два сборщика 104 000"],
        ["Доп. операционные расходы 2, ₽/мес", 142000, "надбавка второй смены + электроэнергия"],
        ["Доп. операционные расходы 3, ₽/мес", 78000, "раскрой 110 ₽/лист и доставка кроя"],
    ]
    for row in rows:
        wp.append(row)
    wp["B4"].number_format = "0%"
    for col, w in enumerate([42, 16, 55], start=1):
        wp.column_dimensions[get_column_letter(col)].width = w
    for r in range(1, 16):
        for c in range(1, 4):
            cell = wp.cell(r, c)
            cell.font = font_h if r == 1 else font_n
            cell.border = thin
            cell.alignment = wrap
            if r == 1:
                cell.fill = fill_h

    # --- спрос ---
    ws = wb.create_sheet("Спрос")
    ws.append(["Месяц", "Текущий спрос, шт.", "Прогноз +35 %, шт.", "С1 доп., шт.", "С2 доп., шт.", "С3 доп., шт."])
    for i, m in enumerate(MONTHS, start=2):
        ws.cell(i, 1, m)
        ws.cell(i, 2, CURRENT[i - 2])
        ws.cell(i, 3, FORECAST[i - 2])
        ws.cell(i, 4, f"=MIN(C{i},Параметры!$B$7)-B{i}")
        ws.cell(i, 5, f"=MIN(C{i},Параметры!$B$8)-B{i}")
        ws.cell(i, 6, f"=MIN(C{i},Параметры!$B$9)-B{i}")
    ws.append(
        [
            "Итого за год",
            "=SUM(B2:B13)",
            "=SUM(C2:C13)",
            "=SUM(D2:D13)",
            "=SUM(E2:E13)",
            "=SUM(F2:F13)",
        ]
    )
    for col, w in enumerate([22, 22, 22, 16, 16, 16], start=1):
        ws.column_dimensions[get_column_letter(col)].width = w
    for r in range(1, 15):
        for c in range(1, 7):
            cell = ws.cell(r, c)
            cell.font = font_h if r in (1, 14) else font_n
            cell.border = thin
            cell.alignment = wrap
            if r == 1:
                cell.fill = fill_h

    # --- денежные потоки ---
    wd = wb.create_sheet("Потоки")
    wd.append(
        [
            "Показатель",
            "Сценарий 1 экстенсивный",
            "Сценарий 2 интенсивный",
            "Сценарий 3 гибридный",
        ]
    )
    # row2 extra boxes
    wd["A2"] = "Дополнительный выпуск, шт./год"
    wd["B2"] = "=Спрос!D14"
    wd["C2"] = "=Спрос!E14"
    wd["D2"] = "=Спрос!F14"
    wd["A3"] = "Дополнительная выручка, ₽/год"
    wd["B3"] = "=B2*Параметры!B2"
    wd["C3"] = "=C2*Параметры!B2"
    wd["D3"] = "=D2*Параметры!B2"
    wd["A4"] = "Доп. операционные расходы, ₽/год"
    wd["B4"] = "=Параметры!B13*12"
    wd["C4"] = "=Параметры!B14*12"
    wd["D4"] = "=Параметры!B15*12"
    wd["A5"] = "Амортизация, ₽/год"
    wd["B5"] = "=Параметры!B10/Параметры!B6"
    wd["C5"] = "=Параметры!B11/Параметры!B6"
    wd["D5"] = "=Параметры!B12/Параметры!B6"
    wd["A6"] = "Доп. расходы всего, ₽/год"
    wd["B6"] = "=B4+B5"
    wd["C6"] = "=C4+C5"
    wd["D6"] = "=D4+D5"
    wd["A7"] = "Доп. прибыль, ₽/год"
    wd["B7"] = "=B3-B6"
    wd["C7"] = "=C3-C6"
    wd["D7"] = "=D3-D6"
    wd["A8"] = "Денежный поток, ₽/год"
    wd["B8"] = "=B3-B4"
    wd["C8"] = "=C3-C4"
    wd["D8"] = "=D3-D4"
    wd["A9"] = "Капитальные затраты, ₽"
    wd["B9"] = "=Параметры!B10"
    wd["C9"] = "=Параметры!B11"
    wd["D9"] = "=Параметры!B12"
    wd["A10"] = "Поток года 0"
    wd["B10"] = "=-B9"
    wd["C10"] = "=-C9"
    wd["D10"] = "=-D9"
    wd["A11"] = "Поток года 1"
    wd["B11"] = "=B8"
    wd["C11"] = "=C8"
    wd["D11"] = "=D8"
    wd["A12"] = "Поток года 2"
    wd["B12"] = "=B8"
    wd["C12"] = "=C8"
    wd["D12"] = "=D8"
    wd["A13"] = "Поток года 3"
    wd["B13"] = "=B8"
    wd["C13"] = "=C8"
    wd["D13"] = "=D8"
    wd["A14"] = "NPV за 3 года, ₽"
    wd["B14"] = "=B10+NPV(Параметры!B4,B11:B13)"
    wd["C14"] = "=C10+NPV(Параметры!B4,C11:C13)"
    wd["D14"] = "=D10+NPV(Параметры!B4,D11:D13)"
    wd["A15"] = "Простой срок окупаемости, лет"
    wd["B15"] = '=IF(B8<=0,"не окупается",B9/B8)'
    wd["C15"] = '=IF(C8<=0,"не окупается",C9/C8)'
    wd["D15"] = '=IF(D9=0,0,IF(D8<=0,"не окупается",D9/D8))'
    wd["A16"] = "Победитель по NPV"
    wd["B16"] = '=IF(AND(D14>=B14,D14>=C14),"Сценарий 3",IF(C14>=B14,"Сценарий 2","Сценарий 1"))'

    for col, w in enumerate([36, 28, 28, 26], start=1):
        wd.column_dimensions[get_column_letter(col)].width = w
    for r in range(1, 17):
        for c in range(1, 5):
            cell = wd.cell(r, c)
            cell.font = font_h if r in (1, 14, 16) else font_n
            cell.border = thin
            cell.alignment = wrap
            if r == 1:
                cell.fill = fill_h
            if r in range(3, 15) and c >= 2:
                cell.number_format = "#,##0"
            if r == 15 and c >= 2:
                cell.number_format = "0.00"

    wb.save(OUT)
    print(OUT)


if __name__ == "__main__":
    main()
