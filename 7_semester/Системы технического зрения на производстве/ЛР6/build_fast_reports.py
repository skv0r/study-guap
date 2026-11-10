# -*- coding: utf-8 -*-
"""Сборка отчётов ЛР5–ЛР7 по фотографиям гайки и шайбы. В текст отчётов не входит."""

import shutil
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = Path(__file__).resolve().parent.parent
TEMPLATE = Path(r"C:\Users\skvor\stz-lr67\fast\results\title_template.docx")
SRC = BASE / "ЛР5" / "ЛР5_Отчет_Буренков_СТЗ.docx"
OUT5 = SRC
OUT6 = BASE / "ЛР6" / "ЛР6_Отчет_Буренков_СТЗ.docx"
OUT7 = BASE / "ЛР7" / "ЛР7_Отчет_Буренков_СТЗ.docx"
RES = Path(r"C:\Users\skvor\stz-lr67\fast\results")
RUNS = Path(r"C:\Users\skvor\stz-lr67\fast\runs")
BLACK = RGBColor(0, 0, 0)

H5 = {
    "ВВЕДЕНИЕ": "‹В5›",
    "1 Цель работы": "‹51›",
    "2 Вариант, объект и классы": "‹52›",
    "3 Набор данных": "‹53›",
    "4 Размеченные изображения": "‹54›",
    "5 Число изображений и рамок": "‹55›",
    "6 Правила маркировки": "‹56›",
    "7 Листинг алгоритма": "‹57›",
    "8 Результаты алгоритма": "‹58›",
    "9 Проблемы разметки": "‹59›",
    "10 Версия набора для следующей работы": "‹5A›",
    "11 Выводы": "‹5B›",
    "ЗАКЛЮЧЕНИЕ": "‹5Z›",
}
H6 = {
    "ВВЕДЕНИЕ": "‹В6›",
    "1 Цель, вариант, объект и классы": "‹61›",
    "2 Версия набора и число рамок": "‹62›",
    "3 Программная среда": "‹63›",
    "4 Файл data.yaml и предобработка": "‹64›",
    "5 Конфигурация сети и фактор": "‹65›",
    "6 Параметры опытов и программа обучения": "‹66›",
    "7 Кривые обучения и результаты на valid": "‹67›",
    "8 Кадры проверочной выборки": "‹68›",
    "9 Выбор модели и файлы для следующей работы": "‹69›",
    "10 Влияние фактора и ограничения": "‹6A›",
    "ЗАКЛЮЧЕНИЕ": "‹6Z›",
}
H7 = {
    "ВВЕДЕНИЕ": "‹В7›",
    "1 Цель, вариант, объект и классы": "‹71›",
    "2 Модель и версия набора": "‹72›",
    "3 Подбор порога на valid": "‹73›",
    "4 Источник test": "‹74›",
    "5 Предобработка и координаты": "‹75›",
    "6 Сопоставление рамок": "‹76›",
    "7 Метрики test": "‹77›",
    "8 Разбор кадров": "‹78›",
    "9 Время обработки": "‹79›",
    "10 Оценки по классам": "‹7A›",
    "11 Файлы результата и режим": "‹7B›",
    "12 Пригодность модели и следующие опыты": "‹7C›",
    "ЗАКЛЮЧЕНИЕ": "‹7Z›",
}

LISTING_CHECK = """from pathlib import Path

NAMES = ["nut", "washer"]
MIN_AREA = 0.001

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

def audit_labels(label_dir, n_classes):
    problems = []
    for path in sorted(Path(label_dir).glob("*.txt")):
        rows = [line.split() for line in path.read_text().splitlines()
                if line.strip()]
        if not rows:
            problems.append((path.name, "пустой файл разметки"))
            continue
        for row in rows:
            status = check_label_line(" ".join(row), n_classes)
            if status != "ok":
                problems.append((path.name, status))
            bw, bh = float(row[3]), float(row[4])
            if bw * bh < MIN_AREA:
                problems.append((path.name, "слишком малая рамка"))
    return problems""".splitlines()

LISTING_TRAIN = """common = dict(
    epochs=40, batch=4, optimizer="AdamW",
    lr0=0.001, lrf=0.01, seed=42,
    deterministic=True, amp=False, patience=0,
    workers=0, device=0, freeze=0,
    mosaic=0.0, mixup=0.0, close_mosaic=0,
    hsv_h=0.0, hsv_s=0.0, hsv_v=0.0,
    degrees=0.0, translate=0.0, scale=0.0,
    shear=0.0, perspective=0.0,
    flipud=0.0, fliplr=0.0,
)
for name, imgsz in (("A", 416), ("B", 640)):
    model = YOLO("yolov8n.pt")
    model.train(data="data.yaml", name=name, imgsz=imgsz, **common)
    YOLO(name + "/weights/best.pt").val(
        data="data.yaml", split="val", imgsz=imgsz,
        conf=0.001, iou=0.7,
    )""".splitlines()

LISTING_PREP = """def prepare_frame(frame):
    return frame.copy()""".splitlines()

LISTING_YAML = """train: train/images
val: valid/images
test: test/images
names:
  0: nut
  1: washer""".splitlines()

LISTING_MATCH = """def iou(a, b):
    width = max(0.0, min(a[2], b[2]) - max(a[0], b[0]))
    height = max(0.0, min(a[3], b[3]) - max(a[1], b[1]))
    intersection = width * height
    area_a = max(0.0, a[2] - a[0]) * max(0.0, a[3] - a[1])
    area_b = max(0.0, b[2] - b[0]) * max(0.0, b[3] - b[1])
    union = area_a + area_b - intersection
    return intersection / union if union > 0 else 0.0

def match_counts(truth, predicted, match_iou=0.5):
    used = set()
    tp = 0
    ordered = sorted(predicted, key=lambda item: item[1], reverse=True)
    for cls, confidence, box in ordered:
        candidates = [(iou(box, other), index)
                      for index, (label, other) in enumerate(truth)
                      if index not in used and label == cls]
        overlap, index = max(candidates, default=(0.0, -1))
        if index >= 0 and overlap >= match_iou:
            used.add(index)
            tp += 1
    return tp, len(predicted) - tp, len(truth) - tp""".splitlines()


def rfonts(run, name):
    run.font.name = name
    run.font.color.rgb = BLACK
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.append(fonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        fonts.set(qn(attr), name)


def style_run(run, size=14, bold=False, font="Times New Roman"):
    rfonts(run, font)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = False


def set_spacing(paragraph, before=0, after=0, line=360):
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    paragraph.paragraph_format.line_spacing = line / 240.0


def set_indent(paragraph, first=None, left=None):
    ppr = paragraph._p.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    if first is not None:
        ind.set(qn("w:firstLine"), str(first))
    if left is not None:
        ind.set(qn("w:left"), str(left))


def add_text(doc, text, *, first=709, before=0, after=0, line=360, center=False, bold=False, size=14):
    paragraph = doc.add_paragraph()
    set_spacing(paragraph, before, after, line)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.JUSTIFY
    if not center and first:
        set_indent(paragraph, first=first)
    run = paragraph.add_run(text)
    style_run(run, size=size, bold=bold)
    return paragraph


def add_heading(doc, text, *, center=False, break_before=False):
    paragraph = doc.add_paragraph()
    set_spacing(paragraph, before=0, after=12, line=360)
    paragraph.paragraph_format.keep_with_next = True
    if center:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        set_indent(paragraph, first=709)
    if break_before:
        run_br = paragraph.add_run()
        style_run(run_br, bold=True)
        br = OxmlElement("w:br")
        br.set(qn("w:type"), "page")
        run_br._r.append(br)
    run = paragraph.add_run(text)
    style_run(run, bold=True)
    return paragraph


def add_toc(doc, title, token):
    paragraph = doc.add_paragraph()
    set_spacing(paragraph, line=360)
    ppr = paragraph._p.get_or_add_pPr()
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:leader"), "dot")
    tab.set(qn("w:pos"), "9071")
    tabs.append(tab)
    ppr.append(tabs)
    run = paragraph.add_run(title)
    style_run(run)
    paragraph.add_run().add_tab()
    num = paragraph.add_run(token)
    style_run(num)


def add_code(doc, line):
    paragraph = doc.add_paragraph()
    set_spacing(paragraph, before=0, after=0, line=240)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_indent(paragraph, first=0, left=0)
    run = paragraph.add_run(line if line else " ")
    style_run(run, size=9, font="Courier New")
    return paragraph


def add_listing(doc, caption, lines):
    add_text(doc, caption, first=0, before=6, after=6, line=240)
    for line in lines:
        add_code(doc, line.rstrip("\n"))


def set_cell(cell, text, *, bold=False, header=False):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    set_spacing(paragraph, before=2, after=2, line=240)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.first_line_indent = Cm(0)
    run = paragraph.add_run(text)
    style_run(run, size=14, bold=bold)
    tcpr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for edge in ("top", "left", "bottom", "right"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:w"), "40")
        node.set(qn("w:type"), "dxa")
        mar.append(node)
    tcpr.append(mar)
    if header:
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "F2F2F2")
        tcpr.append(shd)


def set_table_widths(table, widths):
    table.autofit = False
    table.allow_autofit = False
    tbl = table._tbl
    tblpr = tbl.tblPr
    tblw = tblpr.find(qn("w:tblW"))
    if tblw is None:
        tblw = OxmlElement("w:tblW")
        tblpr.append(tblw)
    tblw.set(qn("w:w"), str(sum(widths)))
    tblw.set(qn("w:type"), "dxa")
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for child in list(grid):
            grid.remove(child)
    else:
        grid = OxmlElement("w:tblGrid")
        tblpr.addnext(grid)
    for width in widths:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(width))
        grid.append(col)
    for row in table.rows:
        for cell, width in zip(row.cells, widths):
            tcpr = cell._tc.get_or_add_tcPr()
            tcw = tcpr.find(qn("w:tcW"))
            if tcw is None:
                tcw = OxmlElement("w:tcW")
                tcpr.append(tcw)
            tcw.set(qn("w:w"), str(width))
            tcw.set(qn("w:type"), "dxa")
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "4")
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), "000000")
        borders.append(node)
    tblpr.append(borders)


def add_table(doc, caption, headers, rows, widths):
    inset = max(0, (9355 - sum(widths)) // 2)
    add_text(doc, caption, first=0, before=6, after=6, line=240)
    cap = doc.paragraphs[-1]
    cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
    cap.paragraph_format.keep_with_next = True
    set_indent(cap, first=0, left=inset)
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    set_table_widths(table, widths)
    for i, header in enumerate(headers):
        set_cell(table.rows[0].cells[i], header, bold=True, header=True)
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            set_cell(table.rows[r + 1].cells[c], value)


def add_picture(doc, path, caption, width_cm):
    paragraph = doc.add_paragraph()
    set_spacing(paragraph, before=6, after=0, line=240)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.add_run().add_picture(str(path), width=Cm(width_cm))
    add_text(doc, caption, first=0, before=6, after=10, line=360, center=True)


def fresh(number, theme):
    doc = Document(str(TEMPLATE))
    body = doc.element.body
    children = list(body)
    for child in children[12:-1]:
        body.remove(child)
    if number == 5:
        return doc
    replacements = {
        "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 5": f"ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № {number}",
        "Разметка изображений и видеокадров производственных объектов": theme,
    }
    found = set()
    for node in doc.element.body.iter(qn("w:t")):
        if node.text in replacements:
            node.text = replacements[node.text]
            found.add(node.text)
    if found != set(replacements.values()):
        raise SystemExit(f"Титул заменён не полностью: {found}")
    return doc


def fill_pages(word, path, headings):
    doc = word.Documents.Open(str(path))
    doc.Repaginate()
    pages = {}
    for para in doc.Paragraphs:
        text = (
            para.Range.Text.replace("\r", "")
            .replace("\x07", "")
            .replace("\x0c", "")
            .strip()
        )
        if text in headings:
            pages[text] = int(para.Range.Information(3))
    doc.Close(False)
    missing = [title for title in headings if title not in pages]
    if missing:
        raise SystemExit(f"Нет страниц для {path.name}: {missing}")
    return pages


def write_page_numbers(path, headings, pages):
    doc = Document(str(path))
    tokens = {token: str(pages[title]) for title, token in headings.items()}
    for node in doc.element.body.iter(qn("w:t")):
        if node.text in tokens:
            node.text = tokens[node.text]
    doc.save(str(path))


def build_lr5():
    doc = fresh(5, "")
    add_heading(doc, "СОДЕРЖАНИЕ", center=True)
    for title, token in H5.items():
        add_toc(doc, title, token)

    add_heading(doc, "ВВЕДЕНИЕ", center=True, break_before=True)
    add_text(doc, "Работа готовит кадры двух крепёжных деталей к обучению детектора. На каждом кадре стоит одна рамка в формате YOLO. Дальше набор делится на train, valid и test, а разметка проверяется по площади рамки.")

    add_heading(doc, "1 Цель работы", break_before=True)
    add_text(doc, "Цель работы — разметить изображения производственных деталей рамками, проверить разметку и подготовить версию набора для обучения нейронной сети.")

    add_heading(doc, "2 Вариант, объект и классы", break_before=True)
    add_text(doc, "Номер варианта — 1. В таблице заданий объект назван деталями на конвейере, а классы — gear, washer и nut. Индивидуальная часть — алгоритм 1, проверка (а): рамка, у которой произведение ширины и высоты меньше 0,001.")
    add_text(doc, "В работу взяты фотографии одной шестигранной гайки и одной плоской шайбы. В файле имён класс 0 — nut, класс 1 — washer. Кадров шестерни gear в этом наборе нет, поэтому третий класс не заводился.")

    add_heading(doc, "3 Набор данных", break_before=True)
    add_text(doc, "Фотографии взяты из набора The Fasteners на Kaggle: https://www.kaggle.com/datasets/alexandrparkhomenko/the-fasteners. Это снимки телефона от 17 апреля 2022 года. Детали лежат на бирюзовой ткани при комнатном свете. Роликов в архиве нет: каждый файл — отдельный кадр 1600×1200.")
    add_text(doc, "Гайка снята 11 раз, с 09:44:46 до 09:47:19. На торце видны клейма M и 5I, по ним это одна и та же деталь: сверху, сбоку и в промежуточных поворотах. Шайба снята 3 раза, с 09:48:36 до 09:48:45, тоже одна и та же, кольцо с отверстием. Остальные папки архива содержат другие изделия и в набор не входили. Всего 14 кадров и 14 рамок: 11 гаек и 3 шайбы.")

    add_heading(doc, "4 Размеченные изображения", break_before=True)
    add_text(doc, "Классы набора — nut и washer. На рисунках 1–3 рамки восстановлены из файлов разметки. На каждом кадре один объект, подпись над рамкой совпадает с классом в файле.")
    add_picture(doc, RES / "mark_nut_top.jpg", "Рисунок 1 – Гайка сверху, класс nut", 12)
    add_picture(doc, RES / "mark_nut_side.jpg", "Рисунок 2 – Гайка сбоку, класс nut", 12)
    add_picture(doc, RES / "mark_washer.jpg", "Рисунок 3 – Плоская шайба, класс washer", 12)

    add_heading(doc, "5 Число изображений и рамок", break_before=True)
    add_text(doc, "Кадры разложены по времени съёмки, чтобы соседние секунды одной позы не попали сразу и в обучение, и в проверку. Train — первые восемь видов гайки и первый кадр шайбы, 9 кадров. Valid — два следующих вида гайки и средний кадр шайбы, 3 кадра. Test — последний вид гайки и последний кадр шайбы, 2 кадра. Доли 9/14, 3/14 и 2/14 близки к 70/20/10. Шайба стоит в каждой выборке, иначе оценка этого класса на valid или test была бы пустой. В train у шайбы один кадр.")
    add_table(
        doc,
        "Таблица 1 – Изображения и рамки по выборкам",
        ["Выборка", "Кадры", "nut", "washer", "Рамок"],
        [
            ["train", "9", "8", "1", "9"],
            ["valid", "3", "2", "1", "3"],
            ["test", "2", "1", "1", "2"],
            ["Всего", "14", "11", "3", "14"],
        ],
        [1871, 1871, 1871, 1871, 1871],
    )
    add_text(doc, "Коэффициент дисбаланса равен 11/3 = 3,67. Ориентир методички — не выше 3. Шайб в наборе меньше, чем гаек, и это ограничение остаётся в версии для обучения.")

    add_heading(doc, "6 Правила маркировки", break_before=True)
    add_text(doc, "Правило 1 выполнено на всех 14 кадрах: интересующий объект один, и у него есть ровно одна рамка. Пустых файлов разметки нет.")
    add_text(doc, "Правило 2 выполнено: рамка закрывает металл целиком. На рисунке 1 виден весь шестигранник и резьба, на рисунке 2 — вся боковая грань и отверстие, на рисунке 3 — всё кольцо шайбы.")
    add_text(doc, "Правило 3 выдержано с запасом по ткани. Нижний край рамки захватывает тень под деталью, а на боковых видах по сторонам остаётся полоса фона. Металл при этом не обрезан. Самый заметный запас — у бокового кадра, который позже попадёт в valid.")
    add_text(doc, "Правило 4 в этих кадрах не к чему применить. На каждом снимке одна деталь, вторая её не закрывает, пересечения рамок нет.")
    add_text(doc, "Правило 5 выполнено именами nut и washer: латиница, без пробела, отдельно гайка и отдельно шайба. Общий класс fastener не использовался.")

    add_heading(doc, "7 Листинг алгоритма", break_before=True)
    add_text(doc, "Алгоритм 1 читает строку YOLO, проверяет пять чисел, известный класс и диапазон от 0 до 1, затем возвращает рамку на кадр. Проверка (а) отмечает рамку, если произведение нормированных ширины и высоты меньше 0,001. Листинг 1 содержит обе части.")
    add_listing(doc, "Листинг 1 – Проверка строки и малых рамок", LISTING_CHECK)

    add_heading(doc, "8 Результаты алгоритма", break_before=True)
    add_text(doc, "На кадрах рисунков 1–3 программа печатает ok. По всему каталогу из 14 файлов список замечаний пуст: малых рамок, пустых файлов, неизвестного класса и чисел вне диапазона нет. Площади рабочих рамок лежат примерно от 0,11 до 0,49 кадра.")
    add_text(doc, "Чтобы увидеть срабатывание проверки (а), в отдельную копию записана строка 0 0,500000 0,500000 0,010000 0,010000. Площадь этой рамки 0,000100. Программа вывела: demo_small.txt — слишком малая рамка. Эта копия в train, valid и test не входит.")

    add_heading(doc, "9 Проблемы разметки", break_before=True)
    add_text(doc, "Проверка (а) рабочий каталог не забраковала. Оставшийся недостаток — ширина рамки по правилу 3. На боковых видах гайки рамка вместе с металлом забирает тень и полосу ткани. Для обучения это значит, что фон рядом с гайкой частично входит в пример класса nut. Обрезанных деталей нет, поэтому рамки не сужались ценой потери края металла.")
    add_text(doc, "Второе ограничение — состав набора. Шайба представлена тремя кадрами, шестерни нет, перекрытия двух деталей нет. Коэффициент 3,67 выше ориентира 3. Для лабораторной работы №6 оставлена эта версия: 14 кадров, классы nut и washer, разбиение таблицы 1.")

    add_heading(doc, "10 Версия набора для следующей работы", break_before=True)
    add_text(doc, "Фотографии опубликованы по адресу https://www.kaggle.com/datasets/alexandrparkhomenko/the-fasteners. Версия для лабораторной работы №6 — те же 14 кадров с рамками YOLO и разбиением train 9, valid 3, test 2. Имена классов в data.yaml: 0 — nut, 1 — washer. Test в выборе модели не участвует.")

    add_heading(doc, "11 Выводы", break_before=True)
    add_text(doc, "Размечены 11 кадров одной гайки и 3 кадра одной шайбы, размер 1600×1200. Рамки покрывают деталь целиком. Проверка (а) на рабочих файлах замечаний не дала и нашла искусственную рамку площадью 0,000100 в отдельной копии. Набор передан в обучение с дисбалансом 3,67 и с запасом фона на боковых рамках.")

    add_heading(doc, "ЗАКЛЮЧЕНИЕ", center=True, break_before=True)
    add_text(doc, "Для обучения подготовлены 14 кадров классов nut и washer. Train содержит 9 кадров, valid — 3, test — 2. Малых рамок в рабочей разметке нет.")
    doc.save(str(OUT5))


def build_lr6():
    doc = fresh(6, "Выбор и обучение нейронной сети по размеченным данным")
    add_heading(doc, "СОДЕРЖАНИЕ", center=True)
    for title, token in H6.items():
        add_toc(doc, title, token)

    add_heading(doc, "ВВЕДЕНИЕ", center=True, break_before=True)
    add_text(doc, "Работа обучает детектор по разметке лабораторной работы №5 и сравнивает два значения одного фактора. Качество считается на valid. Test в сравнении не используется.")

    add_heading(doc, "1 Цель, вариант, объект и классы", break_before=True)
    add_text(doc, "Цель работы — выбрать конфигурацию нейронной сети и обучить её по размеченным кадрам. Номер варианта — 1. Объект и классы те же, что в лабораторной работе №5: гайка nut и шайба washer. Шестерня gear в наборе отсутствует.")
    add_text(doc, "Фактор варианта 1 — сторона входа imgsz. Опыт A обучается при imgsz = 416, опыт B — при imgsz = 640. Остальные параметры общие.")

    add_heading(doc, "2 Версия набора и число рамок", break_before=True)
    add_text(doc, "Версия набора — 14 кадров 1600×1200 из лабораторной работы №5. Train — 9 кадров, valid — 3, test — 2. Разбиение по времени съёмки одно и то же в обоих опытах. Числа рамок приведены в таблице 1.")
    add_table(
        doc,
        "Таблица 1 – Рамки в train, valid и test",
        ["Выборка", "Кадры", "nut", "washer", "Рамок"],
        [
            ["train", "9", "8", "1", "9"],
            ["valid", "3", "2", "1", "3"],
            ["test", "2", "1", "1", "2"],
        ],
        [1871, 1871, 1871, 1871, 1871],
    )
    add_text(doc, "В каждой выборке есть и гайка, и шайба. У шайбы в train один кадр. Test сохранён для лабораторной работы №7.")

    add_heading(doc, "3 Программная среда", break_before=True)
    add_text(doc, "Обучение выполнено на Python 3.12.10. Использованы PyTorch 2.6.0+cu124, torchvision 0.21.0+cu124, Ultralytics 8.4.201, OpenCV 5.0.0, NumPy 2.5.4 и PyYAML 6.0.3. Стартовые веса — файл yolov8n.pt.")
    add_text(doc, "Процессор — AMD Ryzen 5 5600X, оперативная память — 16 ГБ. Видеокарта — NVIDIA GeForce RTX 2060, 6 ГБ. Оба опыта считались на устройстве CUDA 0.")

    add_heading(doc, "4 Файл data.yaml и предобработка", break_before=True)
    add_text(doc, "Файл data.yaml задаёт три выборки и имена классов. Ключ val указывает на кадры valid. Строка с корнем каталога в листинг не входит. Оба опыта читают один и тот же файл.")
    add_listing(doc, "Листинг 1 – Файл data.yaml", LISTING_YAML)
    add_text(doc, "Кадры имеют размер 1600×1200, разметка стоит на этих пикселях. Функция подготовки возвращает копию кадра того же размера. Сеть сама приводит вход к стороне imgsz. Отдельный фильтр яркости к размеченным кадрам не применяется: он сдвинул бы смысл кадра относительно рамки. Листинг 2 показывает подготовку.")
    add_listing(doc, "Листинг 2 – Подготовка кадра", LISTING_PREP)

    add_heading(doc, "5 Конфигурация сети и фактор", break_before=True)
    add_text(doc, "Базовая сеть — YOLOv8n со стартовыми весами yolov8n.pt. На 9 обучающих кадрах компактная сеть позволяет довести оба опыта до 40 эпох и сравнить один фактор.")
    add_text(doc, "Фактор — сторона входа. Кадр 1600×1200 при входе 416 уменьшается сильнее, чем при входе 640. Метрика mAP@0,5 оценивает грубое попадание в объект. Метрика mAP@0,5:0,95 строже к границам рамки, по ней выбирается модель.")
    add_text(doc, "Аугментации выключены: mosaic, mixup, цвет, поворот, перенос, масштаб, сдвиг, перспектива и отражения равны нулю. Различие опытов относится к imgsz. Одинаковы зерно 42, оптимизатор, число эпох и устройство.")

    add_heading(doc, "6 Параметры опытов и программа обучения", break_before=True)
    add_text(doc, "Общие параметры сведены в таблице 2. Меняется только imgsz: 416 в опыте A и 640 в опыте B. Параметр patience = 0 отключает раннюю остановку, оба опыта проходят 40 эпох. Программа приведена в листинге 3.")
    add_table(
        doc,
        "Таблица 2 – Общие параметры опытов A и B",
        ["Параметр", "Значение"],
        [
            ["Модель и веса", "YOLOv8n, yolov8n.pt"],
            ["Эпохи, batch", "40, 4"],
            ["Оптимизатор, lr0, lrf", "AdamW, 0,001, 0,01"],
            ["seed, deterministic", "42, включён"],
            ["patience, amp, workers", "0, выключен, 0"],
            ["device, freeze", "CUDA 0, 0"],
            ["mosaic, mixup, close_mosaic", "0, 0, 0"],
            ["Цвет, геометрия, отражения", "0"],
        ],
        [4200, 5155],
    )
    add_text(doc, "После обучения отдельно проверяется файл best.pt на valid. Сначала среди предсказаний модели выполняется подавление дубликатов NMS: параметр iou = 0,7 задаёт, насколько сильно должны перекрываться две рамки одного класса, чтобы считать их одним и тем же объектом и оставить более уверенную. Параметр conf = 0,001 задаёт низкий порог уверенности, чтобы в эту проверку попали почти все предсказания. Затем оставшиеся рамки сравниваются с эталонной разметкой: метрика mAP считается по совпадению предсказания и правильной рамки при порогах IoU от 0,5 до 0,95. Это не рабочий порог детектора. Рабочий порог выбирается в лабораторной работе №7.")
    add_listing(doc, "Листинг 3 – Программа двух опытов", LISTING_TRAIN)

    add_heading(doc, "7 Кривые обучения и результаты на valid", break_before=True)
    add_text(doc, "Кривые опыта A показаны на рисунке 1, кривые опыта B — на рисунке 2. В журнал попадает метрика текущей эпохи. Файл best.pt выбирается по проверке на valid и может не совпасть с последней эпохой. В таблицу 3 входят результаты отдельной проверки best.pt.")
    add_picture(doc, RUNS / "v01_a" / "results.png", "Рисунок 1 – Кривые обучения опыта A, imgsz = 416", 15.5)
    add_picture(doc, RUNS / "v01_b" / "results.png", "Рисунок 2 – Кривые обучения опыта B, imgsz = 640", 15.5)
    add_text(doc, "У опыта A метрика mAP@0,5:0,95 в журнале равна 0,659 на эпохе 33 и остаётся 0,659 на эпохе 40. Кривая train/box_loss на эпохе 40 равна 0,427. Проверка best.pt даёт те же 0,659.")
    add_text(doc, "У опыта B mAP@0,5:0,95 достигает 0,543 на эпохе 33 и к эпохе 40 снижается до 0,494. Кривая train/box_loss на эпохе 40 равна 0,567. Проверка best.pt даёт 0,543, то есть файл соответствует лучшему участку журнала, а не последней эпохе. К концу обучения строгая метрика valid падает, пока ошибка обучения остаётся высокой: на 9 кадрах вход 640 за 40 эпох не дал более устойчивой проверки, чем вход 416.")
    add_table(
        doc,
        "Таблица 3 – Проверка best.pt на valid",
        ["Опыт", "imgsz", "Эпохи", "mAP@0,5", "mAP@0,5:0,95", "Время, с", "Веса, МБ"],
        [
            ["A", "416", "40", "0,995", "0,659", "35,78", "6,470"],
            ["B", "640", "40", "0,829", "0,543", "31,80", "6,493"],
        ],
        [900, 1100, 1100, 1455, 1900, 1500, 1400],
    )
    add_text(doc, "Разница mAP@0,5:0,95 составляет 0,116 в пользу опыта A. Время обучения 35,78 с и 31,80 с, размеры файлов весов 6,470 МБ и 6,493 МБ. Valid содержит три кадра, поэтому обе оценки mAP относятся к этой короткой проверке.")

    add_heading(doc, "8 Кадры проверочной выборки", break_before=True)
    add_text(doc, "Valid состоит из трёх кадров. Ниже разобраны все три при весах best.pt опыта A, входе 416 и пороге уверенности 0,55. Выбор порога по valid описан в лабораторной работе №7. Зелёная рамка — предсказание, совпавшее с разметкой того же класса при IoU не ниже 0,5. Красная рамка — лишнее предсказание. Оранжевая рамка на этих трёх кадрах не понадобилась: пропусков нет.")
    add_picture(doc, RES / "frames" / "valid_nut_20220417_094712.jpg", "Рисунок 3 – Кадр valid, гайка сбоку", 12)
    add_text(doc, "На рисунке 3 в разметке одна гайка. Сеть дала две рамки класса nut. Рамка с уверенностью 0,750 совпала с разметкой при IoU 0,732 и заняла этот объект. Вторая рамка, уверенность 0,670, пересекает ту же разметку при IoU 0,761, но объект уже сопоставлен, поэтому она считается лишней. На рисунке лишняя рамка шире и выступает влево. Итог кадра: одно верное обнаружение и одна лишняя рамка.")
    add_picture(doc, RES / "frames" / "valid_nut_20220417_094716.jpg", "Рисунок 4 – Кадр valid, гайка в три четверти", 12)
    add_text(doc, "На рисунке 4 одна гайка. Предсказание nut с уверенностью 0,692 совпало с разметкой при IoU 0,640. Лишних рамок и пропусков нет.")
    add_picture(doc, RES / "frames" / "valid_washer_20220417_094841.jpg", "Рисунок 5 – Кадр valid, шайба", 12)
    add_text(doc, "На рисунке 5 одна шайба. Предсказание washer с уверенностью 0,906 совпало с разметкой при IoU 0,813. Лишних рамок и пропусков нет.")
    add_text(doc, "По трём кадрам вместе: 3 верных обнаружения, 1 лишняя рамка, 0 пропусков. Это согласуется с подсчётом при пороге 0,55 в лабораторной работе №7.")

    add_heading(doc, "9 Выбор модели и файлы для следующей работы", break_before=True)
    add_text(doc, "Модель выбирается по mAP@0,5:0,95 на valid. Если разница не больше 0,01, берётся опыт с меньшим входом. Здесь разница 0,116, поэтому выбран опыт A: YOLOv8n, imgsz = 416, файл весов best.pt объёмом 6,470 МБ. На valid у этого файла mAP@0,5 = 0,995 и mAP@0,5:0,95 = 0,659.")
    add_text(doc, "В лабораторную работу №7 передаются эти веса, вход 416, тот же data.yaml и то же разбиение. Test по-прежнему не использован для выбора.")

    add_heading(doc, "10 Влияние фактора и ограничения", break_before=True)
    add_text(doc, "На этой версии набора больший вход не повысил строгую метрику. Опыт B на valid дал mAP@0,5 = 0,829 и mAP@0,5:0,95 = 0,543 против 0,995 и 0,659 у опыта A. К эпохе 40 журнал опыта B снизился с 0,543 до 0,494. Время опыта B при этом меньше, 31,80 с против 35,78 с, но выбор сделан по качеству valid, а не по времени.")
    add_text(doc, "Ограничение результата — объём проверки. Valid содержит три кадра и три объекта, train — девять кадров и одну шайбу. Разница 0,116 получена на этой короткой выборке. Широкая рамка на рисунке 3 уже дала лишнее обнаружение при рабочем пороге. Переносить вывод о преимуществе входа 416 на другие детали и на другие условия съёмки нельзя.")

    add_heading(doc, "ЗАКЛЮЧЕНИЕ", center=True, break_before=True)
    add_text(doc, "Сравнены входы 416 и 640 при общих 40 эпохах и выключенных аугментациях. По mAP@0,5:0,95 на valid выбран вход 416: 0,659 против 0,543, разница 0,116. Веса best.pt этого опыта переданы в лабораторную работу №7. Test в выборе не участвовал.")
    doc.save(str(OUT6))


def build_lr7():
    doc = fresh(7, "Тестирование нейронной сети на видеокадрах производственных объектов")
    add_heading(doc, "СОДЕРЖАНИЕ", center=True)
    for title, token in H7.items():
        add_toc(doc, title, token)

    add_heading(doc, "ВВЕДЕНИЕ", center=True, break_before=True)
    add_text(doc, "Работа проверяет детектор, выбранный в лабораторной работе №6, на отложенных кадрах. Порог уверенности фиксируется по valid. Test в подборе порога не участвует.")

    add_heading(doc, "1 Цель, вариант, объект и классы", break_before=True)
    add_text(doc, "Цель работы — проверить выбранный детектор и зафиксировать режим, в котором получены оценки. Номер варианта — 1. Объект и классы те же, что в лабораторных работах №5 и №6: гайка nut и шайба washer. Индивидуальная часть — подсчёт TP, FP, FN, точности, полноты и F1 отдельно по классам и указание наиболее частой ошибки.")

    add_heading(doc, "2 Модель и версия набора", break_before=True)
    add_text(doc, "Используются веса best.pt опыта A, сеть YOLOv8n, вход imgsz = 416. На valid у этого файла mAP@0,5 = 0,995 и mAP@0,5:0,95 = 0,659. Версия набора — 14 кадров 1600×1200. Train — 9 кадров, valid — 3, test — 2. Подготовка кадра — копия без изменения размера. Эти сведения совпадают с выбором лабораторной работы №6 и во время теста не меняются.")

    add_heading(doc, "3 Подбор порога на valid", break_before=True)
    add_text(doc, "Рабочий порог выбирается на valid среди значений 0,15, 0,35 и 0,55. Сначала берётся максимум F1, при равном F1 — большая точность, при равной точности — больший порог. Подавление пересечений выполняется при iou = 0,7. Совпадение с разметкой требует IoU не ниже 0,5 и того же класса.")
    add_text(doc, "На трёх кадрах valid три объекта. Результат приведён в таблице 1. При пороге 0,15 полнота равна 1, но лишних рамок четыре, точность 0,429 и F1 0,600. При порогах 0,35 и 0,55 суммы одинаковы: 3 верных обнаружения, 1 лишняя рамка, 0 пропусков, точность 0,750, полнота 1, F1 0,857. По правилу равенства выбран больший порог 0,55.")
    add_table(
        doc,
        "Таблица 1 – Подбор порога на valid",
        ["Порог", "TP", "FP", "FN", "P", "R", "F1"],
        [
            ["0,15", "3", "4", "0", "0,429", "1,000", "0,600"],
            ["0,35", "3", "1", "0", "0,750", "1,000", "0,857"],
            ["0,55", "3", "1", "0", "0,750", "1,000", "0,857"],
        ],
        [1336, 1336, 1336, 1336, 1337, 1337, 1337],
    )
    add_text(doc, "Дальше везде используется порог 0,55. На valid он не менялся после просмотра test.")

    add_heading(doc, "4 Источник test", break_before=True)
    add_text(doc, "Test — два последних по времени кадра той же съёмки 17 апреля 2022 года: гайка в 09:47:19 и шайба в 09:48:45. Оба кадра 1600×1200, в обучении и в valid их не было. На гайке одна рамка класса nut, на шайбе одна рамка класса washer. Архив The Fasteners содержит отдельные фотографии, отдельного ролика с этими двумя деталями в нём нет, поэтому вторая видеопроверка на новом ролике здесь не строится.")

    add_heading(doc, "5 Предобработка и координаты", break_before=True)
    add_text(doc, "Кадр читается в том же размере, в каком стоит разметка, 1600×1200. Функция подготовки возвращает копию, как в лабораторной работе №6. Сеть внутри вызова приводит изображение к стороне 416. Рамки предсказания переводятся обратно в пиксели исходного кадра, и IoU считается уже в этих пикселях. Порог, вход и способ подготовки на test те же, что зафиксированы по valid.")
    add_listing(doc, "Листинг 1 – Подготовка кадра", LISTING_PREP)

    add_heading(doc, "6 Сопоставление рамок", break_before=True)
    add_text(doc, "Предсказания сортируются по уверенности. Каждое получает разметку того же класса с наибольшим IoU, если этот объект ещё свободен и IoU не ниже 0,5. Такое совпадение — TP. Предсказание без пары — FP. Разметка без пары — FN. Один объект разметки учитывается один раз. Листинг 2 содержит это сопоставление.")
    add_listing(doc, "Листинг 2 – Сопоставление рамок", LISTING_MATCH)
    add_text(doc, "Проверка на трёх простых случаях прошла. Одна рамка поверх единственной разметки того же класса даёт (TP, FP, FN) = (1, 0, 0). Та же пара плюс второе предсказание даёт (1, 1, 0). Предсказание другого класса даёт (0, 1, 1).")

    add_heading(doc, "7 Метрики test", break_before=True)
    add_text(doc, "Точность равна TP/(TP+FP), а при отсутствии предсказаний — 0. Полнота и F1 считаются по сумме TP, FP и FN по кадрам, а не как среднее кадров. На двух кадрах test при пороге 0,55 результат приведён в таблице 2: 1 верное обнаружение, 0 лишних рамок, 1 пропуск. Точность равна 1, полнота — 0,500, F1 — 0,667.")
    add_table(
        doc,
        "Таблица 2 – Test при пороге 0,55",
        ["Выборка", "Кадры", "TP", "FP", "FN", "P", "R", "F1"],
        [["test", "2", "1", "0", "1", "1,000", "0,500", "0,667"]],
        [1200, 1100, 1000, 1000, 1000, 1350, 1350, 1355],
    )
    add_text(doc, "Отдельный расчёт mAP на test выполняется при conf = 0,001: iou = 0,7 здесь относится к подавлению дублирующихся предсказаний, а сама метрика mAP получается из сравнения оставшихся рамок с эталоном при порогах IoU от 0,5 до 0,95. На test это даёт mAP@0,5 = 0,746 и mAP@0,5:0,95 = 0,250. В таблицу 2 эти величины не подставляются: там рабочий порог 0,55.")

    add_heading(doc, "8 Разбор кадров", break_before=True)
    add_text(doc, "Разобраны все три кадра valid и оба кадра test. Цвета те же: зелёный — совпадение, красный — лишняя рамка, оранжевый — пропуск.")
    add_picture(doc, RES / "frames" / "valid_nut_20220417_094712.jpg", "Рисунок 1 – Valid, гайка сбоку", 12)
    add_text(doc, "На рисунке 1 две рамки nut. Совпадение имеет уверенность 0,750 и IoU 0,732. Вторая рамка, уверенность 0,670, пересекает ту же разметку при IoU 0,761. Она шире и выступает влево. В TP она не входит: единственный объект уже занят первой рамкой. Это единственная лишняя рамка valid.")
    add_picture(doc, RES / "frames" / "valid_nut_20220417_094716.jpg", "Рисунок 2 – Valid, гайка в три четверти", 12)
    add_text(doc, "На рисунке 2 гайка найдена одной рамкой, уверенность 0,692, IoU 0,640.")
    add_picture(doc, RES / "frames" / "valid_washer_20220417_094841.jpg", "Рисунок 3 – Valid, шайба", 12)
    add_text(doc, "На рисунке 3 шайба найдена, уверенность 0,906, IoU 0,813.")
    add_picture(doc, RES / "frames" / "test_nut_20220417_094719.jpg", "Рисунок 4 – Test, гайка спереди", 12)
    add_text(doc, "На рисунке 4 гайка снята со стороны торца. При пороге 0,55 предсказания нет. Оранжевая рамка — разметка, её площадь 0,488 кадра. Это единственный пропуск test.")
    add_picture(doc, RES / "frames" / "test_washer_20220417_094845.jpg", "Рисунок 5 – Test, шайба", 12)
    add_text(doc, "На рисунке 5 шайба найдена, уверенность 0,755, IoU 0,676. Лишней рамки нет.")

    add_heading(doc, "9 Время обработки", break_before=True)
    add_text(doc, "Время измерено на двух кадрах test. Сначала выполнены 5 прогревочных вызовов, они в таблицу не входят. Затем каждый кадр прогнан 15 раз, всего 30 проходов. Цикл включает чтение кадра, копию и вызов детектора. В таблице 3 среднее, медиана и 95-й процентиль.")
    add_table(
        doc,
        "Таблица 3 – Время на кадрах test, мс",
        ["Участок", "Среднее", "Медиана", "95-й процентиль"],
        [
            ["predict", "11,08", "10,43", "14,78"],
            ["Полный цикл", "21,47", "21,25", "25,16"],
        ],
        [2338, 2339, 2339, 2339],
    )
    add_text(doc, "Для потока 25 кадров/с на один кадр остаётся 40 мс. Измеренный средний цикл 21,47 мс и 95-й процентиль 25,16 мс оба укладываются в этот интервал. Измерение сделано повторными проходами двух фотографий, а не по длинному ролику.")

    add_heading(doc, "10 Оценки по классам", break_before=True)
    add_text(doc, "По классам на test при пороге 0,55 результат приведён в таблице 4. Шайба найдена: TP = 1, FP = 0, FN = 0, точность, полнота и F1 равны 1. Гайка пропущена: TP = 0, FP = 0, FN = 1. Предсказаний класса nut нет, поэтому точность равна 0. Полнота и F1 тоже равны 0.")
    add_table(
        doc,
        "Таблица 4 – Классы на test при пороге 0,55",
        ["Класс", "TP", "FP", "FN", "P", "R", "F1"],
        [
            ["nut", "0", "0", "1", "0,000", "0,000", "0,000"],
            ["washer", "1", "0", "0", "1,000", "1,000", "1,000"],
        ],
        [1455, 1200, 1200, 1200, 1433, 1433, 1434],
    )
    add_text(doc, "На valid при том же пороге гайка имеет TP = 2, FP = 1, FN = 0, точность 0,667, полноту 1 и F1 0,800. Шайба на valid имеет TP = 1, FP = 0, FN = 0 и единичные точность, полноту и F1. Наиболее частая ошибка на test — пропуск гайки на фронтальном кадре. На valid к этому добавляется одна лишняя рамка гайки внутри широкой разметки. Ошибки класса washer на этих пяти кадрах нет.")

    add_heading(doc, "11 Файлы результата и режим", break_before=True)
    add_text(doc, "Режим, при котором получены оценки: веса best.pt опыта A, imgsz = 416, conf = 0,55, iou подавления = 0,7, порог IoU сопоставления = 0,5, устройство CUDA 0 на NVIDIA GeForce RTX 2060, подготовка кадра без изменения размера. Порог взят по valid и на test не менялся.")
    add_text(doc, "Разбор кадров сохранён рисунками 1–5. Отдельный файл ролика с наложенными рамками не формировался: в версии набора нет видеозаписи этих деталей.")

    add_heading(doc, "12 Пригодность модели и следующие опыты", break_before=True)
    add_text(doc, "На трёх кадрах valid при пороге 0,55 F1 равен 0,857: все три объекта найдены, одна рамка гайки лишняя. На двух кадрах test F1 равен 0,667: шайба найдена, фронтальная гайка пропущена. mAP@0,5:0,95 на test равен 0,250. Средний цикл 21,47 мс ниже интервала 40 мс, отведённого кадру при 25 кадрах/с.")
    add_text(doc, "Этого мало, чтобы считать детектор готовым к смене ракурса. В обучении одна шайба и восемь гаек, фронтальный вид гайки оставлен в test и при пороге 0,55 не обнаруживается. Следующий опыт стоит пополнить фронтальными кадрами той же гайки и сузить боковые рамки до металла, затем заново зафиксировать порог по valid до просмотра test.")

    add_heading(doc, "ЗАКЛЮЧЕНИЕ", center=True, break_before=True)
    add_text(doc, "Проверена модель лабораторной работы №6 с входом 416. По valid зафиксирован порог 0,55. На test F1 равен 0,667, mAP@0,5:0,95 равен 0,250. Шайба найдена, гайка на фронтальном кадре пропущена. Средний цикл обработки 21,47 мс.")
    doc.save(str(OUT7))


def main():
    TEMPLATE.parent.mkdir(parents=True, exist_ok=True)
    if not TEMPLATE.exists():
        shutil.copyfile(SRC, TEMPLATE)
    needed = [
        TEMPLATE,
        RES / "mark_nut_top.jpg",
        RES / "mark_nut_side.jpg",
        RES / "mark_washer.jpg",
        RES / "frames" / "valid_nut_20220417_094712.jpg",
        RES / "frames" / "valid_nut_20220417_094716.jpg",
        RES / "frames" / "valid_washer_20220417_094841.jpg",
        RES / "frames" / "test_nut_20220417_094719.jpg",
        RES / "frames" / "test_washer_20220417_094845.jpg",
        RUNS / "v01_a" / "results.png",
        RUNS / "v01_b" / "results.png",
    ]
    missing = [str(path) for path in needed if not path.exists()]
    if missing:
        raise SystemExit("Нет файлов:\n" + "\n".join(missing))
    build_lr5()
    build_lr6()
    build_lr7()
    import win32com.client

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        pages5 = fill_pages(word, OUT5, H5)
        pages6 = fill_pages(word, OUT6, H6)
        pages7 = fill_pages(word, OUT7, H7)
    finally:
        word.Quit()
    write_page_numbers(OUT5, H5, pages5)
    write_page_numbers(OUT6, H6, pages6)
    write_page_numbers(OUT7, H7, pages7)
    print("LR5", pages5)
    print("LR6", pages6)
    print("LR7", pages7)


if __name__ == "__main__":
    main()
