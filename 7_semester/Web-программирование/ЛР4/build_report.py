# -*- coding: utf-8 -*-
"""Сборка отчёта ЛР4. В текст отчёта не входит."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(r"C:\Users\skvor\OneDrive\Desktop\GitHub\study-guap\7_semester\Web-программирование")
LAB = ROOT / "ЛР4"
SRC_DOC = ROOT / "ЛР1_Отчет_Буренков_PairWise.docx"
OUT = LAB / "ЛР4_Отчет_Буренков_PairWise.docx"
SHOTS = Path(r"C:\Users\skvor\AppData\Local\Temp\cursor\screenshots")
BLACK = RGBColor(0, 0, 0)


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
    if center:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if first:
            set_indent(paragraph, first=first)
    run = paragraph.add_run(text)
    style_run(run, size=size, bold=bold)
    return paragraph


def add_heading(doc, text, *, center=False, break_before=False, bookmark=None, bid=None):
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
    if bookmark is not None:
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), str(bid))
        start.set(qn("w:name"), bookmark)
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), str(bid))
        p_pr = paragraph._p.find(qn("w:pPr"))
        if p_pr is not None:
            p_pr.addnext(start)
        else:
            paragraph._p.insert(0, start)
        paragraph._p.append(end)
    return paragraph


def field_run(paragraph, kind, payload=None):
    run = paragraph.add_run()
    style_run(run, bold=False)
    for child in list(run._r):
        if child.tag == qn("w:t"):
            run._r.remove(child)
    if kind == "text":
        node = OxmlElement("w:t")
        node.text = payload
        run._r.append(node)
        return
    if kind == "instr":
        node = OxmlElement("w:instrText")
        node.set(qn("xml:space"), "preserve")
        node.text = payload
    else:
        node = OxmlElement("w:fldChar")
        node.set(qn("w:fldCharType"), kind)
    run._r.append(node)


def add_toc(doc, title, bookmark):
    paragraph = doc.add_paragraph()
    set_spacing(paragraph, line=360)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    ppr = paragraph._p.get_or_add_pPr()
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:leader"), "dot")
    tab.set(qn("w:pos"), "9071")
    tabs.append(tab)
    ppr.append(tabs)
    run = paragraph.add_run(title)
    style_run(run, bold=False)
    paragraph.add_run().add_tab()
    field_run(paragraph, "begin")
    field_run(paragraph, "instr", f" PAGEREF {bookmark} \\h ")
    field_run(paragraph, "separate")
    field_run(paragraph, "text", "1")
    field_run(paragraph, "end")
    return paragraph


def shade_cell(cell):
    tc = cell._tc
    tcpr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "F2F2F2")
    tcpr.append(shd)


def set_cell(cell, text, *, bold=False, header=False, size=14):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    set_spacing(paragraph, before=2, after=2, line=240)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.first_line_indent = Cm(0)
    run = paragraph.add_run(text)
    style_run(run, size=size, bold=bold)
    tc = cell._tc
    tcpr = tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for edge in ("top", "left", "bottom", "right"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:w"), "60")
        node.set(qn("w:type"), "dxa")
        mar.append(node)
    tcpr.append(mar)
    if header:
        shade_cell(cell)


def set_table_widths(table, widths):
    table.autofit = False
    table.allow_autofit = False
    tbl = table._tbl
    tblpr = tbl.tblPr
    tblw = tblpr.find(qn("w:tblW"))
    if tblw is None:
        tblw = OxmlElement("w:tblW")
        tblpr.append(tblw)
    total = sum(widths)
    tblw.set(qn("w:w"), str(total))
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
            tc = cell._tc
            tcpr = tc.get_or_add_tcPr()
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
    jc = tblpr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        tblpr.append(jc)
    jc.set(qn("w:val"), "left")


def add_table(doc, caption, headers, rows, widths, size=14):
    content = 9355
    inset = max(0, (content - sum(widths)) // 2)
    add_text(doc, caption, first=0, before=6, after=6, line=240, center=False, size=14)
    cap = doc.paragraphs[-1]
    cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
    cap.paragraph_format.keep_with_next = True
    set_indent(cap, first=0, left=inset)
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    set_table_widths(table, widths)
    for i, header in enumerate(headers):
        set_cell(table.rows[0].cells[i], header, bold=True, header=True, size=size)
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            set_cell(table.rows[r + 1].cells[c], value, size=size)
    return table


def add_picture(doc, path, caption, width_cm):
    paragraph = doc.add_paragraph()
    set_spacing(paragraph, before=6, after=0, line=240)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    run.add_picture(str(path), width=Cm(width_cm))
    add_text(doc, caption, first=0, before=6, after=10, line=360, center=True)


def add_code(doc, caption, text):
    add_text(doc, caption, first=0, before=10, after=2, line=360)
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.LEFT
    doc.paragraphs[-1].paragraph_format.keep_with_next = True
    for line in text.strip("\n").splitlines():
        paragraph = doc.add_paragraph()
        set_spacing(paragraph, line=240)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        paragraph.paragraph_format.keep_together = True
        run = paragraph.add_run(line if line else " ")
        style_run(run, size=10, bold=False, font="Courier New")
        node = run._element.find(qn("w:t"))
        if node is not None:
            node.set(qn("xml:space"), "preserve")


def add_qa(doc, question, answer):
    paragraph = doc.add_paragraph()
    set_spacing(paragraph, line=360)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_indent(paragraph, first=709)
    q = paragraph.add_run(question + " ")
    style_run(q, bold=True)
    a = paragraph.add_run(answer)
    style_run(a, bold=False)


def replace_title(doc):
    replacements = {
        "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 1": "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 4",
        "Создание проекта Next.js с использованием React, TypeScript и App Router":
            "Реализация адаптивности клиентской части и контроль качества с помощью Lighthouse",
    }
    found = set()
    for node in doc.element.body.iter(qn("w:t")):
        if node.text in replacements:
            node.text = replacements[node.text]
            found.add(node.text)
    expected = set(replacements.values())
    if found != expected:
        raise SystemExit(f"Титул заменён не полностью: {found}")


def update_page_fields(path):
    import subprocess
    import tempfile

    script = Path(tempfile.gettempdir()) / "lr4_web_pages.ps1"
    target = str(path).replace("'", "''")
    script.write_text(
        "\n".join(
            [
                "$ErrorActionPreference = 'Stop'",
                "$word = New-Object -ComObject Word.Application",
                "$word.Visible = $false",
                "$word.DisplayAlerts = 0",
                f"$doc = $word.Documents.Open('{target}')",
                "$doc.Repaginate()",
                "$doc.Fields.Update() | Out-Null",
                "$doc.Repaginate()",
                "$doc.Save()",
                "$doc.Close()",
                "$word.Quit()",
            ]
        ),
        encoding="utf-8-sig",
    )
    subprocess.run(["powershell", "-NoProfile", "-File", str(script)], check=True)


def build():
    pictures = {
        "wide": SHOTS / "gallery-1440-chrome.png",
        "narrow": SHOTS / "gallery-375.png",
        "modal": SHOTS / "modal-375-crop.png",
        "menu": SHOTS / "menu-375-crop.png",
        "map": SHOTS / "map-375.png",
        "zoom": SHOTS / "gallery-720.png",
        "lh_base": SHOTS / "lh-base.png",
        "lh_final": SHOTS / "lh-final.png",
    }
    for path in pictures.values():
        if not path.exists():
            raise SystemExit(f"Нет снимка: {path}")

    doc = Document(str(SRC_DOC))
    body = doc.element.body
    children = list(body)
    for child in children[12:-1]:
        body.remove(child)
    replace_title(doc)

    add_heading(doc, "СОДЕРЖАНИЕ", center=True, bookmark="toc", bid=20)
    toc = [
        ("1 Цель работы", "h1"),
        ("2 Исходное состояние", "h2"),
        ("3 Контрольные точки", "h3"),
        ("4 Viewport", "h4"),
        ("5 Галерея и размеры изображений", "h5"),
        ("6 Модальное окно", "h6"),
        ("7 Бургер-меню", "h7"),
        ("8 Карта и уменьшение движения", "h8"),
        ("9 Ручная проверка ширин", "h9"),
        ("10 Исходный Lighthouse", "h10"),
        ("11 Исправления", "h11"),
        ("12 Итоговое сравнение", "h12"),
        ("13 Ответы на контрольные вопросы", "h13"),
        ("ЗАКЛЮЧЕНИЕ", "h14"),
    ]
    for title, bookmark in toc:
        add_toc(doc, title, bookmark)

    add_heading(doc, "1 Цель работы", break_before=True, bookmark="h1", bid=21)
    add_text(
        doc,
        "Цель работы — сделать клиентскую часть PairWise устойчивой к разной ширине окна и проверить страницу галереи в Lighthouse. Сохраняются галерея площадок, модальное окно, бургер-меню, учебное состояние пользователя и карта Яндекса из предыдущей работы.",
    )

    add_heading(doc, "2 Исходное состояние", break_before=True, bookmark="h2", bid=22)
    add_text(
        doc,
        "К началу работы страница /gallery уже показывала шесть карточек, диалог площадки, меню и карту. Сетка была задана как repeat(auto-fit, minmax(15rem, 1fr)): колонка не умела стать уже 15rem, поэтому на ширине меньше 15rem страница получала горизонтальную прокрутку. У изображений стояло sizes=\"(max-width: 700px) 100vw, 33vw\", то есть оценка колонки была привязана к 700px, а не к шагу сетки. Высота карты была фиксированной, 30rem. Панель меню имела ширину min(26rem, 100%) и min-height: 100%, без прокрутки на высоту видимого экрана. Тень карточки включалась по :hover без проверки типа указателя. В корневых стилях не было правила prefers-reduced-motion. Экспорт viewport в layout.tsx отсутствовал, файл favicon.ico тоже.",
    )

    add_heading(doc, "3 Контрольные точки", break_before=True, bookmark="h3", bid=23)
    add_text(
        doc,
        "Точки выбраны по содержимому PairWise. Карточка читается от 15rem: название и три строки описания помещаются, фото держит пропорцию 3/2. Контейнер страницы ограничен 75rem. Панель меню рассчитана на 26rem: уже этой ширины список разделов остаётся одной колонкой, шире — панель не должна закрывать всю страницу.",
    )
    add_text(
        doc,
        "48rem — это ширина, на которой в контейнере с полями помещается вторая колонка карточек. С этого порога поля контейнера увеличиваются с 1rem до 1.5rem, а отступ подложки диалога — с 1rem до 1.5rem.",
    )
    add_text(
        doc,
        "40rem — порог меню и радиуса карты. Панель 26rem на экране уже 40rem оставляет рядом меньше одной карточки 15rem, поэтому панель занимает 100% ширины, а скругление карты уменьшается с 1rem до 0.5rem. 32rem по высоте в альбомной ориентации — порог диалога: фото 3/2 вместе с заголовком и текстом уже не входит в короткий экран, поэтому высота фото ограничена 40dvh, а диалог получает max-height: calc(100dvh - 1rem).",
    )

    add_heading(doc, "4 Viewport", break_before=True, bookmark="h4", bid=24)
    add_text(
        doc,
        "Next.js сам выводит метатег viewport. В собранной странице его значение равно width=device-width, initial-scale=1. Масштаб пользователя не ограничен: maximum-scale и user-scalable в метатеге нет. В экспорт viewport попали только цвет темы и цветовая схема, как в листинге 1. На ширине 320 CSS px строка метатега проверена в браузере и совпала с этим значением.",
    )
    add_code(
        doc,
        "Листинг 1 — Экспорт viewport в app/layout.tsx",
        """
export const viewport: Viewport = {
    themeColor: "#111111",
    colorScheme: "light",
};
""",
    )
    add_text(
        doc,
        "В globals.css для всех элементов сохранён box-sizing: border-box. У img, video и svg стоит max-width: 100%. Контейнер имеет ширину min(75rem, 100%), чтобы не вылезать за окно. Горизонтальный overflow у страницы не скрывается: прокрутка, если она появится, должна остаться видимой.",
    )
    add_code(
        doc,
        "Листинг 2 — Ширина контейнера и шаг полей",
        """
.container {
    width: min(75rem, 100%);
    margin: 0 auto;
    padding-right: 1rem;
    padding-left: 1rem;
}

@media (width >= 48rem) {
    .container {
        padding-right: 1.5rem;
        padding-left: 1.5rem;
    }
}
""",
    )

    add_heading(doc, "5 Галерея и размеры изображений", break_before=True, bookmark="h5", bid=25)
    add_text(
        doc,
        "Сетка мобильная: сначала одна колонка, дополнительные колонки появляются, когда для них есть место. Формула minmax(min(100%, 15rem), 1fr) оставляет желаемую ширину колонки 15rem и разрешает ей сжаться до ширины контейнера. На 320 CSS px получилась одна колонка, scrollWidth совпал с clientWidth и был равен 320. Текст карточки переносится через overflow-wrap: break-word. Тень при наведении включена только внутри (hover: hover) and (pointer: fine). Обводка :focus-visible остаётся на всех ширинах. Область фото держит aspect-ratio: 3 / 2 и object-fit: cover.",
    )
    add_code(
        doc,
        "Листинг 3 — Сетка галереи",
        """
.gallery__grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(min(100%, 15rem), 1fr));
    gap: 1.125rem;
}

@media (hover: hover) and (pointer: fine) {
    .gallery__card:hover {
        box-shadow: 0 8px 24px var(--color-shadow-hover);
    }
}
""",
    )
    add_text(
        doc,
        "У Image с fill атрибут sizes оценивает ширину колонки, чтобы браузер выбрал источник из srcset. Ниже 48rem колонка одна, поэтому берётся 100vw. От 48rem в контейнере две или три колонки, оценка колонки 31vw. Когда контейнер упирается в 75rem, карточка около 18rem. Первая карточка, она же элемент LCP, загружается сразу: preload и fetchPriority=\"high\". Остальные карточки остаются с обычным приоритетом.",
    )
    add_code(
        doc,
        "Листинг 4 — sizes и приоритет первой фотографии",
        """
<Image
    src={spot.image}
    alt={spot.imageAlt}
    fill
    preload={index === 0}
    fetchPriority={index === 0 ? "high" : "auto"}
    sizes="(min-width: 75rem) 18rem, (min-width: 48rem) 31vw, 100vw"
/>
""",
    )

    add_heading(doc, "6 Модальное окно", break_before=True, bookmark="h6", bid=26)
    add_text(
        doc,
        "Диалог имеет ширину min(45rem, 100%) и max-height: calc(100dvh - 2rem), содержимое прокручивается. Подложка даёт поле 1rem, с 48rem — 1.5rem, и тогда высота диалога считается уже от calc(100dvh - 3rem). Ряд действия переносится через flex-wrap. Кнопка закрытия остаётся в правом верхнем углу. В альбомной ориентации при высоте не больше 32rem фото ограничено 40dvh. На ширине 375 CSS px диалог имел ширину 343 px и помещался в окно, фокус сразу стоял на кнопке «Закрыть окно». В альбомном окне 700×320 CSS px высота диалога была 304 px при высоте окна 320 px, max-height фото — 128 px, горизонтальной прокрутки не было. Клавиатурная модель окна из предыдущей работы сохранена: Escape, цикл Tab и возврат фокуса на карточку.",
    )
    add_code(
        doc,
        "Листинг 5 — Высота диалога",
        """
.modal__dialog {
    width: min(45rem, 100%);
    max-height: calc(100dvh - 2rem);
    overflow-y: auto;
}

.modal__actions {
    display: flex;
    flex-wrap: wrap;
    gap: 0.75rem;
}

@media (orientation: landscape) and (height <= 32rem) {
    .modal__dialog {
        max-height: calc(100dvh - 1rem);
    }

    .modal__media {
        max-height: 40dvh;
    }
}
""",
    )
    add_text(
        doc,
        "У фото внутри окна sizes=\"(min-width: 48rem) 42rem, 100vw\": на узком экране источник близок к ширине окна, на широком — к ширине диалога 45rem с учётом внутренних полей.",
    )

    add_heading(doc, "7 Бургер-меню", break_before=True, bookmark="h7", bid=27)
    add_text(
        doc,
        "Панель — колонка flex. На широком экране её ширина min(26rem, 100%), она прижата к правому краю. На ширине не больше 40rem ширина становится 100%. Высота и max-height равны 100dvh, длинный список прокручивается. На 375 CSS px панель измерилась как 375×812 px, то есть на весь видимый экран. После открытия фокус стоял на кнопке «Закрыть меню», у гостя внизу видна ссылка «Войти». Атрибуты aria-expanded, aria-controls, role=\"dialog\" и aria-modal, а также закрытие по Escape, подложке и Tab сохранены.",
    )
    add_code(
        doc,
        "Листинг 6 — Ширина и высота панели меню",
        """
.menu__panel {
    display: flex;
    flex-direction: column;
    width: min(26rem, 100%);
    height: 100dvh;
    max-height: 100dvh;
    margin-left: auto;
    overflow-y: auto;
}

@media (width <= 40rem) {
    .menu__panel {
        width: 100%;
    }
}
""",
    )

    add_heading(doc, "8 Карта и уменьшение движения", break_before=True, bookmark="h8", bid=28)
    add_text(
        doc,
        "Высота карты — clamp(14rem, 48dvh, 30rem). На окне высотой 700 CSS px это 336 px, на 812 CSS px — 390 px, на 900 CSS px — 432 px, то есть предпочитаемое значение 48dvh, пока оно между 14rem и 30rem. После создания карты ResizeObserver вызывает map.container.fitToViewport() при изменении контейнера. В cleanup наблюдатель отключается, карта уничтожается. При сужении с 1440 до 375 CSS px внутренний слой карты был 341×388 px при контейнере 343×390 px, серой полосы не осталось, метки на месте. Скругление на узкой ширине 0.5rem, на 768 CSS px и шире — 1rem.",
    )
    add_code(
        doc,
        "Листинг 7 — Наблюдатель размера карты",
        """
observer = new ResizeObserver(() => {
    map?.container.fitToViewport();
});
observer.observe(container);
map.container.fitToViewport();

return () => {
    observer?.disconnect();
    labels?.disconnect();
    map?.destroy();
};
""",
    )
    add_text(
        doc,
        "У div в общих стилях есть переход фона и цвета. При prefers-reduced-motion: reduce длительности переходов и анимаций сокращаются до 0.01 ms, прокрутка становится мгновенной. Смена состояния, например открытие окна, при этом остаётся.",
    )
    add_code(
        doc,
        "Листинг 8 — prefers-reduced-motion",
        """
@media (prefers-reduced-motion: reduce) {
    *,
    *::before,
    *::after {
        transition-duration: 0.01ms;
        animation-duration: 0.01ms;
        animation-iteration-count: 1;
        scroll-behavior: auto;
    }
}
""",
    )

    add_heading(doc, "9 Ручная проверка ширин", break_before=True, bookmark="h9", bid=29)
    add_text(
        doc,
        "Проверка выполнена в Chrome на странице /gallery. Считались колонки сетки, scrollWidth и clientWidth, размер карты и фокус. Горизонтальной прокрутки не было ни на одной ширине из таблицы 1. Текст карточек читается, кнопки шапки помещаются, фокус диалога и меню виден.",
    )
    add_table(
        doc,
        "Таблица 1 — Ручная проверка ширин",
        ["Ширина", "Колонки", "Прокрутка", "Карта", "Меню"],
        [
            ["320 CSS px", "1", "нет", "288×336 px, радиус 0,5rem", "на всю ширину"],
            ["375 CSS px", "1", "нет", "343×390 px, радиус 0,5rem", "375×812 px"],
            ["768 CSS px", "2", "нет", "705×432 px, радиус 1rem", "до 26rem"],
            ["1024 CSS px", "3", "нет", "961×432 px, радиус 1rem", "до 26rem"],
            ["1440 CSS px", "4", "нет", "1152×432 px, радиус 1rem", "до 26rem"],
        ],
        [1500, 1200, 1400, 2855, 2400],
    )
    add_text(
        doc,
        "Масштаб 200% перестраивает страницу: окно 1440 CSS px при таком масштабе даёт область вёрстки около 720 CSS px. На этой ширине колонок две, scrollWidth не больше clientWidth. Страница на 1440 CSS px показана на рисунке 1, на 375 CSS px — на рисунке 2. Диалог и меню на узком экране — на рисунках 3 и 4. Карта после сужения до 375 CSS px — на рисунке 5. Раскладка области 720 CSS px — на рисунке 6.",
    )
    add_picture(doc, pictures["wide"], "Рисунок 1 — Галерея на ширине 1440 CSS px", 16)
    add_picture(doc, pictures["narrow"], "Рисунок 2 — Галерея на ширине 375 CSS px", 8)
    add_picture(doc, pictures["modal"], "Рисунок 3 — Модальное окно на ширине 375 CSS px", 8)
    add_picture(doc, pictures["menu"], "Рисунок 4 — Открытое бургер-меню на ширине 375 CSS px", 8)
    add_picture(doc, pictures["map"], "Рисунок 5 — Карта после сужения окна до 375 CSS px", 8)
    add_picture(doc, pictures["zoom"], "Рисунок 6 — Раскладка при масштабе 200%, область 720 CSS px", 14)

    add_heading(doc, "10 Исходный Lighthouse", break_before=True, bookmark="h10", bid=30)
    add_text(
        doc,
        "Дата проверки — 31 октября 2026 года. Браузер — Chrome 147.0.7727.138. Lighthouse 13.6.0, режим Navigation, устройство Mobile, категории Performance, Accessibility, Best Practices и SEO. Троттлинг — simulated, экран эмуляции 412×823 CSS px, deviceScaleFactor 1,75. Адрес — http://localhost:3000/gallery, сервер production-сборки next start. Хранилище для первого захода очищается самим режимом Navigation. Один и тот же набор флагов использован для всех шести запусков. Проходной балл заранее не задавался.",
    )
    add_text(
        doc,
        "Три исходных запуска сделаны до правок этого раздела. Медиана Performance равна 88, Accessibility — 96, Best Practices — 73, SEO — 100. Медиана LCP — 3,1 с, TBT — 250 мс, CLS — 0. Числа сведены в таблице 2. Рисунок 7 — второй запуск: Performance 89, Accessibility 96, Best Practices 73, SEO 100, LCP 3,1 с, TBT 240 мс, CLS 0.",
    )
    add_table(
        doc,
        "Таблица 2 — Исходные запуски Lighthouse",
        ["Запуск", "Произв.", "Доступн.", "Практики", "SEO", "LCP", "TBT", "CLS"],
        [
            ["1", "87", "96", "73", "100", "3,1 с", "280 мс", "0"],
            ["2", "89", "96", "73", "100", "3,1 с", "240 мс", "0"],
            ["3", "88", "96", "73", "100", "3,1 с", "250 мс", "0"],
            ["Медиана", "88", "96", "73", "100", "3,1 с", "250 мс", "0"],
        ],
        [1400, 1100, 1200, 1200, 800, 1100, 1200, 1355],
    )
    add_picture(doc, pictures["lh_base"], "Рисунок 7 — Второй исходный отчёт Lighthouse", 16)

    add_heading(doc, "11 Исправления", break_before=True, bookmark="h11", bid=31)
    add_text(
        doc,
        "Для правки взяты три аудита, которые в исходных запусках были провалены и относились к коду страницы. Сводка — в таблице 3.",
    )
    add_table(
        doc,
        "Таблица 3 — Три подтверждённых исправления",
        ["Аудит", "Причина", "Изменение", "Ожидаемый эффект"],
        [
            [
                "link-name",
                "У логотипа Яндекс Карт внутри виджета не было доступного имени",
                "Пустым ссылкам внутри карты задаётся aria-label «Яндекс Карты»",
                "Аудит link-name проходит, балл Accessibility растёт",
            ],
            [
                "errors-in-console",
                "Браузер запрашивал /favicon.ico и получал 404",
                "В app добавлен favicon.ico",
                "Ошибка 404 из консоли уходит, балл Best Practices растёт",
            ],
            [
                "lcp-discovery",
                "Первое фото было loading=lazy, у preload не было fetchpriority=high",
                "У первой карточки стоят preload и fetchPriority=\"high\"",
                "Чеклист LCP становится истинным, LCP немного короче",
            ],
        ],
        [1700, 2550, 2700, 2405],
    )
    add_text(
        doc,
        "Аудиты third-party-cookies и Issues panel после правок остаются. Их источник — скрипт и тайлы Яндекс Карт. Карта входит в работу, поэтому эти записи не снимались удалением виджета. Стили Next.js по-прежнему попадают в render-blocking requests: это общая таблица стилей приложения, отдельным правилом галереи она не выносится.",
    )

    add_heading(doc, "12 Итоговое сравнение", break_before=True, bookmark="h12", bid=32)
    add_text(
        doc,
        "После правок сборка повторена, сервер снова запущен через next start, три запуска сняты с той же конфигурацией. Результаты — в таблице 4. Рисунок 8 — второй итоговый запуск: Performance 89, Accessibility 100, Best Practices 77, SEO 100, LCP 3,0 с, TBT 250 мс, CLS 0. В списке insights этого отчёта пункта LCP request discovery уже нет.",
    )
    add_table(
        doc,
        "Таблица 4 — Итоговые запуски Lighthouse",
        ["Запуск", "Произв.", "Доступн.", "Практики", "SEO", "LCP", "TBT", "CLS"],
        [
            ["1", "87", "100", "77", "100", "3,2 с", "260 мс", "0"],
            ["2", "89", "100", "77", "100", "3,0 с", "250 мс", "0"],
            ["3", "90", "100", "77", "100", "3,0 с", "250 мс", "0"],
            ["Медиана", "89", "100", "77", "100", "3,0 с", "250 мс", "0"],
        ],
        [1400, 1100, 1200, 1200, 800, 1100, 1200, 1355],
    )
    add_picture(doc, pictures["lh_final"], "Рисунок 8 — Второй итоговый отчёт Lighthouse", 16)
    add_text(
        doc,
        "Медианы до и после собраны в таблице 5. Accessibility выросла на 4 пункта: link-name стал пройденным. Best Practices выросли на 4 пункта: из консоли ушла ошибка favicon. Performance выросла на 1 пункт, LCP сократился с 3,1 с до 3,0 с. TBT по медиане остался 250 мс. CLS остался 0. Первый итоговый запуск дал Performance 87, ниже исходной медианы 88. Это разброс одного прогона: сравниваются медианы трёх запусков при одной конфигурации.",
    )
    add_table(
        doc,
        "Таблица 5 — Медианы до и после",
        ["Показатель", "До", "После"],
        [
            ["Performance", "88", "89"],
            ["Accessibility", "96", "100"],
            ["Best Practices", "73", "77"],
            ["SEO", "100", "100"],
            ["LCP", "3,1 с", "3,0 с"],
            ["TBT", "250 мс", "250 мс"],
            ["CLS", "0", "0"],
        ],
        [4000, 2677, 2678],
    )
    add_text(
        doc,
        "Лабораторный отчёт — один компьютер, один профиль троттлинга и одна загрузка страницы. Полевые данные собираются у реальных пользователей на разных устройствах и сетях. Балл 89 на localhost не описывает посетителей сайта. Opportunities и Diagnostics, например render-blocking requests, показывают возможную экономию. Сам балл Performance считается по метрикам загрузки, среди которых здесь видны LCP, TBT и CLS.",
    )
    add_text(
        doc,
        "Команда npm run lint завершилась без ошибок. Команда npm run build завершилась без ошибок: Next.js 16.3.4, TypeScript проверен, страница /gallery собрана как статическая. После правок ручная проверка ширин, фокуса диалога и меню и масштаба 200% повторена: колонки и отсутствие горизонтальной прокрутки сохранились.",
    )

    add_heading(doc, "13 Ответы на контрольные вопросы", break_before=True, bookmark="h13", bid=33)
    answers = [
        (
            "Что означает responsive web design?",
            "Это вёрстка, которая перестраивается под ширину окна: колонки, отступы и размеры блоков меняются, а содержимое остаётся доступным. В PairWise одна и та же галерея даёт 1, 2, 3 или 4 колонки.",
        ),
        (
            "Для чего нужен viewport meta tag?",
            "Он задаёт ширину layout viewport. Со значением width=device-width браузер телефона верстает страницу в CSS-пикселях экрана, а не в ширине около 980 px с последующим уменьшением.",
        ),
        (
            "Какой viewport Next.js добавляет по умолчанию?",
            "width=device-width, initial-scale=1. В проекте эта строка проверена на собранной странице. В экспорт viewport добавлены только themeColor и colorScheme.",
        ),
        (
            "В чём состоит mobile-first подход?",
            "Базовые стили описывают узкое окно. Media query добавляет изменения, когда места становится больше: вторую колонку, более широкие поля, панель меню не на всю ширину.",
        ),
        (
            "Почему breakpoint выбирают по содержимому, а не по модели телефона?",
            "Модель телефона не говорит, влезает ли карточка. Порог 48rem связан с двумя карточками по 15rem, порог 40rem — с панелью 26rem, порог 32rem по высоте — с фото диалога.",
        ),
        (
            "Чем CSS Grid отличается от Flexbox?",
            "Grid строит двумерные дорожки. Им сверстана сетка карточек. Flexbox выстраивает элементы в одном направлении: колонка карточки, ряд кнопок диалога и колонка панели меню.",
        ),
        (
            "Как работают auto-fit и minmax() в сетке карточек?",
            "auto-fit создаёт столько колонок, сколько помещается. minmax(min(100%, 15rem), 1fr) задаёт нижнюю границу как меньшее из 15rem и ширины контейнера, верхнюю — как долю свободного места. Поэтому на 320 CSS px колонка сжимается и страница не уезжает вбок.",
        ),
        (
            "Для чего используется clamp()?",
            "clamp(минимум, предпочтительно, максимум) держит значение в коридоре. Высота карты clamp(14rem, 48dvh, 30rem) растёт с окном и не становится ниже 14rem и выше 30rem.",
        ),
        (
            "Когда нужна единица dvh?",
            "dvh — высота динамического viewport, с учётом панелей мобильного браузера. Ею заданы высота диалога, высота меню и предпочитаемая высота карты, чтобы блок помещался в видимую область.",
        ),
        (
            "Что проверяют hover и pointer в media query?",
            "hover: hover означает, что устройство умеет наведение. pointer: fine означает точный указатель. Тень карточки включена только при обоих условиях, чтобы сенсорный экран не зависел от состояния hover.",
        ),
        (
            "Для чего применяется prefers-reduced-motion?",
            "Запрос сообщает, что пользователь просит меньше движения. В проекте тогда почти обнуляются transition и animation. Открытие окна и смена страницы остаются.",
        ),
        (
            "Почему Image с fill должен иметь sizes?",
            "fill растягивает картинку на родителя и отдаёт браузеру набор srcset. Без sizes браузер считает ширину близкой к 100vw и может скачать слишком крупный файл. sizes сообщает долю окна, которую фото займёт.",
        ),
        (
            "Как aspect-ratio помогает визуальной стабильности?",
            "Пропорция 3/2 резервирует место до загрузки файла. Соседние блоки не прыгают вниз. В замере CLS равен 0.",
        ),
        (
            "Для чего карте нужен ResizeObserver?",
            "Виджет один раз считает размер контейнера. Когда окно меняет ширину, наблюдатель вызывает пересчёт, иначе тайлы остаются прежнего размера и появляется пустая полоса.",
        ),
        (
            "Что делает map.container.fitToViewport()?",
            "Метод подгоняет холст карты под текущий размер контейнера. Он вызывается сразу после создания карты и при каждом срабатывании ResizeObserver. В cleanup наблюдатель отключается.",
        ),
        (
            "Какие категории формирует Lighthouse в Chrome DevTools?",
            "Performance, Accessibility, Best Practices и SEO. В этой работе сняты все четыре, в режиме Navigation и на устройстве Mobile.",
        ),
        (
            "Чем Opportunities и Diagnostics отличаются от метрик Performance?",
            "Метрики — измеренные величины, по ним считается балл. Opportunities и Diagnostics — подсказки, где можно сократить байты или время. Подсказка render-blocking requests не заменяет LCP, TBT и CLS.",
        ),
        (
            "Почему Performance score может меняться между запусками?",
            "Симуляция сети и загрузка главного потока чуть различаются от прогона к прогону. Поэтому сняты три запуска и сравнивается медиана: 88 до правок и 89 после.",
        ),
        (
            "Что измеряют LCP, INP и CLS?",
            "LCP — время до отрисовки крупнейшего элемента, здесь это первое фото, 3,0 с по итоговой медиане. INP — задержка отклика на действия за визит. CLS — сумма неожиданных сдвигов вёрстки, в обоих сериях она равна 0.",
        ),
        (
            "Почему Lighthouse не измеряет INP напрямую?",
            "INP считается по реальным действиям за время визита. Один автоматический заход без нажатий не даёт этой выборки.",
        ),
        (
            "Для чего в лаборатории используется TBT?",
            "TBT суммирует время, когда главный поток был занят дольше 50 мс в ходе загрузки. Это лабораторный показатель отзывчивости на том же прогоне, где INP получить нельзя. Медиана TBT — 250 мс.",
        ),
        (
            "Почему локальный Lighthouse-аудит не заменяет полевые данные?",
            "Локальный аудит — одна машина и один профиль троттлинга. Полевые данные — распределение по пользователям, устройствам и сетям. Они отвечают на другой вопрос.",
        ),
        (
            "Почему Accessibility score не заменяет ручную проверку?",
            "Балл собирает автоматические правила. Он не нажимает Tab, не открывает диалог и не ставит масштаб 200%. Поэтому отдельно проверены фокус кнопок закрытия, пять ширин и перестроение при 200%.",
        ),
        (
            "Зачем сравнивать production-сборки с одинаковой конфигурацией?",
            "Иначе разница баллов смешивает правку страницы с режимом dev, другим устройством или другим троттлингом. Все шесть запусков сняты с next start и с одним набором флагов Lighthouse.",
        ),
    ]
    for question, answer in answers:
        add_qa(doc, question, answer)

    add_heading(doc, "ЗАКЛЮЧЕНИЕ", center=True, break_before=True, bookmark="h14", bid=34)
    add_text(
        doc,
        "Галерея PairWise перестраивается от одной колонки на 320 и 375 CSS px до четырёх колонок на 1440 CSS px. Горизонтальной прокрутки на пяти заданных ширинах нет. Диалог помещается в 100dvh и получает фокус на кнопку закрытия, меню на узком экране занимает всю ширину и высоту 100dvh, карта после смены ширины заполняет контейнер. Медианы Lighthouse изменились с 88, 96, 73 и 100 до 89, 100, 77 и 100. LCP сократился с 3,1 с до 3,0 с, CLS остался равен 0. Три подтверждённых аудита закрыты: имя ссылки в виджете карты, отсутствовавший значок и обнаружение LCP-фото. Записи о сторонних cookie Яндекс Карт остались вместе с самой картой.",
    )

    LAB.mkdir(parents=True, exist_ok=True)
    doc.save(OUT)
    update_page_fields(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
