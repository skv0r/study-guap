# -*- coding: utf-8 -*-
"""Сборка отчёта ЛР2. В текст отчёта не входит."""

from __future__ import annotations

import shutil
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Mm, Pt, RGBColor
from lxml import etree

import generate_figures as figs
from idef_model import DIAGRAMS, GLOSSARY, check_balance

BASE = Path(__file__).resolve().parent
SRC = Path(
    r"C:\Users\skvor\OneDrive\Desktop\GitHub\study-guap"
    r"\7_semester\Web-программирование\ЛР1_Отчет_Буренков_PairWise.docx"
)
OUT = BASE / "ЛР2_Отчет_Буренков_КИС.docx"
FIG = BASE / "figures"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

TITLE_REPLACEMENTS = {
    "ассистент": "профессор, д-р техн. наук, профессор",
    "В.В. Жукалин": "В.В. Фомин",
    "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 1": "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 2",
    "Создание проекта Next.js с использованием React, TypeScript и App Router":
        "Изучение методологии моделирования бизнес-процессов IDEF0",
    "по курсу: Web-программирование": "по курсу: Корпоративные информационные системы",
}

TOC = [
    ("1 Задание", "z1", 0),
    ("2 Предметная область", "z2", 0),
    ("3 Модель IDEF0", "z3", 0),
    ("3.1 Контекстная диаграмма А-0", "z31", 1),
    ("3.2 Декомпозиция А0", "z32", 1),
    ("3.3 Приём заказа", "z33", 1),
    ("3.4 Диагностика", "z34", 1),
    ("3.5 Ремонт", "z35", 1),
    ("3.6 Глоссарий стрелок", "z36", 1),
    ("4 Выводы", "z4", 0),
]


def prepare_title(path: Path) -> None:
    shutil.copy(SRC, path)
    import zipfile

    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find(f"{W}body")
    children = list(body)
    if len(children) < 12:
        raise RuntimeError("В образце меньше 12 элементов body")
    p11 = children[11]
    sect = p11.find(f"{W}pPr/{W}sectPr")
    if sect is None:
        raise RuntimeError("В абзаце 11 нет sectPr")
    sect = etree.fromstring(etree.tostring(sect))
    for child in children:
        body.remove(child)
    for child in children[:11]:
        body.append(child)
    body.append(sect)
    replaced = []
    for node in body.iter(f"{W}t"):
        if node.text in TITLE_REPLACEMENTS:
            replaced.append(node.text)
            node.text = TITLE_REPLACEMENTS[node.text]
    missing = [k for k in TITLE_REPLACEMENTS if k not in replaced]
    if missing:
        raise RuntimeError(f"Не найдены строки титула: {missing}")
    report = None
    for node in body.iter(f"{W}t"):
        if node.text and node.text.startswith("ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ"):
            report = node
            break
    rpr = report.getparent().find(f"{W}rPr")
    if rpr is not None and rpr.find(f"{W}b") is not None:
        raise RuntimeError("Строка отчёта на титуле стала жирной")
    xml_out = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = path.with_suffix(".tmp.zip")
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            data = xml_out if item.filename == "word/document.xml" else zin.read(item.filename)
            zout.writestr(item, data)
    tmp.replace(path)


def set_run_font(run, size=14, bold=False):
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = False
    run.font.color.rgb = RGBColor(0, 0, 0)


def format_paragraph(p, *, indent=True, align="justify", before=0, after=0, line=1.5, left=0):
    pf = p.paragraph_format
    if line == 1:
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    else:
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.line_spacing = line
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.first_line_indent = Cm(1.25 if indent else 0)
    pf.left_indent = Cm(left)
    pf.right_indent = Cm(0)
    p.alignment = {
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "left": WD_ALIGN_PARAGRAPH.LEFT,
    }[align]


def add_body(doc, text):
    p = doc.add_paragraph()
    format_paragraph(p)
    set_run_font(p.add_run(text))
    return p


def add_heading(doc, text, bookmark, bid):
    p = doc.add_paragraph()
    format_paragraph(p, indent=True, align="left", before=12, after=6)
    p.paragraph_format.keep_together = True
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text), bold=True)
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bid))
    start.set(qn("w:name"), bookmark)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bid))
    p._p.insert(0, start)
    p._p.append(end)
    return p


def _field_run(paragraph, kind, payload=None):
    run = paragraph.add_run()
    set_run_font(run)
    if kind == "text":
        run.text = payload
        return
    for child in list(run._r):
        if child.tag == qn("w:t"):
            run._r.remove(child)
    if kind == "instr":
        el = OxmlElement("w:instrText")
        el.set(qn("xml:space"), "preserve")
        el.text = payload
    else:
        el = OxmlElement("w:fldChar")
        el.set(qn("w:fldCharType"), kind)
    run._r.append(el)


def add_toc_line(doc, title, bookmark, level):
    p = doc.add_paragraph()
    format_paragraph(p, indent=False, align="left", left=0.75 * level)
    pPr = p._p.get_or_add_pPr()
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:leader"), "dot")
    tab.set(qn("w:pos"), "9071")
    tabs.append(tab)
    pPr.append(tabs)
    set_run_font(p.add_run(title))
    set_run_font(p.add_run("\t"))
    _field_run(p, "begin")
    _field_run(p, "instr", f" PAGEREF {bookmark} \\h ")
    _field_run(p, "separate")
    _field_run(p, "text", "0")
    _field_run(p, "end")
    return p


def add_caption_table(doc, text, table_cm):
    content_twips = 9355
    table_twips = int(round(table_cm / 2.54 * 1440))
    indent_twips = max(0, (content_twips - table_twips) // 2)
    p = doc.add_paragraph()
    format_paragraph(p, indent=False, align="left", before=12, after=6, line=1)
    p.paragraph_format.left_indent = Emu(indent_twips * 635)
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text))
    p._indent_twips = indent_twips
    return p


def add_caption_figure(doc, text):
    p = doc.add_paragraph()
    format_paragraph(p, indent=False, align="center", before=6, after=0)
    set_run_font(p.add_run(text))
    return p


def add_picture(doc, path, width_cm):
    p = doc.add_paragraph()
    format_paragraph(p, indent=False, align="center", before=6, after=0)
    p.paragraph_format.keep_with_next = True
    run = p.add_run()
    run.add_picture(str(path), width=Cm(width_cm))
    return p


def _set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tcPr.append(borders)


def add_table(doc, headers, rows, widths_cm, indent_twips=0):
    table = doc.add_table(rows=1 + len(rows), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    twips = [str(int(round(cm / 2.54 * 1440))) for cm in widths_cm]
    widths = [Cm(cm) for cm in widths_cm]
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    jc = tblPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        tblPr.append(jc)
    jc.set(qn("w:val"), "left")
    if indent_twips:
        ind = OxmlElement("w:tblInd")
        ind.set(qn("w:w"), str(indent_twips))
        ind.set(qn("w:type"), "dxa")
        tblPr.append(ind)
    grid = tbl.find(qn("w:tblGrid"))
    for i, child in enumerate(list(grid)):
        child.set(qn("w:w"), twips[i])
    total = sum(int(x) for x in twips)
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(total))
    tblW.set(qn("w:type"), "dxa")
    data = [headers] + rows
    for r, row in enumerate(data):
        for c, text in enumerate(row):
            cell = table.rows[r].cells[c]
            cell.width = widths[c]
            _set_cell_border(cell)
            cell.text = ""
            p = cell.paragraphs[0]
            format_paragraph(p, indent=False, align="left", line=1.5)
            set_run_font(p.add_run(text), bold=(r == 0))
    tr = table.rows[0]._tr
    trPr = tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    trPr.append(tbl_header)
    return table


def apply_section(section, *, landscape: bool, start: int | None):
    if landscape:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Mm(297)
        section.page_height = Mm(210)
    else:
        section.orientation = WD_ORIENT.PORTRAIT
        section.page_width = Mm(210)
        section.page_height = Mm(297)
    section.left_margin = Mm(30)
    section.right_margin = Mm(15)
    section.top_margin = Mm(20)
    section.bottom_margin = Mm(20)
    section.header_distance = Cm(1.25)
    section.footer_distance = Cm(1.25)
    sect = section._sectPr
    for tag in ("w:footerReference", "w:headerReference"):
        for el in list(sect.findall(qn(tag))):
            sect.remove(el)
    ref = OxmlElement("w:footerReference")
    ref.set(qn("w:type"), "default")
    ref.set(qn("r:id"), "rId11")
    sect.insert(0, ref)
    pg = sect.find(qn("w:pgNumType"))
    if start is None:
        if pg is not None:
            sect.remove(pg)
    else:
        if pg is None:
            pg = OxmlElement("w:pgNumType")
            sect.append(pg)
        pg.set(qn("w:start"), str(start))



def plain(label: str) -> str:
    return label.replace("\n", " ")


def heading(doc, text, bookmark, bid, new_page=False):
    p = add_heading(doc, text, bookmark, bid)
    if new_page:
        p.paragraph_format.page_break_before = True
    return p


def build():
    boxes, errors = check_balance()
    if errors:
        raise SystemExit("ICOM:\n" + "\n".join(errors))
    figs.main()
    prepare_title(OUT)
    doc = Document(str(OUT))
    body = doc.add_section(WD_SECTION.NEW_PAGE)
    apply_section(body, landscape=False, start=2)

    p = doc.add_paragraph()
    format_paragraph(p, indent=False, align="center", before=0, after=12)
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run("СОДЕРЖАНИЕ"), bold=True)

    toc = [
        ("1 Задание", "z1", 0),
        ("2 Предметная область", "z2", 0),
        ("3 Модель IDEF0", "z3", 0),
    ]
    for i, spec in enumerate(DIAGRAMS, start=1):
        toc.append((f"3.{i} {spec['section']}", f"d{i:02d}", 1))
    toc.append((f"3.{len(DIAGRAMS) + 1} Глоссарий стрелок", "gloss", 1))
    toc.append(("4 Выводы", "z4", 0))
    for title, mark, level in toc:
        add_toc_line(doc, title, mark, level)

    heading(doc, "1 Задание", "z1", 1, True)
    add_body(
        doc,
        "Нужно изучить методологию IDEF0 и показать её на сквозном примере своей предметной области. "
        "Модель начинается с контекстной диаграммы А-0, где зафиксированы цель и точка зрения. "
        "Дальше каждый блок, в названии которого ещё спрятано несколько действий, раскрывается дочерней "
        "диаграммой. На одной диаграмме от трёх до шести блоков. Название блока — глагол, название стрелки — "
        "существительное. Вход подводится слева, управление сверху, выход справа, механизм снизу. "
        "У каждого блока есть управление, выход и механизм. Стрелки родителя не пропадают на дочерней диаграмме. "
        "Номер в правом нижнем углу блока согласован со ссылкой на дочернюю диаграмму. Для стрелок составлен глоссарий.",
    )

    heading(doc, "2 Предметная область", "z2", 2, True)
    add_body(
        doc,
        "Мастерская «Обод» ремонтирует городские и спортивные велосипеды по отдельным заказам. "
        "Клиент приносит велосипед и говорит, что нужно исправить. Приёмщик записывает данные клиента, "
        "модель и жалобу, осматривает видимые повреждения и комплектность, по прайс-листу называет "
        "предварительную сумму и только после согласия клиента ставит велосипед на стеллаж с биркой. "
        "Мастер по регламенту отдельно проверяет рулевую, колёса и покрышки, затем колодки и привод тормоза, "
        "затем цепь, переключатели и каретку. Замечания сводит в дефектную ведомость и отмечает узлы к замене.",
    )
    add_body(
        doc,
        "Ремонт начинается со склада: ведомость сверяют с наличием, детали выдают в работу, отсутствие "
        "фиксируют внутри комплекта. Дальше узел снимают, ставят новую деталь, затягивают крепёж и отмечают "
        "замену. После замены настраивают передачи и тормоза и проверяют свободный накат. Список замен и "
        "регулировок подписывают как перечень работ. При выдаче приёмщик сверяет перечень с прайсом, считает "
        "работы и детали, оформляет акт, принимает оплату через кассу и только с чеком снимает бирку и отдаёт "
        "велосипед. В границу модели входит один заказ. Закупка партий у поставщиков, зарплата и реклама не рассматриваются.",
    )

    heading(doc, "3 Модель IDEF0", "z3", 3, True)
    add_body(
        doc,
        f"Модель описывает один заказ. Цель — показать, из каких действий он состоит и какие документы "
        f"связывают эти действия. Точка зрения — владелец мастерской. Дочерние диаграммы сохраняют ту же цель "
        f"и ту же точку зрения. Всего {len(DIAGRAMS)} диаграммы и {boxes} функциональных блоков. "
        f"Контекст А-0 — один блок. На А0 четыре функции: приём, диагностика, ремонт и выдача. "
        f"Каждая из них раскрыта, а внутри приёма, диагностики, ремонта и выдачи раскрыты и те блоки, "
        f"название которых само по себе ещё не говорит, что именно делает человек.",
    )
    add_body(
        doc,
        "Нумерация узлов такая. А-0 ссылается на А0. На А0 блоки 1–4 ссылаются на А1–А4. "
        "А1 делится на А11–А14: регистрация обращения, осмотр, согласование стоимости, хранение. "
        "А2 делится на А21–А24: ходовая часть, тормоза, трансмиссия, дефектная ведомость. "
        "А3 делится на А31–А34: подбор комплектующих, замена, регулировка, запись работ. "
        "А4 делится на А41–А44: расчёт суммы, акт, оплата, передача велосипеда. "
        "Листья этих диаграмм — отдельные действия, их дальше не дробят.",
    )
    add_body(
        doc,
        "Связь уровней видна по стрелкам. На контекст входят велосипед клиента, заявка и комплектующие, "
        "выходят отремонтированный велосипед и акт. Эти же дуги есть на А0: заявка и велосипед входят в приём, "
        "комплектующие — в ремонт, акт и велосипед выходят из выдачи. Карточка заказа рождается в приёме, "
        "идёт в диагностику и сверху ограничивает выдачу, потому что в ней уже лежит согласованная сумма. "
        "Прайс-лист управляет приёмом и выдачей, регламент — всеми четырьмя функциями. Касса подключена только к выдаче. "
        "На рисунках 1–22 каждая дочерняя диаграмма повторяет входы, управления, выходы и механизмы своего блока.",
    )
    for i, spec in enumerate(DIAGRAMS, start=1):
        add_body(doc, f"Рисунок {i}. {spec['blurb']}")

    for i, spec in enumerate(DIAGRAMS, start=1):
        sec = doc.add_section(WD_SECTION.NEW_PAGE)
        apply_section(sec, landscape=True, start=None)
        heading(doc, f"3.{i} {spec['section']}", f"d{i:02d}", 10 + i, False)
        add_picture(doc, FIG / spec["file"], 24.8)
        add_caption_figure(doc, f"Рисунок {i} — {spec['caption']}")

    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    apply_section(sec, landscape=False, start=None)
    gloss_no = len(DIAGRAMS) + 1
    heading(doc, f"3.{gloss_no} Глоссарий стрелок", "gloss", 80, False)
    add_body(
        doc,
        "В таблице 1 собраны дуги, которые переходят с диаграммы на диаграмму или без пояснения читаются двояко. "
        "Короткие подписи на листах означают те же предметы.",
    )
    widths = (4.8, 10.2)
    cap = add_caption_table(doc, "Таблица 1 — Глоссарий интерфейсных дуг", sum(widths))
    rows = [[plain(name), text] for name, text in GLOSSARY]
    add_table(doc, ["Дуга", "Содержание"], rows, widths, cap._indent_twips)

    heading(doc, "4 Выводы", "z4", 90, True)
    add_body(
        doc,
        f"Построена модель IDEF0 мастерской «Обод» из {len(DIAGRAMS)} диаграмм и {boxes} блоков. "
        "Контекст А-0 ограничивает модель одним заказом и фиксирует цель и точку зрения владельца. "
        "А0 делит заказ на приём, диагностику, ремонт и выдачу. Дальше раскрыты регистрация, осмотр, "
        "согласование цены и хранение; проверка рулевой, колёс, покрышек, колодок, привода тормоза, "
        "цепи, переключателей и каретки; сбор ведомости; подбор со склада, замена, затяжка, регулировка "
        "и подпись перечня; расчёт суммы, акт, оплата и передача велосипеда клиенту.",
    )
    add_body(
        doc,
        "На каждой диаграмме не больше шести блоков. У каждого блока есть управление, выход и механизм, "
        "стрелки родителя сохранены на дочернем листе. Глоссарий поясняет карточку заказа, черновики, "
        "ведомость, своды замечаний, перечень работ, прайс-лист, регламент, акт, итоговую сумму и чек.",
    )
    doc.save(str(OUT))
    print(OUT)
    print("diagrams", len(DIAGRAMS), "boxes", boxes)


if __name__ == "__main__":
    build()
