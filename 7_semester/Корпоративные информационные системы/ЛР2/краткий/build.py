# -*- coding: utf-8 -*-
"""Сборка короткого отчёта ЛР2. В текст отчёта не входит. Старый docx не перезаписывает."""

import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION

PARENT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PARENT))

import build_report as br  # noqa: E402
import draw  # noqa: E402
from model import DIAGRAMS, GLOSSARY, check_balance  # noqa: E402

OUT = PARENT / "ЛР2_Отчет_Буренков_КИС_короткий.docx"
FIG = Path(__file__).resolve().parent / "figures"


def plain(label: str) -> str:
    return label.replace("\n", " ")


def build():
    boxes, errors = check_balance()
    if errors:
        raise SystemExit("ICOM:\n" + "\n".join(errors))
    draw.main()
    br.prepare_title(OUT)
    doc = Document(str(OUT))
    body = doc.add_section(WD_SECTION.NEW_PAGE)
    br.apply_section(body, landscape=False, start=2)

    p = doc.add_paragraph()
    br.format_paragraph(p, indent=False, align="center", before=0, after=12)
    p.paragraph_format.keep_with_next = True
    br.set_run_font(p.add_run("СОДЕРЖАНИЕ"), bold=True)

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
        br.add_toc_line(doc, title, mark, level)

    br.heading(doc, "1 Задание", "z1", 1, True)
    br.add_body(
        doc,
        "Нужно показать методологию IDEF0 на короткой модели своей предметной области. "
        "Модель начинается с контекстной диаграммы А-0, где зафиксированы цель и точка зрения. "
        "Декомпозиция доходит до диаграмм А1, А2 и А3 и на этом останавливается: блоки внутри них "
        "дальше не раскрываются. На одной диаграмме от трёх до шести блоков. Название блока — глагол, "
        "название стрелки — существительное. Вход подводится слева, управление сверху, выход справа, "
        "механизм снизу. Стрелки родителя не пропадают на дочерней диаграмме.",
    )

    br.heading(doc, "2 Предметная область", "z2", 2, True)
    br.add_body(
        doc,
        "Обувная мастерская «Каблук» ремонтирует одну пару по отдельному заказу. Клиент приносит обувь "
        "и говорит, что нужно сделать. Приёмщик осматривает пару, по прайс-листу называет работы и только "
        "после согласия выписывает квитанцию заказа. Оплаты в этой квитанции ещё нет.",
    )
    br.add_body(
        doc,
        "Мастер разбирает пару так, как записано в квитанции, ставит расходную деталь со склада и проверяет, "
        "что согласованная работа сделана. При выдаче приёмщик сверяет пару с квитанцией, касса принимает "
        "сумму по прайсу, и только с отметкой об оплате обувь отдают клиенту вместе с квитанцией об оплате. "
        "В границу модели входит один заказ. Закупка партий у поставщиков и зарплата не рассматриваются.",
    )

    br.heading(doc, "3 Модель IDEF0", "z3", 3, True)
    br.add_body(
        doc,
        f"Модель описывает один заказ. Цель — показать, из каких действий он состоит и какие документы "
        f"их связывают. Точка зрения — владелец мастерской. Всего {len(DIAGRAMS)} диаграмм и {boxes} "
        f"функциональных блоков. Контекст А-0 — один блок. На А0 три функции: приём, ремонт и выдача. "
        f"Каждая раскрыта одним листом. Узлы А11–А13, А21–А23 и А31–А33 остаются действиями на этих листах "
        f"и отдельных диаграмм не имеют.",
    )
    br.add_body(
        doc,
        "Связь уровней видна по стрелкам. На контекст входят обувь клиента, заявка и расходная деталь, "
        "выходят отремонтированная обувь и квитанция об оплате. На А0 заявка и обувь входят в приём, "
        "деталь — в ремонт, оба выхода снимаются с выдачи. Квитанция заказа рождается в приёме и сверху "
        "ограничивает ремонт и выдачу. Прайс-лист управляет всеми тремя функциями, правила приёмки — "
        "приёмом и выдачей. Мастер подключён только к ремонту, касса — только к выдаче.",
    )

    for i, spec in enumerate(DIAGRAMS, start=1):
        br.add_body(doc, f"Рисунок {i}. {spec['blurb']}")

    for i, spec in enumerate(DIAGRAMS, start=1):
        sec = doc.add_section(WD_SECTION.NEW_PAGE)
        br.apply_section(sec, landscape=True, start=None)
        br.heading(doc, f"3.{i} {spec['section']}", f"d{i:02d}", 10 + i, False)
        br.add_picture(doc, FIG / spec["file"], 22)
        br.add_caption_figure(doc, f"Рисунок {i} — {spec['caption']}")

    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    br.apply_section(sec, landscape=False, start=None)
    gloss_no = len(DIAGRAMS) + 1
    br.heading(doc, f"3.{gloss_no} Глоссарий стрелок", "gloss", 80, False)
    br.add_body(
        doc,
        "В таблице 1 собраны дуги, которые переходят с диаграммы на диаграмму или без пояснения читаются двояко.",
    )
    widths = (5.2, 11.0)
    cap = br.add_caption_table(doc, "Таблица 1 — Глоссарий интерфейсных дуг", sum(widths))
    rows = [[plain(name), text] for name, text in GLOSSARY]
    br.add_table(doc, ["Дуга", "Содержание"], rows, widths, cap._indent_twips)

    br.heading(doc, "4 Выводы", "z4", 90, True)
    br.add_body(
        doc,
        f"Построена модель IDEF0 мастерской «Каблук» из {len(DIAGRAMS)} диаграмм и {boxes} блоков. "
        "Контекст А-0 ограничивает модель одним заказом и фиксирует цель и точку зрения владельца. "
        "А0 делит заказ на приём, ремонт и выдачу. Дальше раскрыты осмотр, согласование и квитанция; "
        "разбор, замена детали и проверка; сверка, оплата и передача обуви. Глубже этих листов модель не идёт.",
    )
    br.add_body(
        doc,
        "На каждой диаграмме три блока. У каждого блока есть управление, выход и механизм, стрелки родителя "
        "сохранены на дочернем листе. Квитанция заказа связывает приём с ремонтом и выдачей, а обувь "
        "не покидает мастерскую без отметки об оплате.",
    )
    doc.save(str(OUT))
    print(OUT)


if __name__ == "__main__":
    build()
