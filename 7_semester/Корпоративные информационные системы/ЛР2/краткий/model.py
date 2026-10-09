# -*- coding: utf-8 -*-
"""Короткая модель IDEF0. В текст отчёта не входит."""

PRICE = "Прайс-лист"
RULE = "Правила приёмки"
PART = "Расходная\nдеталь"
SHOE = "Обувь клиента"
ORDER = "Заявка клиента"
DONE = "Отремонтированная\nобувь"
PAID = "Квитанция\nоб оплате"
ACCEPTED = "Принятая\nобувь"
TICKET = "Квитанция\nзаказа"
AFTER = "Обувь после\nремонта"
REC = "Приёмщик"
MAS = "Мастер"
CASH = "Касса"


def b(title, child=None, num=None):
    box = {"title": title, "child": child}
    if num is not None:
        box["num"] = num
    return box


def arr(label, to=None, frm=None):
    item = {"label": label}
    if to is not None:
        item["to"] = to if isinstance(to, list) else [to]
    if frm is not None:
        item["frm"] = frm
    return item


def flow(label, frm, to):
    return {"label": label, "frm": frm, "to": to}


def _labels(items):
    return {item["label"] for item in items}


def boundary(spec):
    return {
        "in": _labels(spec.get("in", [])),
        "out": _labels(spec.get("out", [])),
        "ctrl": _labels(spec.get("ctrl", [])),
        "mech": _labels(spec.get("mech", [])),
    }


def touch(spec, index):
    ins, outs, ctrls, mechs = set(), set(), set(), set()
    for item in spec.get("in", []):
        if index in item["to"]:
            ins.add(item["label"])
    for item in spec.get("out", []):
        if item["frm"] == index:
            outs.add(item["label"])
    for item in spec.get("ctrl", []):
        if index in item["to"]:
            ctrls.add(item["label"])
    for item in spec.get("mech", []):
        if index in item["to"]:
            mechs.add(item["label"])
    for item in spec.get("flow", []):
        if item["frm"] == index:
            outs.add(item["label"])
        if item["to"] == index:
            ins.add(item["label"])
    for item in spec.get("up", []):
        if item["frm"] == index:
            outs.add(item["label"])
        if item["to"] == index:
            ctrls.add(item["label"])
    return {"in": ins, "out": outs, "ctrl": ctrls, "mech": mechs}


DIAGRAMS = [
    {
        "id": "А-0",
        "name": "Отремонтировать обувь",
        "file": "fig01_a-0.png",
        "section": "Контекстная диаграмма А-0",
        "caption": "Контекстная диаграмма А-0",
        "blurb": (
            "Модель ограничена одним заказом мастерской «Каблук». На контекст входят обувь, заявка и "
            "расходная деталь, выходят отремонтированная обувь и квитанция об оплате. Цель — описать "
            "выполнение заказа. Точка зрения — владелец мастерской."
        ),
        "purpose": "Цель: описать выполнение одного заказа на ремонт обуви.\nТочка зрения: владелец мастерской.",
        "boxes": [b("Отремонтировать\nобувь", "А0", "0")],
        "in": [arr(SHOE, 0), arr(ORDER, 0), arr(PART, 0)],
        "out": [arr(DONE, frm=0), arr(PAID, frm=0)],
        "ctrl": [arr(PRICE, 0), arr(RULE, 0)],
        "mech": [arr(REC, 0), arr(MAS, 0), arr(CASH, 0)],
    },
    {
        "id": "А0",
        "name": "Отремонтировать обувь",
        "file": "fig02_a0.png",
        "section": "Декомпозиция А0 «Отремонтировать обувь»",
        "caption": "Декомпозиция А0 «Отремонтировать обувь»",
        "blurb": (
            "Заказ делится на приём, ремонт и выдачу. Обувь и заявка входят в приём, расходная деталь — "
            "в ремонт, оба выхода контекста выходят из выдачи. Квитанция заказа рождается в приёме и сверху "
            "ограничивает ремонт и выдачу: по ней мастер знает согласованные работы, а касса сверяет сумму."
        ),
        "boxes": [
            b("Принять\nобувь", "А1"),
            b("Отремонтировать\nпару", "А2"),
            b("Выдать\nобувь", "А3"),
        ],
        "in": [arr(SHOE, 0), arr(ORDER, 0), arr(PART, 1)],
        "out": [arr(DONE, frm=2), arr(PAID, frm=2)],
        "ctrl": [arr(PRICE, [0, 1, 2]), arr(RULE, [0, 2])],
        "mech": [arr(REC, [0, 2]), arr(MAS, 1), arr(CASH, 2)],
        "flow": [flow(ACCEPTED, 0, 1), flow(AFTER, 1, 2)],
        "up": [flow(TICKET, 0, 1), flow(TICKET, 0, 2)],
    },
    {
        "id": "А1",
        "name": "Принять обувь",
        "file": "fig03_a1.png",
        "section": "Диаграмма А1 «Принять обувь»",
        "caption": "Диаграмма А1 «Принять обувь»",
        "blurb": (
            "Приём состоит из осмотра, согласования работ и квитанции. Это узлы А11–А13, отдельными листами "
            "они не раскрываются. Обувь клиента и заявка остаются входами, на выходе — принятая обувь и "
            "квитанция заказа. Прайс-лист управляет только согласованием, правила приёмки — осмотром и квитанцией."
        ),
        "boxes": [
            b("Осмотреть\nпару"),
            b("Согласовать\nработы"),
            b("Выписать\nквитанцию"),
        ],
        "in": [arr(SHOE, 0), arr(ORDER, 1)],
        "out": [arr(ACCEPTED, frm=2), arr(TICKET, frm=2)],
        "ctrl": [arr(RULE, [0, 2]), arr(PRICE, 1)],
        "mech": [arr(REC, [0, 1, 2])],
        "flow": [
            flow("Описание\nдефекта", 0, 1),
            flow("Осмотренная\nобувь", 0, 2),
            flow("Согласованные\nработы", 1, 2),
        ],
    },
    {
        "id": "А2",
        "name": "Отремонтировать пару",
        "file": "fig04_a2.png",
        "section": "Диаграмма А2 «Отремонтировать пару»",
        "caption": "Диаграмма А2 «Отремонтировать пару»",
        "blurb": (
            "Ремонт состоит из разбора, замены детали и проверки. Это узлы А21–А23, глубже они не раскрываются. "
            "Принятая обувь входит в разбор, расходная деталь — в замену, на выходе одна дуга: обувь после ремонта. "
            "Квитанция заказа ограничивает все три действия, прайс-лист — только замену, потому что деталь берётся из перечня цен."
        ),
        "boxes": [
            b("Разобрать\nпару"),
            b("Заменить\nдеталь"),
            b("Проверить\nпару"),
        ],
        "in": [arr(ACCEPTED, 0), arr(PART, 1)],
        "out": [arr(AFTER, frm=2)],
        "ctrl": [arr(TICKET, [0, 1, 2]), arr(PRICE, 1)],
        "mech": [arr(MAS, [0, 1, 2])],
        "flow": [
            flow("Разобранная\nпара", 0, 1),
            flow("Собранная\nпара", 1, 2),
        ],
    },
    {
        "id": "А3",
        "name": "Выдать обувь",
        "file": "fig05_a3.png",
        "section": "Диаграмма А3 «Выдать обувь»",
        "caption": "Диаграмма А3 «Выдать обувь»",
        "blurb": (
            "Выдача состоит из сверки квитанции, оплаты и передачи пары. Это узлы А31–А33, отдельными листами "
            "они не раскрываются. Обувь после ремонта входит в сверку. Сверенная пара идёт на передачу, сумма — "
            "в кассу. Клиент получает обувь и квитанцию об оплате только после отметки об оплате."
        ),
        "boxes": [
            b("Сверить\nквитанцию"),
            b("Принять\nоплату"),
            b("Передать\nобувь"),
        ],
        "in": [arr(AFTER, 0)],
        "out": [arr(DONE, frm=2), arr(PAID, frm=2)],
        "ctrl": [arr(TICKET, 0), arr(RULE, [0, 2]), arr(PRICE, 1)],
        "mech": [arr(REC, [0, 2]), arr(CASH, 1)],
        "flow": [
            flow("Сверенная\nпара", 0, 2),
            flow("Сумма\nк оплате", 0, 1),
            flow("Отметка\nоб оплате", 1, 2),
        ],
    },
]


def check_balance():
    by_id = {spec["id"]: spec for spec in DIAGRAMS}
    errors = []
    boxes = 0
    for spec in DIAGRAMS:
        boxes += len(spec["boxes"])
        if spec["id"] != "А-0" and not 3 <= len(spec["boxes"]) <= 6:
            errors.append(f"{spec['id']}: блоков {len(spec['boxes'])}, нужно 3–6")
        for i, box in enumerate(spec["boxes"]):
            touched = touch(spec, i)
            if not touched["ctrl"]:
                errors.append(f"{spec['id']} блок {i + 1}: нет управления")
            if not touched["out"]:
                errors.append(f"{spec['id']} блок {i + 1}: нет выхода")
            if not touched["mech"]:
                errors.append(f"{spec['id']} блок {i + 1}: нет механизма")
            child = box.get("child")
            if not child:
                continue
            if child not in by_id:
                errors.append(f"{spec['id']} ссылается на отсутствующую {child}")
                continue
            got = boundary(by_id[child])
            for role in ("in", "out", "ctrl", "mech"):
                if touched[role] != got[role]:
                    errors.append(
                        f"{spec['id']}.{i + 1} → {child} {role}: "
                        f"родитель {sorted(touched[role])} / дочь {sorted(got[role])}"
                    )
    return boxes, errors


GLOSSARY = [
    ("Квитанция заказа", "Согласованные работы, срок и предварительная сумма. Оплаты в ней ещё нет. По ней мастер ремонтирует и приёмщик сверяет выдачу."),
    ("Квитанция об оплате", "Подтверждение, что сумма по прайсу внесена. Без неё обувь клиенту не отдают."),
    ("Прайс-лист", "Перечень работ и цен мастерской. По нему согласуют ремонт, берут деталь и считают оплату."),
    ("Правила приёмки", "Что можно принять, как описать дефект и при каких условиях пару отдают клиенту."),
    ("Расходная деталь", "Набойка, подошва или клей со склада мастерской. Входит в замену, а не в приём."),
    ("Принятая обувь", "Пара после осмотра и квитанции, переданная мастеру."),
    ("Обувь после ремонта", "Пара, в которой согласованная деталь уже стоит и шов проверен."),
    ("Описание дефекта", "Что приёмщик увидел на паре до разговора о цене."),
    ("Согласованные работы", "Работы из прайса, которые клиент подтвердил по заявке и дефекту."),
    ("Сверенная пара", "Обувь после ремонта, совпавшая с квитанцией заказа."),
    ("Сумма к оплате", "Стоимость согласованных работ по прайсу."),
    ("Отметка об оплате", "Запись кассы, что сумма внесена. Пока её нет, пару не передают."),
]
