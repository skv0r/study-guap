"""Календарный план разработки КИС «ЛифтКонтур»: работы, ресурсы, CPM."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import Optional

START = date(2026, 10, 6)  # вторник
HOLIDAYS = {
    date(2026, 11, 4),  # День народного единства
    # Новогодние каникулы 2027 (1–8 января)
    date(2027, 1, 1),
    date(2027, 1, 2),
    date(2027, 1, 3),
    date(2027, 1, 4),
    date(2027, 1, 5),
    date(2027, 1, 6),
    date(2027, 1, 7),
    date(2027, 1, 8),
}


def is_workday(d: date) -> bool:
    return d.weekday() < 5 and d not in HOLIDAYS


def next_workday(d: date) -> date:
    cur = d
    while not is_workday(cur):
        cur += timedelta(days=1)
    return cur


def prev_workday(d: date) -> date:
    cur = d - timedelta(days=1)
    while not is_workday(cur):
        cur -= timedelta(days=1)
    return cur


def start_from_finish(lf: date, duration: int) -> date:
    """Поздний старт: duration рабочих дней, окончание в lf включительно."""
    if duration <= 0:
        return lf
    cur = lf
    left = duration - 1
    while left > 0:
        cur -= timedelta(days=1)
        if is_workday(cur):
            left -= 1
    return cur


def add_workdays(start: date, n: int) -> date:
    """Дата окончания работы длительностью n рабочих дней (включительно)."""
    if n <= 0:
        return start
    cur = start
    left = n - 1
    while left > 0:
        cur += timedelta(days=1)
        if is_workday(cur):
            left -= 1
    return cur


def workdays_between(a: date, b: date) -> int:
    """Число рабочих дней в полуинтервале [a, b)."""
    n = 0
    cur = a
    while cur < b:
        if is_workday(cur):
            n += 1
        cur += timedelta(days=1)
    return n


def shift_workdays(start: date, n: int) -> date:
    """Сместить дату на n рабочих дней вперёд (n=0 → та же дата, если рабочая)."""
    cur = next_workday(start) if n >= 0 else start
    if n == 0:
        return next_workday(start)
    left = n
    while left > 0:
        cur += timedelta(days=1)
        if is_workday(cur):
            left -= 1
    return cur


RESOURCES = {
    "volkov": {"fio": "Волков С.Н.", "role": "Руководитель проекта", "rate": 7200, "email": "volkov@liftkontur.local", "phone": "+7 (812) 310-42-01"},
    "lebedeva": {"fio": "Лебедева И.А.", "role": "Бизнес-аналитик", "rate": 5800, "email": "lebedeva@liftkontur.local", "phone": "+7 (812) 310-42-02"},
    "chernov": {"fio": "Чернов П.Д.", "role": "Архитектор КИС", "rate": 8400, "email": "chernov@liftkontur.local", "phone": "+7 (812) 310-42-03"},
    "morozov": {"fio": "Морозов А.В.", "role": "Backend-разработчик", "rate": 7100, "email": "morozov@liftkontur.local", "phone": "+7 (812) 310-42-04"},
    "kuznetsova": {"fio": "Кузнецова О.Л.", "role": "Frontend-разработчик", "rate": 6900, "email": "kuznetsova@liftkontur.local", "phone": "+7 (812) 310-42-05"},
    "grigoriev": {"fio": "Григорьев Н.С.", "role": "Мобильный разработчик", "rate": 7000, "email": "grigoriev@liftkontur.local", "phone": "+7 (812) 310-42-06"},
    "belyaev": {"fio": "Беляев Р.И.", "role": "Тестировщик", "rate": 5300, "email": "belyaev@liftkontur.local", "phone": "+7 (812) 310-42-07"},
    "sokolova": {"fio": "Соколова Т.М.", "role": "Специалист по 1С", "rate": 6400, "email": "sokolova@liftkontur.local", "phone": "+7 (812) 310-42-08"},
}


@dataclass
class Task:
    code: str
    name: str
    duration: int
    preds: list[str] = field(default_factory=list)
    kind: str = "leaf"  # leaf | summary | milestone
    parent: Optional[str] = None
    alloc: list[tuple[str, float]] = field(default_factory=list)
    parallel_note: str = ""
    es: Optional[date] = None
    ef: Optional[date] = None
    ls: Optional[date] = None
    lf: Optional[date] = None
    slack: int = 0
    critical: bool = False

    @property
    def short(self) -> str:
        return f"{self.code} {self.name}"


# Иерархия WBS: 4 этапа, 16 листьев/вех (требование — не менее 8 пунктов).
TASKS: dict[str, Task] = {
    "1": Task("1", "Инициация и анализ", 0, kind="summary"),
    "1.1": Task("1.1", "Обследование диспетчерской и бригад", 4, parent="1",
                alloc=[("lebedeva", 1.0), ("volkov", 0.4)],
                parallel_note="параллельно с 1.2"),
    "1.2": Task("1.2", "Разбор журналов заявок и Excel-учёта", 4, parent="1",
                alloc=[("chernov", 0.5), ("volkov", 0.35)],
                parallel_note="параллельно с 1.1"),
    "1.3": Task("1.3", "Формирование требований к КИС", 5, preds=["1.1", "1.2"], parent="1",
                alloc=[("lebedeva", 1.0), ("volkov", 0.5)]),
    "1.4": Task("1.4", "Написание технического задания", 6, preds=["1.3"], parent="1",
                alloc=[("lebedeva", 1.0), ("volkov", 0.3)]),
    "1.5": Task("1.5", "Утверждение ТЗ", 0, preds=["1.4"], parent="1", kind="milestone",
                alloc=[("volkov", 1.0)]),
    "2": Task("2", "Проектирование системы", 0, kind="summary"),
    "2.1": Task("2.1", "Разработка архитектуры КИС", 7, preds=["1.5"], parent="2",
                alloc=[("chernov", 1.0), ("volkov", 0.25)]),
    "2.2": Task("2.2", "Проектирование модели данных", 8, preds=["2.1"], parent="2",
                alloc=[("chernov", 0.8), ("morozov", 0.45)],
                parallel_note="параллельно с 2.3 и 2.4"),
    "2.3": Task("2.3", "Проектирование интерфейсов АРМ", 8, preds=["2.1"], parent="2",
                alloc=[("kuznetsova", 1.0), ("grigoriev", 0.55), ("chernov", 0.2)],
                parallel_note="параллельно с 2.2 и 2.4"),
    "2.4": Task("2.4", "Проектирование интеграции с 1С", 5, preds=["2.1"], parent="2",
                alloc=[("sokolova", 1.0), ("morozov", 0.2)],
                parallel_note="параллельно с 2.2 и 2.3"),
    "3": Task("3", "Разработка", 0, kind="summary"),
    "3.1": Task("3.1", "Кодирование серверной части и API", 14, preds=["2.2"], parent="3",
                alloc=[("morozov", 1.0)],
                parallel_note="параллельно с 3.2 и 3.3"),
    "3.2": Task("3.2", "Кодирование веб-АРМ диспетчера", 12, preds=["2.3"], parent="3",
                alloc=[("kuznetsova", 1.0)],
                parallel_note="параллельно с 3.1 и 3.3"),
    "3.3": Task("3.3", "Кодирование мобильного клиента механика", 12, preds=["2.3"], parent="3",
                alloc=[("grigoriev", 1.0)],
                parallel_note="параллельно с 3.1 и 3.2"),
    "3.4": Task("3.4", "Реализация обмена с 1С:Бухгалтерия", 8, preds=["2.4", "3.1"], parent="3",
                alloc=[("sokolova", 1.0), ("morozov", 0.3)]),
    "4": Task("4", "Испытания и сдача", 0, kind="summary"),
    "4.1": Task("4.1", "Модульное и интеграционное тестирование", 6, preds=["3.1", "3.2", "3.3", "3.4"], parent="4",
                alloc=[("belyaev", 1.0), ("morozov", 0.35), ("kuznetsova", 0.25), ("grigoriev", 0.25)]),
    "4.2": Task("4.2", "Пилот на двух домах Выборгского района", 8, preds=["4.1"], parent="4",
                alloc=[("belyaev", 1.0), ("volkov", 0.5), ("lebedeva", 0.4)]),
    "4.3": Task("4.3", "Доработка по замечаниям пилота", 5, preds=["4.2"], parent="4",
                alloc=[("morozov", 0.7), ("kuznetsova", 0.5), ("grigoriev", 0.5), ("belyaev", 0.8)]),
    "4.4": Task("4.4", "Сдача проекта заказчику", 0, preds=["4.3"], parent="4", kind="milestone",
                alloc=[("volkov", 1.0)]),
}

ORDER = [
    "1", "1.1", "1.2", "1.3", "1.4", "1.5",
    "2", "2.1", "2.2", "2.3", "2.4",
    "3", "3.1", "3.2", "3.3", "3.4",
    "4", "4.1", "4.2", "4.3", "4.4",
]

LEAF_CODES = [c for c in ORDER if TASKS[c].kind in ("leaf", "milestone")]
PHASE_COLOR = {
    "1": "#4C78A8",
    "2": "#F58518",
    "3": "#54A24B",
    "4": "#E45756",
}


def _topo() -> list[str]:
    remaining = set(LEAF_CODES)
    out: list[str] = []
    while remaining:
        ready = [c for c in remaining if all(p not in remaining for p in TASKS[c].preds)]
        if not ready:
            raise RuntimeError("цикл в зависимостях")
        ready.sort(key=lambda x: LEAF_CODES.index(x))
        c = ready[0]
        remaining.remove(c)
        out.append(c)
    return out


def schedule() -> dict[str, Task]:
    topo = _topo()
    project_start = next_workday(START)

    for code in topo:
        t = TASKS[code]
        if t.preds:
            es = max(TASKS[p].ef + timedelta(days=1) for p in t.preds)  # type: ignore[operator]
            es = next_workday(es)
        else:
            es = project_start
        t.es = es
        if t.duration == 0:
            t.ef = es
        else:
            t.ef = add_workdays(es, t.duration)

    project_end = max(TASKS[c].ef for c in topo)  # type: ignore[type-var]

    for code in reversed(topo):
        t = TASKS[code]
        succs = [s for s in topo if code in TASKS[s].preds]
        if not succs:
            t.lf = project_end
        else:
            t.lf = min(prev_workday(TASKS[s].ls) for s in succs)  # type: ignore[arg-type]
        t.ls = start_from_finish(t.lf, t.duration)
        t.slack = workdays_between(t.es, t.ls) if t.ls >= t.es else 0  # type: ignore[operator]
        t.critical = t.slack == 0

    for code in ORDER:
        t = TASKS[code]
        if t.kind != "summary":
            continue
        kids = [TASKS[c] for c in ORDER if TASKS[c].parent == code and TASKS[c].es]
        t.es = min(k.es for k in kids)
        t.ef = max(k.ef for k in kids)
        t.critical = any(k.critical for k in kids if k.kind != "summary")
        t.duration = workdays_between(t.es, t.ef + timedelta(days=1))  # type: ignore[operator]

    return TASKS


def task_cost(t: Task) -> float:
    if t.kind == "summary" or t.duration == 0:
        return 0.0
    return sum(RESOURCES[rid]["rate"] * t.duration * frac for rid, frac in t.alloc)


def total_cost() -> float:
    return sum(task_cost(TASKS[c]) for c in LEAF_CODES)


def sequential_duration() -> int:
    """Длительность, если все листья идут строго друг за другом (без вех)."""
    return sum(TASKS[c].duration for c in LEAF_CODES if TASKS[c].kind == "leaf")


def planned_duration() -> int:
    t0 = min(TASKS[c].es for c in LEAF_CODES)
    t1 = max(TASKS[c].ef for c in LEAF_CODES)
    return workdays_between(t0, t1 + timedelta(days=1))  # type: ignore[operator]


def fmt(d: Optional[date]) -> str:
    return d.strftime("%d.%m.%Y") if d else "—"


def resource_load_by_day() -> dict[str, dict[date, float]]:
    load: dict[str, dict[date, float]] = {rid: {} for rid in RESOURCES}
    for code in LEAF_CODES:
        t = TASKS[code]
        if t.duration == 0 or t.es is None or t.ef is None:
            continue
        cur = t.es
        while cur <= t.ef:
            if is_workday(cur):
                for rid, frac in t.alloc:
                    load[rid][cur] = load[rid].get(cur, 0.0) + frac
            cur += timedelta(days=1)
    return load


def iter_days(a: date, b: date):
    cur = a
    while cur <= b:
        yield cur
        cur += timedelta(days=1)


if __name__ == "__main__":
    schedule()
    print("start", fmt(START), "end", fmt(max(TASKS[c].ef for c in LEAF_CODES)))
    print("planned", planned_duration(), "sequential", sequential_duration())
    print("saved", sequential_duration() - planned_duration())
    print("cost", round(total_cost()))
    for c in ORDER:
        t = TASKS[c]
        mark = "*" if t.critical else " "
        print(f"{mark} {t.code:4} {t.duration:3} {fmt(t.es):10} {fmt(t.ef):10} slack={t.slack:2}  {t.name}")
    print("--- costs ---")
    for c in LEAF_CODES:
        t = TASKS[c]
        if t.duration:
            print(c, round(task_cost(t)), t.name)
