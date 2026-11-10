# -*- coding: utf-8 -*-
"""Сборка отчётов ЛР5–ЛР7. В текст отчётов не входит."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(r"C:\Users\skvor\OneDrive\Desktop\GitHub\study-guap\7_semester\Web-программирование")
PROJECT = Path(r"C:\Users\skvor\OneDrive\Desktop\GitHub\pairwise")
EVIDENCE = PROJECT / "lab-evidence"
SRC_DOC = ROOT / "ЛР1_Отчет_Буренков_PairWise.docx"
if not SRC_DOC.exists():
    SRC_DOC = ROOT / "ЛР4" / "ЛР4_Отчет_Буренков_PairWise.docx"

OUT5 = ROOT / "ЛР5" / "ЛР5_Отчет_Буренков_PairWise.docx"
OUT6 = ROOT / "ЛР6" / "ЛР6_Отчет_Буренков_PairWise.docx"
OUT7 = ROOT / "ЛР7" / "ЛР7_Отчет_Буренков_PairWise.docx"
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


def set_cell(cell, text, *, bold=False, header=False, size=12):
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


def add_table(doc, caption, headers, rows, widths, size=11):
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
    if path is None or not Path(path).exists():
        return False
    paragraph = doc.add_paragraph()
    set_spacing(paragraph, before=6, after=0, line=240)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    run.add_picture(str(path), width=Cm(width_cm))
    add_text(doc, caption, first=0, before=6, after=10, line=360, center=True)
    return True


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
        style_run(run, size=9, bold=False, font="Courier New")
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


def read_evidence(name: str) -> str:
    path = EVIDENCE / name
    if not path.exists():
        return f"[нет файла {name}]"
    return path.read_text(encoding="utf-8", errors="replace").strip()


def png(name: str) -> Path | None:
    path = EVIDENCE / name
    return path if path.exists() else None


def replace_title(doc, number: str, topic: str):
    replacements = {
        "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 1": f"ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № {number}",
        "Создание проекта Next.js с использованием React, TypeScript и App Router": topic,
    }
    # Если титул уже с другой лабы — подхватить известные строки
    extra_topics = {
        "Реализация адаптивности клиентской части и контроль качества с помощью Lighthouse": topic,
        "Создание интерактивных элементов интерфейса на JavaScript": topic,
        "Работа с DOM и событиями в клиентской части": topic,
    }
    found = set()
    for node in doc.element.body.iter(qn("w:t")):
        text = node.text or ""
        if text in replacements:
            node.text = replacements[text]
            found.add(node.text)
        elif text in extra_topics:
            node.text = extra_topics[text]
            found.add(node.text)
        elif text.startswith("ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ №") and text != replacements[
            "ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 1"
        ]:
            node.text = replacements["ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 1"]
            found.add(node.text)
    expected_title = replacements["ОТЧЕТ О ЛАБОРАТОРНОЙ РАБОТЕ № 1"]
    if expected_title not in found and topic not in found:
        raise SystemExit(f"Титул заменён не полностью для ЛР{number}: {found}")


def fresh_title(number: str, topic: str) -> Document:
    doc = Document(str(SRC_DOC))
    body = doc.element.body
    children = list(body)
    for child in children[12:-1]:
        body.remove(child)
    replace_title(doc, number, topic)
    return doc


def fill_pages(path: Path) -> None:
    import win32com.client as win32

    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = word.Documents.Open(str(path))
    try:
        doc.Repaginate()
        for i in range(1, doc.Fields.Count + 1):
            doc.Fields.Item(i).Update()
        doc.Repaginate()
        doc.Save()
    finally:
        doc.Close(False)
        word.Quit()


def write_page_numbers(path: Path) -> None:
    """Обновляет поля PAGEREF содержания после пересчёта страниц."""
    fill_pages(path)


def save_report(doc: Document, out: Path) -> None:
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    write_page_numbers(out)
    print(out)


# --- Листинги из проекта (без секретов) ---

SCHEMA_LR5 = """
// Генерирует Prisma Client для установленной версии Prisma ORM 6.19.2.
generator client {
  provider = "prisma-client-js"
}

// Читает локальное подключение PostgreSQL из DATABASE_URL.
datasource db {
  provider = "postgresql"
  url      = env("DATABASE_URL")
}

// Временная модель позволяет проверить генерацию клиента.
// В ЛР6 она удаляется до создания первой миграции.
model SetupCheck {
  id Int @id
}
"""

ENV_EXAMPLE_LR5 = """
# Документирует формат подключения без настоящего пароля.
DATABASE_URL="postgresql://study_app:CHANGE_ME@localhost:5432/study_space?schema=public"
"""

PRISMA_TS = """
import { PrismaClient } from "@prisma/client";

const globalForPrisma = globalThis as unknown as {
    prisma: PrismaClient | undefined;
};

export const prisma = globalForPrisma.prisma ?? new PrismaClient();

if (process.env.NODE_ENV !== "production") {
    globalForPrisma.prisma = prisma;
}
"""

CHECK_PRISMA = """
import "dotenv/config";
import { prisma } from "../app/lib/prisma";

type ConnectionInfo = {
    database: string;
    role: string;
    port: number;
};

async function main() {
    const rows = await prisma.$queryRaw<ConnectionInfo[]>`
        SELECT
            current_database() AS database,
            current_user AS role,
            inet_server_port() AS port
    `;

    console.log("Prisma connected:", rows[0]);
}
"""

SCHEMA_LR6 = Path(PROJECT / "prisma" / "schema.prisma").read_text(encoding="utf-8")
MIGRATION_SQL = Path(
    PROJECT / "prisma" / "migrations" / "20261108233000_init_study_space" / "migration.sql"
).read_text(encoding="utf-8")

SEED_SAFE = """
import "dotenv/config";
import { PrismaClient } from "@prisma/client";
import { hash } from "bcrypt";

const prisma = new PrismaClient();

// Массив из шести площадок PairWise (slug, title, summary, address, image...).
const places = [ /* шесть upsert-записей */ ];

async function main() {
    // Учебный пароль хешируется bcrypt с cost 12 и в отчёт не выводится.
    const passwordHash = await hash("<учебный_пароль>", 12);

    const user = await prisma.user.upsert({
        where: { email: "student@example.test" },
        update: { name: "Демонстрационный студент", passwordHash },
        create: {
            name: "Демонстрационный студент",
            email: "student@example.test",
            passwordHash,
        },
    });

    for (const place of places) {
        await prisma.studyPlace.upsert({
            where: { slug: place.slug },
            update: place,
            create: place,
        });
    }

    const home = await prisma.studyPlace.findUniqueOrThrow({
        where: { slug: "tochka-kipeniya" },
    });

    await prisma.favoritePlace.upsert({
        where: { userId_placeId: { userId: user.id, placeId: home.id } },
        update: {},
        create: { userId: user.id, placeId: home.id },
    });

    console.log("Seed complete: 1 user, 6 places, 1 favorite");
}
"""

VALIDATION_TS = Path(PROJECT / "app" / "lib" / "validation.ts").read_text(encoding="utf-8")
AUTH_TS = Path(PROJECT / "app" / "lib" / "auth.ts").read_text(encoding="utf-8")
CURRENT_USER_TS = Path(PROJECT / "app" / "lib" / "current-user.ts").read_text(encoding="utf-8")
REGISTER_TS = Path(PROJECT / "app" / "api" / "auth" / "register" / "route.ts").read_text(encoding="utf-8")
LOGIN_TS = Path(PROJECT / "app" / "api" / "auth" / "login" / "route.ts").read_text(encoding="utf-8")
ME_TS = Path(PROJECT / "app" / "api" / "auth" / "me" / "route.ts").read_text(encoding="utf-8")
LOGOUT_TS = Path(PROJECT / "app" / "api" / "auth" / "logout" / "route.ts").read_text(encoding="utf-8")
FAVORITES_TS = Path(PROJECT / "app" / "api" / "favorites" / "[placeId]" / "route.ts").read_text(encoding="utf-8")
PLACES_TS = Path(PROJECT / "app" / "api" / "places" / "route.ts").read_text(encoding="utf-8")
AUTH_PROVIDER = Path(PROJECT / "app" / "providers" / "AuthProvider.tsx").read_text(encoding="utf-8")
GALLERY = Path(PROJECT / "app" / "gallery" / "_components" / "PlacesGallery.tsx").read_text(encoding="utf-8")
ENV_EXAMPLE = Path(PROJECT / ".env.example").read_text(encoding="utf-8")
VITEST_CFG = Path(PROJECT / "vitest.config.mts").read_text(encoding="utf-8")
PLAYWRIGHT_CFG = Path(PROJECT / "playwright.config.ts").read_text(encoding="utf-8")
UNIT_TEST = Path(PROJECT / "tests" / "unit" / "validation.test.ts").read_text(encoding="utf-8")
E2E_TEST = Path(PROJECT / "tests" / "e2e" / "auth-and-favorites.spec.ts").read_text(encoding="utf-8")
LOGIN_FORM = Path(PROJECT / "app" / "ui" / "LoginMenu" / "LoginMenu.tsx").read_text(encoding="utf-8")
REGISTER_FORM = Path(PROJECT / "app" / "ui" / "RegisterMenu" / "RegisterMenu.tsx").read_text(encoding="utf-8")
PLACE_TS = Path(PROJECT / "app" / "lib" / "place.ts").read_text(encoding="utf-8")


def build_lr5():
    doc = fresh_title("5", "Установка PostgreSQL и подключение Prisma ORM")

    add_heading(doc, "СОДЕРЖАНИЕ", center=True, bookmark="toc", bid=20)
    toc = [
        ("1 Цель работы", "h1"),
        ("2 Среда выполнения", "h2"),
        ("3 Сервер, база, схема, роль и клиент", "h3"),
        ("4 Создание роли и базы", "h4"),
        ("5 Проверка подключения через psql", "h5"),
        ("6 Установка Prisma ORM", "h6"),
        ("7 Новые файлы проекта", "h7"),
        ("8 Схема, окружение и клиент", "h8"),
        ("9 Проверка validate, generate и TypeScript-запроса", "h9"),
        ("10 Результаты lint и build", "h10"),
        ("11 Возникшая проблема и диагностика", "h11"),
        ("12 Ответы на контрольные вопросы", "h12"),
        ("ЗАКЛЮЧЕНИЕ", "h13"),
    ]
    for title, bookmark in toc:
        add_toc(doc, title, bookmark)

    add_heading(doc, "1 Цель работы", break_before=True, bookmark="h1", bid=21)
    add_text(
        doc,
        "Цель работы — установить и проверить локальный PostgreSQL, создать изолированную роль study_app и базу study_space для проекта PairWise, подключить Prisma ORM 6.19.2 к существующему Next.js-проекту, безопасно настроить строку подключения, сгенерировать Prisma Client и выполнить диагностический запрос к серверу базы данных. Физические таблицы предметной области в этой работе ещё не создаются.",
    )

    add_heading(doc, "2 Среда выполнения", break_before=True, bookmark="h2", bid=22)
    add_text(
        doc,
        "Работа выполнена в Windows 10/11. Версии средств зафиксированы в таблице 1. Служба PostgreSQL — postgresql-x64-17, порт 5432. Каталог данных и установочный каталог заданы установщиком EDB по умолчанию. Установлены сервер PostgreSQL и Command Line Tools; pgAdmin доступен как дополнительный клиент и сервер не заменяет.",
    )
    add_table(
        doc,
        "Таблица 1 — Версии среды",
        ["Компонент", "Версия"],
        [
            ["Windows", "10/11"],
            ["Node.js", "v24.20.0"],
            ["npm", "11.19.0"],
            ["PostgreSQL", "17.11"],
            ["Порт", "5432"],
            ["Служба", "postgresql-x64-17"],
            ["prisma", "6.19.2"],
            ["@prisma/client", "6.19.2"],
        ],
        [4677, 4678],
    )
    add_code(doc, "Листинг 1 — Вывод psql --version и версий Node/npm/Prisma", read_evidence("versions.txt"))

    add_heading(doc, "3 Сервер, база, схема, роль и клиент", break_before=True, bookmark="h3", bid=23)
    add_text(
        doc,
        "Сервер PostgreSQL — запущенная служба, которая слушает порт и обслуживает несколько баз. Кластер — набор баз одного экземпляра. База study_space — отдельное логическое пространство учебного проекта PairWise. Схема public — пространство имён объектов внутри этой базы. Роль study_app — учётная запись, от имени которой подключается приложение. Клиент psql отправляет SQL серверу. pgAdmin является графическим клиентом и не заменяет сервер.",
    )

    add_heading(doc, "4 Создание роли и базы", break_before=True, bookmark="h4", bid=24)
    add_text(
        doc,
        "Под ролью postgres созданы прикладная роль без прав суперпользователя и отдельная база с владельцем study_app. Пароль роли в отчёт не помещён.",
    )
    add_code(
        doc,
        "Листинг 2 — SQL создания роли и базы (пароль удалён)",
        """
CREATE ROLE study_app
  WITH LOGIN
  PASSWORD '<удалено>';

CREATE DATABASE study_space
  WITH OWNER = study_app
       ENCODING = 'UTF8';
""",
    )

    add_heading(doc, "5 Проверка подключения через psql", break_before=True, bookmark="h5", bid=25)
    add_text(
        doc,
        "Подключение выполнено командой psql -h localhost -p 5432 -U study_app -d study_space. Запрос SELECT current_database(), current_user, inet_server_port() вернул study_space, study_app и 5432. На этом этапе \\dt показывает отсутствие прикладных таблиц.",
    )
    add_code(doc, "Листинг 3 — Результат прикладного SELECT", read_evidence("select_role.txt"))
    shot = png("psql-version.png") or png("select-role.png")
    if shot:
        add_picture(doc, shot, "Рисунок 1 — Проверка psql и роли", 14)

    add_heading(doc, "6 Установка Prisma ORM", break_before=True, bookmark="h6", bid=26)
    add_text(
        doc,
        "В корне проекта установлены согласованные версии prisma@6.19.2 и @prisma/client@6.19.2, а также dotenv и tsx. Команда npx prisma --version подтвердила совпадение CLI и клиента (листинг 1). Инициализация выполнена для провайдера postgresql.",
    )

    add_heading(doc, "7 Новые файлы проекта", break_before=True, bookmark="h7", bid=27)
    add_text(
        doc,
        "Добавлены файлы конфигурации и проверки соединения. Дерево новых путей:",
    )
    add_code(
        doc,
        "Листинг 4 — Дерево новых файлов",
        """
prisma/
  schema.prisma
app/lib/
  prisma.ts
scripts/
  check-prisma.ts
.env.example
.gitignore  (+ .env*)
""",
    )

    add_heading(doc, "8 Схема, окружение и клиент", break_before=True, bookmark="h8", bid=28)
    add_text(
        doc,
        "В schema.prisma описаны generator client, datasource postgresql через env(\"DATABASE_URL\") и временная модель SetupCheck только для проверки генерации клиента. Предметные таблицы появятся в следующей работе. Реальный .env исключён из Git правилом .env*; в репозиторий попадает только .env.example без секрета.",
    )
    add_code(doc, "Листинг 5 — prisma/schema.prisma (ветка lr5)", SCHEMA_LR5)
    add_code(doc, "Листинг 6 — .env.example", ENV_EXAMPLE_LR5)
    add_code(doc, "Листинг 7 — app/lib/prisma.ts", PRISMA_TS)
    add_code(doc, "Листинг 8 — scripts/check-prisma.ts", CHECK_PRISMA)

    add_heading(doc, "9 Проверка validate, generate и TypeScript-запроса", break_before=True, bookmark="h9", bid=29)
    add_text(
        doc,
        "Команды prisma format и prisma validate завершились успешно. Prisma Client сгенерирован. Скрипт db:check выполнил диагностический SELECT через Prisma Client и подтвердил study_space, study_app, порт 5432.",
    )
    add_code(doc, "Листинг 9 — prisma validate", read_evidence("validate.txt"))
    add_code(doc, "Листинг 10 — prisma generate", read_evidence("generate.txt"))
    add_code(doc, "Листинг 11 — Результат TypeScript-запроса", read_evidence("check-prisma.txt"))
    for name, caption in (
        ("validate.png", "Рисунок — prisma validate"),
        ("generate.png", "Рисунок — prisma generate"),
        ("check-prisma.png", "Рисунок — TypeScript-запрос"),
    ):
        p = png(name)
        if p:
            add_picture(doc, p, caption, 14)

    add_heading(doc, "10 Результаты lint и build", break_before=True, bookmark="h10", bid=30)
    add_text(
        doc,
        "Команда npm run lint завершилась без ошибок ESLint (имеется предупреждение о deprecated-правиле jsx-indent плагина stylistic, на результат проверки не влияет). Команда npm run build завершилась без ошибок.",
    )
    add_code(doc, "Листинг 12 — npm run lint", read_evidence("lint.txt"))

    add_heading(doc, "11 Возникшая проблема и диагностика", break_before=True, bookmark="h11", bid=31)
    add_text(
        doc,
        "После установки PostgreSQL команда psql в уже открытом окне PowerShell не находилась в PATH. Диагностика: where.exe psql не возвращал путь, тогда как SQL Shell из меню Пуск работал. Решение — открыть новый сеанс PowerShell после обновления пользовательского PATH каталогом bin установленного PostgreSQL 17 и повторить psql --version. Ошибка подключения к неверной роли или порту исключалась проверкой \\conninfo и SELECT current_database(), current_user, inet_server_port().",
    )

    add_heading(doc, "12 Ответы на контрольные вопросы", break_before=True, bookmark="h12", bid=32)
    answers = [
        (
            "Чем сервер PostgreSQL отличается от базы данных?",
            "Сервер — процесс и служба, принимающие подключения и управляющие файлами данных. База данных — одно логическое пространство внутри сервера; один сервер обслуживает несколько баз.",
        ),
        (
            "Что такое кластер баз данных PostgreSQL?",
            "Это набор баз, управляемых одним экземпляром сервера (общие каталог данных, настройки и порт).",
        ),
        (
            "Для чего нужна роль PostgreSQL?",
            "Роль задаёт учётную запись и права доступа. Приложение PairWise подключается как study_app, а не как администратор.",
        ),
        (
            "Почему приложение не должно работать под ролью postgres?",
            "Роль postgres имеет избыточные административные права. Ошибка приложения под ней опаснее, чем под ограниченной прикладной ролью.",
        ),
        (
            "Чем база данных отличается от схемы public?",
            "Клиент сначала подключается к базе study_space, а уже внутри неё обращается к объектам схемы (здесь public). База и схема — разные уровни пространства имён.",
        ),
        (
            "Какие параметры использует psql для подключения?",
            "Ключевые параметры: -h хост, -p порт, -U роль, -d база. В работе: localhost, 5432, study_app, study_space.",
        ),
        (
            "Что показывает \\conninfo?",
            "Метакоманда выводит сведения о текущем соединении: база, пользователь, хост и порт.",
        ),
        (
            "Из каких частей состоит Prisma ORM?",
            "Из Prisma Schema, CLI (validate, generate, migrate), Prisma Client для TypeScript и (при миграциях) движка миграций.",
        ),
        (
            "Что описывают datasource и generator?",
            "datasource задаёт провайдер и URL подключения. generator указывает, какой клиент генерировать (здесь prisma-client-js).",
        ),
        (
            "Для чего используется DATABASE_URL?",
            "Это строка подключения PostgreSQL для Prisma: роль, хост, порт, база и schema=public.",
        ),
        (
            "Почему DATABASE_URL нельзя объявлять с NEXT_PUBLIC_?",
            "Префикс NEXT_PUBLIC_ попадает в клиентский бандл. Секрет подключения не должен быть доступен браузеру.",
        ),
        (
            "Почему специальные символы пароля могут потребовать percent-encoding?",
            "В URL символы вроде @, :, /, % имеют особое значение. Их кодируют percent-encoding, иначе строка подключения разбирается неверно.",
        ),
        (
            "Почему версии prisma и @prisma/client должны совпадать?",
            "CLI генерирует клиент под конкретную версию API. Расхождение версий приводит к несовместимому клиенту и ошибкам выполнения.",
        ),
        (
            "Что делает prisma validate?",
            "Проверяет синтаксис и согласованность schema.prisma без изменения базы.",
        ),
        (
            "Что делает prisma generate?",
            "Генерирует Prisma Client в node_modules по текущей схеме.",
        ),
        (
            "Почему PrismaClient сохраняют в globalThis в development?",
            "Чтобы при горячей перезагрузке Next.js не создавать новый экземпляр на каждый reload и не исчерпывать пул соединений.",
        ),
        (
            "Чем migrate dev отличается от migrate deploy?",
            "migrate dev создаёт и применяет миграции в разработке. migrate deploy применяет уже существующие миграции, обычно в production.",
        ),
        (
            "Почему в этой работе миграция не выполняется?",
            "Физические таблицы предметной области проектируются в ЛР6. Здесь проверяются только установка, конфигурация, генерация клиента и соединение.",
        ),
        (
            "Что означают типовые ошибки подключения P1000 и P1001?",
            "P1000 — отказ аутентификации (неверные учётные данные). P1001 — сервер недоступен (не запущен, неверный хост или порт).",
        ),
        (
            "Какие данные нельзя помещать в отчёт и репозиторий?",
            "Реальные пароли ролей, DATABASE_URL с паролем, AUTH_SECRET, содержимое cookie/JWT и любые другие секреты.",
        ),
    ]
    for question, answer in answers:
        add_qa(doc, question, answer)

    add_heading(doc, "ЗАКЛЮЧЕНИЕ", center=True, break_before=True, bookmark="h13", bid=33)
    add_text(
        doc,
        "Установлен PostgreSQL 17.11 (служба postgresql-x64-17, порт 5432). Созданы роль study_app и база study_space. Подключена Prisma ORM 6.19.2 с совпадающим @prisma/client. Строка подключения хранится локально, .env.example без секрета. Временная модель SetupCheck позволила проверить generate; предметные таблицы не создавались. Диагностический TypeScript-запрос подтвердил study_space | study_app | 5432. Команды lint и build завершились без ошибок.",
    )
    save_report(doc, OUT5)


def build_lr6():
    doc = fresh_title(
        "6",
        "Физическая реализация локальной базы данных, заполнение БД и серверная авторизация",
    )

    add_heading(doc, "СОДЕРЖАНИЕ", center=True, bookmark="toc", bid=20)
    toc = [
        ("1 Цель работы", "h1"),
        ("2 Схема связей и назначение таблиц", "h2"),
        ("3 Столбцы, ключи и ограничения", "h3"),
        ("4 Составной ключ и onDelete Cascade", "h4"),
        ("5 schema.prisma и migration.sql", "h5"),
        ("6 Проверка таблиц и migrate status", "h6"),
        ("7 Seed и заполнение базы", "h7"),
        ("8 Валидация, сессия и current user", "h8"),
        ("9 Маршруты авторизации и избранного", "h9"),
        ("10 Матрица HTTP-сценариев", "h10"),
        ("11 Отсутствие passwordHash в ответах", "h11"),
        ("12 Lint, build и ограничения учебной авторизации", "h12"),
        ("13 Ответы на контрольные вопросы", "h13"),
        ("ЗАКЛЮЧЕНИЕ", "h14"),
    ]
    for title, bookmark in toc:
        add_toc(doc, title, bookmark)

    add_heading(doc, "1 Цель работы", break_before=True, bookmark="h1", bid=21)
    add_text(
        doc,
        "Цель работы — спроектировать и физически создать локальную реляционную базу study_space в PostgreSQL для PairWise, объяснить назначение таблиц User, StudyPlace и FavoritePlace, заполнить базу идемпотентным seed и реализовать учебную серверную авторизацию на Next.js: bcrypt, подписанная cookie-сессия jose, маршруты register/login/me/logout и защита избранного.",
    )

    add_heading(doc, "2 Схема связей и назначение таблиц", break_before=True, bookmark="h2", bid=22)
    add_text(
        doc,
        "Связь предметной области: User → FavoritePlace ← StudyPlace. User хранит зарегистрированных пользователей. StudyPlace — площадки галереи и карты PairWise. FavoritePlace — явная таблица связи многие-ко-многим между пользователем и площадкой.",
    )
    add_code(
        doc,
        "Листинг 1 — Схема связей",
        """
User (1) ----< FavoritePlace >---- (1) StudyPlace
                 @@id(userId, placeId)
""",
    )

    add_heading(doc, "3 Столбцы, ключи и ограничения", break_before=True, bookmark="h3", bid=23)
    add_text(
        doc,
        "Ниже приведены столбцы с типами PostgreSQL, запретом NULL и ограничениями.",
    )
    add_table(
        doc,
        "Таблица 1 — User",
        ["Столбец", "Тип", "NULL", "Ограничения"],
        [
            ["id", "uuid", "NOT NULL", "PK, default uuid()"],
            ["name", "varchar(80)", "NOT NULL", ""],
            ["email", "varchar(254)", "NOT NULL", "UNIQUE"],
            ["passwordHash", "varchar(255)", "NOT NULL", "хеш, не пароль"],
            ["createdAt", "timestamptz(3)", "NOT NULL", "DEFAULT now()"],
            ["updatedAt", "timestamptz(3)", "NOT NULL", "@updatedAt"],
        ],
        [2000, 2500, 1600, 3255],
    )
    add_table(
        doc,
        "Таблица 2 — StudyPlace",
        ["Столбец", "Тип", "NULL", "Ограничения"],
        [
            ["id", "uuid", "NOT NULL", "PK"],
            ["slug", "varchar(80)", "NOT NULL", "UNIQUE"],
            ["title", "varchar(160)", "NOT NULL", ""],
            ["summary", "text", "NOT NULL", ""],
            ["address", "varchar(240)", "NOT NULL", ""],
            ["image", "varchar(240)", "NOT NULL", "путь из public"],
            ["imageAlt", "varchar(240)", "NOT NULL", ""],
            ["latitude", "numeric(9,6)", "NOT NULL", ""],
            ["longitude", "numeric(9,6)", "NOT NULL", ""],
            ["createdAt", "timestamptz(3)", "NOT NULL", "DEFAULT now()"],
            ["updatedAt", "timestamptz(3)", "NOT NULL", "@updatedAt"],
        ],
        [2000, 2500, 1600, 3255],
    )
    add_table(
        doc,
        "Таблица 3 — FavoritePlace",
        ["Столбец", "Тип", "NULL", "Ограничения"],
        [
            ["userId", "uuid", "NOT NULL", "PK, FK → User"],
            ["placeId", "uuid", "NOT NULL", "PK, FK → StudyPlace, INDEX"],
            ["createdAt", "timestamptz(3)", "NOT NULL", "DEFAULT now()"],
        ],
        [2000, 2500, 1600, 3255],
    )

    add_heading(doc, "4 Составной ключ и onDelete Cascade", break_before=True, bookmark="h4", bid=24)
    add_text(
        doc,
        "Составной первичный ключ (userId, placeId) не позволяет одному пользователю дважды добавить одну площадку. onDelete: Cascade удаляет строки FavoritePlace при удалении пользователя или площадки, чтобы не оставлять «висячие» связи. Для placeId создан отдельный индекс FavoritePlace_placeId_idx: поиск избранного по площадке не покрывается ведущим столбцом составного PK.",
    )

    add_heading(doc, "5 schema.prisma и migration.sql", break_before=True, bookmark="h5", bid=25)
    add_text(
        doc,
        "Временная модель SetupCheck удалена. Первая миграция 20261108233000_init_study_space создаёт три таблицы, уникальные индексы, индекс placeId и внешние ключи с ON DELETE CASCADE.",
    )
    add_code(doc, "Листинг 2 — prisma/schema.prisma", SCHEMA_LR6)
    add_code(doc, "Листинг 3 — migration.sql", MIGRATION_SQL)

    add_heading(doc, "6 Проверка таблиц и migrate status", break_before=True, bookmark="h6", bid=26)
    add_text(
        doc,
        "В схеме public после миграции присутствуют FavoritePlace, StudyPlace, User и служебная _prisma_migrations. Команда prisma migrate status сообщает, что найдена одна миграция и схема актуальна.",
    )
    add_code(doc, "Листинг 4 — Список таблиц", read_evidence("tables.txt"))
    add_code(doc, "Листинг 5 — \\d трёх таблиц", read_evidence("describe.txt"))
    add_code(doc, "Листинг 6 — migrate status", read_evidence("migrate-status.txt"))
    for name, caption in (
        ("describe-user.png", "Рисунок — \\d User"),
        ("describe-studyplace.png", "Рисунок — \\d StudyPlace"),
        ("describe-favorite.png", "Рисунок — \\d FavoritePlace"),
        ("prisma-studio.png", "Рисунок — Prisma Studio"),
    ):
        p = png(name)
        if p:
            add_picture(doc, p, caption, 14)

    add_heading(doc, "7 Seed и заполнение базы", break_before=True, bookmark="h7", bid=27)
    add_text(
        doc,
        "Скрипт prisma/seed.ts идемпотентен: повторный запуск использует upsert и не плодит дубликаты. Результат двух запусков: 1 пользователь, 6 площадок, 1 избранное. Пароль в отчёт не выводится — в базу пишется только bcrypt-хеш.",
    )
    add_code(doc, "Листинг 7 — seed.ts (фрагмент без секрета)", SEED_SAFE)
    add_code(doc, "Листинг 8 — count(*) после seed", read_evidence("counts.txt"))

    add_heading(doc, "8 Валидация, сессия и current user", break_before=True, bookmark="h8", bid=28)
    add_text(
        doc,
        "Модуль validation.ts нормализует email и проверяет имя, пароль и совпадение подтверждения. auth.ts подписывает JWT (HS256) через jose, кладёт в cookie study_session атрибуты httpOnly, sameSite=lax, path=/, maxAge семь дней и secure в production. current-user.ts читает сессию и повторно загружает пользователя из базы без passwordHash.",
    )
    add_code(doc, "Листинг 9 — validation.ts", VALIDATION_TS)
    add_code(doc, "Листинг 10 — auth.ts", AUTH_TS)
    add_code(doc, "Листинг 11 — current-user.ts", CURRENT_USER_TS)

    add_heading(doc, "9 Маршруты авторизации и избранного", break_before=True, bookmark="h9", bid=29)
    add_code(doc, "Листинг 12 — POST /api/auth/register", REGISTER_TS)
    add_code(doc, "Листинг 13 — POST /api/auth/login", LOGIN_TS)
    add_code(doc, "Листинг 14 — GET /api/auth/me", ME_TS)
    add_code(doc, "Листинг 15 — POST /api/auth/logout", LOGOUT_TS)
    add_code(doc, "Листинг 16 — POST/DELETE /api/favorites/[placeId]", FAVORITES_TS)

    add_heading(doc, "10 Матрица HTTP-сценариев", break_before=True, bookmark="h10", bid=30)
    add_table(
        doc,
        "Таблица 4 — Проверенные HTTP-сценарии",
        ["Сценарий", "Метод и путь", "Статус"],
        [
            ["Регистрация", "POST /api/auth/register", "201"],
            ["Дубликат email", "POST /api/auth/register", "409"],
            ["Невалидные поля", "POST /api/auth/register", "400"],
            ["Вход", "POST /api/auth/login", "200"],
            ["Неверный пароль/email", "POST /api/auth/login", "401"],
            ["Текущий пользователь (гость)", "GET /api/auth/me", "200, user: null"],
            ["Текущий пользователь (сессия)", "GET /api/auth/me", "200, user"],
            ["Выход", "POST /api/auth/logout", "200"],
            ["Добавить избранное", "POST /api/favorites/:id", "201"],
            ["Удалить избранное", "DELETE /api/favorites/:id", "200"],
            ["Избранное без сессии", "POST/DELETE /api/favorites/:id", "401"],
            ["Несуществующее место", "POST /api/favorites/:id", "404"],
        ],
        [2800, 3555, 3000],
    )

    add_heading(doc, "11 Отсутствие passwordHash в ответах", break_before=True, bookmark="h11", bid=31)
    add_text(
        doc,
        "В register и login выборка пользователя задаётся select: { id, name, email }. getCurrentUser также выбирает только id, name, email. Поле passwordHash в JSON-ответах не возвращается. Cookie содержит только подписанный токен сессии; сам JWT и его значение в отчёт не включаются.",
    )

    add_heading(doc, "12 Lint, build и ограничения учебной авторизации", break_before=True, bookmark="h12", bid=32)
    add_text(
        doc,
        "npm run lint и npm run build выполняются без ошибок. Учебная авторизация имеет ограничения: один секрет AUTH_SECRET на процесс, нет refresh-токенов и отзыва сессий на сервере списком, нет rate limit и 2FA, cookie sameSite=lax без отдельной CSRF-защиты форм, bcrypt cost 12 достаточен для учёбы, но политика паролей упрощена, нет подтверждения email.",
    )
    add_code(doc, "Листинг 17 — npm run lint", read_evidence("lint.txt"))
    add_code(doc, "Листинг 18 — .env.example (без секретов)", ENV_EXAMPLE)

    add_heading(doc, "13 Ответы на контрольные вопросы", break_before=True, bookmark="h13", bid=33)
    answers = [
        (
            "Что входит в физическую схему реляционной базы?",
            "Таблицы, столбцы, типы, первичные и внешние ключи, уникальные ограничения, индексы и правила удаления.",
        ),
        (
            "Чем модель Prisma отличается от физической таблицы?",
            "Модель описывает сущность приложения в schema.prisma. Физическая таблица — объект PostgreSQL, который создаёт SQL миграции.",
        ),
        (
            "Для чего нужен первичный ключ?",
            "Он однозначно идентифицирует строку. У User и StudyPlace это id (uuid), у FavoritePlace — пара (userId, placeId).",
        ),
        (
            "Чем уникальное ограничение отличается от обычного индекса?",
            "Уникальное ограничение запрещает дубликаты значения. Обычный индекс ускоряет поиск, но сам по себе уникальность не гарантирует.",
        ),
        (
            "Что гарантирует внешний ключ?",
            "Что значение ссылается на существующую строку родительской таблицы (и задаёт поведение при UPDATE/DELETE).",
        ),
        (
            "Почему FavoritePlace имеет составной первичный ключ?",
            "Пара пользователь–площадка должна быть уникальна: одно место нельзя добавить в избранное дважды одним пользователем.",
        ),
        (
            "Что делает onDelete: Cascade?",
            "При удалении родителя автоматически удаляет зависимые строки FavoritePlace.",
        ),
        (
            "Почему внешний ключ не всегда получает индекс автоматически?",
            "В PostgreSQL индекс по FK не создаётся сам по себе для всех случаев. Для placeId добавлен явный @@index, потому что ведущий столбец составного PK — userId.",
        ),
        (
            "Для чего нужен migration.sql?",
            "Это версионируемый SQL, который приводит базу к состоянию, описанному схемой Prisma.",
        ),
        (
            "Чем migrate dev отличается от migrate status?",
            "migrate dev создаёт/применяет миграции в разработке. migrate status только показывает, применены ли уже имеющиеся миграции.",
        ),
        (
            "Почему нельзя редактировать уже применённую миграцию без понимания последствий?",
            "История миграций должна совпадать у всех копий базы. Правка применённого файла ломает checksum и воспроизводимость.",
        ),
        (
            "Что означает идемпотентный seed?",
            "Повторный запуск даёт то же состояние данных без лишних дубликатов (здесь через upsert).",
        ),
        (
            "Чем upsert отличается от create?",
            "create всегда вставляет. upsert обновляет существующую строку по уникальному ключу или создаёт новую.",
        ),
        (
            "Почему пароль нельзя хранить в исходном виде?",
            "Утечка базы сразу раскрыла бы все пароли. Хранят необратимый хеш.",
        ),
        (
            "Что хранит bcrypt-хеш?",
            "Результат хеширования пароля с солью и cost factor (здесь 12), пригодный для compare, но не для восстановления пароля.",
        ),
        (
            "Чем аутентификация отличается от авторизации?",
            "Аутентификация подтверждает личность (login). Авторизация проверяет право выполнить действие (например, менять своё избранное).",
        ),
        (
            "Что такое stateless session?",
            "Сервер не хранит серверную таблицу сессий: состояние пользователя проверяется по подписанному токену в cookie.",
        ),
        (
            "Зачем JWT подписывается секретом?",
            "Подпись гарантирует, что payload не подделан без знания AUTH_SECRET.",
        ),
        (
            "Какие данные допустимо хранить в payload сессии?",
            "Минимум идентификаторов: userId и name. Пароль и passwordHash в payload не кладут.",
        ),
        (
            "Для чего нужны httpOnly, secure и sameSite?",
            "httpOnly закрывает доступ из JavaScript, secure ограничивает cookie HTTPS в production, sameSite снижает риск CSRF при межсайтовых запросах.",
        ),
        (
            "Почему скрытая кнопка не является авторизацией?",
            "Клиентский UI можно обойти прямым HTTP-запросом. Проверка прав должна выполняться на сервере.",
        ),
        (
            "Почему current user повторно проверяется в базе?",
            "Чтобы учесть удаление или изменение учётной записи после выдачи cookie и не доверять только содержимому токена.",
        ),
        (
            "Почему passwordHash нельзя возвращать клиенту?",
            "Хеш — секрет аутентификации. Его утечка на клиент упрощает офлайн-атаки.",
        ),
        (
            "Зачем для неверного email и неверного пароля используется одно сообщение?",
            "Чтобы не раскрывать, существует ли учётная запись с данным email.",
        ),
        (
            "Какие ограничения имеет учебная реализация?",
            "Нет refresh/отзыва сессий, rate limit, 2FA и подтверждения email; упрощённая политика паролей; один AUTH_SECRET на процесс.",
        ),
    ]
    for question, answer in answers:
        add_qa(doc, question, answer)

    add_heading(doc, "ЗАКЛЮЧЕНИЕ", center=True, break_before=True, bookmark="h14", bid=34)
    add_text(
        doc,
        "Создана физическая схема User, StudyPlace и FavoritePlace с каскадным удалением и составным ключом избранного. Миграция применена, seed идемпотентно даёт 1 пользователя, 6 площадок и 1 избранное. Реализованы register/login/me/logout на bcrypt и jose cookie, избранное защищено серверной проверкой пользователя. passwordHash в JSON не возвращается. Lint и build проходят без ошибок.",
    )
    save_report(doc, OUT6)


def build_lr7():
    doc = fresh_title(
        "7",
        "Соединение и тестирование клиентской и серверной частей веб-приложения",
    )

    add_heading(doc, "СОДЕРЖАНИЕ", center=True, bookmark="toc", bid=20)
    toc = [
        ("1 Цель работы", "h1"),
        ("2 Поток клиент → API → Prisma → PostgreSQL", "h2"),
        ("3 HTTP-контракты и DTO", "h3"),
        ("4 GET /api/places", "h4"),
        ("5 AuthProvider и формы", "h5"),
        ("6 Галерея, состояния и toggle избранного", "h6"),
        ("7 Сохранение после reload и SQL-проверка", "h7"),
        ("8 Матрица ручных сценариев", "h8"),
        ("9 Vitest и unit-тесты", "h9"),
        ("10 Playwright и E2E", "h10"),
        ("11 Итоговые команды проверки", "h11"),
        ("12 Повторная проверка ЛР3–ЛР4", "h12"),
        ("13 Ответы на контрольные вопросы", "h13"),
        ("ЗАКЛЮЧЕНИЕ", "h14"),
    ]
    for title, bookmark in toc:
        add_toc(doc, title, bookmark)

    add_heading(doc, "1 Цель работы", break_before=True, bookmark="h1", bid=21)
    add_text(
        doc,
        "Цель работы — соединить клиентские компоненты PairWise с серверными Route Handlers и PostgreSQL, заменить демонстрационное состояние пользователя и избранного реальными запросами, реализовать состояния загрузки и ошибок, проверить HTTP-контракты и добавить модульные и end-to-end тесты.",
    )

    add_heading(doc, "2 Поток клиент → API → Prisma → PostgreSQL", break_before=True, bookmark="h2", bid=22)
    add_code(
        doc,
        "Листинг 1 — Потоки чтения и избранного",
        """
PlacesGallery → GET /api/places → Prisma Client → PostgreSQL
             ← JSON places + isFavorite ←

кнопка → POST/DELETE /api/favorites/:placeId → cookie + User
       → upsert/deleteMany FavoritePlace → JSON → карточка
""",
    )

    add_heading(doc, "3 HTTP-контракты и DTO", break_before=True, bookmark="h3", bid=23)
    add_table(
        doc,
        "Таблица 1 — HTTP-контракты",
        ["Операция", "Метод", "Успех", "Ошибки"],
        [
            ["Список площадок", "GET /api/places", "200", "500"],
            ["Текущий пользователь", "GET /api/auth/me", "200", "—"],
            ["Регистрация", "POST /api/auth/register", "201", "400, 409, 500"],
            ["Вход", "POST /api/auth/login", "200", "400, 401, 500"],
            ["Выход", "POST /api/auth/logout", "200", "—"],
            ["Добавить избранное", "POST /api/favorites/:id", "201", "401, 404"],
            ["Удалить избранное", "DELETE /api/favorites/:id", "200", "401"],
        ],
        [2400, 2600, 1400, 2955],
    )
    add_text(
        doc,
        "UserDTO на клиенте — id, name, email (без passwordHash). StudyPlaceDTO / PlaceCard — id, slug, title, summary, address, image, imageAlt, latitude и longitude как number, isFavorite. Decimal координат преобразуется через Number(...) до JSON.",
    )
    add_code(doc, "Листинг 2 — PlaceCard (DTO площадки)", PLACE_TS)

    add_heading(doc, "4 GET /api/places", break_before=True, bookmark="h4", bid=24)
    add_text(
        doc,
        "Маршрут объявлен dynamic = \"force-dynamic\". isFavorite вычисляется на сервере относительно текущего пользователя и множества FavoritePlace; значение с клиента не принимается.",
    )
    add_code(doc, "Листинг 3 — app/api/places/route.ts", PLACES_TS)

    add_heading(doc, "5 AuthProvider и формы", break_before=True, bookmark="h5", bid=25)
    add_text(
        doc,
        "AuthProvider при монтировании вызывает GET /api/auth/me, держит статусы loading/authenticated/guest и отдаёт login/register/logout/refresh. Формы LoginMenu и RegisterMenu вызывают эти методы и при успехе переходят в /cabinet.",
    )
    add_code(doc, "Листинг 4 — AuthProvider.tsx", AUTH_PROVIDER)
    add_code(
        doc,
        "Листинг 5 — Фрагмент LoginMenu (вызов login)",
        """
const message = await login(email.trim(), password);
if (message) {
    setError(message);
    setPending(false);
    return;
}
router.push("/cabinet");
""",
    )
    add_code(
        doc,
        "Листинг 6 — Фрагмент RegisterMenu (вызов register)",
        """
const name = `${firstName.trim()} ${lastName.trim()}`.trim();
const message = await register({
    name,
    email: email.trim(),
    password,
    confirmPassword: passwordRepeat,
});
""",
    )

    add_heading(doc, "6 Галерея, состояния и toggle избранного", break_before=True, bookmark="h6", bid=26)
    add_text(
        doc,
        "PlacesGallery после готовности auth загружает GET /api/places. Состояния: loading, error, empty, ready. Избранное меняется пессимистично: запрос POST/DELETE, обновление Set только при response.ok; при 401 вызывается refresh. AbortController отменяет запрос при размонтировании.",
    )
    add_code(doc, "Листинг 7 — PlacesGallery.tsx", GALLERY)

    add_heading(doc, "7 Сохранение после reload и SQL-проверка", break_before=True, bookmark="h7", bid=27)
    add_text(
        doc,
        "После добавления избранного и перезагрузки страницы кнопка остаётся в состоянии «Убрать из избранного»: данные читаются из FavoritePlace. SQL-проверка count(*) подтверждает наличие строк User, StudyPlace и FavoritePlace. Скриншоты Network с Set-Cookie и значениями токена в отчёт не включаются.",
    )
    add_code(doc, "Листинг 8 — count(*)", read_evidence("counts.txt"))
    for name, caption in (
        ("network-places.png", "Рисунок — Network GET /api/places"),
        ("favorite-reload.png", "Рисунок — избранное после reload"),
    ):
        p = png(name)
        if p:
            add_picture(doc, p, caption, 14)

    add_heading(doc, "8 Матрица ручных сценариев", break_before=True, bookmark="h8", bid=28)
    add_table(
        doc,
        "Таблица 2 — Ручные положительные и отрицательные сценарии",
        ["Сценарий", "Ожидание", "Результат"],
        [
            ["Гость открывает /gallery", "places, isFavorite=false", "ок"],
            ["Регистрация", "201, кабинет", "ок"],
            ["Вход", "200, сессия", "ок"],
            ["Добавить избранное", "201, кнопка меняется", "ок"],
            ["Reload после избранного", "состояние сохранено", "ок"],
            ["Выход", "гость, избранное недоступно", "ок"],
            ["Гость POST favorites", "401", "ок"],
            ["Неверный login", "401, общее сообщение", "ок"],
            ["Дубликат email", "409", "ок"],
            ["Невалидная регистрация", "400", "ок"],
        ],
        [3200, 3200, 2955],
    )

    add_heading(doc, "9 Vitest и unit-тесты", break_before=True, bookmark="h9", bid=29)
    add_text(
        doc,
        "Unit-тесты проверяют чистую логику normalizeEmail и validateRegistration без обращения к базе. Запуск: 4 теста успешно.",
    )
    add_code(doc, "Листинг 9 — vitest.config.mts", VITEST_CFG)
    add_code(doc, "Листинг 10 — tests/unit/validation.test.ts", UNIT_TEST)
    add_code(doc, "Листинг 11 — Результат vitest", read_evidence("vitest.txt"))

    add_heading(doc, "10 Playwright и E2E", break_before=True, bookmark="h10", bid=30)
    add_text(
        doc,
        "E2E запускает production-сервер через webServer (npm run start), один worker, baseURL http://127.0.0.1:3000. Сценарий регистрирует пользователя с уникальным email, добавляет избранное, проверяет сохранение после reload и выход. afterAll удаляет тестового пользователя из базы.",
    )
    add_code(doc, "Листинг 12 — playwright.config.ts", PLAYWRIGHT_CFG)
    add_code(doc, "Листинг 13 — E2E-тест", E2E_TEST)
    add_code(doc, "Листинг 14 — Результат Playwright", read_evidence("playwright.txt"))

    add_heading(doc, "11 Итоговые команды проверки", break_before=True, bookmark="h11", bid=31)
    add_text(
        doc,
        "Последовательно подтверждены prisma validate, migrate status (схема актуальна), lint, unit (4/4), build и e2e (1/1).",
    )
    add_code(doc, "Листинг 15 — validate", read_evidence("validate.txt"))
    add_code(doc, "Листинг 16 — migrate status", read_evidence("migrate-status.txt"))
    add_code(doc, "Листинг 17 — lint", read_evidence("lint.txt"))

    add_heading(doc, "12 Повторная проверка ЛР3–ЛР4", break_before=True, bookmark="h12", bid=32)
    add_text(
        doc,
        "После соединения с API сохранены модальное окно, карта Яндекса, бургер-меню и адаптивная сетка галереи. На узкой и широкой ширине карточки читаются, диалог получает фокус, меню открывается на всю высоту видимой области. Горизонтальная прокрутка не появилась. Ручные проверки ЛР3–ЛР4 повторены на production-сборке.",
    )

    add_heading(doc, "13 Ответы на контрольные вопросы", break_before=True, bookmark="h13", bid=33)
    answers = [
        (
            "Как проходит запрос от Client Component до PostgreSQL?",
            "Компонент делает fetch к Route Handler, сервер проверяет данные и пользователя, Prisma выполняет SQL, ответ JSON возвращается в React-состояние.",
        ),
        (
            "Что такое HTTP-контракт?",
            "Соглашение о методе, пути, теле запроса, статусах и форме JSON между клиентом и сервером.",
        ),
        (
            "Где располагается Route Handler в App Router?",
            "В файле route.ts внутри сегмента app, без page.tsx на том же сегменте.",
        ),
        (
            "Какие HTTP-методы поддерживает route.ts?",
            "GET, POST, PUT, PATCH, DELETE, HEAD, OPTIONS. Неподдерживаемый метод даёт 405.",
        ),
        (
            "Почему response.ok нужно проверять вручную?",
            "fetch не бросает исключение на 4xx/5xx. Без проверки клиент примет ошибочное тело как успех.",
        ),
        (
            "Что такое DTO?",
            "Объект передачи данных: урезанное и безопасное представление сущности для границы HTTP.",
        ),
        (
            "Почему UserDTO не должен повторять всю строку User?",
            "В строке есть passwordHash и служебные поля. Клиенту нужны только id, name, email.",
        ),
        (
            "Зачем Decimal преобразуется перед JSON?",
            "JSON не имеет Decimal. Number(latitude/longitude) даёт сериализуемые числа для карты.",
        ),
        (
            "Где вычисляется isFavorite и почему?",
            "На сервере относительно текущего пользователя. Клиенту нельзя доверять собственный флаг избранного.",
        ),
        (
            "Как httpOnly cookie участвует в fetch-запросе?",
            "При credentials: \"same-origin\" браузер автоматически прикрепляет cookie сессии к запросу на тот же origin.",
        ),
        (
            "Чем loading отличается от empty?",
            "loading — запрос ещё идёт. empty — ответ успешен, но список площадок пуст.",
        ),
        (
            "Для чего нужен AbortController?",
            "Чтобы отменить незавершённый fetch при размонтировании или смене условий и не обновлять состояние устаревшим ответом.",
        ),
        (
            "Почему AuthProvider имеет состояние loading?",
            "Пока /api/auth/me не ответил, нельзя считать пользователя гостем или авторизованным — иначе будет мигание UI.",
        ),
        (
            "Чем пессимистичное обновление отличается от оптимистичного?",
            "Пессимистичное ждёт успех сервера, затем меняет UI. Оптимистичное меняет UI сразу и откатывает при ошибке. В галерее использован пессимистичный toggle.",
        ),
        (
            "Почему userId нельзя принимать от клиента для избранного?",
            "Клиент может подставить чужой идентификатор. userId берётся только из проверенной серверной сессии.",
        ),
        (
            "Чем unit test отличается от integration и E2E?",
            "Unit проверяет чистую функцию без сети и БД. Integration связывает несколько модулей. E2E проходит сценарий в браузере против реального сервера.",
        ),
        (
            "Почему async Server Components удобнее проверять E2E?",
            "Их логика выполняется на сервере при рендере страницы; проще подтвердить результат через UI и HTTP, чем изолированно мокать весь стек.",
        ),
        (
            "Что проверяет Vitest в этой работе?",
            "Нормализацию email и правила validateRegistration: четыре модульных теста.",
        ),
        (
            "Что проверяет Playwright?",
            "Сквозной сценарий: регистрация, добавление избранного, сохранение после reload, выход.",
        ),
        (
            "Для чего нужны baseURL и webServer?",
            "baseURL задаёт адрес приложения. webServer поднимает production-сервер перед тестами.",
        ),
        (
            "Почему E2E выполняется после production build?",
            "Чтобы проверять тот же режим, что ближе к сдаче, а не особенности dev-сервера.",
        ),
        (
            "Почему выбран один worker?",
            "Один worker исключает гонки нескольких тестов за одну локальную базу и cookie.",
        ),
        (
            "Зачем тесту уникальный email?",
            "Чтобы не конфликтовать с уже существующими пользователями и повторными прогонами.",
        ),
        (
            "Почему тестовые данные нужно удалять?",
            "Иначе база засоряется, а повторные запуски зависят от прошлых прогонов.",
        ),
        (
            "Почему waitForTimeout является плохой синхронизацией?",
            "Фиксированная пауза хрупка: то медленнее окружение, то зря тратит время. Нужны ожидания условий (locator, URL, response).",
        ),
        (
            "Какие проверки остаются ручными?",
            "Визуальная адаптивность, фокус диалога и меню, карта, удобство UI и негативные сценарии вне автотестов.",
        ),
        (
            "Почему проходящие тесты не доказывают полную безопасность?",
            "Тесты покрывают выбранные сценарии. Они не заменяют аудит угроз, rate limit, CSRF-стратегию и эксплуатационные проверки.",
        ),
    ]
    for question, answer in answers:
        add_qa(doc, question, answer)

    add_heading(doc, "ЗАКЛЮЧЕНИЕ", center=True, break_before=True, bookmark="h14", bid=34)
    add_text(
        doc,
        "Клиент PairWise читает площадки из PostgreSQL через GET /api/places, восстанавливает сессию через AuthProvider и /api/auth/me, сохраняет избранное серверными POST/DELETE. Состояния loading/error/empty обработаны. Vitest: 4/4. Playwright: 1/1 с очисткой тестового пользователя. validate, migrate status, lint, build и e2e завершились успешно. Модал, карта, меню и адаптивность ЛР3–ЛР4 сохранены. По курсу: от статического Next.js-клиента до локальной БД, серверной авторизации и проверенного соединения клиента с API.",
    )
    save_report(doc, OUT7)


def main():
    build_lr5()
    build_lr6()
    build_lr7()
    missing = []
    for folder, name in (
        (OUT5.parent, OUT5.name),
        (OUT6.parent, OUT6.name),
        (OUT7.parent, OUT7.name),
    ):
        path = folder / name
        if not path.exists():
            missing.append(str(path))
    pngs = list(EVIDENCE.glob("*.png"))
    if not pngs:
        print("NOTE: в lab-evidence нет png — в отчётах использованы txt-листинги")
    if missing:
        raise SystemExit("Не созданы: " + ", ".join(missing))
    print("OK", OUT5)
    print("OK", OUT6)
    print("OK", OUT7)


if __name__ == "__main__":
    main()
