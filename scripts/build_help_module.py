# -*- coding: utf-8 -*-
"""H-направление плана luminous-humming-curry: сборка контента помощи в BSL-модуль.

Читает docs/help/*.md (кроме README), собирает в один общий модуль
src/CommonModules/TL_ПомощьКонтент.bsl с функцией ПолучитьВсеСтатьи().

Структура результата (в BSL):
    Соответствие(Ключ → Структура(Заголовок, Группа, Markdown))

Где:
    Ключ      - имя файла без .md (например "загрузка_пакета")
    Заголовок - первая строка # ... из .md
    Группа    - префикс ключа: "checklist", "faq", "gloss" или "process"
    Markdown  - всё содержимое файла (.md как есть, парсится в HTML на лету)

Запуск: py -3.13 build_help_module.py
"""
import io
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HELP_DIR = os.path.join(ROOT, "docs", "help")
OUT_BSL = os.path.join(ROOT, "src", "CommonModules", "TL_ПомощьКонтент.bsl")


def detect_group(key: str) -> str:
    if key.startswith("checklist_"):
        return "checklist"
    if key.startswith("faq_"):
        return "faq"
    if key.startswith("gloss_"):
        return "gloss"
    return "process"


def first_title(md: str) -> str:
    for line in md.splitlines():
        line = line.strip()
        if line.startswith("# "):
            return line[2:].strip()
    return "(без заголовка)"


def escape_bsl_line(s: str) -> str:
    """Экранировать ОДНУ строку без переноса для BSL-литерала: удвоить кавычки."""
    return s.replace('"', '""')


def split_long_line(s: str, max_len: int = 600) -> list[str]:
    """Если одна строка длиннее max_len — разбить по пробелам на куски."""
    if len(s) <= max_len:
        return [s]
    out = []
    remaining = s
    while len(remaining) > max_len:
        cut = remaining.rfind(" ", 0, max_len)
        if cut < max_len // 2:
            cut = max_len
        out.append(remaining[:cut])
        remaining = remaining[cut:]
    if remaining:
        out.append(remaining)
    return out


def md_to_bsl_expression(md: str) -> str:
    """Преобразовать Markdown-текст в BSL-выражение конкатенации строк через Символы.ПС."""
    if not md:
        return '""'
    lines = md.split("\n")
    chunks = []
    for ln in lines:
        for piece in split_long_line(ln, max_len=600):
            chunks.append(f'"{escape_bsl_line(piece)}"')
    if len(chunks) == 1:
        return chunks[0]
    # Соединяем через + Символы.ПС +
    return "\n\t\t+ Символы.ПС + ".join(chunks)


def main():
    files = sorted(
        f for f in os.listdir(HELP_DIR)
        if f.endswith(".md") and f.lower() != "readme.md"
    )
    print(f"Найдено .md файлов: {len(files)}")

    articles = []
    for fname in files:
        key = fname[:-3]  # без .md
        with open(os.path.join(HELP_DIR, fname), "r", encoding="utf-8") as f:
            md = f.read()
        title = first_title(md)
        group = detect_group(key)
        articles.append({"key": key, "title": title, "group": group, "md": md})
        print(f"  {group:9} {key:45} = {title}")

    # Сборка BSL
    lines = [
        "////////////////////////////////////////////////////////////////////////////////",
        "// Общий модуль TL_ПомощьКонтент",
        "// Расширение TradeLedger для 1С:БП 3.0",
        "// Контекст: Сервер",
        "//",
        "// АВТОГЕНЕРИРУЕМЫЙ МОДУЛЬ — не править вручную.",
        "// Источник: docs/help/*.md (26 файлов).",
        "// Пересборка: py -3.13 scripts/build_help_module.py",
        "////////////////////////////////////////////////////////////////////////////////",
        "",
        "#Область ПрограммныйИнтерфейс",
        "",
        "// Получить все статьи помощи как Соответствие(Ключ → Структура).",
        "//",
        "// Возвращаемое значение:",
        "//   Соответствие: Ключ (Строка) → Структура(Заголовок, Группа, Markdown)",
        "//",
        "Функция ПолучитьВсеСтатьи() Экспорт",
        "",
        "\tСтатьи = Новый Соответствие;",
        "",
    ]

    for art in articles:
        key = art["key"]
        title = art["title"]
        group = art["group"]
        md = art["md"]
        md_expr = md_to_bsl_expression(md)

        lines.append(f'\t// --- {key} ---')
        lines.append(f'\tСтатьи.Вставить("{key}", Новый Структура(')
        lines.append(f'\t\t"Заголовок, Группа, Markdown",')
        lines.append(f'\t\t"{escape_bsl_line(title)}",')
        lines.append(f'\t\t"{group}",')
        lines.append(f'\t\t{md_expr}')
        lines.append(f'\t));')
        lines.append("")

    lines.extend([
        "\tВозврат Статьи;",
        "",
        "КонецФункции",
        "",
        "// Получить одну статью по ключу.",
        "//",
        "Функция ПолучитьСтатью(Ключ) Экспорт",
        "\tВсе = ПолучитьВсеСтатьи();",
        "\tВозврат Все.Получить(Ключ);",
        "КонецФункции",
        "",
        "#КонецОбласти",
        "",
    ])

    # Запись с UTF-8 BOM (для BSL-модулей с кириллицей)
    bsl_text = "\n".join(lines)
    with open(OUT_BSL, "w", encoding="utf-8-sig", newline="\r\n") as f:
        f.write(bsl_text)

    print()
    print(f"Записано: {OUT_BSL}")
    print(f"Размер:   {os.path.getsize(OUT_BSL):,} байт")
    print(f"Статей:   {len(articles)}")


if __name__ == "__main__":
    main()
