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


def escape_bsl(s: str) -> str:
    """Экранирование строки для BSL-литерала: удвоить кавычки, разбить по строкам."""
    # В BSL многострочные литералы — через | в начале каждой строки
    # либо конкатенация. Используем массив строк + СтрСоединить.
    return s.replace('"', '""')


def split_to_bsl_lines(s: str, max_len: int = 800) -> list[str]:
    """Разбить длинную строку на куски ≤ max_len, чтобы не упереться в лимит BSL-литерала."""
    out = []
    remaining = s
    while len(remaining) > max_len:
        # Найти разрыв по \n как можно ближе к max_len
        cut = remaining.rfind("\n", 0, max_len)
        if cut < max_len // 2:
            cut = max_len
        out.append(remaining[:cut])
        remaining = remaining[cut:]
    if remaining:
        out.append(remaining)
    return out


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
        # Разбить md на части не длиннее ~800 символов для безопасных BSL-литералов
        parts = split_to_bsl_lines(md, max_len=800)
        # Сформировать конкатенацию строк
        bsl_chunks = []
        for p in parts:
            escaped = escape_bsl(p)
            bsl_chunks.append(f'"{escaped}"')
        if len(bsl_chunks) == 1:
            md_expr = bsl_chunks[0]
        else:
            # Многострочный литерал через " + СимвПС + " + ...
            # Соединяем через + для конкатенации
            md_expr = "\n\t\t+ ".join(bsl_chunks)

        lines.append(f'\t// --- {key} ---')
        lines.append(f'\tСтатьи.Вставить("{key}", Новый Структура(')
        lines.append(f'\t\t"Заголовок, Группа, Markdown",')
        lines.append(f'\t\t"{escape_bsl(title)}",')
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
