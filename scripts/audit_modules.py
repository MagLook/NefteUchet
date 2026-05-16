# -*- coding: utf-8 -*-
"""
Аудит вызовов общих модулей TL_*: для каждого вызова TL_X.Method проверяем
что Method реально определён в файле CommonModules/TL_X.bsl.
"""
import re
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(r"D:\Users\magsp\ELSYPLUS\NefteUchet\src")

RE_DEF = re.compile(r"^\s*(?:Процедура|Функция)\s+([\w]+)", re.UNICODE)
RE_CALL = re.compile(r"(TL_[\w]+)\s*\.\s*([\w]+)\s*\(", re.UNICODE)
RE_COMMENT = re.compile(r"^\s*//")

# 1. Собрать таблицу: имя модуля → множество определённых процедур/функций
defined = {}
for bsl in ROOT.rglob("*.bsl"):
    parent = bsl.parent
    # Имя модуля
    if parent.name == "CommonModules":
        module = bsl.stem  # TL_Маппинг
    elif parent.parent.name == "CommonModules":
        # CommonModules/TL_X/Module.bsl — этот вариант тоже учитываем
        module = parent.name
    else:
        # Это форма/документ — определяем отдельно для своей зоны
        if "DataProcessors" in str(parent):
            # DataProcessors\X\Forms\Y\Module.bsl
            parts = parent.parts
            try:
                dp_idx = parts.index("DataProcessors")
                module = "DP." + parts[dp_idx + 1]
                if "Forms" in parts:
                    f_idx = parts.index("Forms")
                    module += ".Form." + parts[f_idx + 1]
            except (ValueError, IndexError):
                module = bsl.stem
        else:
            module = bsl.stem

    txt = bsl.read_text(encoding="utf-8", errors="replace")
    for ln in txt.splitlines():
        m = RE_DEF.match(ln)
        if m:
            defined.setdefault(module, set()).add(m.group(1))

# 2. Найти все вызовы TL_X.Method и проверить
broken = {}
for bsl in ROOT.rglob("*.bsl"):
    txt = bsl.read_text(encoding="utf-8", errors="replace")
    for lineno, ln in enumerate(txt.splitlines(), 1):
        if RE_COMMENT.match(ln):
            continue
        for mc in RE_CALL.finditer(ln):
            mod = mc.group(1)
            method = mc.group(2)
            # Найти определение
            if mod in defined and method in defined[mod]:
                continue
            # Также модуль может быть представлен как `TL_X.bsl` файл — там уже есть
            key = mod + "." + method
            broken.setdefault(key, []).append(
                (str(bsl.relative_to(ROOT)), lineno, ln.strip())
            )

# 3. Вывод
total = 0
for key in sorted(broken):
    print("\nBROKEN:", key)
    for f, lineno, line in broken[key]:
        print(f"  {f}:{lineno}  {line}")
        total += 1

print(f"\n=== Total broken call sites: {total} ===")
print(f"=== Total broken module.method pairs: {len(broken)} ===")

# Выведу также список определённых модулей TL_*
print("\n=== Defined TL_* modules (with proc counts): ===")
for mod in sorted(defined):
    if not mod.startswith("TL_"):
        continue
    print(f"  {mod}: {len(defined[mod])} processes")
