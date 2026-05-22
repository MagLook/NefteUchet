# -*- coding: utf-8 -*-
"""Проверить какие номенклатуры топлива есть в Catalog.Номенклатура БП ГИГ."""
import os, sys
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _tl_config import connect

conn = connect()


def s(v):
    if v is None: return "—"
    try: return str(v)
    except Exception: return "?"


print("Топлива в Catalog.Номенклатура (поиск по словам):")
ПОИСК = ["АИ-92", "АИ-95", "АИ-98", "АИ-100", "Дизельн", "ДТ", "СУГ", "Газ", "Бензин"]
for термин in ПОИСК:
    q = conn.NewObject("Запрос")
    q.УстановитьПараметр("Т", "%" + термин + "%")
    q.Текст = """
    ВЫБРАТЬ Наименование, Артикул
    ИЗ Справочник.Номенклатура
    ГДЕ Наименование ПОДОБНО &Т
      И НЕ ПометкаУдаления
    УПОРЯДОЧИТЬ ПО Наименование
    """
    в = q.Выполнить().Выбрать()
    нашли = []
    while в.Следующий():
        нашли.append(s(в.Наименование))
    print(f"\n  '{термин}' → {len(нашли)} наим.")
    for н in нашли[:10]:
        print(f"    • {н}")
    if len(нашли) > 10:
        print(f"    ... ещё {len(нашли)-10}")
