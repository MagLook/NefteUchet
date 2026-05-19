# -*- coding: utf-8 -*-
"""Прочитать и переустановить КаталогOwnCloud_ЦБ в БП ГИГ."""
import sys, os
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _tl_config import connect

conn = connect()

# 1. Прочитать
текущ = conn.TL_Настройки.ПолучитьЗначение("КаталогOwnCloud_ЦБ", "")
print(f"Текущее значение КаталогOwnCloud_ЦБ: {текущ!r}")

# 2. Где сейчас наши свежие файлы
for каталог in (r"C:\TL_BP_Export", r"D:\TL_BP_Export"):
    if os.path.exists(каталог):
        файлы = [f for f in os.listdir(каталог) if f.endswith(".json")]
        print(f"  {каталог}: существует, .json={len(файлы)}")
    else:
        print(f"  {каталог}: НЕТ")

# 3. Установить на C:\TL_BP_Export\
новый = r"C:\TL_BP_Export\\"
print(f"\nУстанавливаю: {новый!r}")
conn.TL_Настройки.УстановитьЗначение("КаталогOwnCloud_ЦБ", новый,
    "Каталог обмена ЦБ ↔ БП ГИГ (стенд: локальный C:\\)")

# 4. Прочитать обратно
текущ2 = conn.TL_Настройки.ПолучитьЗначение("КаталогOwnCloud_ЦБ", "")
print(f"Прочитано после Set: {текущ2!r}")
