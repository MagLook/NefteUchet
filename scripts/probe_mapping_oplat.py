# -*- coding: utf-8 -*-
"""Проверка содержимого регистра TL_МаппингОплат + TL_Настройки виды оплат."""
import sys, io, win32com.client
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _tl_config import BASE_PATH, USER, PWD
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
com = win32com.client.Dispatch("V83.COMConnector")
conn = com.Connect(f'File="{BASE_PATH}";Usr="{USER}";Pwd="{PWD}";')

# 1. Регистр TL_МаппингОплат
print("="*80)
print("TL_МаппингОплат:")
print("="*80)
try:
    q = conn.NewObject("Запрос")
    q.Текст = ("ВЫБРАТЬ * ИЗ РегистрСведений.TL_МаппингОплат "
               "УПОРЯДОЧИТЬ ПО Приоритет")
    sel = q.Выполнить().Выбрать()
    n = 0
    while sel.Следующий():
        n += 1
        print(f"  Приоритет={sel.Приоритет}  Образец='{sel.ОбразецИмени}'  Канал='{sel.КаналОплаты}'")
    if n == 0:
        print("  (пусто)")
except Exception as e:
    print(f"  ОШИБКА: {e}")

# 2. TL_Настройки — все ключи Оплата_*
print()
print("="*80)
print("TL_Настройки — ключи про оплаты/каналы:")
print("="*80)
q = conn.NewObject("Запрос")
q.Текст = "ВЫБРАТЬ Ключ, Значение ИЗ РегистрСведений.TL_Настройки УПОРЯДОЧИТЬ ПО Ключ"
sel = q.Выполнить().Выбрать()
while sel.Следующий():
    k = str(sel.Ключ)
    if any(s in k.lower() for s in ("оплат", "канал", "карт", "склад", "виртуал")):
        v = str(sel.Значение)
        print(f"  {k:50} = {v}")
