# -*- coding: utf-8 -*-
"""Удалить запись 'безнал' из TL_МаппингОплат (опасный шаблон, ловит и корп-имена)."""
import sys, io, win32com.client
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _tl_config import BASE_PATH, USER, PWD
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

com = win32com.client.Dispatch("V83.COMConnector")
conn = com.Connect(f'File="{BASE_PATH}";Usr="{USER}";Pwd="{PWD}";')

мн = conn.РегистрыСведений.TL_МаппингОплат.СоздатьМенеджерЗаписи()
мн.ОбразецИмени = "безнал"
мн.Прочитать()
if мн.Выбран():
    мн.Удалить()
    print("Удалена запись: безнал")
else:
    print("Запись безнал не найдена")

# Дополнительно проверю: нет ли других опасных шаблонов в регистре
print("\nТекущий состав маппинга:")
q = conn.NewObject("Запрос")
q.Текст = "ВЫБРАТЬ ОбразецИмени, КаналОплаты, Склад ИЗ РегистрСведений.TL_МаппингОплат УПОРЯДОЧИТЬ ПО ОбразецИмени"
sel = q.Выполнить().Выбрать()
while sel.Следующий():
    склад = "" if not sel.Склад or (hasattr(sel.Склад, "Пустая") and sel.Склад.Пустая()) else str(sel.Склад)
    print(f"  {sel.ОбразецИмени:14} → {str(sel.КаналОплаты):14} → {склад}")
