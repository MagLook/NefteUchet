# -*- coding: utf-8 -*-
"""Debug: проверка переопределения склада."""
import sys, io, win32com.client
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _tl_config import BASE_PATH, USER, PWD
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
com = win32com.client.Dispatch("V83.COMConnector")
conn = com.Connect(f'File="{BASE_PATH}";Usr="{USER}";Pwd="{PWD}";')

# 1) Меняем viacard.Склад на VIAcard_TEST
тест = conn.Справочники.Склады.НайтиПоНаименованию("VIAcard_TEST", True)
пуст = True
try: пуст = bool(тест.Пустая())
except: pass
print(f"VIAcard_TEST.Пустая() = {пуст}")
if пуст:
    obj = conn.Справочники.Склады.СоздатьЭлемент()
    obj.Наименование = "VIAcard_TEST"
    obj.Записать()
    тест = obj.Ссылка
print(f"VIAcard_TEST: {тест}, наименование: {тест.Наименование}")

мн = conn.РегистрыСведений.TL_МаппингОплат.СоздатьМенеджерЗаписи()
мн.ОбразецИмени = "viacard"
мн.Прочитать()
исх = мн.Склад
print(f"viacard.Склад исходный: {исх}")
мн.Склад = тест
мн.Записать()
print(f"viacard.Склад установлен: {мн.Склад}")

# 2) Проверяем через НайтиЗаписьМаппинга
зап = conn.TL_СозданиеДокументов.НайтиЗаписьМаппинга("viacard")
print(f"\nЧерез НайтиЗаписьМаппинга:")
print(f"  Канал: {зап.КаналОплаты}")
print(f"  Склад: {зап.Склад}")
print(f"  Склад.Наименование: {зап.Склад.Наименование}")
print(f"  Склад.УникальныйИдентификатор: {зап.Склад.УникальныйИдентификатор()}")

# 3) Дополнительно — Прочитать назад
мн2 = conn.РегистрыСведений.TL_МаппингОплат.СоздатьМенеджерЗаписи()
мн2.ОбразецИмени = "viacard"
мн2.Прочитать()
print(f"\nЧерез менеджер записи: viacard.Склад = {мн2.Склад}")
print(f"  Наименование: {мн2.Склад.Наименование}")

# Откат
мн.Склад = исх
мн.Записать()
print(f"\nВосстановлено: viacard.Склад = {мн.Склад}")
