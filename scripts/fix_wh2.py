# -*- coding: utf-8 -*-
"""Try to change warehouse type via DataExchange. Run: py -3.13-32 scripts/fix_wh2.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\fix_wh.txt", "w", encoding="utf-8")

wh = ib.Справочники.Склады.НайтиПоНаименованию("АКЗС Витебский", True)
obj = wh.ПолучитьОбъект()

out.write(f"Текущий тип: {obj.ТипСклада}\n")
obj.ТипСклада = ib.Перечисления.ТипыСкладов.ОптовыйСклад

# Попробуем через ОбменДанными.Загрузка = Истина (обходит проверки)
try:
    obj.ОбменДанными.Загрузка = True
    obj.Записать()
    out.write("Записано через ОбменДанными!\n")
except Exception as e:
    out.write(f"ОбменДанными: {e}\n")
    # Вернём назад
    try:
        obj2 = wh.ПолучитьОбъект()
        obj2.ТипСклада = ib.Перечисления.ТипыСкладов.ОптовыйСклад
        obj2.ДополнительныеСвойства.Вставить("ПропуститьПроверкуТипаСклада", True)
        obj2.Записать()
        out.write("Записано с доп.свойством!\n")
    except Exception as e2:
        out.write(f"Доп.свойство: {e2}\n")

out.close()
print("Done")
