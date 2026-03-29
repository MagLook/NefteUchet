# -*- coding: utf-8 -*-
"""Check and fix station mapping. Run: py -3.13-32 scripts/fix_station.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\station_check.txt", "w", encoding="utf-8")

# 1. Current station settings
out.write("=== Настройки станции ===\n")
q = ib.NewObject("Query")
q.Text = "ВЫБРАТЬ Ключ, Значение ИЗ РегистрСведений.TL_Настройки ГДЕ Ключ ПОДОБНО &М"
q.УстановитьПараметр("М", "%Станция%")
r = q.Выполнить().Выбрать()
while r.Следующий():
    out.write(f"  {r.Ключ} = [{r.Значение}]\n")

# 2. TL_АЗС справочник
out.write("\n=== Справочник TL_АЗС ===\n")
try:
    q2 = ib.NewObject("Query")
    q2.Text = "ВЫБРАТЬ КодСтанции, Наименование, Склад.Наименование КАК СкладНаим ИЗ Справочник.TL_АЗС ГДЕ НЕ ПометкаУдаления"
    r2 = q2.Выполнить().Выбрать()
    while r2.Следующий():
        out.write(f"  Код:{r2.КодСтанции} Имя:[{r2.Наименование}] Склад:[{r2.СкладНаим}]\n")
except Exception as e:
    out.write(f"  Ошибка: {e}\n")

# 3. All warehouses with "5" or "Витебский" or "АЗС"
out.write("\n=== Склады ===\n")
q3 = ib.NewObject("Query")
q3.Text = """ВЫБРАТЬ Наименование, ТипСклада ИЗ Справочник.Склады
ГДЕ (Наименование ПОДОБНО &Н1 ИЛИ Наименование ПОДОБНО &Н2 ИЛИ Наименование ПОДОБНО &Н3)
И НЕ ПометкаУдаления"""
q3.УстановитьПараметр("Н1", "%Витебский%")
q3.УстановитьПараметр("Н2", "%АЗС 5%")
q3.УстановитьПараметр("Н3", "%АЗС-5%")
r3 = q3.Выполнить().Выбрать()
while r3.Следующий():
    out.write(f"  [{r3.Наименование}] тип: {r3.ТипСклада}\n")

out.close()

# 4. FIX: set correct warehouse in settings
rec = ib.РегистрыСведений.TL_Настройки.СоздатьМенеджерЗаписи()
rec.Ключ = "Станция_5_Склад"
rec.Значение = "АКЗС Витебский"
rec.Описание = "Склад станции 5"
rec.Записать(True)
print("Fixed: Станция_5_Склад = АКЗС Витебский")

print("Details saved to build/station_check.txt")
