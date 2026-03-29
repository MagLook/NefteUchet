# -*- coding: utf-8 -*-
"""Check warehouse type and retail price setup. Run: py -3.13-32 scripts/check_warehouse_type.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client
from datetime import datetime

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\wh_type.txt", "w", encoding="utf-8")

# 1. Warehouse details
q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ Наименование, Код, ТипСклада
ИЗ Справочник.Склады ГДЕ Наименование ПОДОБНО &Н"""
q.УстановитьПараметр("Н", "%АКЗС Витебский%")
r = q.Выполнить().Выбрать()
while r.Следующий():
    out.write(f"Склад: [{r.Наименование}] код:{r.Код} тип:{r.ТипСклада}\n")

# 2. All ТипСклада values
out.write("\n=== Перечисление ТипыСкладов ===\n")
try:
    enum = ib.Перечисления.ТипыСкладов
    for i in range(10):
        try:
            v = enum.Получить(i)
            # Compare with our warehouse type
            out.write(f"  [{i}] {v}\n")
        except:
            break
except Exception as e:
    out.write(f"  Ошибка: {e}\n")

# 3. ЦеныНоменклатуры - retail prices at different dates
out.write("\n=== ЦеныНоменклатуры (все записи) ===\n")
q2 = ib.NewObject("Query")
q2.Text = """ВЫБРАТЬ Период, Номенклатура.Наименование КАК Ном, ТипЦен.Наименование КАК ТипЦен, Цена
ИЗ РегистрСведений.ЦеныНоменклатуры
ГДЕ Номенклатура.Наименование ПОДОБНО &Н
УПОРЯДОЧИТЬ ПО Период УБЫВ"""
q2.УстановитьПараметр("Н", "%Витебский%")
try:
    r2 = q2.Выполнить().Выбрать()
    while r2.Следующий():
        out.write(f"  {r2.Период} [{r2.Ном}] тип:[{r2.ТипЦен}] цена:{r2.Цена}\n")
except Exception as e:
    out.write(f"  Ошибка: {e}\n")

# 4. Check if there's a "Установка цен" or "розничные цены" register
out.write("\n=== Регистры с 'Розн' или 'Наценк' ===\n")
md = ib.Метаданные.РегистрыСведений
for i in range(md.Количество()):
    reg = md.Получить(i)
    if "Розн" in reg.Имя or "Наценк" in reg.Имя or "розн" in reg.Имя or "НТТ" in reg.Имя or "АТТ" in reg.Имя:
        out.write(f"  {reg.Имя}\n")

md2 = ib.Метаданные.РегистрыНакопления
for i in range(md2.Количество()):
    reg = md2.Получить(i)
    if "Розн" in reg.Имя or "Наценк" in reg.Имя or "НТТ" in reg.Имя or "АТТ" in reg.Имя:
        out.write(f"  (накопл) {reg.Имя}\n")

# 5. Which warehouses are "Розничный" and which are "Оптовый"
out.write("\n=== Все склады по типам ===\n")
q3 = ib.NewObject("Query")
q3.Text = "ВЫБРАТЬ Наименование, ТипСклада ИЗ Справочник.Склады ГДЕ НЕ ПометкаУдаления"
r3 = q3.Выполнить().Выбрать()
while r3.Следующий():
    out.write(f"  [{r3.Наименование}]: {r3.ТипСклада}\n")

out.close()
print("Done")
