# -*- coding: utf-8 -*-
"""Verify all init chain docs and stock. Run: py -3.13-32 scripts/verify_all.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\verify.txt", "w", encoding="utf-8")

# 1. Find INIT documents
out.write("=== Документы TL|INIT ===\n")
q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ Ссылка, Номер, Дата, Проведен, Комментарий
ИЗ Документ.ПоступлениеТоваровУслуг
ГДЕ Комментарий ПОДОБНО &М И НЕ ПометкаУдаления
УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q.УстановитьПараметр("М", "%TL|INIT%")
r = q.Выполнить().Выбрать()
while r.Следующий():
    out.write(f"  Поступление: {r.Номер} {r.Дата} Проведен={r.Проведен} [{r.Комментарий}]\n")

q2 = ib.NewObject("Query")
q2.Text = """ВЫБРАТЬ Ссылка, Номер, Дата, Проведен, Комментарий
ИЗ Документ.ПеремещениеТоваров
ГДЕ Комментарий ПОДОБНО &М И НЕ ПометкаУдаления
УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q2.УстановитьПараметр("М", "%TL|INIT%")
r2 = q2.Выполнить().Выбрать()
while r2.Следующий():
    out.write(f"  Перемещение: {r2.Номер} {r2.Дата} Проведен={r2.Проведен} [{r2.Комментарий}]\n")

q3 = ib.NewObject("Query")
q3.Text = """ВЫБРАТЬ Ссылка, Номер, Дата, Проведен, Комментарий
ИЗ Документ.КомплектацияНоменклатуры
ГДЕ Комментарий ПОДОБНО &М И НЕ ПометкаУдаления
УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q3.УстановитьПараметр("М", "%TL|INIT%")
r3 = q3.Выполнить().Выбрать()
while r3.Следующий():
    out.write(f"  Комплектация: {r3.Номер} {r3.Дата} Проведен={r3.Проведен} [{r3.Комментарий}]\n")

# 2. Stock on warehouses
out.write("\n=== Остатки на складах ===\n")
for wh_name in ["Основной склад", "АКЗС Витебский"]:
    out.write(f"\n  Склад: {wh_name}\n")
    q4 = ib.NewObject("Query")
    q4.Text = """ВЫБРАТЬ
        Номенклатура.Наименование КАК Наим,
        КоличествоОстаток КАК Кол,
        СтоимостьОстаток КАК Стоимость
    ИЗ РегистрНакопления.ТоварыНаСкладах.Остатки(, Склад.Наименование = &С)
    ГДЕ КоличествоОстаток <> 0"""
    q4.УстановитьПараметр("С", wh_name)
    try:
        r4 = q4.Выполнить().Выбрать()
        found = False
        while r4.Следующий():
            found = True
            out.write(f"    {r4.Наим}: {r4.Кол} (стоимость: {r4.Стоимость})\n")
        if not found:
            out.write("    ПУСТО\n")
    except Exception as e:
        out.write(f"    Ошибка: {e}\n")

out.close()
print("Results saved to build/verify.txt")
