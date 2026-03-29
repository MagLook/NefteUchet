# -*- coding: utf-8 -*-
"""Verify init_stock document in DB. Run: py -3.13-32 scripts/verify_stock.py"""
import pythoncom
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

# 1. Find document by comment
q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 5
    Ссылка, Номер, Дата, Организация.Наименование КАК Орг,
    Контрагент.Наименование КАК Контр, Комментарий, Проведен
ИЗ Документ.ПоступлениеТоваровУслуг
ГДЕ Комментарий ПОДОБНО &Маска
УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q.УстановитьПараметр("Маска", "%INIT_STOCK%")
r = q.Выполнить().Выбрать()
found = 0
while r.Следующий():
    found += 1
    print(f"Doc #{found}: No.{r.Номер} Date:{r.Дата} Org:{r.Орг} Contr:{r.Контр} Posted:{r.Проведен}")
    print(f"  Comment: {r.Комментарий}")

    # Check rows
    doc = r.Ссылка.ПолучитьОбъект()
    for i in range(doc.Товары.Количество()):
        row = doc.Товары.Получить(i)
        print(f"  Row {i+1}: {row.Номенклатура} qty={row.Количество} price={row.Цена} sum={row.Сумма}")

    # Check warehouse
    try:
        print(f"  Склад: {doc.Склад}")
    except:
        print("  Склад: NOT SET")

    try:
        print(f"  ВидОперации: {doc.ВидОперации}")
    except:
        print("  ВидОперации: NOT SET")

if found == 0:
    print("Document TL|INIT_STOCK NOT FOUND!")
    print("\nSearching last 3 documents...")
    q2 = ib.NewObject("Query")
    q2.Text = """ВЫБРАТЬ ПЕРВЫЕ 3
        Номер, Дата, Контрагент.Наименование КАК Контр, Комментарий, Проведен
    ИЗ Документ.ПоступлениеТоваровУслуг
    УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
    r2 = q2.Выполнить().Выбрать()
    while r2.Следующий():
        print(f"  No.{r2.Номер} {r2.Дата} {r2.Контр} [{r2.Комментарий}] posted={r2.Проведен}")

# 2. Check stock on warehouse
print("\n--- Stock on АКЗС Витебский ---")
q3 = ib.NewObject("Query")
q3.Text = """ВЫБРАТЬ
    Номенклатура.Наименование КАК Наим,
    КоличествоОстаток КАК Остаток
ИЗ РегистрНакопления.ТоварыНаСкладах.Остатки(, Склад.Наименование = &Склад)"""
q3.УстановитьПараметр("Склад", "АКЗС Витебский")

try:
    r3 = q3.Выполнить().Выбрать()
    any_stock = False
    while r3.Следующий():
        if r3.Остаток != 0:
            print(f"  {r3.Наим}: {r3.Остаток}")
            any_stock = True
    if not any_stock:
        print("  EMPTY - no stock!")
except Exception as e:
    print(f"  Query error: {e}")
