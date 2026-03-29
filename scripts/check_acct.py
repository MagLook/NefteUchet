# -*- coding: utf-8 -*-
"""Check account settings for fuel. Run: py -3.13-32 scripts/check_acct.py"""
import pythoncom
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

# Check СчетаУчетаНоменклатуры register
q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ
    Номенклатура.Наименование КАК Номенклатура,
    Склад.Наименование КАК Склад,
    ТипСклада,
    СчетУчетаБУ.Код КАК СчетУчета,
    СчетПередачиБУ.Код КАК СчетПередачи
ИЗ РегистрСведений.СчетаУчетаНоменклатуры
ГДЕ Номенклатура.Наименование ПОДОБНО &Маска
    ИЛИ Номенклатура = ЗНАЧЕНИЕ(Справочник.Номенклатура.ПустаяСсылка)"""
q.УстановитьПараметр("Маска", "%92%")

try:
    r = q.Выполнить().Выбрать()
    print("=== СчетаУчетаНоменклатуры ===")
    while r.Следующий():
        print(f"  {r.Номенклатура or '(all)'} | Склад:{r.Склад or '(all)'} | Тип:{r.ТипСклада} | Счёт:{r.СчетУчета} | Передача:{r.СчетПередачи}")
except Exception as e:
    print(f"Query 1 error: {e}")

# Check warehouse type
print("\n=== Тип склада АКЗС Витебский ===")
q2 = ib.NewObject("Query")
q2.Text = """ВЫБРАТЬ Наименование, ТипСклада
ИЗ Справочник.Склады
ГДЕ Наименование ПОДОБНО &Н"""
q2.УстановитьПараметр("Н", "%Витебский%")
try:
    r2 = q2.Выполнить().Выбрать()
    while r2.Следующий():
        print(f"  {r2.Наименование}: ТипСклада = {r2.ТипСклада}")
except Exception as e:
    print(f"Query 2 error: {e}")

# Check what account resolves for АИ-92 (л) on АКЗС Витебский
print("\n=== Checking specific account resolution ===")
nom = ib.Справочники.Номенклатура.НайтиПоНаименованию("АИ-92 (л)", True)
wh = ib.Справочники.Склады.НайтиПоНаименованию("АКЗС Витебский", True)
print(f"  Nomenclature: {nom}, empty: {nom.Пустая() if nom else True}")
print(f"  Warehouse: {wh}, empty: {wh.Пустая() if wh else True}")

# Try to get account via БухгалтерскийУчет module
try:
    acct = ib.БухгалтерскийУчет.СчетУчетаПоУмолчанию("ТоварыНаСкладах", nom, wh)
    print(f"  Default account: {acct}")
except Exception as e:
    print(f"  БухгалтерскийУчет error: {e}")
