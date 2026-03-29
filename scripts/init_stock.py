# -*- coding: utf-8 -*-
"""
Инициализация начальных остатков топлива на складе АЗС.
Запуск: py -3.13-32 scripts/init_stock.py
"""
import sys, os
os.environ['PYTHONIOENCODING'] = 'utf-8'

import pythoncom
pythoncom.CoInitialize()
import win32com.client

DB_PATH = r"D:\Users\magsp\GIG Base2"
DB_USER = r"Гайворонская Татьяна"
DB_PWD = "12345"

FUEL = [
    ("АИ-92 (л)",              50000, 55.00),
    ("АИ-95 (л)",              30000, 60.00),
    ("Дизельное топливо (л)",  40000, 62.00),
]
WAREHOUSE = "АКЗС Витебский"
ORG_NAME  = "ГАЗИНВЕСТГРУПП ООО"

print("1. Connecting...")
conn = win32com.client.Dispatch("V83.COMConnector")
connStr = 'File="{}";Usr="{}";Pwd="{}"'.format(DB_PATH, DB_USER, DB_PWD)
print("   ConnStr:", connStr)
ib = conn.Connect(connStr)
print("   OK, connected")

# --- Org ---
org = ib.Справочники.Организации.НайтиПоНаименованию(ORG_NAME, True)
print("2. Org:", str(org) if org else "NOT FOUND")

# --- Warehouse ---
wh = ib.Справочники.Склады.НайтиПоНаименованию(WAREHOUSE, True)
print("3. Warehouse:", str(wh) if wh else "NOT FOUND")

# --- Counterparty ---
ca = None
for n in ["СУРГУТНЕФТЕГАЗ ПАО", "СУРГУТНЕФТЕГАЗ", "БАЛТОП АО", "БАЛТОП"]:
    ca = ib.Справочники.Контрагенты.НайтиПоНаименованию(n, False)
    if ca and not ca.Пустая():
        break
    ca = None
if not ca:
    q = ib.NewObject("Query")
    q.Text = "SELECT TOP 1 Ref FROM Catalog.Контрагенты WHERE NOT DeletionMark"
    r = q.Execute().Choose()
    if r.Next():
        ca = r.Ref
print("4. Counterparty:", str(ca) if ca else "NOT FOUND")

# --- VAT ---
vat = None
try:
    vat = ib.Перечисления.СтавкиНДС.НДС22
except:
    try:
        vat = ib.Перечисления.СтавкиНДС.НДС20
    except:
        pass
print("5. VAT:", str(vat) if vat else "NOT FOUND")

# --- Create document ---
print("\n6. Creating document...")
doc = ib.Документы.ПоступлениеТоваровУслуг.СоздатьДокумент()
from datetime import datetime
doc.Дата = datetime(2026, 3, 28, 12, 0, 0)
doc.Организация = org
doc.Контрагент = ca
doc.Комментарий = "TL|INIT_STOCK"

try:
    doc.ВидОперации = ib.Перечисления.ВидыОперацийПоступлениеТоваровУслуг.Покупка
    print("   ВидОперации: Покупка")
except Exception as e:
    print("   ВидОперации skip:", e)

try:
    doc.Склад = wh
    print("   Склад:", WAREHOUSE)
except Exception as e:
    print("   Склад skip:", e)

for name, qty, price in FUEL:
    nom = ib.Справочники.Номенклатура.НайтиПоНаименованию(name, True)
    if not nom or nom.Пустая():
        print("   SKIP:", name, "- not found")
        continue

    row = doc.Товары.Добавить()
    row.Номенклатура = nom
    row.Количество = qty
    row.Цена = price
    row.Сумма = qty * price
    if vat:
        try:
            row.СтавкаНДС = vat
            row.СуммаНДС = round(qty * price * 22 / 122, 2)
        except:
            pass
    print("   +", name, qty, "x", price)

print("\n7. Writing document...")
try:
    doc.Записать()
    ref = doc.Ссылка
    print("   Written OK. Ref:", str(ref))
    print("   IsEmpty:", ref.Пустая())

    # Try to read back
    q = ib.NewObject("Query")
    q.Text = "SELECT TOP 1 Ref, Number, Date FROM Document.ПоступлениеТоваровУслуг WHERE Comment LIKE '%INIT_STOCK%' ORDER BY Date DESC"
    r = q.Execute().Choose()
    if r.Next():
        print("   Found in DB: No.", r.Number, "Date:", r.Date)
    else:
        print("   NOT FOUND in DB after write!")
except Exception as e:
    print("   WRITE ERROR:", e)

print("\nDone.")
