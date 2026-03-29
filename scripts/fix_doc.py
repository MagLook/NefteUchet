# -*- coding: utf-8 -*-
"""Diagnose and fix INIT_STOCK doc. Run: py -3.13-32 scripts/fix_doc.py"""
import pythoncom
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

# 1. List available ВидОперации values
print("=== ВидыОперацийПоступлениеТоваровУслуг ===")
try:
    enum = ib.Перечисления.ВидыОперацийПоступлениеТоваровУслуг
    for i in range(20):
        try:
            val = enum.Получить(i)
            print(f"  [{i}] {val}")
        except:
            break
except Exception as e:
    print(f"  Enum error: {e}")
    # Try metadata
    try:
        md = ib.Метаданные.Перечисления
        for i in range(md.Количество()):
            en = md.Получить(i)
            if "Поступлени" in str(en.Имя) or "Операци" in str(en.Имя):
                print(f"  Found enum: {en.Имя}")
                vals = en.ЗначенияПеречисления
                for j in range(vals.Количество()):
                    print(f"    [{j}] {vals.Получить(j).Имя}")
    except Exception as e2:
        print(f"  Metadata error: {e2}")

# 2. Delete bad documents and create a fresh one with ВидОперации
print("\n=== Deleting old INIT_STOCK docs ===")
q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ Ссылка ИЗ Документ.ПоступлениеТоваровУслуг
ГДЕ Комментарий ПОДОБНО &Маска"""
q.УстановитьПараметр("Маска", "%INIT_STOCK%")
r = q.Выполнить().Выбрать()
while r.Следующий():
    try:
        obj = r.Ссылка.ПолучитьОбъект()
        obj.УстановитьПометкуУдаления(True)
        print(f"  Marked for deletion: {r.Ссылка}")
    except Exception as e:
        print(f"  Delete error: {e}")

# 3. Create new document properly
print("\n=== Creating new document ===")
doc = ib.Документы.ПоступлениеТоваровУслуг.СоздатьДокумент()

# Set ВидОперации FIRST (before anything else)
try:
    покупка = ib.Перечисления.ВидыОперацийПоступлениеТоваровУслуг.Покупка
    doc.ВидОперации = покупка
    print("  ВидОперации = Покупка (direct)")
except:
    try:
        # Try by index
        покупка = ib.Перечисления.ВидыОперацийПоступлениеТоваровУслуг.Получить(0)
        doc.ВидОперации = покупка
        print(f"  ВидОперации = index 0: {покупка}")
    except Exception as e:
        print(f"  ВидОперации FAILED: {e}")

from datetime import datetime
doc.Дата = datetime(2026, 3, 1)
doc.Организация = ib.Справочники.Организации.НайтиПоНаименованию("ГАЗИНВЕСТГРУПП ООО", True)

# Counterparty - try СУРГУТНЕФТЕГАЗ with partial match
ca = ib.Справочники.Контрагенты.НайтиПоНаименованию("СУРГУТНЕФТЕГАЗ", False)
if ca and not ca.Пустая():
    doc.Контрагент = ca
    print(f"  Контрагент: partial match")
else:
    # Exact search for full name
    q2 = ib.NewObject("Query")
    q2.Text = "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.Контрагенты ГДЕ Наименование ПОДОБНО &Н"
    q2.УстановитьПараметр("Н", "%СУРГУТНЕФТЕГАЗ%")
    r2 = q2.Выполнить().Выбрать()
    if r2.Следующий():
        doc.Контрагент = r2.Ссылка
        print(f"  Контрагент: query match")
    else:
        print("  Контрагент: NOT FOUND")

doc.Комментарий = "TL|INIT_STOCK|v2"

wh = ib.Справочники.Склады.НайтиПоНаименованию("АКЗС Витебский", True)
try:
    doc.Склад = wh
    print(f"  Склад: set")
except:
    print(f"  Склад: skip")

# НДС
vat = None
try: vat = ib.Перечисления.СтавкиНДС.НДС22
except:
    try: vat = ib.Перечисления.СтавкиНДС.НДС20
    except: pass

# Rows
for name, qty, price in [("АИ-92 (л)", 50000, 55), ("АИ-95 (л)", 30000, 60), ("Дизельное топливо (л)", 40000, 62)]:
    nom = ib.Справочники.Номенклатура.НайтиПоНаименованию(name, True)
    if not nom or nom.Пустая():
        print(f"  SKIP: {name}")
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
        except: pass
    print(f"  + {name}: {qty} x {price}")

# Write then post
print("\n=== Writing ===")
try:
    doc.Записать()
    print(f"  Written OK")
except Exception as e:
    print(f"  Write error: {e}")
    exit()

print("=== Posting ===")
try:
    doc.Записать(ib.РежимЗаписиДокумента.Проведение)
    print("  POSTED OK!")
except Exception as e:
    print(f"  Post error: {e}")
    print("  Document saved but not posted. Open in 1C and post manually.")
