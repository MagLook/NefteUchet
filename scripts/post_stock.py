# -*- coding: utf-8 -*-
"""Fix and post the INIT_STOCK document. Run: py -3.13-32 scripts/post_stock.py"""
import pythoncom
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка
ИЗ Документ.ПоступлениеТоваровУслуг
ГДЕ Комментарий ПОДОБНО &Маска И НЕ Проведен
УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q.УстановитьПараметр("Маска", "%INIT_STOCK%")
r = q.Выполнить().Выбрать()

if not r.Следующий():
    print("No unposted INIT_STOCK found")
    exit()

doc = r.Ссылка.ПолучитьОбъект()
print("Found doc, rows:", doc.Товары.Количество())

# Fix ВидОперации
try:
    doc.ВидОперации = ib.Перечисления.ВидыОперацийПоступлениеТоваровУслуг.Покупка
    print("ВидОперации set to Покупка")
except Exception as e:
    print("ВидОперации error:", e)

# Fix Склад if empty
wh = ib.Справочники.Склады.НайтиПоНаименованию("АКЗС Витебский", True)
try:
    if not doc.Склад or doc.Склад.Пустая():
        doc.Склад = wh
        print("Склад set to АКЗС Витебский")
    else:
        print("Склад already set")
except:
    try:
        doc.Склад = wh
        print("Склад forced to АКЗС Витебский")
    except Exception as e:
        print("Склад error:", e)

# Fill счета учёта in rows if needed
for i in range(doc.Товары.Количество()):
    row = doc.Товары.Получить(i)
    try:
        # Счёт учёта = 41.02 for litres on retail warehouse
        plan = ib.ПланыСчетов.Хозрасчетный
        acct = plan.НайтиПоКоду("41.02")
        if acct and not acct.Пустая():
            row.СчетУчета = acct
            print(f"  Row {i+1}: account set to 41.02")
    except Exception as e:
        print(f"  Row {i+1} account error: {e}")

# Post
try:
    doc.Записать(ib.РежимЗаписиДокумента.Проведение)
    print("\nPOSTED OK!")
except Exception as e:
    print("\nPost error:", e)
    # Try write without posting
    try:
        doc.Записать()
        print("Written (not posted). Open in 1C and post manually.")
    except Exception as e2:
        print("Write error:", e2)
