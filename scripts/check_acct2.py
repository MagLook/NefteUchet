# -*- coding: utf-8 -*-
"""Check accounts via simplified query. Run: py -3.13-32 scripts/check_acct2.py"""
import pythoncom
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

# Simple query - all records from accounts register
q = ib.NewObject("Query")
q.Text = "ВЫБРАТЬ * ИЗ РегистрСведений.СчетаУчетаНоменклатуры"
try:
    tbl = q.Выполнить().Выгрузить()
    cols = tbl.Колонки
    print(f"Columns ({cols.Количество()}):")
    for i in range(cols.Количество()):
        col = cols.Получить(i)
        print(f"  [{i}] {col.Имя}")

    print(f"\nRows: {tbl.Количество()}")
    for i in range(tbl.Количество()):
        row = tbl.Получить(i)
        vals = []
        for j in range(min(cols.Количество(), 6)):
            v = row.Получить(j)
            vals.append(str(v) if v else "-")
        print(f"  [{i}] {' | '.join(vals)}")
except Exception as e:
    print(f"Error: {e}")

# Check the broken doc - what's actually in it
print("\n=== Last INIT_STOCK doc details ===")
q2 = ib.NewObject("Query")
q2.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка
ИЗ Документ.ПоступлениеТоваровУслуг
ГДЕ Комментарий ПОДОБНО &М И НЕ ПометкаУдаления
УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q2.УстановитьПараметр("М", "%INIT_STOCK%")
r = q2.Выполнить().Выбрать()
if r.Следующий():
    doc = r.Ссылка.ПолучитьОбъект()
    print(f"  ВидОперации: {doc.ВидОперации}")
    print(f"  Склад: {doc.Склад}")
    print(f"  Контрагент: {doc.Контрагент}")

    # Check each row for account
    for i in range(doc.Товары.Количество()):
        row = doc.Товары.Получить(i)
        print(f"  Row {i}: nom={row.Номенклатура} qty={row.Количество}")
        # List all properties
        try:
            for attr in ["СчетУчета", "СчетУчетаБУ", "СчетРасчетовСКонтрагентом"]:
                try:
                    v = getattr(row, attr, None)
                    if v is not None:
                        print(f"    {attr} = {v}")
                except:
                    pass
        except:
            pass
