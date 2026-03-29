# -*- coding: utf-8 -*-
"""Dump all field names for document tabular sections. Run: py -3.13-32 scripts/dump_fields.py"""
import pythoncom
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

docs = [
    "ПоступлениеТоваровУслуг",
    "ПеремещениеТоваров",
    "КомплектацияНоменклатуры",
    "ОтчетОРозничныхПродажах",
]

for doc_name in docs:
    print(f"\n=== {doc_name} ===")
    try:
        md = ib.Метаданные.Документы.Найти(doc_name)
        if not md:
            print("  NOT FOUND")
            continue

        # Header attributes
        print("  Реквизиты заголовка:")
        attrs = md.Реквизиты
        for i in range(attrs.Количество()):
            a = attrs.Получить(i)
            name = a.Имя
            # Filter for account-related
            if "Счет" in name or "счет" in name or "Account" in name or "Учет" in name:
                print(f"    * {name}")
            elif "Склад" in name or "Контр" in name or "Договор" in name or "Операци" in name:
                print(f"      {name}")

        # Tabular sections
        tss = md.ТабличныеЧасти
        for j in range(tss.Количество()):
            ts = tss.Получить(j)
            print(f"  ТЧ: {ts.Имя}")
            ts_attrs = ts.Реквизиты
            for k in range(ts_attrs.Количество()):
                ta = ts_attrs.Получить(k)
                name = ta.Имя
                if "Счет" in name or "счет" in name or "Account" in name or "Учет" in name:
                    print(f"    * {name}")

    except Exception as e:
        print(f"  Error: {e}")
