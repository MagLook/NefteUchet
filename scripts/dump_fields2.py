# -*- coding: utf-8 -*-
"""Dump fields to file. Run: py -3.13-32 scripts/dump_fields2.py"""
import pythoncom, sys, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\fields.txt", "w", encoding="utf-8")

for doc_name in ["ПеремещениеТоваров", "ПоступлениеТоваровУслуг", "ОтчетОРозничныхПродажах"]:
    md = ib.Метаданные.Документы.Найти(doc_name)
    if not md:
        out.write(f"{doc_name}: NOT FOUND\n")
        continue
    out.write(f"\n=== {doc_name} ===\n")

    out.write("  Header:\n")
    for i in range(md.Реквизиты.Количество()):
        a = md.Реквизиты.Получить(i)
        out.write(f"    {a.Имя}\n")

    for j in range(md.ТабличныеЧасти.Количество()):
        ts = md.ТабличныеЧасти.Получить(j)
        out.write(f"  TC: {ts.Имя}\n")
        for k in range(ts.Реквизиты.Количество()):
            ta = ts.Реквизиты.Получить(k)
            out.write(f"    {ta.Имя}\n")

out.close()
print("Saved to build/fields.txt")
