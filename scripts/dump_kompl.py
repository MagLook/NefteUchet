# -*- coding: utf-8 -*-
"""Dump КомплектацияНоменклатуры fields. Run: py -3.13-32 scripts/dump_kompl.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\kompl_fields.txt", "w", encoding="utf-8")

md = ib.Метаданные.Документы.Найти("КомплектацияНоменклатуры")
out.write("=== КомплектацияНоменклатуры ===\n")
out.write("  Header:\n")
for i in range(md.Реквизиты.Количество()):
    out.write(f"    {md.Реквизиты.Получить(i).Имя}\n")

for j in range(md.ТабличныеЧасти.Количество()):
    ts = md.ТабличныеЧасти.Получить(j)
    out.write(f"  TC: {ts.Имя}\n")
    for k in range(ts.Реквизиты.Количество()):
        out.write(f"    {ts.Реквизиты.Получить(k).Имя}\n")

out.close()
print("Saved to build/kompl_fields.txt")
