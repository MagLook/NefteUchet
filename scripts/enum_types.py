# -*- coding: utf-8 -*-
"""Dump warehouse type enum names. Run: py -3.13-32 scripts/enum_types.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\enum_types.txt", "w", encoding="utf-8")

# Get metadata of ТипыСкладов enum
md = ib.Метаданные.Перечисления.Найти("ТипыСкладов")
if md:
    vals = md.ЗначенияПеречисления
    for i in range(vals.Количество()):
        v = vals.Получить(i)
        out.write(f"  [{i}] {v.Имя}\n")

wh = ib.Справочники.Склады.НайтиПоНаименованию("АКЗС Витебский", True)
typ = wh.ТипСклада

# Match
for i in range(vals.Количество()):
    v = vals.Получить(i)
    enum_val = getattr(ib.Перечисления.ТипыСкладов, v.Имя)
    match = typ == enum_val
    out.write(f"  Match [{v.Имя}]: {match}\n")

out.close()
print("Done")
