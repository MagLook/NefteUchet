# -*- coding: utf-8 -*-
"""Dump 41.02 balance columns. Run: py -3.13-32 scripts/check_bu3.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\bu_check3.txt", "w", encoding="utf-8")

from datetime import datetime

# Just get columns first
q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ * ИЗ РегистрБухгалтерии.Хозрасчетный.ОстаткиИОбороты(
    &ДатаНач, &ДатаКон,,,Счет = &Счет,,,,)"""
q.УстановитьПараметр("ДатаНач", datetime(2026, 3, 1))
q.УстановитьПараметр("ДатаКон", datetime(2026, 3, 31))
q.УстановитьПараметр("Счет", ib.ПланыСчетов.Хозрасчетный.НайтиПоКоду("41.02"))

try:
    tbl = q.Выполнить().Выгрузить()
    cols = tbl.Колонки
    out.write(f"Columns ({cols.Количество()}):\n")
    for i in range(cols.Количество()):
        out.write(f"  [{i}] {cols.Получить(i).Имя}\n")

    out.write(f"\nRows: {tbl.Количество()}\n")
    for i in range(tbl.Количество()):
        row = tbl.Получить(i)
        out.write(f"\n  Row [{i}]:\n")
        for j in range(cols.Количество()):
            col_name = cols.Получить(j).Имя
            val = row.Получить(j)
            out.write(f"    {col_name} = {val}\n")
except Exception as e:
    out.write(f"Error: {e}\n")

out.close()
print("Saved to build/bu_check3.txt")
