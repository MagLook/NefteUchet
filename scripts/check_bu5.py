# -*- coding: utf-8 -*-
"""Dump columns of Хозрасчетный register. Run: py -3.13-32 scripts/check_bu5.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\bu_columns.txt", "w", encoding="utf-8")

from datetime import datetime

# Dump columns
q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 * ИЗ РегистрБухгалтерии.Хозрасчетный
ГДЕ Период >= &Д И (СчетДт.Код = "41.02" ИЛИ СчетКт.Код = "41.02")"""
q.УстановитьПараметр("Д", datetime(2026, 3, 1))

try:
    tbl = q.Выполнить().Выгрузить()
    cols = tbl.Колонки
    out.write(f"Columns ({cols.Количество()}):\n")
    for i in range(cols.Количество()):
        out.write(f"  [{i}] {cols.Получить(i).Имя}\n")

    if tbl.Количество() > 0:
        out.write(f"\nSample row:\n")
        row = tbl.Получить(0)
        for i in range(cols.Количество()):
            col_name = cols.Получить(i).Имя
            val = row.Получить(i)
            out.write(f"  {col_name} = {val}\n")
    else:
        out.write("\nНет проводок по 41.02 с марта 2026!\n")
except Exception as e:
    out.write(f"Error: {e}\n")

# Also check - any entries at all for our INIT docs?
out.write("\n=== Проводки документов TL|INIT ===\n")
q2 = ib.NewObject("Query")
q2.Text = """ВЫБРАТЬ ПЕРВЫЕ 10
    Период, СчетДт.Код КАК Дебет, СчетКт.Код КАК Кредит, Сумма
ИЗ РегистрБухгалтерии.Хозрасчетный
ГДЕ Период >= &Д И СчетДт.Код ПОДОБНО "41%"
УПОРЯДОЧИТЬ ПО Период УБЫВ"""
q2.УстановитьПараметр("Д", datetime(2026, 3, 28))
try:
    r2 = q2.Выполнить().Выбрать()
    while r2.Следующий():
        out.write(f"  {r2.Период} Дт:{r2.Дебет} Кт:{r2.Кредит} Сумма:{r2.Сумма}\n")
except Exception as e:
    out.write(f"  Error: {e}\n")

out.close()
print("Saved to build/bu_columns.txt")
