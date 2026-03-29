# -*- coding: utf-8 -*-
"""Check 41.02 balance NOW. Run: py -3.13-32 scripts/check_balance_now.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client
from datetime import datetime

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\balance_now.txt", "w", encoding="utf-8")

q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ
    Субконто1.Наименование КАК Номенклатура,
    Субконто2.Наименование КАК Склад,
    СуммаОстаток КАК Сумма,
    КоличествоОстаток КАК Кол
ИЗ РегистрБухгалтерии.Хозрасчетный.Остатки(&Дата, Счет = &Счет, ,)"""
q.УстановитьПараметр("Дата", datetime(2026, 3, 30))
q.УстановитьПараметр("Счет", ib.ПланыСчетов.Хозрасчетный.НайтиПоКоду("41.02"))
try:
    r = q.Выполнить().Выбрать()
    while r.Следующий():
        cost_per = r.Сумма / r.Кол if r.Кол and r.Кол != 0 else 0
        out.write(f"  [{r.Номенклатура}] [{r.Склад}]: {r.Кол} шт, {r.Сумма} руб (за ед: {cost_per:.2f})\n")
except Exception as e:
    out.write(f"Error: {e}\n")

out.close()
print("Done")
