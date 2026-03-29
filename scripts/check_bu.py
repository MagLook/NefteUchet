# -*- coding: utf-8 -*-
"""Check accounting entries (проводки). Run: py -3.13-32 scripts/check_bu.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\bu_check.txt", "w", encoding="utf-8")

# 1. Остатки по 41.02 на АКЗС Витебский
out.write("=== Остатки по счёту 41.02 (бухгалтерский) ===\n")
try:
    q = ib.NewObject("Query")
    q.Text = """ВЫБРАТЬ
        СубконтоДт1.Наименование КАК Номенклатура,
        СубконтоДт2.Наименование КАК Склад,
        СуммаОстатокДт КАК Сумма,
        КоличествоОстатокДт КАК Количество
    ИЗ РегистрБухгалтерии.Хозрасчетный.Остатки(,
        Счет = ПланыСчетов.Хозрасчетный.НайтиПоКоду("41.02"),)
    ГДЕ КоличествоОстатокДт <> 0"""
    r = q.Выполнить().Выбрать()
    while r.Следующий():
        cost_per = r.Сумма / r.Количество if r.Количество != 0 else 0
        out.write(f"  [{r.Номенклатура}] на [{r.Склад}]: {r.Количество} шт, {r.Сумма} руб (за ед: {cost_per:.2f})\n")
except Exception as e:
    out.write(f"  Ошибка: {e}\n")

# 2. Остатки по 41.01
out.write("\n=== Остатки по счёту 41.01 (бухгалтерский) ===\n")
try:
    q2 = ib.NewObject("Query")
    q2.Text = """ВЫБРАТЬ
        СубконтоДт1.Наименование КАК Номенклатура,
        СубконтоДт2.Наименование КАК Склад,
        СуммаОстатокДт КАК Сумма,
        КоличествоОстатокДт КАК Количество
    ИЗ РегистрБухгалтерии.Хозрасчетный.Остатки(,
        Счет = ПланыСчетов.Хозрасчетный.НайтиПоКоду("41.01"),)
    ГДЕ КоличествоОстатокДт <> 0"""
    r2 = q2.Выполнить().Выбрать()
    while r2.Следующий():
        out.write(f"  [{r2.Номенклатура}] на [{r2.Склад}]: {r2.Количество} шт, {r2.Сумма} руб\n")
except Exception as e:
    out.write(f"  Ошибка: {e}\n")

out.close()
print("Saved to build/bu_check.txt")
