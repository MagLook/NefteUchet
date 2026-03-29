# -*- coding: utf-8 -*-
"""Check бухгалтер's Комплектация. Run: py -3.13-32 scripts/check_buh_kompl.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\buh_kompl.txt", "w", encoding="utf-8")

# Find бухгалтер's Комплектация (not TL|)
q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 3 Ссылка, Номер, Дата, Проведен, Комментарий
ИЗ Документ.КомплектацияНоменклатуры
ГДЕ Проведен И Комментарий НЕ ПОДОБНО &М
УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q.УстановитьПараметр("М", "%TL|%")
r = q.Выполнить().Выбрать()
while r.Следующий():
    out.write(f"\n=== {r.Номер} {r.Дата} [{r.Комментарий}] ===\n")
    doc = r.Ссылка.ПолучитьОбъект()

    # Header fields
    out.write(f"  Склад: {doc.Склад}\n")
    out.write(f"  Номенклатура (выпуск): {doc.Номенклатура}\n")
    out.write(f"  Количество (выпуск): {doc.Количество}\n")
    try: out.write(f"  СчетУчета (заголовок): {doc.СчетУчета.Код}\n")
    except: out.write(f"  СчетУчета (заголовок): ПУСТО\n")

    out.write(f"  Комплектующие ({doc.Комплектующие.Количество()}):\n")
    for i in range(doc.Комплектующие.Количество()):
        row = doc.Комплектующие.Получить(i)
        acct = ""
        try: acct = row.СчетУчета.Код
        except: acct = "ПУСТО"
        out.write(f"    [{i}] {row.Номенклатура} кол={row.Количество} счёт={acct}\n")

    # Entries
    q2 = ib.NewObject("Query")
    q2.Text = """ВЫБРАТЬ НомерСтроки, СчетДт.Код КАК Дт, СчетКт.Код КАК Кт, Сумма, КоличествоДт, КоличествоКт
    ИЗ РегистрБухгалтерии.Хозрасчетный ГДЕ Регистратор = &Д УПОРЯДОЧИТЬ ПО НомерСтроки"""
    q2.УстановитьПараметр("Д", r.Ссылка)
    r2 = q2.Выполнить().Выбрать()
    cnt = 0
    while r2.Следующий():
        cnt += 1
        out.write(f"  Проводка [{r2.НомерСтроки}] Дт:{r2.Дт} Кт:{r2.Кт} Сумма:{r2.Сумма} КолДт:{r2.КоличествоДт} КолКт:{r2.КоличествоКт}\n")
    if cnt == 0:
        out.write("  НЕТ ПРОВОДОК!\n")

# Also check OUR kompl document details
out.write(f"\n\n=== НАШ документ Комплектации ===\n")
q3 = ib.NewObject("Query")
q3.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка, Номер ИЗ Документ.КомплектацияНоменклатуры
ГДЕ Комментарий ПОДОБНО &М И НЕ ПометкаУдаления УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q3.УстановитьПараметр("М", "%TL|INIT|Комплектация АИ-92%")
r3 = q3.Выполнить().Выбрать()
if r3.Следующий():
    doc = r3.Ссылка.ПолучитьОбъект()
    out.write(f"  Номер: {r3.Номер}\n")
    out.write(f"  Склад: {doc.Склад}\n")
    out.write(f"  Номенклатура (выпуск): {doc.Номенклатура}\n")
    out.write(f"  Количество (выпуск): {doc.Количество}\n")
    try: out.write(f"  СчетУчета (заголовок): {doc.СчетУчета.Код}\n")
    except: out.write(f"  СчетУчета (заголовок): ПУСТО\n")

    for i in range(doc.Комплектующие.Количество()):
        row = doc.Комплектующие.Получить(i)
        acct = ""
        try: acct = row.СчетУчета.Код
        except: acct = "ПУСТО"
        out.write(f"  Комплектующие [{i}] {row.Номенклатура} кол={row.Количество} счёт={acct}\n")

out.close()
print("Done. Check build/buh_kompl.txt")
