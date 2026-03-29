# -*- coding: utf-8 -*-
"""Check our Комплектация documents detail. Run: py -3.13-32 scripts/check_kompl_detail.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\kompl_detail.txt", "w", encoding="utf-8")

q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ Ссылка, Номер, Комментарий
ИЗ Документ.КомплектацияНоменклатуры
ГДЕ Комментарий ПОДОБНО &М И НЕ ПометкаУдаления
УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q.УстановитьПараметр("М", "%TL|INIT%")
r = q.Выполнить().Выбрать()
while r.Следующий():
    doc = r.Ссылка.ПолучитьОбъект()
    out.write(f"\n=== {r.Номер} [{r.Комментарий}] ===\n")
    out.write(f"  Выпуск: [{doc.Номенклатура.Наименование}] код:{doc.Номенклатура.Код} группа:{doc.Номенклатура.Родитель}\n")
    out.write(f"  Кол: {doc.Количество}\n")
    out.write(f"  Склад: [{doc.Склад.Наименование}]\n")
    try: out.write(f"  СчетУчета: {doc.СчетУчета.Код}\n")
    except: out.write(f"  СчетУчета: ПУСТО\n")

    for i in range(doc.Комплектующие.Количество()):
        row = doc.Комплектующие.Получить(i)
        out.write(f"  Комплектующие [{i}]: [{row.Номенклатура.Наименование}] код:{row.Номенклатура.Код} группа:{row.Номенклатура.Родитель}\n")
        out.write(f"    Кол: {row.Количество}\n")
        try: out.write(f"    СчетУчета: {row.СчетУчета.Код}\n")
        except: out.write(f"    СчетУчета: ПУСТО\n")

# Сравним с бухгалтерским
out.write(f"\n\n=== БУХГАЛТЕРСКИЙ ===\n")
q2 = ib.NewObject("Query")
q2.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка, Номер ИЗ Документ.КомплектацияНоменклатуры
ГДЕ Проведен И Комментарий НЕ ПОДОБНО &М УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q2.УстановитьПараметр("М", "%TL|%")
r2 = q2.Выполнить().Выбрать()
if r2.Следующий():
    doc2 = r2.Ссылка.ПолучитьОбъект()
    out.write(f"  {r2.Номер}\n")
    out.write(f"  Выпуск: [{doc2.Номенклатура.Наименование}] код:{doc2.Номенклатура.Код} группа:{doc2.Номенклатура.Родитель}\n")
    out.write(f"  Склад: [{doc2.Склад.Наименование}]\n")
    try: out.write(f"  СчетУчета: {doc2.СчетУчета.Код}\n")
    except: out.write(f"  СчетУчета: ПУСТО\n")
    row2 = doc2.Комплектующие.Получить(0)
    out.write(f"  Комплектующие [0]: [{row2.Номенклатура.Наименование}] код:{row2.Номенклатура.Код} группа:{row2.Номенклатура.Родитель}\n")
    try: out.write(f"    СчетУчета: {row2.СчетУчета.Код}\n")
    except: out.write(f"    СчетУчета: ПУСТО\n")

out.close()
print("Done. Check build/kompl_detail.txt")
