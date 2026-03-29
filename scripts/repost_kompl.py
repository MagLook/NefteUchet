# -*- coding: utf-8 -*-
"""Repost our Комплектация + try ОбработкаЗаполнения. Run: py -3.13-32 scripts/repost_kompl.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\repost.txt", "w", encoding="utf-8")

# Get our latest Комплектация АИ-92
q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка, Номер
ИЗ Документ.КомплектацияНоменклатуры
ГДЕ Комментарий ПОДОБНО &М И НЕ ПометкаУдаления И Проведен
УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q.УстановитьПараметр("М", "%TL|INIT|Комплектация АИ-92%")
r = q.Выполнить().Выбрать()
if not r.Следующий():
    out.write("Not found\n")
    out.close()
    exit()

doc = r.Ссылка.ПолучитьОбъект()
out.write(f"Doc: {r.Номер}\n")
out.write(f"Склад: {doc.Склад}\n")
out.write(f"Номенклатура: {doc.Номенклатура}\n")
out.write(f"Количество: {doc.Количество}\n")
try: out.write(f"СчетУчета: {doc.СчетУчета.Код}\n")
except: out.write(f"СчетУчета: ?\n")
out.write(f"Комплектующие: {doc.Комплектующие.Количество()}\n")
for i in range(doc.Комплектующие.Количество()):
    row = doc.Комплектующие.Получить(i)
    try: out.write(f"  [{i}] {row.Номенклатура} кол={row.Количество} счёт={row.СчетУчета.Код}\n")
    except: out.write(f"  [{i}] {row.Номенклатура} кол={row.Количество} счёт=?\n")

# Repost: отмена + повторное проведение
out.write(f"\nОтмена проведения...\n")
try:
    doc.Записать(ib.РежимЗаписиДокумента.ОтменаПроведения)
    out.write("OK\n")
except Exception as e:
    out.write(f"Error: {e}\n")

out.write(f"Повторное проведение...\n")
try:
    doc.Записать(ib.РежимЗаписиДокумента.Проведение)
    out.write("OK\n")
except Exception as e:
    out.write(f"Error: {e}\n")

# Check entries after repost
q2 = ib.NewObject("Query")
q2.Text = """ВЫБРАТЬ НомерСтроки, СчетДт.Код КАК Дт, СчетКт.Код КАК Кт, Сумма, КоличествоДт, КоличествоКт
ИЗ РегистрБухгалтерии.Хозрасчетный ГДЕ Регистратор = &Д"""
q2.УстановитьПараметр("Д", r.Ссылка)
r2 = q2.Выполнить().Выбрать()
cnt = 0
while r2.Следующий():
    cnt += 1
    out.write(f"  [{r2.НомерСтроки}] Дт:{r2.Дт} Кт:{r2.Кт} Сумма:{r2.Сумма} КолДт:{r2.КоличествоДт} КолКт:{r2.КоличествоКт}\n")
out.write(f"Проводок после перепроведения: {cnt}\n")

out.close()
print("Done. Check build/repost.txt")
