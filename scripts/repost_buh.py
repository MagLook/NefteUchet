# -*- coding: utf-8 -*-
"""Repost бухгалтерский Комплектация via COM. Run: py -3.13-32 scripts/repost_buh.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\repost_buh.txt", "w", encoding="utf-8")

q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка, Номер
ИЗ Документ.КомплектацияНоменклатуры
ГДЕ Проведен И Комментарий НЕ ПОДОБНО &М
УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q.УстановитьПараметр("М", "%TL|%")
r = q.Выполнить().Выбрать()
r.Следующий()

out.write(f"Документ: {r.Номер}\n")

# Count entries before
q2 = ib.NewObject("Query")
q2.Text = "ВЫБРАТЬ НомерСтроки, СчетДт.Код КАК Дт, СчетКт.Код КАК Кт, Сумма ИЗ РегистрБухгалтерии.Хозрасчетный ГДЕ Регистратор = &Д"
q2.УстановитьПараметр("Д", r.Ссылка)
r2 = q2.Выполнить().Выбрать()
out.write("ДО перепроведения:\n")
cnt = 0
while r2.Следующий():
    cnt += 1
    out.write(f"  [{r2.НомерСтроки}] Дт:{r2.Дт} Кт:{r2.Кт} Сумма:{r2.Сумма}\n")
out.write(f"  Проводок: {cnt}\n")

# Repost
doc = r.Ссылка.ПолучитьОбъект()
out.write("\nОтмена проведения...\n")
doc.Записать(ib.РежимЗаписиДокумента.ОтменаПроведения)
out.write("Повторное проведение...\n")
doc.Записать(ib.РежимЗаписиДокумента.Проведение)

# Count after
r3 = q2.Выполнить().Выбрать()
out.write("\nПОСЛЕ перепроведения:\n")
cnt2 = 0
while r3.Следующий():
    cnt2 += 1
    out.write(f"  [{r3.НомерСтроки}] Дт:{r3.Дт} Кт:{r3.Кт} Сумма:{r3.Сумма}\n")
out.write(f"  Проводок: {cnt2}\n")

out.close()
print("Done. Check build/repost_buh.txt")
