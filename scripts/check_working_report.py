# -*- coding: utf-8 -*-
"""Check accounting entries of a WORKING posted report. Run: py -3.13-32 scripts/check_working_report.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\working_report.txt", "w", encoding="utf-8")

# 1. Find a working posted report on АКЗС Витебский (made by бухгалтер)
q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка, Номер, Дата
ИЗ Документ.ОтчетОРозничныхПродажах
ГДЕ Проведен И Склад.Наименование = &С И Комментарий НЕ ПОДОБНО &М
УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q.УстановитьПараметр("С", "АКЗС Витебский")
q.УстановитьПараметр("М", "%TL|%")
r = q.Выполнить().Выбрать()
if not r.Следующий():
    out.write("No working report found!\n")
    out.close()
    exit()

out.write(f"=== Рабочий проведённый отчёт бухгалтера ===\n")
out.write(f"Номер: {r.Номер} Дата: {r.Дата}\n")
ref = r.Ссылка

# 2. Get ALL accounting entries (проводки) for this document
out.write(f"\n=== Проводки этого документа ===\n")
q2 = ib.NewObject("Query")
q2.Text = """ВЫБРАТЬ
    НомерСтроки,
    СчетДт.Код КАК Дебет,
    СчетКт.Код КАК Кредит,
    Сумма,
    КоличествоДт,
    КоличествоКт,
    Содержание
ИЗ РегистрБухгалтерии.Хозрасчетный
ГДЕ Регистратор = &Док
УПОРЯДОЧИТЬ ПО НомерСтроки"""
q2.УстановитьПараметр("Док", ref)
r2 = q2.Выполнить().Выбрать()
cnt = 0
while r2.Следующий():
    cnt += 1
    out.write(f"  [{r2.НомерСтроки}] Дт:{r2.Дебет} Кт:{r2.Кредит} Сумма:{r2.Сумма} КолДт:{r2.КоличествоДт} КолКт:{r2.КоличествоКт} [{r2.Содержание}]\n")
out.write(f"Итого проводок: {cnt}\n")

# 3. Now find OUR TL document
out.write(f"\n=== НАШ документ TL|СМЕНА ===\n")
q3 = ib.NewObject("Query")
q3.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка, Номер, Дата, Проведен
ИЗ Документ.ОтчетОРозничныхПродажах
ГДЕ Комментарий ПОДОБНО &М И НЕ ПометкаУдаления
УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q3.УстановитьПараметр("М", "%TL|СМЕНА%")
r3 = q3.Выполнить().Выбрать()
if r3.Следующий():
    out.write(f"Номер: {r3.Номер} Дата: {r3.Дата} Проведен: {r3.Проведен}\n")
    our_ref = r3.Ссылка

    # Try to post it and catch exact error
    out.write(f"\n=== Попытка проведения ===\n")
    try:
        doc = our_ref.ПолучитьОбъект()
        doc.Записать(ib.РежимЗаписиДокумента.Проведение)
        out.write("ПРОВЕДЕНО УСПЕШНО!\n")

        # Show entries
        q4 = ib.NewObject("Query")
        q4.Text = """ВЫБРАТЬ НомерСтроки, СчетДт.Код КАК Дебет, СчетКт.Код КАК Кредит, Сумма
        ИЗ РегистрБухгалтерии.Хозрасчетный ГДЕ Регистратор = &Док"""
        q4.УстановитьПараметр("Док", our_ref)
        r4 = q4.Выполнить().Выбрать()
        while r4.Следующий():
            out.write(f"  [{r4.НомерСтроки}] Дт:{r4.Дебет} Кт:{r4.Кредит} Сумма:{r4.Сумма}\n")
    except Exception as e:
        out.write(f"ОШИБКА: {e}\n")
else:
    out.write("НЕ НАЙДЕН\n")

out.close()
print("Saved to build/working_report.txt")
