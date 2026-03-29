# -*- coding: utf-8 -*-
"""Check entries from our INIT documents. Run: py -3.13-32 scripts/check_init_entries.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\init_entries.txt", "w", encoding="utf-8")

# Find all our INIT documents and their entries
for doc_type in ["ПоступлениеТоваровУслуг", "ПеремещениеТоваров", "КомплектацияНоменклатуры"]:
    out.write(f"\n=== {doc_type} (TL|INIT) ===\n")
    q = ib.NewObject("Query")
    q.Text = f"""ВЫБРАТЬ Ссылка, Номер, Дата, Проведен, Комментарий
    ИЗ Документ.{doc_type}
    ГДЕ Комментарий ПОДОБНО &М И НЕ ПометкаУдаления
    УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
    q.УстановитьПараметр("М", "%TL|INIT%")
    r = q.Выполнить().Выбрать()
    while r.Следующий():
        out.write(f"\n  {r.Номер} {r.Дата} Проведен:{r.Проведен} [{r.Комментарий}]\n")
        if r.Проведен:
            q2 = ib.NewObject("Query")
            q2.Text = """ВЫБРАТЬ НомерСтроки, СчетДт.Код КАК Дебет, СчетКт.Код КАК Кредит,
                Сумма, КоличествоДт, КоличествоКт, Содержание
            ИЗ РегистрБухгалтерии.Хозрасчетный ГДЕ Регистратор = &Док
            УПОРЯДОЧИТЬ ПО НомерСтроки"""
            q2.УстановитьПараметр("Док", r.Ссылка)
            r2 = q2.Выполнить().Выбрать()
            cnt = 0
            while r2.Следующий():
                cnt += 1
                out.write(f"    [{r2.НомерСтроки}] Дт:{r2.Дебет} Кт:{r2.Кредит} Сумма:{r2.Сумма} КолДт:{r2.КоличествоДт} КолКт:{r2.КоличествоКт}\n")
            if cnt == 0:
                out.write("    НЕТ ПРОВОДОК!\n")

out.close()
print("Saved to build/init_entries.txt")
