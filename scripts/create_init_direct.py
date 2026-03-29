# -*- coding: utf-8 -*-
"""Create init stock by calling TL_Маппинг directly. Run: py -3.13-32 scripts/create_init_direct.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client
from datetime import datetime

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\init_direct.txt", "w", encoding="utf-8")

org = ib.Справочники.Организации.НайтиПоНаименованию("ГАЗИНВЕСТГРУПП ООО", True)
осн = ib.Справочники.Склады.НайтиПоНаименованию("Основной склад", True)
азс = ib.Справочники.Склады.НайтиПоНаименованию("АКЗС Витебский", True)

# Контрагент
q = ib.NewObject("Query")
q.Text = "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.Контрагенты ГДЕ Наименование ПОДОБНО &Н И НЕ ПометкаУдаления"
q.УстановитьПараметр("Н", "%СУРГУТНЕФТЕГАЗ%")
r = q.Выполнить().Выбрать()
контр = r.Ссылка if r.Следующий() else None

out.write(f"Org: {'OK' if org and not org.Пустая() else 'FAIL'}\n")
out.write(f"Main WH: {'OK' if осн and not осн.Пустая() else 'FAIL'}\n")
out.write(f"AZS WH: {'OK' if азс and not азс.Пустая() else 'FAIL'}\n")
out.write(f"Contr: {'OK' if контр else 'FAIL'}\n")

try:
    ib.TL_Маппинг.ИнициализироватьНачальныеОстатки(org, осн, азс, контр, datetime(2026, 3, 20))
    out.write("\nИнициализация УСПЕШНА!\n")
except Exception as e:
    out.write(f"\nОшибка: {e}\n")

# Verify entries
out.write("\n=== Проверка проводок ===\n")
for doc_type in ["ПоступлениеТоваровУслуг", "ПеремещениеТоваров", "КомплектацияНоменклатуры"]:
    q2 = ib.NewObject("Query")
    q2.Text = f"""ВЫБРАТЬ Ссылка, Номер, Проведен, Комментарий
    ИЗ Документ.{doc_type}
    ГДЕ Комментарий ПОДОБНО &М И НЕ ПометкаУдаления"""
    q2.УстановитьПараметр("М", "%TL|INIT%")
    r2 = q2.Выполнить().Выбрать()
    while r2.Следующий():
        out.write(f"\n  {doc_type}: {r2.Номер} Проведен:{r2.Проведен}\n")
        if r2.Проведен:
            q3 = ib.NewObject("Query")
            q3.Text = """ВЫБРАТЬ НомерСтроки, СчетДт.Код КАК Дт, СчетКт.Код КАК Кт, Сумма, КоличествоДт, КоличествоКт
            ИЗ РегистрБухгалтерии.Хозрасчетный ГДЕ Регистратор = &Д УПОРЯДОЧИТЬ ПО НомерСтроки"""
            q3.УстановитьПараметр("Д", r2.Ссылка)
            r3 = q3.Выполнить().Выбрать()
            cnt = 0
            while r3.Следующий():
                cnt += 1
                out.write(f"    [{r3.НомерСтроки}] Дт:{r3.Дт} Кт:{r3.Кт} Сумма:{r3.Сумма} КолДт:{r3.КоличествоДт} КолКт:{r3.КоличествоКт}\n")
            if cnt == 0:
                out.write("    НЕТ ПРОВОДОК!\n")

out.close()
print("Done. Check build/init_direct.txt")
