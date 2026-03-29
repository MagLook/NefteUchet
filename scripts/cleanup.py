# -*- coding: utf-8 -*-
"""Clean up all TL test documents and registry entries. Run: py -3.13-32 scripts/cleanup.py"""
import pythoncom
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

# 1. Clear TL_СтатусыЗагрузки registry (all TL| entries)
print("Cleaning TL_СтатусыЗагрузки...")
q = ib.NewObject("Query")
q.Text = "ВЫБРАТЬ КлючЗагрузки ИЗ РегистрСведений.TL_СтатусыЗагрузки ГДЕ КлючЗагрузки ПОДОБНО &М"
q.УстановитьПараметр("М", "TL|%")
r = q.Выполнить().Выбрать()
count = 0
while r.Следующий():
    try:
        rec = ib.РегистрыСведений.TL_СтатусыЗагрузки.СоздатьМенеджерЗаписи()
        rec.КлючЗагрузки = r.КлючЗагрузки
        rec.Удалить()
        count += 1
    except:
        pass
print(f"  Deleted {count} registry entries")

# 2. Mark TL documents for deletion
for doc_type in ["ОтчетОРозничныхПродажах", "ПеремещениеТоваров", "КомплектацияНоменклатуры", "ПоступлениеТоваровУслуг"]:
    q2 = ib.NewObject("Query")
    q2.Text = f"ВЫБРАТЬ Ссылка ИЗ Документ.{doc_type} ГДЕ Комментарий ПОДОБНО &М И НЕ ПометкаУдаления"
    q2.УстановитьПараметр("М", "TL|%")
    r2 = q2.Выполнить().Выбрать()
    cnt = 0
    while r2.Следующий():
        try:
            obj = r2.Ссылка.ПолучитьОбъект()
            if obj.Проведен:
                obj.Записать(ib.РежимЗаписиДокумента.ОтменаПроведения)
            obj.УстановитьПометкуУдаления(True)
            cnt += 1
        except:
            pass
    if cnt > 0:
        print(f"  {doc_type}: marked {cnt} for deletion")

print("\nCleanup done. Now reload TradeLedger and try fresh.")
