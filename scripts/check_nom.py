# -*- coding: utf-8 -*-
"""Check what nomenclature is used. Run: py -3.13-32 scripts/check_nom.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\nom_check.txt", "w", encoding="utf-8")

# 1. Nomenclature in existing posted reports on АКЗС Витебский
out.write("=== Номенклатура в проведённых отчётах (АКЗС Витебский) ===\n")
q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ РАЗЛИЧНЫЕ
    Т.Номенклатура.Наименование КАК Наим
ИЗ Документ.ОтчетОРозничныхПродажах.Товары КАК Т
ГДЕ Т.Ссылка.Проведен И Т.Ссылка.Склад.Наименование = &С"""
q.УстановитьПараметр("С", "АКЗС Витебский")
try:
    r = q.Выполнить().Выбрать()
    while r.Следующий():
        out.write(f"  Бухгалтер использует: [{r.Наим}]\n")
except Exception as e:
    out.write(f"  Ошибка: {e}\n")

# 2. Nomenclature in OUR (TL) documents
out.write("\n=== Номенклатура в документах TL ===\n")
q2 = ib.NewObject("Query")
q2.Text = """ВЫБРАТЬ РАЗЛИЧНЫЕ
    Т.Номенклатура.Наименование КАК Наим
ИЗ Документ.ОтчетОРозничныхПродажах.Товары КАК Т
ГДЕ Т.Ссылка.Комментарий ПОДОБНО &М"""
q2.УстановитьПараметр("М", "%TL|СМЕНА%")
try:
    r2 = q2.Выполнить().Выбрать()
    while r2.Следующий():
        out.write(f"  TradeLedger использует: [{r2.Наим}]\n")
except Exception as e:
    out.write(f"  Ошибка: {e}\n")

# 3. What's in настройки for fuel mapping
out.write("\n=== Настройки маппинга топлива ===\n")
q3 = ib.NewObject("Query")
q3.Text = """ВЫБРАТЬ Ключ, Значение ИЗ РегистрСведений.TL_Настройки
ГДЕ Ключ ПОДОБНО &М УПОРЯДОЧИТЬ ПО Ключ"""
q3.УстановитьПараметр("М", "%Топливо%Литры%")
try:
    r3 = q3.Выполнить().Выбрать()
    while r3.Следующий():
        out.write(f"  {r3.Ключ} = [{r3.Значение}]\n")
except Exception as e:
    out.write(f"  Ошибка: {e}\n")

# 4. Stock on АКЗС Витебский - what nomenclature has stock
out.write("\n=== Остатки на АКЗС Витебский (через оборотку) ===\n")
q4 = ib.NewObject("Query")
q4.Text = """ВЫБРАТЬ
    Номенклатура.Наименование КАК Наим,
    КоличествоКонечныйОстаток КАК Кол
ИЗ РегистрНакопления.ТоварыНаСкладах.ОстаткиИОбороты(, , , , Склад.Наименование = &С)
ГДЕ КоличествоКонечныйОстаток <> 0"""
q4.УстановитьПараметр("С", "АКЗС Витебский")
try:
    r4 = q4.Выполнить().Выбрать()
    found = False
    while r4.Следующий():
        out.write(f"  [{r4.Наим}]: {r4.Кол}\n")
        found = True
    if not found:
        out.write("  ПУСТО!\n")
except Exception as e:
    out.write(f"  Ошибка: {e}\n")

out.close()
print("Saved to build/nom_check.txt")
