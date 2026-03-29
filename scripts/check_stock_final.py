# -*- coding: utf-8 -*-
"""Check actual stock. Run: py -3.13-32 scripts/check_stock_final.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\stock_final.txt", "w", encoding="utf-8")

# Find all accumulation registers
md = ib.Метаданные.РегистрыНакопления
out.write("=== ВСЕ регистры накопления ===\n")
for i in range(md.Количество()):
    out.write(f"  {md.Получить(i).Имя}\n")

# Try each likely register for stock
out.write("\n=== Поиск остатков по регистрам ===\n")
for reg in ["ТоварыОрганизаций", "ТоварыВНТТ", "ТоварыВРознице", "ОстаткиТоваров"]:
    try:
        q = ib.NewObject("Query")
        q.Text = f"""ВЫБРАТЬ ПЕРВЫЕ 20
            Номенклатура.Наименование КАК Наим,
            КоличествоОстаток КАК Кол
        ИЗ РегистрНакопления.{reg}.Остатки
        ГДЕ КоличествоОстаток <> 0"""
        r = q.Выполнить().Выбрать()
        out.write(f"\n  {reg}:\n")
        found = False
        while r.Следующий():
            out.write(f"    [{r.Наим}]: {r.Кол}\n")
            found = True
        if not found:
            out.write("    ПУСТО\n")
    except Exception as e:
        pass

# Try ТоварыОрганизаций with warehouse dimension
out.write("\n=== ТоварыОрганизаций с измерениями ===\n")
try:
    q2 = ib.NewObject("Query")
    q2.Text = """ВЫБРАТЬ
        Номенклатура.Наименование КАК Наим,
        Склад.Наименование КАК СкладНаим,
        КоличествоОстаток КАК Кол,
        СтоимостьОстаток КАК Стоимость
    ИЗ РегистрНакопления.ТоварыОрганизаций.Остатки
    ГДЕ КоличествоОстаток <> 0
    И (Номенклатура.Наименование ПОДОБНО &Н1
        ИЛИ Номенклатура.Наименование ПОДОБНО &Н2
        ИЛИ Номенклатура.Наименование ПОДОБНО &Н3)"""
    q2.УстановитьПараметр("Н1", "%Витебский%")
    q2.УстановитьПараметр("Н2", "%92%")
    q2.УстановитьПараметр("Н3", "%ДТ%")
    r2 = q2.Выполнить().Выбрать()
    while r2.Следующий():
        cost_per = r2.Стоимость / r2.Кол if r2.Кол != 0 else 0
        out.write(f"    [{r2.Наим}] на [{r2.СкладНаим}]: {r2.Кол} (стоимость: {r2.Стоимость}, за ед: {cost_per:.2f})\n")
except Exception as e:
    out.write(f"    Ошибка: {e}\n")

out.close()
print("Saved to build/stock_final.txt")
