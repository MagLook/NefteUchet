# -*- coding: utf-8 -*-
"""Check retail prices and warehouse type. Run: py -3.13-32 scripts/check_retail.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\retail_check.txt", "w", encoding="utf-8")

# 1. Warehouse type
out.write("=== Тип склада ===\n")
q = ib.NewObject("Query")
q.Text = "ВЫБРАТЬ Наименование, ТипСклада ИЗ Справочник.Склады ГДЕ Наименование ПОДОБНО &Н"
q.УстановитьПараметр("Н", "%Витебский%")
r = q.Выполнить().Выбрать()
while r.Следующий():
    out.write(f"  {r.Наименование}: {r.ТипСклада}\n")

# 2. Check if there are retail price registers
out.write("\n=== Регистры с ценами ===\n")
md = ib.Метаданные.РегистрыСведений
for i in range(md.Количество()):
    reg = md.Получить(i)
    name = reg.Имя
    if "Цен" in name or "цен" in name or "Рознич" in name or "рознич" in name:
        out.write(f"  {name}\n")

# 3. Check ЦеныНоменклатуры register
out.write("\n=== ЦеныНоменклатуры (если есть) ===\n")
for reg_name in ["ЦеныНоменклатуры", "ТоварыВРозничныхЦенах"]:
    try:
        q2 = ib.NewObject("Query")
        q2.Text = f"ВЫБРАТЬ ПЕРВЫЕ 5 * ИЗ РегистрСведений.{reg_name}"
        tbl = q2.Выполнить().Выгрузить()
        out.write(f"  {reg_name}: {tbl.Количество()} записей\n")
        cols = tbl.Колонки
        col_names = []
        for j in range(cols.Количество()):
            col_names.append(cols.Получить(j).Имя)
        out.write(f"    Колонки: {', '.join(col_names)}\n")
    except Exception as e:
        out.write(f"  {reg_name}: не найден ({e})\n")

# 4. Check existing posted ОтчетОРозничныхПродажах - how they were created
out.write("\n=== Существующие ОтчетыОРозничныхПродажах ===\n")
q3 = ib.NewObject("Query")
q3.Text = """ВЫБРАТЬ ПЕРВЫЕ 3 Ссылка, Номер, Дата, Проведен, Склад.Наименование КАК СкладНаим
ИЗ Документ.ОтчетОРозничныхПродажах
ГДЕ Проведен
УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
try:
    r3 = q3.Выполнить().Выбрать()
    while r3.Следующий():
        out.write(f"  {r3.Номер} {r3.Дата} Склад:{r3.СкладНаим}\n")
        doc = r3.Ссылка.ПолучитьОбъект()
        for i in range(min(doc.Товары.Количество(), 3)):
            row = doc.Товары.Получить(i)
            retail_price = 0
            retail_sum = 0
            try: retail_price = row.ЦенаВРознице
            except: pass
            try: retail_sum = row.СуммаВРознице
            except: pass
            out.write(f"    {row.Номенклатура}: qty={row.Количество} price={row.Цена} sum={row.Сумма} retail_price={retail_price} retail_sum={retail_sum}\n")
except Exception as e:
    out.write(f"  Ошибка: {e}\n")

out.close()
print("Saved to build/retail_check.txt")
