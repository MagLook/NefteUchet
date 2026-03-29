# -*- coding: utf-8 -*-
"""Deep analysis of the posting error. Run: py -3.13-32 scripts/deep_check.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\deep_check.txt", "w", encoding="utf-8")

# 1. Find our latest TL document
out.write("=== Последний документ TL|СМЕНА ===\n")
q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка, Номер, Дата, Проведен, Комментарий,
    Склад.Наименование КАК СкладНаим, Склад.ТипСклада КАК ТипСклада
ИЗ Документ.ОтчетОРозничныхПродажах
ГДЕ Комментарий ПОДОБНО &М И НЕ ПометкаУдаления
УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q.УстановитьПараметр("М", "%TL|СМЕНА%")
r = q.Выполнить().Выбрать()
if r.Следующий():
    out.write(f"  Номер: {r.Номер}\n")
    out.write(f"  Дата: {r.Дата}\n")
    out.write(f"  Проведен: {r.Проведен}\n")
    out.write(f"  Склад: {r.СкладНаим}\n")
    out.write(f"  ТипСклада: {r.ТипСклада}\n")
    out.write(f"  Комментарий: {r.Комментарий}\n")

    doc = r.Ссылка.ПолучитьОбъект()
    out.write(f"\n  Товары ({doc.Товары.Количество()} строк):\n")
    for i in range(doc.Товары.Количество()):
        row = doc.Товары.Получить(i)
        nom_name = ""
        try: nom_name = str(row.Номенклатура)
        except: pass

        vals = {}
        for attr in ["Количество", "Цена", "Сумма", "ЦенаВРознице", "СуммаВРознице",
                      "СчетУчета", "СчетДоходов", "СчетРасходов", "СтавкаНДС", "СуммаНДС",
                      "СтавкаНДСВРознице", "ДокументОприходования", "Себестоимость"]:
            try:
                v = getattr(row, attr)
                vals[attr] = v
            except:
                pass
        out.write(f"    [{i}] {nom_name}\n")
        for k, v in vals.items():
            out.write(f"        {k} = {v}\n")
else:
    out.write("  НЕ НАЙДЕН\n")

# 2. Check what nomenclature is ACTUALLY used now (after mapping fix)
out.write("\n=== Текущий маппинг топлива (key-value) ===\n")
q2 = ib.NewObject("Query")
q2.Text = "ВЫБРАТЬ Ключ, Значение ИЗ РегистрСведений.TL_Настройки ГДЕ Ключ ПОДОБНО &М"
q2.УстановитьПараметр("М", "%Топливо%")
r2 = q2.Выполнить().Выбрать()
while r2.Следующий():
    out.write(f"  {r2.Ключ} = [{r2.Значение}]\n")

# 3. What nomenclature exists with "Витебский" or "92"
out.write("\n=== Номенклатура в справочнике ===\n")
q3 = ib.NewObject("Query")
q3.Text = """ВЫБРАТЬ Наименование ИЗ Справочник.Номенклатура
ГДЕ (Наименование ПОДОБНО &Н1 ИЛИ Наименование ПОДОБНО &Н2)
И НЕ ПометкаУдаления"""
q3.УстановитьПараметр("Н1", "%Витебский%")
q3.УстановитьПараметр("Н2", "%92%")
r3 = q3.Выполнить().Выбрать()
while r3.Следующий():
    out.write(f"  [{r3.Наименование}]\n")

# 4. Check ALL register names that contain "Товар" or "Склад"
out.write("\n=== Регистры накопления с Товар/Склад ===\n")
md = ib.Метаданные.РегистрыНакопления
for i in range(md.Количество()):
    reg = md.Получить(i)
    if "Товар" in reg.Имя or "Склад" in reg.Имя or "Рознич" in reg.Имя:
        out.write(f"  {reg.Имя}\n")

# 5. Check stock using correct register name
out.write("\n=== Попытки найти остатки ===\n")
for reg_name in ["ТоварыОрганизаций", "ТоварыНаСкладах", "ТоварыВНТТ", "ТоварыПереданныеНаКомиссию"]:
    try:
        q4 = ib.NewObject("Query")
        q4.Text = f"""ВЫБРАТЬ ПЕРВЫЕ 10
            Номенклатура.Наименование КАК Наим,
            КоличествоОстаток КАК Кол
        ИЗ РегистрНакопления.{reg_name}.Остатки"""
        r4 = q4.Выполнить().Выбрать()
        out.write(f"\n  {reg_name}:\n")
        while r4.Следующий():
            out.write(f"    {r4.Наим}: {r4.Кол}\n")
    except:
        pass

# 6. Проведённые документы бухгалтера — сравним их детально
out.write("\n=== Детали проведённого отчёта бухгалтера (АКЗС Витебский) ===\n")
q5 = ib.NewObject("Query")
q5.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка
ИЗ Документ.ОтчетОРозничныхПродажах
ГДЕ Проведен И Склад.Наименование = &С И Комментарий НЕ ПОДОБНО &М
УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q5.УстановитьПараметр("С", "АКЗС Витебский")
q5.УстановитьПараметр("М", "%TL|%")
r5 = q5.Выполнить().Выбрать()
if r5.Следующий():
    doc2 = r5.Ссылка.ПолучитьОбъект()
    out.write(f"  Номер: {doc2.Номер}, Дата: {doc2.Дата}\n")
    for i in range(min(doc2.Товары.Количество(), 3)):
        row = doc2.Товары.Получить(i)
        out.write(f"    [{i}] {row.Номенклатура}\n")
        for attr in ["Количество", "Цена", "Сумма", "ЦенаВРознице", "СуммаВРознице",
                      "СчетУчета", "СчетДоходов", "СчетРасходов", "Себестоимость",
                      "ДокументОприходования"]:
            try:
                v = getattr(row, attr)
                out.write(f"        {attr} = {v}\n")
            except:
                pass

out.close()
print("Saved to build/deep_check.txt")
