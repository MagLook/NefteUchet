# -*- coding: utf-8 -*-
"""Simple init: direct Поступление литров на АКЗС Витебский. Run: py -3.13-32 scripts/init_simple.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client
from datetime import datetime

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\init_simple.txt", "w", encoding="utf-8")

org = ib.Справочники.Организации.НайтиПоНаименованию("ГАЗИНВЕСТГРУПП ООО", True)
склад = ib.Справочники.Склады.НайтиПоНаименованию("АКЗС Витебский", True)

q = ib.NewObject("Query")
q.Text = "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.Контрагенты ГДЕ Наименование ПОДОБНО &Н И НЕ ПометкаУдаления"
q.УстановитьПараметр("Н", "%СУРГУТНЕФТЕГАЗ%")
r = q.Выполнить().Выбрать()
контр = r.Ссылка if r.Следующий() else None

# Договор
q2 = ib.NewObject("Query")
q2.Text = "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.ДоговорыКонтрагентов ГДЕ Владелец = &К И НЕ ПометкаУдаления"
q2.УстановитьПараметр("К", контр)
r2 = q2.Выполнить().Выбрать()
договор = r2.Ссылка if r2.Следующий() else None

# Счета
счет4102 = ib.ПланыСчетов.Хозрасчетный.НайтиПоКоду("41.02")
счет6001 = ib.ПланыСчетов.Хозрасчетный.НайтиПоКоду("60.01")
счет1903 = ib.ПланыСчетов.Хозрасчетный.НайтиПоКоду("19.03")
ндс = ib.Перечисления.СтавкиНДС.НДС22

# Литры на станции — прямое поступление
FUEL = [
    ("Аи-92 (Витебский)",  500000, 55),
    ("Аи-95 (Витебский)",  300000, 60),
    ("ДТ (Витебский)",     400000, 62),
]

doc = ib.Документы.ПоступлениеТоваровУслуг.СоздатьДокумент()
doc.Дата = datetime(2026, 3, 20, 10, 0, 0)
doc.Организация = org
doc.Контрагент = контр
doc.Комментарий = "TL|INIT_LITRES|Начальные остатки литров"

try: doc.ВидОперации = ib.Перечисления.ВидыОперацийПоступлениеТоваровУслуг.Покупка
except: pass
try: doc.Склад = склад
except: pass
try: doc.ДоговорКонтрагента = договор
except: pass
try: doc.СчетУчетаРасчетовСКонтрагентом = счет6001
except: pass

for name, qty, price in FUEL:
    nom = ib.Справочники.Номенклатура.НайтиПоНаименованию(name, True)
    if not nom or nom.Пустая():
        out.write(f"SKIP: {name}\n")
        continue
    row = doc.Товары.Добавить()
    row.Номенклатура = nom
    row.Количество = qty
    row.Цена = price
    row.Сумма = qty * price
    try: row.СчетУчета = счет4102
    except: pass
    try: row.СчетРасчетов = счет6001
    except: pass
    try: row.СчетУчетаНДС = счет1903
    except: pass
    try:
        row.СтавкаНДС = ндс
        row.СуммаНДС = round(qty * price * 22 / 122, 2)
    except: pass
    out.write(f"  + {name}: {qty} x {price}\n")

try:
    doc.Записать(ib.РежимЗаписиДокумента.Проведение)
    out.write(f"\nПРОВЕДЕНО!\n")

    # Check entries
    q3 = ib.NewObject("Query")
    q3.Text = "ВЫБРАТЬ НомерСтроки, СчетДт.Код КАК Дт, СчетКт.Код КАК Кт, Сумма, КоличествоДт ИЗ РегистрБухгалтерии.Хозрасчетный ГДЕ Регистратор = &Д"
    q3.УстановитьПараметр("Д", doc.Ссылка)
    r3 = q3.Выполнить().Выбрать()
    while r3.Следующий():
        out.write(f"  [{r3.НомерСтроки}] Дт:{r3.Дт} Кт:{r3.Кт} Сумма:{r3.Сумма} Кол:{r3.КоличествоДт}\n")
except Exception as e:
    out.write(f"\nОшибка: {e}\n")
    try:
        doc.Записать()
        out.write("Записано без проведения\n")
    except: pass

out.close()
print("Done. Check build/init_simple.txt")
