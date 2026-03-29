# -*- coding: utf-8 -*-
"""Check the last created TL document in detail. Run: py -3.13-32 scripts/check_last_doc.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\last_doc.txt", "w", encoding="utf-8")

# 1. Last TL doc
q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка, Номер, Дата, Проведен, Комментарий,
    Склад.Наименование КАК СкладНаим
ИЗ Документ.ОтчетОРозничныхПродажах
ГДЕ Комментарий ПОДОБНО &М И НЕ ПометкаУдаления
УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q.УстановитьПараметр("М", "%TL|СМЕНА%")
r = q.Выполнить().Выбрать()
if r.Следующий():
    out.write(f"Номер: {r.Номер}\n")
    out.write(f"Склад: [{r.СкладНаим}]\n")
    out.write(f"Комментарий: {r.Комментарий}\n")

    doc = r.Ссылка.ПолучитьОбъект()
    out.write(f"\nТовары ({doc.Товары.Количество()}):\n")
    for i in range(doc.Товары.Количество()):
        row = doc.Товары.Получить(i)
        nom_name = str(row.Номенклатура.Наименование) if row.Номенклатура else "?"
        out.write(f"  [{i}] Номенклатура: [{nom_name}]\n")
        out.write(f"       Кол: {row.Количество} Цена: {row.Цена} Сумма: {row.Сумма}\n")
        for attr in ["СчетУчета", "СчетДоходов", "СчетРасходов"]:
            try:
                v = getattr(row, attr)
                code = ""
                try: code = v.Код
                except: pass
                out.write(f"       {attr}: {code}\n")
            except:
                out.write(f"       {attr}: НЕТ\n")

    # 2. Compare with working бухгалтерский doc
    out.write(f"\n=== Рабочий документ бухгалтера (АКЗС Витебский) ===\n")
    q2 = ib.NewObject("Query")
    q2.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка, Склад.Наименование КАК СкладНаим
    ИЗ Документ.ОтчетОРозничныхПродажах
    ГДЕ Проведен И Склад.Наименование = &С И Комментарий НЕ ПОДОБНО &М
    УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
    q2.УстановитьПараметр("С", "АКЗС Витебский")
    q2.УстановитьПараметр("М", "%TL|%")
    r2 = q2.Выполнить().Выбрать()
    if r2.Следующий():
        out.write(f"Склад: [{r2.СкладНаим}]\n")
        doc2 = r2.Ссылка.ПолучитьОбъект()
        for i in range(min(doc2.Товары.Количество(), 3)):
            row2 = doc2.Товары.Получить(i)
            nom2 = str(row2.Номенклатура.Наименование) if row2.Номенклатура else "?"
            out.write(f"  [{i}] Номенклатура: [{nom2}]\n")
            out.write(f"       Кол: {row2.Количество} Цена: {row2.Цена} Сумма: {row2.Сумма}\n")
            for attr in ["СчетУчета", "СчетДоходов", "СчетРасходов"]:
                try:
                    v = getattr(row2, attr)
                    code = ""
                    try: code = v.Код
                    except: pass
                    out.write(f"       {attr}: {code}\n")
                except:
                    out.write(f"       {attr}: НЕТ\n")
else:
    out.write("НЕ НАЙДЕН\n")

out.close()
print("Saved to build/last_doc.txt")
