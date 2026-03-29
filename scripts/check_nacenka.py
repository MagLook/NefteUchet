# -*- coding: utf-8 -*-
"""Check торговая наценка registers. Run: py -3.13-32 scripts/check_nacenka.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client
from datetime import datetime

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\nacenka.txt", "w", encoding="utf-8")

# РасчетТорговойНаценкиАТТ — for our warehouse
for reg_name in ["РасчетТорговойНаценкиАТТ", "РасчетТорговойНаценкиНТТ"]:
    out.write(f"\n=== {reg_name} ===\n")
    try:
        # Structure
        md = ib.Метаданные.РегистрыСведений.Найти(reg_name)
        if md:
            out.write("  Измерения:\n")
            for i in range(md.Измерения.Количество()):
                out.write(f"    {md.Измерения.Получить(i).Имя}\n")
            out.write("  Ресурсы:\n")
            for i in range(md.Ресурсы.Количество()):
                out.write(f"    {md.Ресурсы.Получить(i).Имя}\n")

        q = ib.NewObject("Query")
        q.Text = f"ВЫБРАТЬ ПЕРВЫЕ 20 * ИЗ РегистрСведений.{reg_name}"
        tbl = q.Выполнить().Выгрузить()
        out.write(f"  Записей: {tbl.Количество()}\n")
        if tbl.Количество() > 0:
            cols = tbl.Колонки
            for i in range(min(tbl.Количество(), 5)):
                row = tbl.Получить(i)
                vals = []
                for j in range(cols.Количество()):
                    vals.append(f"{cols.Получить(j).Имя}={row.Получить(j)}")
                out.write(f"  [{i}] {'; '.join(vals)}\n")
    except Exception as e:
        out.write(f"  Ошибка: {e}\n")

# Also: check СчетаУчетаНоменклатуры for our specific warehouse
out.write(f"\n=== СчетаУчетаНоменклатуры для Товары АЗС ===\n")
q2 = ib.NewObject("Query")
q2.Text = """ВЫБРАТЬ * ИЗ РегистрСведений.СчетаУчетаНоменклатуры
ГДЕ ВидНоменклатуры.Наименование ПОДОБНО &Н"""
q2.УстановитьПараметр("Н", "%Товар%")
try:
    tbl2 = q2.Выполнить().Выгрузить()
    out.write(f"  Записей: {tbl2.Количество()}\n")
    for i in range(tbl2.Количество()):
        row = tbl2.Получить(i)
        cols = tbl2.Колонки
        vals = []
        for j in range(cols.Количество()):
            v = row.Получить(j)
            if v is not None and str(v) != "":
                vals.append(f"{cols.Получить(j).Имя}={v}")
        out.write(f"  [{i}] {'; '.join(vals)}\n")
except Exception as e:
    out.write(f"  Ошибка: {e}\n")

out.close()
print("Done")
