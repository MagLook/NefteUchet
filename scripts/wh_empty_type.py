# -*- coding: utf-8 -*-
"""Check if warehouse type is empty. Run: py -3.13-32 scripts/wh_empty_type.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\wh_empty.txt", "w", encoding="utf-8")

for name in ["АКЗС Витебский", "Основной склад", "АГЗС Верево", "ЯНДЕКС"]:
    wh = ib.Справочники.Склады.НайтиПоНаименованию(name, True)
    if wh and not wh.Пустая():
        typ = wh.ТипСклада
        empty = ib.Перечисления.ТипыСкладов.ПустаяСсылка()
        is_empty = (typ == empty)
        out.write(f"  [{name}]: пустой={is_empty}, == Оптовый:{typ == ib.Перечисления.ТипыСкладов.ОптовыйСклад}, == Розничный:{typ == ib.Перечисления.ТипыСкладов.РозничныйМагазин}, == НТТ:{typ == ib.Перечисления.ТипыСкладов.НеавтоматизированнаяТорговаяТочка}\n")

out.close()
print("Done")
