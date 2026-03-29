# -*- coding: utf-8 -*-
"""Change АКЗС Витебский to Оптовый. Run: py -3.13-32 scripts/fix_warehouse_type.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

wh = ib.Справочники.Склады.НайтиПоНаименованию("АКЗС Витебский", True)
obj = wh.ПолучитьОбъект()
obj.ТипСклада = ib.Перечисления.ТипыСкладов.ОптовыйСклад
obj.Записать()
print("АКЗС Витебский -> ОптовыйСклад: OK")
