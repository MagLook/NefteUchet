# -*- coding: utf-8 -*-
"""Fix fuel mapping to use station-specific nomenclature. Run: py -3.13-32 scripts/fix_fuel_mapping.py"""
import pythoncom
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

# Correct mapping: station-specific names used by бухгалтер
fixes = [
    ("Топливо_2_Литры", "Аи-92 (Витебский)"),
    ("Топливо_3_Литры", "Аи-95 (Витебский)"),
    ("Топливо_5_Литры", "ДТ (Витебский)"),
]

for key, value in fixes:
    rec = ib.РегистрыСведений.TL_Настройки.СоздатьМенеджерЗаписи()
    rec.Ключ = key
    rec.Значение = value
    rec.Описание = "Литры для АКЗС Витебский"
    rec.Записать(True)
    print(f"  {key} = {value}")

print("\nDone! Now TradeLedger will use the correct nomenclature.")
