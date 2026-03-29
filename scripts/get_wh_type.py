# -*- coding: utf-8 -*-
"""Get warehouse type as string. Run: py -3.13-32 scripts/get_wh_type.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\wh_types.txt", "w", encoding="utf-8")

# Get warehouse ref and check ТипСклада via eval
wh = ib.Справочники.Склады.НайтиПоНаименованию("АКЗС Витебский", True)
typ = wh.ТипСклада

# Compare with enum values
enum = ib.Перечисления.ТипыСкладов
for i in range(5):
    try:
        v = enum.Получить(i)
        match = "  <<<" if typ == v else ""
        out.write(f"  ТипыСкладов[{i}]: match={typ == v}{match}\n")
    except:
        break

# Check specific values
try:
    out.write(f"\nОптовый: {typ == enum.ОптовыйСклад}\n")
except: pass
try:
    out.write(f"НеавтоматизированнаяТорговаяТочка: {typ == enum.НеавтоматизированнаяТорговаяТочка}\n")
except: pass
try:
    out.write(f"РозничныйМагазин: {typ == enum.РозничныйМагазин}\n")
except: pass
try:
    out.write(f"АвтоматизированнаяТорговаяТочка: {typ == enum.АвтоматизированнаяТорговаяТочка}\n")
except: pass

# Also check Основной склад type
out.write(f"\n--- Основной склад ---\n")
wh2 = ib.Справочники.Склады.НайтиПоНаименованию("Основной склад", True)
typ2 = wh2.ТипСклада
try: out.write(f"ОптовыйСклад: {typ2 == enum.ОптовыйСклад}\n")
except: pass
try: out.write(f"НеавтоматизированнаяТорговаяТочка: {typ2 == enum.НеавтоматизированнаяТорговаяТочка}\n")
except: pass

out.close()
print("Done")
