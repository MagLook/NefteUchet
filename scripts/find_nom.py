# -*- coding: utf-8 -*-
"""Find all nomenclature items named Аи-92 (Витебский). Run: py -3.13-32 scripts/find_nom.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\find_nom.txt", "w", encoding="utf-8")

q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ
    Ссылка.Код КАК Код,
    Наименование,
    Родитель.Наименование КАК Группа,
    Родитель.Родитель.Наименование КАК Группа2
ИЗ Справочник.Номенклатура
ГДЕ Наименование ПОДОБНО &Н И НЕ ПометкаУдаления"""
q.УстановитьПараметр("Н", "%Витебский%")
r = q.Выполнить().Выбрать()
while r.Следующий():
    out.write(f"  Код:{r.Код} [{r.Наименование}] Группа:[{r.Группа}] / [{r.Группа2}]\n")

# Also check what НайтиПоНаименованию returns
out.write("\n=== НайтиПоНаименованию ===\n")
nom = ib.Справочники.Номенклатура.НайтиПоНаименованию("Аи-92 (Витебский)", True)
if nom and not nom.Пустая():
    out.write(f"  ExactMatch: Код={nom.Код} [{nom.Наименование}] Группа=[{nom.Родитель}]\n")
nom2 = ib.Справочники.Номенклатура.НайтиПоНаименованию("Аи-92 (Витебский)", False)
if nom2 and not nom2.Пустая():
    out.write(f"  PartialMatch: Код={nom2.Код} [{nom2.Наименование}] Группа=[{nom2.Родитель}]\n")

out.close()
print("Done. Check build/find_nom.txt")
