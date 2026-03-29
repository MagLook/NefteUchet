# -*- coding: utf-8 -*-
"""Check accounts register with correct field names. Run: py -3.13-32 scripts/check_acct_reg2.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\acct_reg2.txt", "w", encoding="utf-8")

q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ
    Номенклатура.Наименование КАК НомНаим,
    ВидНоменклатуры.Наименование КАК ВидНаим,
    ТипСклада,
    Склад.Наименование КАК СкладНаим,
    СчетУчета.Код КАК СчетУчета,
    СчетУчетаПередачи.Код КАК СчетПередачи,
    СчетДоходовОтРеализации.Код КАК СчетДоходов,
    СчетРасходовОтРеализации.Код КАК СчетРасходов
ИЗ РегистрСведений.СчетаУчетаНоменклатуры"""
r = q.Выполнить().Выбрать()
while r.Следующий():
    out.write(f"  Ном:[{r.НомНаим or '-'}] Вид:[{r.ВидНаим or '-'}] Тип:[{r.ТипСклада}] Склад:[{r.СкладНаим or '-'}] Учёт:{r.СчетУчета or '-'} Перед:{r.СчетПередачи or '-'} Дох:{r.СчетДоходов or '-'} Расх:{r.СчетРасходов or '-'}\n")

out.close()
print("Done. Check build/acct_reg2.txt")
