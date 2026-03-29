# -*- coding: utf-8 -*-
"""Check СчетаУчетаНоменклатуры for Витебский fuel. Run: py -3.13-32 scripts/check_acct_reg.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\acct_reg.txt", "w", encoding="utf-8")

# Dump ALL columns first
q0 = ib.NewObject("Query")
q0.Text = "ВЫБРАТЬ ПЕРВЫЕ 1 * ИЗ РегистрСведений.СчетаУчетаНоменклатуры"
tbl = q0.Выполнить().Выгрузить()
out.write(f"Columns ({tbl.Колонки.Количество()}):\n")
for i in range(tbl.Колонки.Количество()):
    out.write(f"  {tbl.Колонки.Получить(i).Имя}\n")

# Dump all rows (there are only ~17)
out.write(f"\n=== ВСЕ записи СчетаУчетаНоменклатуры ===\n")
q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ
    Номенклатура.Наименование КАК НомНаим,
    ВидНоменклатуры.Наименование КАК ВидНаим,
    ТипСклада,
    Склад.Наименование КАК СкладНаим,
    СчетУчетаБУ.Код КАК СчетУчета,
    СчетПередачиБУ.Код КАК СчетПередачи,
    СчетДоходовОтРеализацииБУ.Код КАК СчетДоходов,
    СчетРасходовРеализацииБУ.Код КАК СчетРасходов
ИЗ РегистрСведений.СчетаУчетаНоменклатуры"""
try:
    r = q.Выполнить().Выбрать()
    while r.Следующий():
        out.write(f"  [{r.НомНаим or 'ВСЕ'}] Вид:[{r.ВидНаим or 'ВСЕ'}] Тип:[{r.ТипСклада}] Склад:[{r.СкладНаим or 'ВСЕ'}] СчётУчёта:{r.СчетУчета} Передачи:{r.СчетПередачи} Доходы:{r.СчетДоходов} Расходы:{r.СчетРасходов}\n")
except Exception as e:
    out.write(f"  Error: {e}\n")

out.close()
print("Done. Check build/acct_reg.txt")
