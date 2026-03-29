# -*- coding: utf-8 -*-
"""Субконто через таблицу субконто."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

# Наше Перемещение
q = conn.NewObject('Query')
q.Text = '''SELECT Ссылка FROM Документ.ПеремещениеТоваров
WHERE Комментарий LIKE "%TL|ТТН|65|5|2994|2|ПЕРЕМ%" AND НЕ ПометкаУдаления'''
r = q.Execute().Choose()
if not r.Next():
    print('Не найден'); sys.exit()
ref = r.Ссылка

# Читаем субконто через таблицу СубконтоДт/Кт
q2 = conn.NewObject('Query')
q2.Text = '''SELECT
    СчетДт.Код AS ДтСчет,
    СчетКт.Код AS КтСчет,
    Сумма,
    КоличествоДт,
    СубконтоДт1 AS ДтНом,
    СубконтоДт2 AS ДтСклад,
    СубконтоКт1 AS КтНом,
    СубконтоКт2 AS КтСклад
FROM РегистрБухгалтерии.Хозрасчетный
WHERE Регистратор = &Р'''
try:
    q2.SetParameter('Р', ref)
    r2 = q2.Execute().Choose()
    print('=== Субконто нашего Перемещения ===')
    while r2.Next():
        dt_nom = str(r2.ДтНом)[:30] if r2.ДтНом else 'None'
        dt_skl = str(r2.ДтСклад)[:30] if r2.ДтСклад else 'None'
        kt_nom = str(r2.КтНом)[:30] if r2.КтНом else 'None'
        kt_skl = str(r2.КтСклад)[:30] if r2.КтСклад else 'None'
        if 'COMObject' in dt_nom:
            try: dt_nom = str(r2.ДтНом.Наименование)[:30]
            except: pass
        if 'COMObject' in dt_skl:
            try: dt_skl = str(r2.ДтСклад.Наименование)[:30]
            except: pass
        if 'COMObject' in kt_nom:
            try: kt_nom = str(r2.КтНом.Наименование)[:30]
            except: pass
        if 'COMObject' in kt_skl:
            try: kt_skl = str(r2.КтСклад.Наименование)[:30]
            except: pass
        print(f'  Дт {str(r2.ДтСчет).strip()} | ном={dt_nom} | склад={dt_skl}')
        print(f'  Кт {str(r2.КтСчет).strip()} | ном={kt_nom} | склад={kt_skl}')
        print(f'  Сумма={float(r2.Сумма or 0):,.2f} | кол={float(r2.КоличествоДт or 0):,.3f}')
except Exception as e:
    print(f'Ошибка: {e}')

# Сравним с INIT Перемещением (который работал)
print()
print('=== Субконто INIT Перемещения (рабочего) ===')
q3 = conn.NewObject('Query')
q3.Text = '''SELECT Ссылка FROM Документ.ПеремещениеТоваров
WHERE Комментарий LIKE "%TL|INIT|Перемещение%" AND Проведен = TRUE AND НЕ ПометкаУдаления'''
r3 = q3.Execute().Choose()
if r3.Next():
    ref3 = r3.Ссылка
    q4 = conn.NewObject('Query')
    q4.Text = '''SELECT
        СчетДт.Код AS ДтСчет, СчетКт.Код AS КтСчет, Сумма, КоличествоДт,
        СубконтоДт1 AS ДтНом, СубконтоДт2 AS ДтСклад,
        СубконтоКт1 AS КтНом, СубконтоКт2 AS КтСклад
    FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор = &Р'''
    q4.SetParameter('Р', ref3)
    r4 = q4.Execute().Choose()
    while r4.Next():
        dt_nom = str(r4.ДтНом)[:30] if r4.ДтНом else 'None'
        dt_skl = str(r4.ДтСклад)[:30] if r4.ДтСклад else 'None'
        kt_nom = str(r4.КтНом)[:30] if r4.КтНом else 'None'
        kt_skl = str(r4.КтСклад)[:30] if r4.КтСклад else 'None'
        if 'COMObject' in dt_nom:
            try: dt_nom = str(r4.ДтНом.Наименование)[:30]
            except: pass
        if 'COMObject' in dt_skl:
            try: dt_skl = str(r4.ДтСклад.Наименование)[:30]
            except: pass
        if 'COMObject' in kt_nom:
            try: kt_nom = str(r4.КтНом.Наименование)[:30]
            except: pass
        if 'COMObject' in kt_skl:
            try: kt_skl = str(r4.КтСклад.Наименование)[:30]
            except: pass
        print(f'  Дт {str(r4.ДтСчет).strip()} | ном={dt_nom} | склад={dt_skl}')
        print(f'  Кт {str(r4.КтСчет).strip()} | ном={kt_nom} | склад={kt_skl}')
        print(f'  Сумма={float(r4.Сумма or 0):,.2f} | кол={float(r4.КоличествоДт or 0):,.3f}')
