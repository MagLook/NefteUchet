# -*- coding: utf-8 -*-
"""Проверка субконто в проводках Перемещения."""
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

# Проводки с субконто
q2 = conn.NewObject('Query')
q2.Text = '''SELECT
    СчетДт.Код AS ДтСчет,
    СчетКт.Код AS КтСчет,
    Сумма,
    КоличествоДт,
    КоличествоКт,
    ЗначениеДт AS ДтЗнач,
    ЗначениеКт AS КтЗнач
FROM РегистрБухгалтерии.Хозрасчетный.ДвиженияССубконто(, , Регистратор = &Р, ,)
ORDER BY НомерСтроки'''
try:
    q2.SetParameter('Р', ref)
    r2 = q2.Execute().Choose()
    print('=== Проводки с субконто (наш ПЕРЕМ) ===')
    while r2.Next():
        print(f'  Дт {str(r2.ДтСчет).strip()} Кт {str(r2.КтСчет).strip()} | {float(r2.Сумма or 0):,.2f}')
        print(f'    Дт значение: {r2.ДтЗнач}')
        print(f'    Кт значение: {r2.КтЗнач}')
except Exception as e:
    print(f'Ошибка ДвиженияССубконто: {e}')

# Попробуем через обычный запрос с явными субконто
print()
print('=== Альтернативный запрос ===')
q3 = conn.NewObject('Query')
q3.Text = '''SELECT
    НомерСтроки,
    СчетДт.Код AS ДтСчет,
    СчетКт.Код AS КтСчет,
    Сумма, КоличествоДт, КоличествоКт,
    СубконтоДт.Склады AS ДтСклад,
    СубконтоКт.Склады AS КтСклад
FROM РегистрБухгалтерии.Хозрасчетный
WHERE Регистратор = &Р'''
try:
    q3.SetParameter('Р', ref)
    r3 = q3.Execute().Choose()
    while r3.Next():
        print(f'  Строка {r3.НомерСтроки}: Дт {str(r3.ДтСчет).strip()} Кт {str(r3.КтСчет).strip()}')
        print(f'    ДтСклад: {r3.ДтСклад}')
        print(f'    КтСклад: {r3.КтСклад}')
except Exception as e:
    print(f'Ошибка: {e}')
    # Совсем простой запрос
    q4 = conn.NewObject('Query')
    q4.Text = 'SELECT * FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор = &Р'
    q4.SetParameter('Р', ref)
    try:
        r4 = q4.Execute()
        cols = r4.Колонки
        print(f'Колонки ({cols.Количество()}):')
        for i in range(cols.Количество()):
            print(f'  {cols.Получить(i).Имя}')
    except Exception as e2:
        print(f'Ошибка2: {e2}')
