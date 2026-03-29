# -*- coding: utf-8 -*-
"""Диагностика: почему Комплектации не удалились."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

# 1. Все активные (не удалённые) Комплектации
print('=== Активные Комплектации (НЕ помечены на удаление) ===')
q = conn.NewObject('Query')
q.Text = '''SELECT Номер, Дата, Проведен, Комментарий
FROM Документ.КомплектацияНоменклатуры
WHERE Комментарий LIKE "%TL|%" AND НЕ ПометкаУдаления
ORDER BY Дата DESC'''
r = q.Execute().Choose()
cnt = 0
while r.Next():
    st = 'Пров.' if r.Проведен else 'Не пр.'
    print(f'  №{str(r.Номер).strip()} {str(r.Дата)[:10]} | {st} | {str(r.Комментарий)[:70]}')
    cnt += 1
print(f'Итого: {cnt}')

# 2. Ключи Комплектаций в регистре
print()
print('=== Ключи КОМПЛ в регистре TL_СтатусыЗагрузки ===')
q2 = conn.NewObject('Query')
q2.Text = 'SELECT КлючЗагрузки, Статус FROM InformationRegister.TL_СтатусыЗагрузки WHERE КлючЗагрузки LIKE "%КОМПЛ%" ORDER BY КлючЗагрузки'
r2 = q2.Execute().Choose()
cnt2 = 0
while r2.Next():
    print(f'  {r2.КлючЗагрузки} | {r2.Статус}')
    cnt2 += 1
print(f'Итого: {cnt2}')

# 3. Ключи ПЕРЕМ в регистре
print()
print('=== Ключи ПЕРЕМ в регистре ===')
q3 = conn.NewObject('Query')
q3.Text = 'SELECT КлючЗагрузки, Статус FROM InformationRegister.TL_СтатусыЗагрузки WHERE КлючЗагрузки LIKE "%ПЕРЕМ%" ORDER BY КлючЗагрузки'
r3 = q3.Execute().Choose()
cnt3 = 0
while r3.Next():
    print(f'  {r3.КлючЗагрузки} | {r3.Статус}')
    cnt3 += 1
print(f'Итого: {cnt3}')

# 4. Базовые ключи ТТН в регистре
print()
print('=== Базовые ключи ТТН (без |КОМПЛ и |ПЕРЕМ) ===')
q4 = conn.NewObject('Query')
q4.Text = '''SELECT КлючЗагрузки, Статус FROM InformationRegister.TL_СтатусыЗагрузки
WHERE КлючЗагрузки LIKE "%ТТН%"
AND НЕ КлючЗагрузки LIKE "%КОМПЛ%"
AND НЕ КлючЗагрузки LIKE "%ПЕРЕМ%"
ORDER BY КлючЗагрузки'''
r4 = q4.Execute().Choose()
while r4.Next():
    # Проверим есть ли дочерние
    base = str(r4.КлючЗагрузки)
    q5 = conn.NewObject('Query')
    q5.Text = 'SELECT COUNT(*) AS К FROM InformationRegister.TL_СтатусыЗагрузки WHERE КлючЗагрузки LIKE &М AND КлючЗагрузки <> &К'
    q5.SetParameter('М', base + '|%')
    q5.SetParameter('К', base)
    r5 = q5.Execute().Choose()
    r5.Next()
    children = int(str(r5.К))
    print(f'  {base} | {r4.Статус} | дочерних: {children}')

# 5. Все записи регистра (полный дамп)
print()
print('=== ВСЕ записи регистра ===')
q6 = conn.NewObject('Query')
q6.Text = 'SELECT КлючЗагрузки, Статус FROM InformationRegister.TL_СтатусыЗагрузки ORDER BY КлючЗагрузки'
r6 = q6.Execute().Choose()
total = 0
while r6.Next():
    print(f'  {r6.КлючЗагрузки} | {r6.Статус}')
    total += 1
print(f'Итого: {total}')
