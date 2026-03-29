# -*- coding: utf-8 -*-
"""Диагностика ТТН 2765."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

# 1. Все ТТН в регистре статусов
q = conn.NewObject('Query')
q.Text = 'SELECT КлючЗагрузки, Статус, ДатаЗагрузки, ТипДокумента FROM InformationRegister.TL_СтатусыЗагрузки WHERE КлючЗагрузки LIKE "%ТТН%" ORDER BY ДатаЗагрузки DESC'
res = q.Execute().Choose()
print('=== Все ТТН в регистре статусов ===')
cnt = 0
while res.Next():
    print(f'  {res.КлючЗагрузки} | {res.Статус} | {res.ДатаЗагрузки} | {res.ТипДокумента}')
    cnt += 1
if cnt == 0:
    print('  (пусто)')

# 2. Документы с 2765 в комментарии
print()
print('=== Документы с "2765" в комментарии ===')
for doc_type in ['ПеремещениеТоваров', 'КомплектацияНоменклатуры', 'ПоступлениеТоваровУслуг']:
    q2 = conn.NewObject('Query')
    q2.Text = f'SELECT Номер, Дата, Проведен, Комментарий FROM Документ.{doc_type} WHERE Комментарий LIKE "%2765%"'
    res2 = q2.Execute().Choose()
    while res2.Next():
        st = 'ПРОВЕДЁН' if res2.Проведен else 'не проведён'
        print(f'  [{doc_type}] №{str(res2.Номер).strip()} {str(res2.Дата)[:10]} | {st} | {str(res2.Комментарий)[:70]}')

# 3. Проверим что вернул бы API для ТТН 2765 — ищем по всем загруженным
print()
print('=== Все ТТН ключи (уникальные) ===')
q3 = conn.NewObject('Query')
q3.Text = 'SELECT РАЗЛИЧНЫЕ КлючЗагрузки FROM InformationRegister.TL_СтатусыЗагрузки WHERE ТипДокумента = "ТТН" OR КлючЗагрузки LIKE "%ТТН%"'
res3 = q3.Execute().Choose()
while res3.Next():
    k = str(res3.КлючЗагрузки)
    if 'КОМПЛ' not in k and 'ПЕРЕМ' not in k:
        print(f'  {k}')

# 4. Проверим статусы с ошибками
print()
print('=== Статусы с ошибками ===')
q4 = conn.NewObject('Query')
q4.Text = 'SELECT КлючЗагрузки, Статус FROM InformationRegister.TL_СтатусыЗагрузки WHERE Статус <> "Загружен" AND Статус <> "Проведён"'
res4 = q4.Execute().Choose()
cnt4 = 0
while res4.Next():
    print(f'  {res4.КлючЗагрузки} | {res4.Статус}')
    cnt4 += 1
if cnt4 == 0:
    print('  (нет ошибок)')
