# -*- coding: utf-8 -*-
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn_str = 'File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";'
conn = connector.Connect(conn_str)

# Настройки
q = conn.NewObject('Query')
q.Text = 'SELECT Ключ, Значение FROM InformationRegister.TL_Настройки ORDER BY Ключ'
result = q.Execute()
sel = result.Choose()
print('=== TL_Настройки ===')
cnt = 0
while sel.Next():
    k = str(sel.Ключ)
    v = str(sel.Значение)[:60]
    if 'ароль' in k:
        v = '***'
    print('  ' + k + ' = ' + v)
    cnt += 1
print('Всего: ' + str(cnt))

# Статусы загрузки
print()
q2 = conn.NewObject('Query')
q2.Text = 'SELECT TOP 20 КлючЗагрузки, Статус, ДатаЗагрузки, ТипДокумента FROM InformationRegister.TL_СтатусыЗагрузки ORDER BY ДатаЗагрузки DESC'
result2 = q2.Execute()
sel2 = result2.Choose()
print('=== TL_СтатусыЗагрузки ===')
cnt2 = 0
while sel2.Next():
    print('  ' + str(sel2.КлючЗагрузки) + ' | ' + str(sel2.Статус) + ' | ' + str(sel2.ДатаЗагрузки) + ' | ' + str(sel2.ТипДокумента))
    cnt2 += 1
if cnt2 == 0:
    print('  (пусто)')
print('Всего: ' + str(cnt2))
