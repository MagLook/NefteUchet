# -*- coding: utf-8 -*-
"""Проверить почему формы не заполняются."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

# 1. Проверяем регистр — есть ли настройки
q = conn.NewObject('Query')
q.Text = 'SELECT COUNT(*) AS Кол FROM InformationRegister.TL_Настройки'
res = q.Execute().Choose()
res.Next()
total = int(str(res.Кол))
print(f'Записей в TL_Настройки: {total}')

# 2. Проверяем критичные ключи
keys = ['URLСервера', 'Логин', 'Пароль', 'КодСистемы', 'Организация', 'Станция_5_Склад']
for k in keys:
    q2 = conn.NewObject('Query')
    q2.Text = 'SELECT Значение FROM InformationRegister.TL_Настройки WHERE Ключ = &К'
    q2.SetParameter('К', k)
    res2 = q2.Execute().Choose()
    if res2.Next():
        v = str(res2.Значение)[:40]
        if 'ароль' in k:
            v = '***'
        print(f'  {k} = "{v}"')
    else:
        print(f'  {k} = НЕ НАЙДЕН!')

# 3. Тестируем модуль TL_Настройки напрямую
print()
print('=== Тест TL_Настройки.ПолучитьЗначение() ===')
try:
    url = conn.TL_Настройки.ПолучитьЗначение('URLСервера', '')
    print(f'  URLСервера = "{url}"')
    login = conn.TL_Настройки.ПолучитьЗначение('Логин', '')
    print(f'  Логин = "{login}"')
except Exception as e:
    print(f'  Ошибка: {e}')

# 4. Тестируем НастройкиЗаполнены()
print()
print('=== Тест TL_Настройки.НастройкиЗаполнены() ===')
try:
    filled = conn.TL_Настройки.НастройкиЗаполнены()
    print(f'  НастройкиЗаполнены = {filled}')
except Exception as e:
    print(f'  Ошибка: {e}')

# 5. Тестируем ИнициализироватьПоУмолчаниюГИГ()
print()
print('=== Тест TL_Настройки.ИнициализироватьПоУмолчаниюГИГ() ===')
try:
    conn.TL_Настройки.ИнициализироватьПоУмолчаниюГИГ()
    print('  Вызов успешен')
    # Перепроверяем
    url2 = conn.TL_Настройки.ПолучитьЗначение('URLСервера', '')
    print(f'  URLСервера после инициализации = "{url2}"')
except Exception as e:
    print(f'  Ошибка: {e}')
