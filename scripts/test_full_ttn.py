# -*- coding: utf-8 -*-
"""Полный тест: загрузить ТТН через расширение, провести, проверить проводки."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

print('=== ПОЛНЫЙ ТЕСТ ТТН ===')

# 1. Авторизация в STS API
print('\n1. Авторизация...')
url = conn.TL_Настройки.ПолучитьЗначение('URLСервера', '')
login = conn.TL_Настройки.ПолучитьЗначение('Логин', '')
pwd = conn.TL_Настройки.ПолучитьЗначение('Пароль', '')
code = int(conn.TL_Настройки.ПолучитьЗначение('КодСистемы', '0'))
token = conn.TL_ApiКлиент.Авторизация(url, login, pwd, code)
print(f'   Токен: {len(str(token))} символов')

# 2. Получить поступления (ТТН) для станции 5, смена 94 (где есть ТТН)
print('\n2. Получение ТТН...')
ttn_json = conn.TL_ApiКлиент.ПолучитьПоступления(url, token, code, 5, 94)
if not ttn_json:
    print('   ТТН не найдены для смены 94')
    # Попробуем смену 90
    ttn_json = conn.TL_ApiКлиент.ПолучитьПоступления(url, token, code, 5, 90)
    if not ttn_json:
        print('   ТТН не найдены и для смены 90')
        sys.exit()

ttn_data = conn.TL_СозданиеДокументов.ПолучитьИзJSON(ttn_json)
print(f'   Получено элементов: {ttn_data.Count() if hasattr(ttn_data, "Count") else "?"}')

# Возьмём первый элемент
elem = None
if hasattr(ttn_data, 'Count'):
    if ttn_data.Count() > 0:
        elem = ttn_data.Get(0)
elif hasattr(ttn_data, 'Количество'):
    if ttn_data.Количество() > 0:
        elem = ttn_data.Получить(0)

if elem is None:
    print('   Нет элементов ТТН')
    sys.exit()

# Получим данные элемента
fuel = str(conn.TL_ApiКлиент.ПолучитьЗначение(elem, 'fuel', ''))
ttn_num = str(conn.TL_ApiКлиент.ПолучитьЗначение(elem, 'ttn', ''))
print(f'   ТТН №{ttn_num}, fuel={fuel}')

# Сериализуем обратно в JSON
elem_json = conn.TL_ApiКлиент.ОбъектВJSON(elem)

# 3. Загрузка через ОбработатьТТН (без проведения)
key = f'TL|ТТН|65|5|{ttn_num}|{fuel}'
print(f'\n3. Загрузка: ключ={key}')

# Проверим нет ли уже
status = conn.TL_РегистрСтатусов.ПолучитьСтатус(key)
if status is not None:
    print(f'   Уже загружен: {status.Статус} — пропускаем')
else:
    result = conn.TL_СозданиеДокументов.ОбработатьТТН(elem_json, key, False)
    docs = result.Документы
    errs = result.Ошибки
    print(f'   Документов: {docs.Количество()}, Ошибок: {errs.Количество()}')
    for i in range(errs.Количество()):
        print(f'   Ошибка: {errs.Получить(i)}')
    for i in range(docs.Количество()):
        print(f'   Документ: {docs.Получить(i)}')

# 4. Проверим что создалось в регистре
print('\n4. Регистр статусов:')
q = conn.NewObject('Query')
q.Text = f'SELECT КлючЗагрузки, Статус FROM InformationRegister.TL_СтатусыЗагрузки WHERE КлючЗагрузки LIKE "%{ttn_num}|{fuel}%" ORDER BY КлючЗагрузки'
r = q.Execute().Choose()
keys_created = []
while r.Next():
    k = str(r.КлючЗагрузки)
    s = str(r.Статус)
    print(f'   {k} | {s}')
    keys_created.append(k)

# 5. Проведение ПЕРЕМ
print('\n5. Проведение ПЕРЕМ...')
for k in keys_created:
    if '|КОМПЛ' in k:
        continue
    doc_ref = conn.TL_РегистрСтатусов.НайтиДокументПоКлючу(k)
    if doc_ref is not None and str(doc_ref) != '':
        try:
            doc = doc_ref.ПолучитьОбъект()
            if not doc.Проведен:
                doc.Записать(conn.NewObject('DocumentWriteMode') if False else None)
                # Прямой вызов проведения через платформенный API
        except:
            pass
        # Попробуем через расширение — просто запишем с проведением
        try:
            doc = doc_ref.ПолучитьОбъект()
            doc.Проведен = True
            doc.Записать()
            print(f'   {k}: записан (Проведен=True)')
        except Exception as e:
            print(f'   {k}: ошибка — {e}')

# Проверим проводки ПЕРЕМ
print('\n6. Проводки ПЕРЕМ:')
for k in keys_created:
    if '|КОМПЛ' in k:
        continue
    doc_ref = conn.TL_РегистрСтатусов.НайтиДокументПоКлючу(k)
    if doc_ref:
        q2 = conn.NewObject('Query')
        q2.Text = 'SELECT СчетДт.Код AS Дт, СчетКт.Код AS Кт, Сумма, КоличествоДт AS КД FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор = &Р'
        q2.SetParameter('Р', doc_ref)
        r2 = q2.Execute().Choose()
        cnt = 0
        while r2.Next():
            print(f'   Дт {str(r2.Дт).strip()} Кт {str(r2.Кт).strip()} | {float(r2.Сумма or 0):,.2f} | кол={float(r2.КД or 0):,.3f}')
            cnt += 1
        if cnt == 0:
            print(f'   {k}: НЕТ ПРОВОДОК')

# 7. Теперь проведение КОМПЛ
print('\n7. Проведение КОМПЛ...')
for k in keys_created:
    if '|КОМПЛ' not in k:
        continue
    doc_ref = conn.TL_РегистрСтатусов.НайтиДокументПоКлючу(k)
    if doc_ref:
        try:
            doc = doc_ref.ПолучитьОбъект()
            doc.Проведен = True
            doc.Записать()
            print(f'   {k}: записан (Проведен=True)')
        except Exception as e:
            print(f'   {k}: ошибка — {e}')

# 8. Проводки КОМПЛ
print('\n8. Проводки КОМПЛ:')
for k in keys_created:
    if '|КОМПЛ' not in k:
        continue
    doc_ref = conn.TL_РегистрСтатусов.НайтиДокументПоКлючу(k)
    if doc_ref:
        q3 = conn.NewObject('Query')
        q3.Text = 'SELECT СчетДт.Код AS Дт, СчетКт.Код AS Кт, Сумма, КоличествоДт AS КД, КоличествоКт AS КК FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор = &Р'
        q3.SetParameter('Р', doc_ref)
        r3 = q3.Execute().Choose()
        cnt = 0
        while r3.Next():
            print(f'   Дт {str(r3.Дт).strip()} Кт {str(r3.Кт).strip()} | {float(r3.Сумма or 0):,.2f} | Дт.кол={float(r3.КД or 0):,.3f} | Кт.кол={float(r3.КК or 0):,.3f}')
            cnt += 1
        if cnt == 0:
            print(f'   {k}: НЕТ ПРОВОДОК (через Проведен=True)')
            # Перепроведём через 1С расширение
            print('   Попытка перепроведения...')
            try:
                doc = doc_ref.ПолучитьОбъект()
                doc.Проведен = False
                doc.Записать()
                doc.Проведен = True
                doc.Записать()
                # Проверим снова
                r3b = q3.Execute().Choose()
                while r3b.Next():
                    print(f'   ПОСЛЕ: Дт {str(r3b.Дт).strip()} Кт {str(r3b.Кт).strip()} | {float(r3b.Сумма or 0):,.2f}')
                    cnt += 1
                if cnt == 0:
                    print('   ПОСЛЕ: всё ещё НЕТ ПРОВОДОК')
            except Exception as e:
                print(f'   Ошибка перепроведения: {e}')

print('\n=== ТЕСТ ЗАВЕРШЁН ===')
