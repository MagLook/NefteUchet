# -*- coding: utf-8 -*-
"""Доказательство: проведение КОМПЛ после ПЕРЕМ в отдельных серверных вызовах COM."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

# Ищем непроведённую Комплектацию и соответствующее Перемещение
print('=== Поиск непроведённых документов ===')
q = conn.NewObject('Query')
q.Text = '''SELECT КлючЗагрузки, Статус FROM InformationRegister.TL_СтатусыЗагрузки
WHERE КлючЗагрузки LIKE "%ТТН%" AND Статус = "Загружен"
ORDER BY КлючЗагрузки'''
r = q.Execute().Choose()
keys = []
while r.Next():
    k = str(r.КлючЗагрузки)
    print(f'  {k} | {r.Статус}')
    keys.append(k)

if not keys:
    print('  Нет непроведённых ТТН')
    print('  Ищем ЛЮБЫЕ ТТН...')
    q2 = conn.NewObject('Query')
    q2.Text = 'SELECT КлючЗагрузки, Статус FROM InformationRegister.TL_СтатусыЗагрузки WHERE КлючЗагрузки LIKE "%ТТН%" ORDER BY КлючЗагрузки'
    r2 = q2.Execute().Choose()
    while r2.Next():
        k = str(r2.КлючЗагрузки)
        print(f'  {k} | {r2.Статус}')
        keys.append(k)

# Берём базовый ключ ТТН (без |ПЕРЕМ и |КОМПЛ)
base_keys = set()
for k in keys:
    if '|ПЕРЕМ' not in k and '|КОМПЛ' not in k:
        base_keys.add(k)

if not base_keys:
    print('\nНет базовых ключей ТТН')
    sys.exit()

base = sorted(base_keys)[0]
perem_key = base + '|ПЕРЕМ'
kompl_key = base + '|КОМПЛ'
print(f'\nТестируем: {base}')
print(f'  ПЕРЕМ: {perem_key}')
print(f'  КОМПЛ: {kompl_key}')

# Найдём документы
perem_ref = conn.TL_РегистрСтатусов.НайтиДокументПоКлючу(perem_key)
kompl_ref = conn.TL_РегистрСтатусов.НайтиДокументПоКлючу(kompl_key)
print(f'\n  ПЕРЕМ документ: {perem_ref}')
print(f'  КОМПЛ документ: {kompl_ref}')

if not perem_ref or not kompl_ref:
    print('  Документы не найдены')
    sys.exit()

# Шаг 1: Распровести оба (если проведены)
print('\n=== Шаг 1: Распроведение ===')
for label, ref in [('ПЕРЕМ', perem_ref), ('КОМПЛ', kompl_ref)]:
    try:
        doc = ref.ПолучитьОбъект()
        if doc.Проведен:
            doc.Проведен = False
            doc.Записать()
            print(f'  {label}: распроведён')
        else:
            print(f'  {label}: уже не проведён')
    except Exception as e:
        print(f'  {label}: ошибка — {e}')

# Шаг 2: Провести ПЕРЕМ
print('\n=== Шаг 2: Проведение ПЕРЕМ ===')
try:
    doc = perem_ref.ПолучитьОбъект()
    doc.Проведен = True
    doc.Записать()
    print('  ПЕРЕМ: проведён (Проведен=True)')
except Exception as e:
    print(f'  ПЕРЕМ: ошибка — {e}')

# Проверим проводки ПЕРЕМ
q3 = conn.NewObject('Query')
q3.Text = 'SELECT СчетДт.Код AS Дт, СчетКт.Код AS Кт, Сумма, КоличествоДт AS КД FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор = &Р'
q3.SetParameter('Р', perem_ref)
r3 = q3.Execute().Choose()
while r3.Next():
    print(f'  Проводка: Дт {str(r3.Дт).strip()} Кт {str(r3.Кт).strip()} | {float(r3.Сумма or 0):,.2f} | кол={float(r3.КД or 0):,.3f}')

# Шаг 3: Проверим остатки 41.01 АКЗС Витебский СЕЙЧАС
print('\n=== Шаг 3: Остатки 41.01 АКЗС Витебский ===')
q4 = conn.NewObject('Query')
q4.Text = '''SELECT Субконто1.Наименование AS Ном, КоличествоОстатокДт AS Кол, СуммаОстатокДт AS Сумма
FROM РегистрБухгалтерии.Хозрасчетный.Остатки(, Счет = &Счет, , Субконто2 = &Склад)'''
q4.SetParameter('Счет', conn.ChartsOfAccounts.Хозрасчетный.FindByCode('41.01'))
q4.SetParameter('Склад', conn.Catalogs.Склады.FindByDescription('АКЗС Витебский', True))
r4 = q4.Execute().Choose()
found = False
while r4.Next():
    kol = float(r4.Кол or 0)
    if kol != 0:
        print(f'  {str(r4.Ном)[:30]} | кол={kol:,.3f} | сумма={float(r4.Сумма or 0):,.2f}')
        found = True
if not found:
    print('  ПУСТО — тонн нет на АКЗС Витебский!')

# Шаг 4: Провести КОМПЛ (в отдельном вызове — кэш должен быть свежий)
print('\n=== Шаг 4: Проведение КОМПЛ ===')
try:
    doc = kompl_ref.ПолучитьОбъект()
    doc.Проведен = True
    doc.Записать()
    print('  КОМПЛ: проведён (Проведен=True)')
except Exception as e:
    print(f'  КОМПЛ: ошибка — {e}')

# Проверим проводки КОМПЛ
q5 = conn.NewObject('Query')
q5.Text = 'SELECT СчетДт.Код AS Дт, СчетКт.Код AS Кт, Сумма, КоличествоДт AS КД, КоличествоКт AS КК FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор = &Р'
q5.SetParameter('Р', kompl_ref)
r5 = q5.Execute().Choose()
cnt = 0
while r5.Next():
    print(f'  Проводка: Дт {str(r5.Дт).strip()} Кт {str(r5.Кт).strip()} | {float(r5.Сумма or 0):,.2f} | литров={float(r5.КД or 0):,.1f} | тонн={float(r5.КК or 0):,.3f}')
    cnt += 1
if cnt == 0:
    print('  НЕТ ПРОВОДОК КОМПЛ даже через COM!')
    print('  Вывод: doc.Проведен=True + doc.Записать() НЕ вызывает ОбработкуПроведения')
    print('  Нужен: doc.Записать(РежимЗаписиДокумента.Проведение)')
else:
    print(f'  ЕСТЬ {cnt} проводок!')

print('\n=== ЗАВЕРШЕНО ===')
