# -*- coding: utf-8 -*-
"""Глубокая диагностика: почему нет проводок при идентичных реквизитах."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

# 1. Остатки на 41.01 (тонны) на складе АКЗС Витебский
print('=== Остатки на 41.01 (тонны) АКЗС Витебский ===')
q = conn.NewObject('Query')
q.Text = '''SELECT
    Субконто1.Наименование AS Номенклатура,
    КоличествоОстатокДт AS Количество,
    СуммаОстатокДт AS Сумма
FROM РегистрБухгалтерии.Хозрасчетный.Остатки(
    ,
    Счет = ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.ТоварыНаСкладах),
    ,
)
WHERE Субконто2.Наименование LIKE "%Витебский%"
'''
try:
    r = q.Execute().Choose()
    while r.Next():
        print(f'  {str(r.Номенклатура):40s} | кол={float(r.Количество or 0):>10,.3f} | сумма={float(r.Сумма or 0):>14,.2f}')
except Exception as e:
    print(f'  Ошибка: {e}')

# Попробуем проще
print()
print('=== Остатки 41.01 + 41.02 (все склады с Витебский) ===')
q2 = conn.NewObject('Query')
q2.Text = '''SELECT
    Счет.Код AS Счет,
    Субконто1.Наименование AS Номенклатура,
    Субконто2.Наименование AS Склад,
    КоличествоОстатокДт AS Кол,
    СуммаОстатокДт AS Сумма
FROM РегистрБухгалтерии.Хозрасчетный.Остатки(, Счет.Код LIKE "41%",,)
WHERE Субконто2.Наименование LIKE "%Витебский%"
ORDER BY Счет.Код, Субконто1.Наименование'''
try:
    r2 = q2.Execute().Choose()
    while r2.Next():
        schet = str(r2.Счет).strip()
        nom = str(r2.Номенклатура)[:35]
        kol = float(r2.Кол or 0)
        summa = float(r2.Сумма or 0)
        print(f'  {schet} | {nom:35s} | кол={kol:>12,.3f} | сумма={summa:>14,.2f}')
except Exception as e:
    print(f'  Ошибка: {e}')

# 2. Перепровести одну Комплектацию и посмотреть результат
print()
print('=== Тест: перепроведение Комплектации ГИ0Г-000321 ===')
q3 = conn.NewObject('Query')
q3.Text = 'SELECT Ссылка FROM Документ.КомплектацияНоменклатуры WHERE Номер = "ГИ0Г-000321"'
r3 = q3.Execute().Choose()
if r3.Next():
    ref = r3.Ссылка
    doc = ref.ПолучитьОбъект()
    try:
        doc.Записать(1)  # РежимЗаписиДокумента.Проведение = 1
        print('  Перепроведение: OK')
    except Exception as e:
        print(f'  Ошибка перепроведения: {e}')

    # Проверим проводки
    q4 = conn.NewObject('Query')
    q4.Text = 'SELECT СчетДт.Код AS Дт, СчетКт.Код AS Кт, Сумма, КоличествоДт, КоличествоКт FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор = &Р'
    q4.SetParameter('Р', ref)
    r4 = q4.Execute().Choose()
    cnt = 0
    while r4.Next():
        dt = str(r4.Дт).strip()
        kt = str(r4.Кт).strip()
        s = float(r4.Сумма or 0)
        kd = float(r4.КоличествоДт or 0)
        kk = float(r4.КоличествоКт or 0)
        print(f'    Дт {dt} Кт {kt} | сумма={s:,.2f} | кол.Дт={kd:,.3f} | кол.Кт={kk:,.3f}')
        cnt += 1
    if cnt == 0:
        print('  ПРОВОДОК НЕТ ДАЖЕ ПОСЛЕ ПЕРЕПРОВЕДЕНИЯ')
