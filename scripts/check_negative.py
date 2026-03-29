# -*- coding: utf-8 -*-
"""Проверка: разрешены ли отрицательные остатки, и есть ли остатки на Основном складе."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

# 1. Остатки на 41.01 по ВСЕМ складам
print('=== Остатки 41.01 (тонны) по ВСЕМ складам ===')
q = conn.NewObject('Query')
q.Text = '''SELECT
    Счет.Код AS Счет,
    Субконто1.Наименование AS Номенклатура,
    Субконто2.Наименование AS Склад,
    КоличествоОстатокДт AS Кол,
    СуммаОстатокДт AS Сумма
FROM РегистрБухгалтерии.Хозрасчетный.Остатки(, Счет.Код = "41.01",,)
ORDER BY Субконто2.Наименование, Субконто1.Наименование'''
try:
    r = q.Execute().Choose()
    cnt = 0
    while r.Next():
        nom = str(r.Номенклатура)[:35]
        skl = str(r.Склад)[:25]
        kol = float(r.Кол or 0)
        summa = float(r.Сумма or 0)
        print(f'  {skl:25s} | {nom:35s} | кол={kol:>10,.3f} | сумма={summa:>14,.2f}')
        cnt += 1
    if cnt == 0:
        print('  (пусто — нет остатков на 41.01)')
except Exception as e:
    print(f'  Ошибка: {e}')

# 2. Остатки на 41.02 по ВСЕМ складам
print()
print('=== Остатки 41.02 (литры) по ВСЕМ складам ===')
q2 = conn.NewObject('Query')
q2.Text = '''SELECT
    Счет.Код AS Счет,
    Субконто1.Наименование AS Номенклатура,
    Субконто2.Наименование AS Склад,
    КоличествоОстатокДт AS Кол,
    СуммаОстатокДт AS Сумма
FROM РегистрБухгалтерии.Хозрасчетный.Остатки(, Счет.Код = "41.02",,)
WHERE КоличествоОстатокДт <> 0
ORDER BY Субконто2.Наименование, Субконто1.Наименование'''
try:
    r2 = q2.Execute().Choose()
    while r2.Next():
        nom = str(r2.Номенклатура)[:35]
        skl = str(r2.Склад)[:25]
        kol = float(r2.Кол or 0)
        summa = float(r2.Сумма or 0)
        print(f'  {skl:25s} | {nom:35s} | кол={kol:>10,.3f} | сумма={summa:>14,.2f}')
except Exception as e:
    print(f'  Ошибка: {e}')

# 3. Проверим контроль остатков
print()
print('=== Настройки контроля остатков ===')
try:
    q3 = conn.NewObject('Query')
    q3.Text = 'SELECT Наименование, Значение FROM Константа.КонтролироватьОстаткиТоваров'
    # Это может не работать, попробуем через функцию
except:
    pass

try:
    val = conn.Константы.КонтрольОстатковТоваров.Получить()
    print(f'  КонтрольОстатковТоваров = {val}')
except:
    try:
        val = conn.Константы.КонтролироватьОстаткиТоваров.Получить()
        print(f'  КонтролироватьОстаткиТоваров = {val}')
    except:
        print('  Не удалось прочитать настройку контроля остатков')

# 4. Тест: создать Комплектацию через COM и провести — проверить что произойдёт
print()
print('=== Тест: ручное проведение через COM ===')
try:
    doc = conn.Documents.КомплектацияНоменклатуры.CreateDocument()
    doc.Дата = conn.NewObject('Дата', 2026, 3, 30)
    doc.Организация = conn.Catalogs.Организации.FindByDescription('ГАЗИНВЕСТГРУПП ООО', True)
    doc.Склад = conn.Catalogs.Склады.FindByDescription('АКЗС Витебский', True)

    schet4102 = conn.ChartsOfAccounts.Хозрасчетный.FindByCode('41.02')
    schet4101 = conn.ChartsOfAccounts.Хозрасчетный.FindByCode('41.01')

    doc.Номенклатура = conn.Catalogs.Номенклатура.FindByDescription('АИ-92 (л)', True)
    doc.Количество = 100
    doc.СчетУчета = schet4102

    row = doc.Комплектующие.Add()
    row.Номенклатура = conn.Catalogs.Номенклатура.FindByDescription('АИ-92(т)', True)
    row.Количество = 0.074
    row.СчетУчета = schet4101
    try:
        row.СтранаПроисхождения = conn.Catalogs.СтраныМира.FindByDescription('РОССИЯ', True)
    except:
        pass

    doc.Комментарий = 'TL|TEST_KOMPL_COM'
    doc.Write(1)  # Проведение
    print(f'  Документ создан и проведён: {doc.Номер}')

    # Проводки
    q5 = conn.NewObject('Query')
    q5.Text = 'SELECT СчетДт.Код AS Дт, СчетКт.Код AS Кт, Сумма, КоличествоДт, КоличествоКт FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор = &Р'
    q5.SetParameter('Р', doc.Ref)
    r5 = q5.Execute().Choose()
    cnt = 0
    while r5.Next():
        print(f'    Дт {str(r5.Дт).strip()} Кт {str(r5.Кт).strip()} | сумма={float(r5.Сумма or 0):,.2f} | кол.Дт={float(r5.КоличествоДт or 0):,.3f} | кол.Кт={float(r5.КоличествоКт or 0):,.3f}')
        cnt += 1
    if cnt == 0:
        print('  ПРОВОДОК НЕТ')

    # Удалим тестовый документ
    doc.Write(2)  # Отмена проведения
    doc.SetDeletionMark(True)
    print('  Тестовый документ удалён')

except Exception as e:
    print(f'  Ошибка: {e}')
