# -*- coding: utf-8 -*-
"""Тест: создать Комплектацию вручную через COM на складе где есть тонны, проверить проводки."""
import sys, win32com.client, datetime
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

print('=== Тест Комплектации через COM ===')

# Найдём объекты
org = conn.Catalogs.Организации.FindByDescription('ГАЗИНВЕСТГРУПП ООО', True)
sklad = conn.Catalogs.Склады.FindByDescription('АКЗС Витебский', True)
nom_litres = conn.Catalogs.Номенклатура.FindByDescription('АИ-92 (л)', True)
nom_tons = conn.Catalogs.Номенклатура.FindByDescription('АИ-92(т)', True)
schet4101 = conn.ChartsOfAccounts.Хозрасчетный.FindByCode('41.01')
schet4102 = conn.ChartsOfAccounts.Хозрасчетный.FindByCode('41.02')
russia = conn.Catalogs.СтраныМира.FindByDescription('РОССИЯ', True)

print(f'  Организация: {org}')
print(f'  Склад: {sklad}')
print(f'  Номенклатура литры: {nom_litres}')
print(f'  Номенклатура тонны: {nom_tons}')
print(f'  Счёт 41.01: {schet4101}')
print(f'  Счёт 41.02: {schet4102}')
print(f'  Россия: {russia}')

# Создаём документ
doc = conn.Documents.КомплектацияНоменклатуры.CreateDocument()
doc.Дата = datetime.datetime(2026, 3, 30, 12, 0, 0)
doc.Организация = org
doc.Склад = sklad
doc.Комментарий = 'TL|TEST_KOMPL_MANUAL'

# Выпуск (шапка)
doc.Номенклатура = nom_litres
doc.Количество = 100
doc.СчетУчета = schet4102

# Комплектующие
row = doc.Комплектующие.Add()
row.Номенклатура = nom_tons
row.Количество = 0.074  # ~100л при плотности 0.74
row.СчетУчета = schet4101
row.СтранаПроисхождения = russia

# Проведение
try:
    doc.Write(conn.Enums.РежимЗаписиДокумента.Проведение if hasattr(conn.Enums, 'РежимЗаписиДокумента') else 1)
    print(f'\n  Документ проведён: №{doc.Номер}')
except Exception as e:
    # Попробуем по-другому
    try:
        from win32com.client import constants
        doc.Write(1)
        print(f'\n  Документ проведён (режим 1): №{doc.Номер}')
    except Exception as e2:
        print(f'\n  Ошибка проведения: {e2}')
        # Попробуем просто записать
        try:
            doc.Write()
            print(f'  Записан без проведения: №{doc.Номер}')
            # Проведём отдельно
            doc.Проведен = True
            doc.Write()
            print('  Проведён после записи')
        except Exception as e3:
            print(f'  Ошибка записи: {e3}')

# Проверим проводки
print('\n  Проводки:')
q = conn.NewObject('Query')
q.Text = '''SELECT СчетДт.Код AS Дт, СчетКт.Код AS Кт, Сумма, КоличествоДт AS КД, КоличествоКт AS КК
FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор = &Р'''
q.SetParameter('Р', doc.Ref)
r = q.Execute().Choose()
cnt = 0
while r.Next():
    print(f'    Дт {str(r.Дт).strip()} Кт {str(r.Кт).strip()} | сумма={float(r.Сумма or 0):,.2f} | Дт.кол={float(r.КД or 0):,.3f} | Кт.кол={float(r.КК or 0):,.3f}')
    cnt += 1
if cnt == 0:
    print('    (НЕТ ПРОВОДОК)')

# Удалим тестовый документ
try:
    doc.SetDeletionMark(True)
    print('\n  Тестовый документ помечен на удаление')
except:
    pass
