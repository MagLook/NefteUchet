# -*- coding: utf-8 -*-
"""Текущие остатки 41.01 и 41.02 на АКЗС Витебский."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

for schet_code in ['41.01', '41.02']:
    print(f'=== {schet_code} на АКЗС Витебский ===')
    q = conn.NewObject('Query')
    q.Text = '''SELECT
        Субконто1.Наименование AS Ном,
        КоличествоОстатокДт AS Кол,
        СуммаОстатокДт AS Сумма
    FROM РегистрБухгалтерии.Хозрасчетный.Остатки(
        , Счет = &Счет, ,
        Субконто2 = &Склад
    )'''
    q.SetParameter('Счет', conn.ChartsOfAccounts.Хозрасчетный.FindByCode(schet_code))
    q.SetParameter('Склад', conn.Catalogs.Склады.FindByDescription('АКЗС Витебский', True))
    r = q.Execute().Choose()
    cnt = 0
    while r.Next():
        nom = str(r.Ном)[:40]
        kol = float(r.Кол or 0)
        summa = float(r.Сумма or 0)
        if kol != 0 or summa != 0:
            print(f'  {nom:40s} | {kol:>12,.3f} | {summa:>14,.2f}')
            cnt += 1
    if cnt == 0:
        print('  (пусто)')
    print()
