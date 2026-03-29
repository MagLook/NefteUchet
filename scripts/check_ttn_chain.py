# -*- coding: utf-8 -*-
"""Проверка полной цепочки ТТН: ПЕРЕМ + КОМПЛ — проводки каждого."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

# Все записи регистра по ТТН 2994 и 2765
for ttn in ['2994', '2765']:
    print(f'=== ТТН {ttn} ===')
    q = conn.NewObject('Query')
    q.Text = f'SELECT КлючЗагрузки, Статус FROM InformationRegister.TL_СтатусыЗагрузки WHERE КлючЗагрузки LIKE "%ТТН|65|5|{ttn}%" ORDER BY КлючЗагрузки'
    r = q.Execute().Choose()
    keys = []
    while r.Next():
        k = str(r.КлючЗагрузки)
        s = str(r.Статус)
        keys.append(k)
        print(f'  {k} | {s}')

    # Для каждого ключа — проверим документ и проводки
    for key in keys:
        doc_ref = conn.TL_РегистрСтатусов.НайтиДокументПоКлючу(key)
        if doc_ref is None or str(doc_ref) == '':
            print(f'    {key}: документ НЕ НАЙДЕН')
            continue

        # Проводки
        q2 = conn.NewObject('Query')
        q2.Text = 'SELECT СчетДт.Код AS Дт, СчетКт.Код AS Кт, Сумма, КоличествоДт AS КД, КоличествоКт AS КК FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор = &Р'
        q2.SetParameter('Р', doc_ref)
        r2 = q2.Execute().Choose()
        postings = []
        while r2.Next():
            postings.append(f'Дт {str(r2.Дт).strip()} Кт {str(r2.Кт).strip()} | сумма={float(r2.Сумма or 0):,.2f} | Дт.кол={float(r2.КД or 0):,.3f} | Кт.кол={float(r2.КК or 0):,.3f}')

        suffix = key.split('|')[-1] if '|' in key else ''
        if postings:
            for p in postings:
                print(f'    {suffix}: {p}')
        else:
            print(f'    {suffix}: НЕТ ПРОВОДОК')
    print()
