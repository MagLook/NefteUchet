# -*- coding: utf-8 -*-
"""Проверка проводок Комплектаций после фикса порядка."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

print('=== Комплектации (проведённые, не удалённые) ===')
q = conn.NewObject('Query')
q.Text = '''SELECT Ссылка, Номер, Дата, Комментарий
FROM Документ.КомплектацияНоменклатуры
WHERE Комментарий LIKE "%TL|ТТН%" AND Проведен = TRUE AND НЕ ПометкаУдаления
ORDER BY Номер DESC'''
r = q.Execute().Choose()

total = 0
with_postings = 0
without = 0

while r.Next():
    ref = r.Ссылка
    num = str(r.Номер).strip()
    komm = str(r.Комментарий)[:60]

    q2 = conn.NewObject('Query')
    q2.Text = 'SELECT СчетДт.Код AS Дт, СчетКт.Код AS Кт, Сумма, КоличествоДт AS КД, КоличествоКт AS КК FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор = &Р'
    q2.SetParameter('Р', ref)
    r2 = q2.Execute().Choose()

    postings = []
    while r2.Next():
        postings.append({
            'dt': str(r2.Дт).strip(),
            'kt': str(r2.Кт).strip(),
            's': float(r2.Сумма or 0),
            'kd': float(r2.КД or 0),
            'kk': float(r2.КК or 0),
        })

    total += 1
    if postings:
        with_postings += 1
        p = postings[0]
        print(f'  OK  №{num} | {komm}')
        print(f'      Дт {p["dt"]} Кт {p["kt"]} | сумма={p["s"]:>12,.2f} | литров={p["kd"]:>10,.1f} | тонн={p["kk"]:>8,.3f}')
    else:
        without += 1
        print(f'  BAD №{num} | {komm} — НЕТ ПРОВОДОК')

print()
print(f'Итого: {total} | с проводками: {with_postings} | без: {without}')
