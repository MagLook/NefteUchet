# -*- coding: utf-8 -*-
"""Полная проверка проведённых документов TradeLedger."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

print('=' * 80)
print('ПОЛНАЯ ПРОВЕРКА ПРОВЕДЁННЫХ ДОКУМЕНТОВ')
print('=' * 80)

# === 1. Все проведённые TL| документы ===
all_refs = []
for doc_type in ['ОтчетОРозничныхПродажах', 'ПеремещениеТоваров', 'КомплектацияНоменклатуры']:
    q = conn.NewObject('Query')
    q.Text = 'SELECT Ссылка, Номер, Дата, Комментарий FROM Документ.' + doc_type + ' WHERE Комментарий LIKE "%TL|%" AND Проведен = TRUE ORDER BY Дата DESC'
    r = q.Execute().Choose()
    docs = []
    while r.Next():
        docs.append((r.Ссылка, str(r.Номер).strip(), str(r.Дата)[:10], str(r.Комментарий)[:70], doc_type))
        all_refs.append((r.Ссылка, str(r.Комментарий)[:60], doc_type))
    if docs:
        print(f'\n[{doc_type}] — {len(docs)} проведён(о):')
        for _, num, dt, komm, _ in docs:
            print(f'  №{num} {dt} | {komm}')

# === 2. Проводки ===
print()
print('=' * 80)
print('ПРОВОДКИ')
print('=' * 80)

total_docs = 0
no_postings = []

for ref, komm, dtype in all_refs:
    q2 = conn.NewObject('Query')
    q2.Text = 'SELECT СчетДт.Код AS Дт, СчетКт.Код AS Кт, Сумма AS С, КоличествоДт AS КД, КоличествоКт AS КК FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор = &Рег ORDER BY НомерСтроки'
    q2.SetParameter('Рег', ref)
    r2 = q2.Execute().Choose()
    lines = []
    doc_sum = 0
    while r2.Next():
        dt = str(r2.Дт).strip()
        kt = str(r2.Кт).strip()
        s = float(r2.С) if r2.С else 0
        kd = float(r2.КД) if r2.КД else 0
        kk = float(r2.КК) if r2.КК else 0
        doc_sum += s
        kol = ''
        if kd: kol += f' Дт={kd:,.3f}'
        if kk: kol += f' Кт={kk:,.3f}'
        lines.append(f'    Дт {dt:8s} Кт {kt:8s} | {s:>12,.2f}{kol}')

    if lines:
        print(f'\n  [{komm}]')
        for l in lines:
            print(l)
        print(f'  ИТОГО: {doc_sum:>14,.2f} р ({len(lines)} проводок)')
        total_docs += 1
    else:
        no_postings.append(f'  {komm} ({dtype})')

print(f'\nВсего документов с проводками: {total_docs}')

# === 3. Проведённые без проводок ===
if no_postings:
    print()
    print('=' * 80)
    print(f'ПРОВЕДЁННЫЕ БЕЗ ПРОВОДОК ({len(no_postings)}):')
    for np in no_postings:
        print(np)
else:
    print('Все проведённые документы имеют проводки.')

print()
print('=' * 80)
