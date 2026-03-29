# -*- coding: utf-8 -*-
"""Проверка проведённых документов TradeLedger — проводки, суммы, счета."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

print('=' * 80)
print('ПРОВЕРКА ПРОВЕДЁННЫХ ДОКУМЕНТОВ TradeLedger')
print('=' * 80)

# === 1. Найти документы с комментарием TL| ===
doc_types = [
    'ОтчетОРозничныхПродажах',
    'ПеремещениеТоваров',
    'КомплектацияНоменклатуры',
    'ПоступлениеТоваровУслуг',
]

for doc_name in doc_types:
    q = conn.NewObject('Query')
    q.Text = (
        'SELECT Ссылка, Номер, Дата, Проведен, Комментарий '
        'FROM Документ.' + doc_name + ' '
        'WHERE Комментарий LIKE "%TL|%" '
        'ORDER BY Дата DESC'
    )
    try:
        result = q.Execute()
        sel = result.Choose()
        docs = []
        while sel.Next():
            docs.append({
                'ref': sel.Ссылка,
                'num': str(sel.Номер).strip(),
                'date': str(sel.Дата)[:10],
                'posted': sel.Проведен,
                'comment': str(sel.Комментарий),
            })
        if not docs:
            continue
        print()
        print(f'--- {doc_name} ({len(docs)} шт.) ---')
        for d in docs:
            st = 'ПРОВЕДЁН' if d['posted'] else 'не проведён'
            print(f"  №{d['num']} от {d['date']} | {st} | {d['comment'][:70]}")
    except Exception as e:
        print(f'  Ошибка {doc_name}: {e}')

# === 2. Проводки проведённых документов (смена 97 + ТТН 3310) ===
print()
print('=' * 80)
print('ПРОВОДКИ (смена 97 + ТТН 3310)')
print('=' * 80)

q2 = conn.NewObject('Query')
q2.Text = '''
SELECT
    Регистратор.Представление AS Документ,
    СчетДт.Код AS ДтСчет,
    СчетКт.Код AS КтСчет,
    Сумма,
    КоличествоДт AS КолДт,
    КоличествоКт AS КолКт
FROM
    РегистрБухгалтерии.Хозрасчетный
WHERE
    Регистратор.Комментарий LIKE "%TL|СМЕНА|65|5|97%"
    OR Регистратор.Комментарий LIKE "%TL|ТТН|65|5|3310%"
ORDER BY
    Регистратор.Дата, Регистратор.Номер, НомерСтроки
'''

try:
    result2 = q2.Execute()
    sel2 = result2.Choose()
    current_doc = ''
    doc_sum = 0
    while sel2.Next():
        doc = str(sel2.Документ)
        if doc != current_doc:
            if current_doc:
                print(f'  ИТОГО: {doc_sum:>14,.2f}')
                print()
            print(f'  [{doc}]')
            current_doc = doc
            doc_sum = 0

        dt = str(sel2.ДтСчет).strip()
        kt = str(sel2.КтСчет).strip()
        s = float(sel2.Сумма) if sel2.Сумма else 0
        kd = float(sel2.КолДт) if sel2.КолДт else 0
        kk = float(sel2.КолКт) if sel2.КолКт else 0
        doc_sum += s

        kol = ''
        if kd: kol += f' кол.Дт={kd:,.3f}'
        if kk: kol += f' кол.Кт={kk:,.3f}'
        print(f'    Дт {dt:8s} Кт {kt:8s} | {s:>14,.2f}{kol}')

    if current_doc:
        print(f'  ИТОГО: {doc_sum:>14,.2f}')
except Exception as e:
    print(f'Ошибка: {e}')

print()
print('=' * 80)
print('ПРОВЕРКА ЗАВЕРШЕНА')
print('=' * 80)
