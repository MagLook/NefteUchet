# -*- coding: utf-8 -*-
"""Проверка массового проведения документов."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

print('=' * 80)
print('ПРОВЕРКА МАССОВОГО ПРОВЕДЕНИЯ')
print('=' * 80)

# 1. Регистр статусов
q = conn.NewObject('Query')
q.Text = 'SELECT Статус, COUNT(*) AS Кол FROM InformationRegister.TL_СтатусыЗагрузки GROUP BY Статус'
r = q.Execute().Choose()
print('\n--- Регистр TL_СтатусыЗагрузки ---')
total = 0
while r.Next():
    print(f'  {r.Статус}: {r.Кол}')
    total += int(str(r.Кол))
print(f'  Итого: {total}')

# 2. Документы по типам
print('\n--- Документы 1С (TL|) ---')
for doc_type in ['ОтчетОРозничныхПродажах', 'ПеремещениеТоваров', 'КомплектацияНоменклатуры']:
    q2 = conn.NewObject('Query')
    q2.Text = f'''SELECT
        SUM(ВЫБОР КОГДА Проведен И НЕ ПометкаУдаления ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) AS Проведено,
        SUM(ВЫБОР КОГДА НЕ Проведен И НЕ ПометкаУдаления ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) AS НеПроведено,
        SUM(ВЫБОР КОГДА ПометкаУдаления ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) AS Удалено
    FROM Документ.{doc_type}
    WHERE Комментарий LIKE "%TL|%"'''
    r2 = q2.Execute().Choose()
    if r2.Next():
        print(f'  {doc_type}: проведено={r2.Проведено or 0}, не пров.={r2.НеПроведено or 0}, удалено={r2.Удалено or 0}')

# 3. Проводки — считаем по типам документов
print('\n--- Проводки ---')
for doc_type, label in [('ОтчетОРозничныхПродажах', 'Розница'), ('ПеремещениеТоваров', 'Перемещения'), ('КомплектацияНоменклатуры', 'Комплектации')]:
    q3 = conn.NewObject('Query')
    q3.Text = f'''SELECT COUNT(РАЗЛИЧНЫЕ Регистратор) AS Док, SUM(Сумма) AS Сумма
    FROM РегистрБухгалтерии.Хозрасчетный
    WHERE Регистратор ССЫЛКА Документ.{doc_type}
        AND Регистратор.Проведен = TRUE
        AND Регистратор.Комментарий LIKE "%TL|%"'''
    try:
        r3 = q3.Execute().Choose()
        if r3.Next():
            docs = int(str(r3.Док or 0))
            summa = float(r3.Сумма or 0)
            print(f'  {label}: {docs} документов с проводками, сумма={summa:,.2f}')
    except Exception as e:
        print(f'  {label}: ошибка — {e}')

# 4. Проверка: проведённые без проводок
print('\n--- Проведённые БЕЗ проводок ---')
for doc_type in ['ОтчетОРозничныхПродажах', 'ПеремещениеТоваров', 'КомплектацияНоменклатуры']:
    q4 = conn.NewObject('Query')
    q4.Text = f'''SELECT Номер, Дата, Комментарий FROM Документ.{doc_type}
    WHERE Комментарий LIKE "%TL|%" AND Проведен = TRUE AND НЕ ПометкаУдаления
        AND НЕ Ссылка В (SELECT РАЗЛИЧНЫЕ Регистратор FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор ССЫЛКА Документ.{doc_type} AND Регистратор.Комментарий LIKE "%TL|%")'''
    try:
        r4 = q4.Execute().Choose()
        cnt = 0
        while r4.Next():
            if cnt == 0:
                print(f'  [{doc_type}]:')
            print(f'    №{str(r4.Номер).strip()} {str(r4.Дата)[:10]} | {str(r4.Комментарий)[:60]}')
            cnt += 1
        if cnt == 0:
            print(f'  [{doc_type}]: все OK')
    except:
        pass

# 5. Пример проводок последнего ОтчётОРозничныхПродажах
print('\n--- Пример проводок (последний ОтчётОРозничныхПродажах) ---')
q5 = conn.NewObject('Query')
q5.Text = '''SELECT ПЕРВЫЕ 1 Ссылка, Комментарий FROM Документ.ОтчетОРозничныхПродажах
WHERE Комментарий LIKE "%TL|СМЕНА|%" AND Проведен = TRUE AND НЕ ПометкаУдаления
ORDER BY Дата DESC'''
r5 = q5.Execute().Choose()
if r5.Next():
    ref = r5.Ссылка
    komm = str(r5.Комментарий)[:60]
    print(f'  Документ: {komm}')
    q6 = conn.NewObject('Query')
    q6.Text = 'SELECT СчетДт.Код AS Дт, СчетКт.Код AS Кт, Сумма AS С, КоличествоКт AS К FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор = &Р ORDER BY НомерСтроки'
    q6.SetParameter('Р', ref)
    r6 = q6.Execute().Choose()
    while r6.Next():
        dt = str(r6.Дт).strip()
        kt = str(r6.Кт).strip()
        s = float(r6.С or 0)
        k = float(r6.К or 0)
        kol = f' кол={k:,.2f}' if k else ''
        print(f'    Дт {dt:8s} Кт {kt:8s} | {s:>12,.2f}{kol}')

print('\n' + '=' * 80)
