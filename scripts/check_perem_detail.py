# -*- coding: utf-8 -*-
"""Детали Перемещения ТТН 2994."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

q = conn.NewObject('Query')
q.Text = '''SELECT Ссылка, Номер, СкладОтправитель.Наименование AS Откуда, СкладПолучатель.Наименование AS Куда, Комментарий
FROM Документ.ПеремещениеТоваров
WHERE Комментарий LIKE "%TL|ТТН|65|5|2994|2|ПЕРЕМ%" AND НЕ ПометкаУдаления'''
r = q.Execute().Choose()
if r.Next():
    print(f'№{str(r.Номер).strip()} | {r.Комментарий}')
    print(f'Откуда: {r.Откуда}')
    print(f'Куда: {r.Куда}')

    # Товары
    doc = r.Ссылка.ПолучитьОбъект()
    tp = doc.Товары
    print(f'\nТовары: {tp.Количество()} строк')
    md_tp = doc.Метаданные().ТабличныеЧасти.Товары.Реквизиты
    cols = []
    for i in range(md_tp.Количество()):
        cols.append(str(md_tp.Получить(i).Имя))
    if tp.Количество() > 0:
        row = tp.Получить(0)
        for c in cols:
            try:
                v = getattr(row, c, '?')
                vs = str(v)[:50]
                if 'COMObject' in vs:
                    try: vs = str(v.Наименование)[:50]
                    except:
                        try: vs = str(v.Код)[:50]
                        except: pass
                if vs and vs != '0' and vs != 'False' and vs != '' and vs != 'None':
                    print(f'  {c} = {vs}')
            except:
                pass

# Теперь ручное перемещение для сравнения
print()
print('=== РУЧНОЕ Перемещение для сравнения ===')
q2 = conn.NewObject('Query')
q2.Text = '''SELECT ПЕРВЫЕ 1 Ссылка, Номер, СкладОтправитель.Наименование AS Откуда, СкладПолучатель.Наименование AS Куда, Комментарий
FROM Документ.ПеремещениеТоваров
WHERE Проведен = TRUE AND НЕ ПометкаУдаления AND НЕ Комментарий LIKE "%TL|%"
    AND СкладПолучатель.Наименование LIKE "%Витебский%"
ORDER BY Дата DESC'''
r2 = q2.Execute().Choose()
if r2.Next():
    print(f'№{str(r2.Номер).strip()} | {r2.Комментарий}')
    print(f'Откуда: {r2.Откуда}')
    print(f'Куда: {r2.Куда}')
    doc2 = r2.Ссылка.ПолучитьОбъект()
    tp2 = doc2.Товары
    if tp2.Количество() > 0:
        row2 = tp2.Получить(0)
        md_tp2 = doc2.Метаданные().ТабличныеЧасти.Товары.Реквизиты
        for i in range(md_tp2.Количество()):
            c = str(md_tp2.Получить(i).Имя)
            try:
                v = getattr(row2, c, '?')
                vs = str(v)[:50]
                if 'COMObject' in vs:
                    try: vs = str(v.Наименование)[:50]
                    except:
                        try: vs = str(v.Код)[:50]
                        except: pass
                if vs and vs != '0' and vs != 'False' and vs != '' and vs != 'None':
                    print(f'  {c} = {vs}')
            except:
                pass
else:
    print('Не найдено')
