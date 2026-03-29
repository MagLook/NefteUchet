# -*- coding: utf-8 -*-
"""Сравнение субконто: наше Перемещение vs INIT Перемещение (рабочее)."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

def show_postings(label, ref):
    print(f'\n=== {label} ===')
    # Движения через набор записей
    doc = ref.ПолучитьОбъект()
    try:
        movements = doc.Движения
        for i in range(movements.Количество()):
            reg_set = movements.Получить(i)
            reg_name = str(reg_set.Метаданные().ПолноеИмя())
            reg_set.Прочитать()
            cnt = reg_set.Количество()
            if cnt == 0:
                continue
            print(f'\n  Регистр: {reg_name} ({cnt} записей)')
            for j in range(min(cnt, 5)):
                row = reg_set.Получить(j)
                # Попробуем основные поля
                fields = {}
                for fname in ['СчетДт', 'СчетКт', 'Сумма', 'КоличествоДт', 'КоличествоКт',
                              'Счет', 'Количество', 'Стоимость',
                              'ВидДвижения', 'Номенклатура', 'Склад']:
                    try:
                        v = getattr(row, fname, None)
                        if v is not None:
                            vs = str(v)[:40]
                            if 'COMObject' in vs:
                                try: vs = str(v.Наименование)[:40]
                                except:
                                    try: vs = str(v.Код)[:40]
                                    except: pass
                            if vs and vs != '0' and vs != '0.0' and vs != 'None':
                                fields[fname] = vs
                    except:
                        pass
                print(f'    [{j}] {fields}')
    except Exception as e:
        print(f'  Ошибка движений: {e}')

# 1. Наше Перемещение ТТН
q1 = conn.NewObject('Query')
q1.Text = 'SELECT ПЕРВЫЕ 1 Ссылка FROM Документ.ПеремещениеТоваров WHERE Комментарий LIKE "%TL|ТТН|65|5|2994|2|ПЕРЕМ%" AND НЕ ПометкаУдаления'
r1 = q1.Execute().Choose()
if r1.Next():
    show_postings('Наше Перемещение ТТН 2994|2|ПЕРЕМ', r1.Ссылка)

# 2. INIT Перемещение (рабочее)
q2 = conn.NewObject('Query')
q2.Text = 'SELECT ПЕРВЫЕ 1 Ссылка FROM Документ.ПеремещениеТоваров WHERE Комментарий LIKE "%TL|INIT|Перемещение%" AND Проведен = TRUE AND НЕ ПометкаУдаления'
r2 = q2.Execute().Choose()
if r2.Next():
    show_postings('INIT Перемещение (рабочее)', r2.Ссылка)

# 3. Наша Комплектация
q3 = conn.NewObject('Query')
q3.Text = 'SELECT ПЕРВЫЕ 1 Ссылка FROM Документ.КомплектацияНоменклатуры WHERE Комментарий LIKE "%TL|ТТН|65|5|2994|2|КОМПЛ%" AND НЕ ПометкаУдаления'
r3 = q3.Execute().Choose()
if r3.Next():
    show_postings('Наша Комплектация ТТН 2994|2|КОМПЛ', r3.Ссылка)

# 4. INIT Комплектация (рабочая)
q4 = conn.NewObject('Query')
q4.Text = 'SELECT ПЕРВЫЕ 1 Ссылка FROM Документ.КомплектацияНоменклатуры WHERE Комментарий LIKE "%TL|INIT|Комплектация АИ-92%" AND Проведен = TRUE AND НЕ ПометкаУдаления'
r4 = q4.Execute().Choose()
if r4.Next():
    show_postings('INIT Комплектация АИ-92 (рабочая)', r4.Ссылка)

# 5. Ручная Комплектация бухгалтера
q5 = conn.NewObject('Query')
q5.Text = 'SELECT ПЕРВЫЕ 1 Ссылка FROM Документ.КомплектацияНоменклатуры WHERE НЕ Комментарий LIKE "%TL|%" AND Проведен = TRUE AND НЕ ПометкаУдаления ORDER BY Дата DESC'
r5 = q5.Execute().Choose()
if r5.Next():
    show_postings('Ручная Комплектация бухгалтера', r5.Ссылка)
