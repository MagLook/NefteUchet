# -*- coding: utf-8 -*-
"""Диагностика: почему Комплектация не создаёт проводок."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

# Берём последнюю проведённую Комплектацию
q = conn.NewObject('Query')
q.Text = '''SELECT ПЕРВЫЕ 1 Ссылка, Номер, Дата, Комментарий, Проведен
FROM Документ.КомплектацияНоменклатуры
WHERE Комментарий LIKE "%TL|ТТН%" AND Проведен = TRUE AND НЕ ПометкаУдаления
ORDER BY Дата DESC'''
r = q.Execute().Choose()

if not r.Next():
    print('Нет проведённых Комплектаций')
    sys.exit()

ref = r.Ссылка
print(f'=== Комплектация №{str(r.Номер).strip()} от {str(r.Дата)[:10]} ===')
print(f'Комментарий: {r.Комментарий}')
print(f'Проведен: {r.Проведен}')

doc = ref.ПолучитьОбъект()
md = doc.Метаданные()

# Все реквизиты шапки
print('\n--- Реквизиты шапки ---')
attrs = md.Реквизиты
for i in range(attrs.Количество()):
    a = attrs.Получить(i)
    name = str(a.Имя)
    try:
        v = getattr(doc, name, '?')
        v_str = str(v)[:80]
        if 'COMObject' in v_str:
            try: v_str = str(v.Наименование)[:80]
            except:
                try: v_str = str(v.Код)[:80]
                except: pass
        print(f'  {name} = {v_str}')
    except:
        print(f'  {name} = (ошибка)')

# Табличные части
print('\n--- Табличные части ---')
tps = md.ТабличныеЧасти
for i in range(tps.Количество()):
    tp = tps.Получить(i)
    tp_name = str(tp.Имя)
    tp_data = getattr(doc, tp_name, None)
    cnt = tp_data.Количество() if tp_data else 0
    print(f'\n  ТЧ "{tp_name}": {cnt} строк')
    if cnt > 0:
        cols = tp.Реквизиты
        col_names = []
        for j in range(cols.Количество()):
            col_names.append(str(cols.Получить(j).Имя))
        print(f'  Колонки: {", ".join(col_names)}')
        for row_idx in range(min(cnt, 3)):
            row = tp_data.Получить(row_idx)
            print(f'  [{row_idx}]')
            for c in col_names:
                try:
                    v = getattr(row, c, '?')
                    v_str = str(v)[:60]
                    if 'COMObject' in v_str:
                        try: v_str = str(v.Наименование)[:60]
                        except:
                            try: v_str = str(v.Код)[:60]
                            except: pass
                    print(f'    {c} = {v_str}')
                except:
                    print(f'    {c} = (ошибка)')

# Проводки
print('\n--- Проводки в регистре бухгалтерии ---')
q2 = conn.NewObject('Query')
q2.Text = 'SELECT СчетДт.Код AS Дт, СчетКт.Код AS Кт, Сумма, КоличествоДт, КоличествоКт FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор = &Р'
q2.SetParameter('Р', ref)
r2 = q2.Execute().Choose()
cnt = 0
while r2.Next():
    dt = str(r2.Дт).strip()
    kt = str(r2.Кт).strip()
    s = float(r2.Сумма or 0)
    kd = float(r2.КоличествоДт or 0)
    kk = float(r2.КоличествоКт or 0)
    print(f'  Дт {dt} Кт {kt} | сумма={s:,.2f} | кол.Дт={kd:,.3f} | кол.Кт={kk:,.3f}')
    cnt += 1
if cnt == 0:
    print('  (НЕТ ПРОВОДОК)')

# Проверим движения по другим регистрам
print('\n--- Движения по регистрам ---')
try:
    movements = doc.ПолучитьОбъект().Движения if hasattr(doc, 'Движения') else None
    if movements is None:
        movements = doc.Движения
    # Перечислим все регистры движений
    for i in range(movements.Количество()):
        reg = movements.Получить(i)
        reg_name = str(reg.Метаданные().ПолноеИмя())
        reg.Прочитать()
        cnt_mov = reg.Количество()
        print(f'  {reg_name}: {cnt_mov} записей')
        if cnt_mov > 0 and cnt_mov <= 5:
            for j in range(cnt_mov):
                row = reg.Получить(j)
                print(f'    [{j}] (детали в регистре)')
except Exception as e:
    print(f'  Ошибка чтения движений: {e}')

# Сравним с рабочей (ручной) Комплектацией из базы
print('\n--- Ищем РУЧНУЮ Комплектацию (не TL|) для сравнения ---')
q3 = conn.NewObject('Query')
q3.Text = '''SELECT ПЕРВЫЕ 1 Ссылка, Номер, Дата, Комментарий FROM Документ.КомплектацияНоменклатуры
WHERE Проведен = TRUE AND НЕ ПометкаУдаления AND НЕ Комментарий LIKE "%TL|%"
ORDER BY Дата DESC'''
r3 = q3.Execute().Choose()
if r3.Next():
    ref3 = r3.Ссылка
    print(f'  №{str(r3.Номер).strip()} от {str(r3.Дата)[:10]} | {r3.Комментарий}')
    doc3 = ref3.ПолучитьОбъект()
    md3 = doc3.Метаданные()

    print('  Реквизиты шапки:')
    attrs3 = md3.Реквизиты
    for i in range(attrs3.Количество()):
        a = attrs3.Получить(i)
        name = str(a.Имя)
        try:
            v = getattr(doc3, name, '?')
            v_str = str(v)[:60]
            if 'COMObject' in v_str:
                try: v_str = str(v.Наименование)[:60]
                except:
                    try: v_str = str(v.Код)[:60]
                    except: pass
            # Показываем только заполненные
            if v_str and v_str != '0' and v_str != 'False' and v_str != '' and v_str != 'None':
                print(f'    {name} = {v_str}')
        except:
            pass

    print('  Табличная часть Комплектующие:')
    tp3 = getattr(doc3, 'Комплектующие', None)
    if tp3 and tp3.Количество() > 0:
        row3 = tp3.Получить(0)
        cols3 = md3.ТабличныеЧасти.Комплектующие.Реквизиты
        for j in range(cols3.Количество()):
            c = str(cols3.Получить(j).Имя)
            try:
                v = getattr(row3, c, '?')
                v_str = str(v)[:60]
                if 'COMObject' in v_str:
                    try: v_str = str(v.Наименование)[:60]
                    except:
                        try: v_str = str(v.Код)[:60]
                        except: pass
                print(f'    {c} = {v_str}')
            except:
                print(f'    {c} = (ошибка)')

    # Проводки ручной
    q4 = conn.NewObject('Query')
    q4.Text = 'SELECT СчетДт.Код AS Дт, СчетКт.Код AS Кт, Сумма, КоличествоДт, КоличествоКт FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор = &Р'
    q4.SetParameter('Р', ref3)
    r4 = q4.Execute().Choose()
    print('  Проводки:')
    while r4.Next():
        dt = str(r4.Дт).strip()
        kt = str(r4.Кт).strip()
        s = float(r4.Сумма or 0)
        kd = float(r4.КоличествоДт or 0)
        kk = float(r4.КоличествоКт or 0)
        print(f'    Дт {dt} Кт {kt} | сумма={s:,.2f} | кол.Дт={kd:,.3f} | кол.Кт={kk:,.3f}')
else:
    print('  Нет ручных Комплектаций для сравнения')
