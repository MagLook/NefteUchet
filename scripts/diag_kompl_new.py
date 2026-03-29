# -*- coding: utf-8 -*-
"""Диагностика новой Комплектации (после фикса СтранаПроисхождения)."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

# Берём последнюю Комплектацию (новую, 305+)
q = conn.NewObject('Query')
q.Text = '''SELECT ПЕРВЫЕ 1 Ссылка, Номер, Дата, Комментарий
FROM Документ.КомплектацияНоменклатуры
WHERE Комментарий LIKE "%TL|ТТН%" AND Проведен = TRUE AND НЕ ПометкаУдаления
ORDER BY Номер DESC'''
r = q.Execute().Choose()
if not r.Next():
    print('Нет')
    sys.exit()

ref = r.Ссылка
doc = ref.ПолучитьОбъект()
print(f'=== №{str(r.Номер).strip()} | {r.Комментарий} ===')

# Шапка — ключевые поля
for name in ['Организация', 'Склад', 'СчетУчета', 'Номенклатура', 'Количество', 'СуммаДокумента', 'ВидОперации']:
    try:
        v = getattr(doc, name, '?')
        v_str = str(v)[:60]
        if 'COMObject' in v_str:
            try: v_str = str(v.Наименование)[:60]
            except:
                try: v_str = str(v.Код)[:60]
                except:
                    try: v_str = str(v.Представление)[:60]
                    except: pass
        print(f'  {name} = {v_str}')
    except:
        print(f'  {name} = (ошибка)')

# Комплектующие — ВСЕ поля
tp = doc.Комплектующие
print(f'\n  Комплектующие: {tp.Количество()} строк')
if tp.Количество() > 0:
    row = tp.Получить(0)
    md_tp = doc.Метаданные().ТабличныеЧасти.Комплектующие.Реквизиты
    for j in range(md_tp.Количество()):
        c = str(md_tp.Получить(j).Имя)
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

# Сравним с ручной
print('\n=== РУЧНАЯ Комплектация для сравнения ===')
q2 = conn.NewObject('Query')
q2.Text = '''SELECT ПЕРВЫЕ 1 Ссылка, Номер, Комментарий FROM Документ.КомплектацияНоменклатуры
WHERE Проведен = TRUE AND НЕ ПометкаУдаления AND НЕ Комментарий LIKE "%TL|%"
ORDER BY Дата DESC'''
r2 = q2.Execute().Choose()
if r2.Next():
    doc2 = r2.Ссылка.ПолучитьОбъект()
    print(f'  №{str(r2.Номер).strip()} | {r2.Комментарий}')
    for name in ['ВидОперации', 'Склад', 'СчетУчета', 'Номенклатура', 'Количество', 'СуммаДокумента']:
        try:
            v = getattr(doc2, name, '?')
            v_str = str(v)[:60]
            if 'COMObject' in v_str:
                try: v_str = str(v.Наименование)[:60]
                except:
                    try: v_str = str(v.Код)[:60]
                    except:
                        try: v_str = str(v.Представление)[:60]
                        except: pass
            # Получим значение нашего документа для сравнения
            v_our = getattr(doc, name, '?')
            v_our_str = str(v_our)[:60]
            if 'COMObject' in v_our_str:
                try: v_our_str = str(v_our.Наименование)[:60]
                except:
                    try: v_our_str = str(v_our.Код)[:60]
                    except: pass
            match = 'OK' if v_str == v_our_str else f'РАЗН: наш="{v_our_str}"'
            print(f'    {name} = {v_str} | {match}')
        except:
            pass

    tp2 = doc2.Комплектующие
    if tp2.Количество() > 0:
        print('  Комплектующие[0]:')
        row2 = tp2.Получить(0)
        row1 = doc.Комплектующие.Получить(0)
        md_tp2 = doc2.Метаданные().ТабличныеЧасти.Комплектующие.Реквизиты
        for j in range(md_tp2.Количество()):
            c = str(md_tp2.Получить(j).Имя)
            try:
                v2 = getattr(row2, c, '?')
                v1 = getattr(row1, c, '?')
                v2s = str(v2)[:40]
                v1s = str(v1)[:40]
                if 'COMObject' in v2s:
                    try: v2s = str(v2.Наименование)[:40]
                    except:
                        try: v2s = str(v2.Код)[:40]
                        except: pass
                if 'COMObject' in v1s:
                    try: v1s = str(v1.Наименование)[:40]
                    except:
                        try: v1s = str(v1.Код)[:40]
                        except: pass
                match = 'OK' if v2s == v1s else 'РАЗН'
                if match == 'РАЗН':
                    print(f'    {c}: руч="{v2s}" vs наш="{v1s}" ← {match}')
                else:
                    print(f'    {c}: {v2s} | {match}')
            except:
                pass
