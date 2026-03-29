# -*- coding: utf-8 -*-
"""Проверить содержимое документа Комплектация ТТН 3310."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

q = conn.NewObject('Query')
q.Text = 'SELECT Ссылка FROM Документ.КомплектацияНоменклатуры WHERE Комментарий LIKE "%TL|ТТН|65|5|3310|КОМПЛ%"'
res = q.Execute().Choose()
if res.Next():
    ref = res.Ссылка
    doc = ref.ПолучитьОбъект()

    print('=== Комплектация ТТН 3310 ===')
    print(f'  Дата: {doc.Дата}')
    print(f'  Номер: {doc.Номер}')
    print(f'  Проведен: {doc.Проведен}')
    try: print(f'  Организация: {doc.Организация}')
    except: pass
    try: print(f'  Склад: {doc.Склад}')
    except: pass

    # Табличные части
    md = doc.Метаданные()
    tps = md.ТабличныеЧасти
    for i in range(tps.Количество()):
        tp = tps.Получить(i)
        tp_name = str(tp.Имя)
        tp_data = getattr(doc, tp_name, None)
        cnt = tp_data.Количество() if tp_data else 0
        print()
        print(f'--- ТЧ "{tp_name}": {cnt} строк ---')
        if cnt > 0:
            cols = tp.Реквизиты
            col_names = []
            for j in range(cols.Количество()):
                col_names.append(str(cols.Получить(j).Имя))
            print(f'  Колонки: {", ".join(col_names[:10])}')
            for row_idx in range(min(cnt, 5)):
                row = tp_data.Получить(row_idx)
                print(f'  [{row_idx}]')
                for c in col_names[:10]:
                    try:
                        v = getattr(row, c, '?')
                        print(f'    {c} = {v}')
                    except:
                        print(f'    {c} = (ошибка)')
else:
    print('Документ не найден')
