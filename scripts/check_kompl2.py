# -*- coding: utf-8 -*-
"""Проверить ВСЕ реквизиты комплектации ТТН 3310."""
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
    md = doc.Метаданные()

    print('=== ВСЕ РЕКВИЗИТЫ шапки ===')
    attrs = md.Реквизиты
    for i in range(attrs.Количество()):
        a = attrs.Получить(i)
        name = str(a.Имя)
        try:
            v = getattr(doc, name, '?')
            v_str = str(v)[:80]
            if 'COMObject' in v_str:
                # Попробуем Представление
                try:
                    v_str = str(v.Наименование)[:80]
                except:
                    try:
                        v_str = str(v.Представление)[:80] if hasattr(v, 'Представление') else v_str
                    except:
                        try:
                            v_str = str(v.Код)[:80]
                        except:
                            pass
            print(f'  {name} = {v_str}')
        except Exception as e:
            print(f'  {name} = (ошибка: {e})')

    # Стандартные реквизиты
    print()
    print('=== Стандартные реквизиты ===')
    for name in ['Номер', 'Дата', 'Проведен', 'ПометкаУдаления', 'Комментарий']:
        try:
            v = getattr(doc, name, '?')
            print(f'  {name} = {v}')
        except:
            pass
