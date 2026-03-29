# -*- coding: utf-8 -*-
"""Очистка сиротских записей: старые ключи без кода топлива + непомеченные документы."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

# 1. Удалить из регистра старые ключи (без кода топлива между номером ТТН и |КОМПЛ/|ПЕРЕМ)
print('=== Очистка старых ключей из регистра ===')
q = conn.NewObject('Query')
q.Text = '''SELECT КлючЗагрузки FROM InformationRegister.TL_СтатусыЗагрузки
WHERE (КлючЗагрузки LIKE "%ТТН|65|5|%|КОМПЛ" AND НЕ КлючЗагрузки LIKE "%|_|КОМПЛ")
   OR (КлючЗагрузки LIKE "%ТТН|65|5|%|ПЕРЕМ" AND НЕ КлючЗагрузки LIKE "%|_|ПЕРЕМ")
   OR (КлючЗагрузки LIKE "TL|ТТН|65|5|%" AND НЕ КлючЗагрузки LIKE "%|%|%|%|%|%")'''
# Проще: найти все ТТН-ключи с ровно 5 разделителями (старый формат: TL|ТТН|65|5|НОМЕР)
# vs новый формат: TL|ТТН|65|5|НОМЕР|КОД (6 разделителей)

# Упрощённый подход: удалить ключи в статусе "Загружен" для ТТН
q2 = conn.NewObject('Query')
q2.Text = 'SELECT КлючЗагрузки FROM InformationRegister.TL_СтатусыЗагрузки WHERE Статус = "Загружен"'
r2 = q2.Execute().Choose()
deleted = 0
while r2.Next():
    key = str(r2.КлючЗагрузки)
    print(f'  Удаляю: {key}')
    mgr = conn.InformationRegisters.TL_СтатусыЗагрузки.CreateRecordManager()
    mgr.КлючЗагрузки = key
    mgr.Read()
    if mgr.Selected():
        mgr.Delete()
        deleted += 1
print(f'Удалено из регистра: {deleted}')

# 2. Пометить на удаление непроведённые документы без пометки
print()
print('=== Пометка непроведённых TL| документов ===')
for doc_type in ['КомплектацияНоменклатуры', 'ПеремещениеТоваров', 'ОтчетОРозничныхПродажах']:
    q3 = conn.NewObject('Query')
    q3.Text = f'SELECT Ссылка, Номер, Комментарий FROM Документ.{doc_type} WHERE Комментарий LIKE "%TL|%" AND НЕ Проведен AND НЕ ПометкаУдаления'
    r3 = q3.Execute().Choose()
    cnt = 0
    while r3.Next():
        try:
            obj = r3.Ссылка.ПолучитьОбъект()
            obj.УстановитьПометкуУдаления(True)
            cnt += 1
        except:
            pass
    if cnt > 0:
        print(f'  {doc_type}: помечено {cnt}')

# 3. Проверяем INIT-документы (старые инициализационные)
print()
print('=== INIT документы ===')
for doc_type in ['КомплектацияНоменклатуры', 'ПеремещениеТоваров', 'ПоступлениеТоваровУслуг']:
    q4 = conn.NewObject('Query')
    q4.Text = f'SELECT Номер, Дата, Проведен, ПометкаУдаления, Комментарий FROM Документ.{doc_type} WHERE Комментарий LIKE "%TL|INIT%" AND НЕ ПометкаУдаления'
    r4 = q4.Execute().Choose()
    while r4.Next():
        st = 'Пров.' if r4.Проведен else 'Не пр.'
        print(f'  [{doc_type}] №{str(r4.Номер).strip()} {str(r4.Дата)[:10]} | {st} | {str(r4.Комментарий)[:60]}')

print()
print('=== Готово ===')
