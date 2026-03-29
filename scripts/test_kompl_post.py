# -*- coding: utf-8 -*-
"""Тест проведения через правильный вызов Write()."""
import sys, win32com.client, pythoncom
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

# Найдём последнюю непроведённую Комплектацию (или тестовую)
q = conn.NewObject('Query')
q.Text = 'SELECT ПЕРВЫЕ 1 Ссылка FROM Документ.КомплектацияНоменклатуры WHERE Комментарий LIKE "%TL|ТТН%" AND Проведен = TRUE AND НЕ ПометкаУдаления ORDER BY Номер DESC'
r = q.Execute().Choose()
if not r.Next():
    print('Нет документов')
    sys.exit()

ref = r.Ссылка
doc = ref.ПолучитьОбъект()
print(f'Документ: №{doc.Номер} | {doc.Комментарий}')
print(f'Проведен: {doc.Проведен}')

# Распроведём
print('\n1. Распроведение...')
try:
    # В COM: Write(DocumentWriteMode) где Проведение=1, ОтменаПроведения=2
    # Через МетаданныеОбъект
    mode_unpost = conn.Enums.get('DocumentWriteMode', None)
    if mode_unpost is None:
        # Попробуем через строковый eval
        mode = conn.Eval('РежимЗаписиДокумента.ОтменаПроведения')
        doc.Write(mode)
        print('  Распроведено через Eval')
    else:
        doc.Write(mode_unpost.ОтменаПроведения)
except Exception as e:
    print(f'  Ошибка: {e}')
    # Альтернатива
    try:
        doc.Проведен = False
        doc.Write()
        print('  Распроведено через Проведен=False')
    except Exception as e2:
        print(f'  Ошибка2: {e2}')

# Перепроведём
print('\n2. Проведение...')
try:
    mode = conn.Eval('РежимЗаписиДокумента.Проведение')
    doc.Write(mode)
    print('  Проведено через Eval')
except Exception as e:
    print(f'  Ошибка: {e}')

# Проверим проводки
print('\n3. Проводки:')
q2 = conn.NewObject('Query')
q2.Text = 'SELECT СчетДт.Код AS Дт, СчетКт.Код AS Кт, Сумма, КоличествоДт AS КД, КоличествоКт AS КК FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор = &Р'
q2.SetParameter('Р', ref)
r2 = q2.Execute().Choose()
cnt = 0
while r2.Next():
    print(f'  Дт {str(r2.Дт).strip()} Кт {str(r2.Кт).strip()} | сумма={float(r2.Сумма or 0):,.2f} | Дт.кол={float(r2.КД or 0):,.3f} | Кт.кол={float(r2.КК or 0):,.3f}')
    cnt += 1
if cnt == 0:
    print('  (НЕТ)')
