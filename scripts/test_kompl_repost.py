# -*- coding: utf-8 -*-
"""Перепроведение через правильный COM-вызов."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

# РежимЗаписиДокумента через NewObject
try:
    # В 1С COM: нет прямого доступа к перечислениям платформы
    # Но можно использовать числовые значения:
    # 0 = Запись, 1 = Проведение, 2 = ОтменаПроведения
    pass
except:
    pass

# Найдём Комплектацию
q = conn.NewObject('Query')
q.Text = 'SELECT ПЕРВЫЕ 1 Ссылка, Номер, Комментарий FROM Документ.КомплектацияНоменклатуры WHERE Комментарий LIKE "%TL|ТТН|65|5|1458|5|КОМПЛ%" AND НЕ ПометкаУдаления ORDER BY Номер DESC'
r = q.Execute().Choose()
if not r.Next():
    print('Нет')
    sys.exit()

ref = r.Ссылка
print(f'Документ: №{str(r.Номер).strip()} | {r.Комментарий}')

# Метод 1: Через обработку данных (Execute запрос)
print('\nМетод: перепроведение через 1С-код')
try:
    # Выполним код 1С напрямую
    proc = conn.NewObject('COMSafeArray', 'VT_VARIANT', 1)
    # Не работает...
except:
    pass

# Метод 2: doc.Записать(1) — где 1 = РежимЗаписиДокумента.Проведение
doc = ref.ПолучитьОбъект()
print(f'Проведен: {doc.Проведен}')

# Сначала отменим проведение
try:
    doc.Записать(2)  # 2 = ОтменаПроведения
    print('Отмена проведения: OK')
except Exception as e:
    print(f'Отмена проведения ошибка: {e}')
    # Пробуем альтернативу
    try:
        doc.Проведен = False
        doc.Записать(0)  # 0 = Запись
        print('Отмена через Проведен=False: OK')
    except Exception as e2:
        print(f'Тоже ошибка: {e2}')

# Теперь проведём
try:
    doc.Записать(1)  # 1 = Проведение
    print('Проведение: OK')
except Exception as e:
    print(f'Проведение ошибка: {e}')

# Проверим
q2 = conn.NewObject('Query')
q2.Text = 'SELECT СчетДт.Код AS Дт, СчетКт.Код AS Кт, Сумма, КоличествоДт AS КД, КоличествоКт AS КК FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор = &Р'
q2.SetParameter('Р', ref)
r2 = q2.Execute().Choose()
print('\nПроводки:')
cnt = 0
while r2.Next():
    print(f'  Дт {str(r2.Дт).strip()} Кт {str(r2.Кт).strip()} | сумма={float(r2.Сумма or 0):,.2f} | кол.Дт={float(r2.КД or 0):,.3f} | кол.Кт={float(r2.КК or 0):,.3f}')
    cnt += 1
if cnt == 0:
    print('  (НЕТ)')
else:
    print(f'  ЕСТЬ {cnt} проводок!')
