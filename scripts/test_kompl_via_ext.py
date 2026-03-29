# -*- coding: utf-8 -*-
"""Перепроведение всех Комплектаций через модуль расширения."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

# Получим все Комплектации без проводок
q = conn.NewObject('Query')
q.Text = '''SELECT Ссылка, Номер, Комментарий FROM Документ.КомплектацияНоменклатуры
WHERE Комментарий LIKE "%TL|ТТН%" AND НЕ ПометкаУдаления
ORDER BY Номер'''
r = q.Execute().Choose()

docs = []
while r.Next():
    docs.append((r.Ссылка, str(r.Номер).strip(), str(r.Комментарий)[:60]))

print(f'Найдено Комплектаций: {len(docs)}')

for ref, num, komm in docs:
    # Проверим есть ли проводки
    q2 = conn.NewObject('Query')
    q2.Text = 'SELECT COUNT(*) AS К FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор = &Р'
    q2.SetParameter('Р', ref)
    r2 = q2.Execute().Choose()
    r2.Next()
    has_postings = int(str(r2.К)) > 0

    if has_postings:
        print(f'  №{num} | OK (есть проводки)')
        continue

    # Нет проводок — перепроведём
    doc = ref.ПолучитьОбъект()
    was_posted = doc.Проведен

    try:
        # Способ: Отмена + Проведение через РежимЗаписиДокумента
        if was_posted:
            doc.Проведен = False
            doc.Записать()

        doc.Проведен = True
        doc.Записать()

        # Проверим проводки
        r2 = q2.Execute().Choose()
        r2.Next()
        now_has = int(str(r2.К)) > 0
        if now_has:
            print(f'  №{num} | ПЕРЕПРОВЕДЁН — проводки появились!')
        else:
            print(f'  №{num} | перепроведён, но проводок НЕТ (нужен РежимЗаписиДокумента.Проведение)')
    except Exception as e:
        print(f'  №{num} | ошибка: {e}')
