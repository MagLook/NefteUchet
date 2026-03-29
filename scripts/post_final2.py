# -*- coding: utf-8 -*-
"""Post with contract. Run: py -3.13-32 scripts/post_final2.py"""
import pythoncom
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

# Find doc
q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка
ИЗ Документ.ПоступлениеТоваровУслуг
ГДЕ Комментарий ПОДОБНО &М И НЕ ПометкаУдаления И НЕ Проведен
УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q.УстановитьПараметр("М", "%INIT_STOCK%")
r = q.Выполнить().Выбрать()
if not r.Следующий():
    print("No doc found")
    exit()

doc = r.Ссылка.ПолучитьОбъект()
print(f"Found doc, rows: {doc.Товары.Количество()}")
print(f"Контрагент: {doc.Контрагент}")

# Find or create contract
контрагент = doc.Контрагент
q2 = ib.NewObject("Query")
q2.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка
ИЗ Справочник.ДоговорыКонтрагентов
ГДЕ Владелец = &К И НЕ ПометкаУдаления"""
q2.УстановитьПараметр("К", контрагент)
r2 = q2.Выполнить().Выбрать()
if r2.Следующий():
    договор = r2.Ссылка
    print(f"Found contract: {договор}")
else:
    # Create contract
    print("No contract found, creating...")
    договор_obj = ib.Справочники.ДоговорыКонтрагентов.СоздатьЭлемент()
    договор_obj.Владелец = контрагент
    договор_obj.Наименование = "Основной (TL)"
    договор_obj.Организация = doc.Организация
    try:
        договор_obj.ВидДоговора = ib.Перечисления.ВидыДоговоровКонтрагентов.СПоставщиком
    except:
        pass
    договор_obj.Записать()
    договор = договор_obj.Ссылка
    print(f"Created contract: {договор}")

# Set contract on doc
doc.ДоговорКонтрагента = договор
print("Contract set on doc")

# Post
print("Posting...")
try:
    doc.Записать(ib.РежимЗаписиДокумента.Проведение)
    print("POSTED OK!")
except Exception as e:
    print(f"Error: {e}")
    doc.Записать()
    print("Saved (not posted)")
