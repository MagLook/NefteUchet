# -*- coding: utf-8 -*-
"""Final attempt to post. Run: py -3.13-32 scripts/post_final.py"""
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

# Set all required accounts
plan = ib.ПланыСчетов.Хозрасчетный
acct_41_02 = plan.НайтиПоКоду("41.02")
acct_60_01 = plan.НайтиПоКоду("60.01")
acct_19_03 = plan.НайтиПоКоду("19.03")

print(f"41.02: {acct_41_02}, empty: {acct_41_02.Пустая() if acct_41_02 else True}")
print(f"60.01: {acct_60_01}, empty: {acct_60_01.Пустая() if acct_60_01 else True}")
print(f"19.03: {acct_19_03}, empty: {acct_19_03.Пустая() if acct_19_03 else True}")

for i in range(doc.Товары.Количество()):
    row = doc.Товары.Получить(i)
    # Set all account fields
    for attr, acct in [
        ("СчетУчета", acct_41_02),
        ("СчетУчетаБУ", acct_41_02),
        ("СчетРасчетовСКонтрагентом", acct_60_01),
        ("СчетРасчетовСКонтрагентомБУ", acct_60_01),
        ("СчетУчетаНДС", acct_19_03),
        ("СчетУчетаНДСБУ", acct_19_03),
    ]:
        try:
            setattr(row, attr, acct)
        except:
            pass
    print(f"  Row {i}: accounts set")

# Set header accounts too
try: doc.СчетРасчетовСКонтрагентом = acct_60_01
except: pass
try: doc.СчетРасчетовСКонтрагентомБУ = acct_60_01
except: pass

# Try posting
print("\nPosting...")
try:
    doc.Записать(ib.РежимЗаписиДокумента.Проведение)
    print("POSTED OK!")
except Exception as e:
    err = str(e)
    print(f"Error: {err}")
    # Save anyway
    try:
        doc.Записать()
        print("Saved (not posted)")
    except:
        pass
