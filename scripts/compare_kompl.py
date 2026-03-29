# -*- coding: utf-8 -*-
"""Compare ALL fields of our vs бухгалтер Комплектация. Run: py -3.13-32 scripts/compare_kompl.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\compare_kompl.txt", "w", encoding="utf-8")

# All header attributes
md = ib.Метаданные.Документы.Найти("КомплектацияНоменклатуры")
header_attrs = []
for i in range(md.Реквизиты.Количество()):
    header_attrs.append(md.Реквизиты.Получить(i).Имя)

kompl_attrs = []
ts = md.ТабличныеЧасти.Найти("Комплектующие")
if ts:
    for i in range(ts.Реквизиты.Количество()):
        kompl_attrs.append(ts.Реквизиты.Получить(i).Имя)

# Get бухгалтер's doc
q1 = ib.NewObject("Query")
q1.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.КомплектацияНоменклатуры
ГДЕ Проведен И Комментарий НЕ ПОДОБНО &М УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q1.УстановитьПараметр("М", "%TL|%")
r1 = q1.Выполнить().Выбрать()
r1.Следующий()
buh = r1.Ссылка.ПолучитьОбъект()

# Get our doc
q2 = ib.NewObject("Query")
q2.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.КомплектацияНоменклатуры
ГДЕ Комментарий ПОДОБНО &М И НЕ ПометкаУдаления УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q2.УстановитьПараметр("М", "%TL|INIT|Комплектация АИ-92%")
r2 = q2.Выполнить().Выбрать()
r2.Следующий()
our = r2.Ссылка.ПолучитьОбъект()

# Compare headers
out.write("=== ЗАГОЛОВОК ===\n")
out.write(f"{'Поле':<40} {'Бухгалтер':<30} {'Наш':<30}\n")
out.write("-" * 100 + "\n")
for attr in header_attrs:
    try:
        bv = getattr(buh, attr)
        ov = getattr(our, attr)
        bvs = str(bv) if bv is not None else "None"
        ovs = str(ov) if ov is not None else "None"
        # Try to get code for accounts
        try: bvs = bv.Код
        except: pass
        try: ovs = ov.Код
        except: pass
        marker = " <<<" if str(bvs) != str(ovs) else ""
        out.write(f"  {attr:<38} {str(bvs):<30} {str(ovs):<30}{marker}\n")
    except:
        pass

# Compare Комплектующие row 0
out.write(f"\n=== КОМПЛЕКТУЮЩИЕ [0] ===\n")
br = buh.Комплектующие.Получить(0)
or_ = our.Комплектующие.Получить(0)
for attr in kompl_attrs:
    try:
        bv = getattr(br, attr)
        ov = getattr(or_, attr)
        bvs = str(bv) if bv is not None else "None"
        ovs = str(ov) if ov is not None else "None"
        try: bvs = bv.Код
        except: pass
        try: ovs = ov.Код
        except: pass
        marker = " <<<" if str(bvs) != str(ovs) else ""
        out.write(f"  {attr:<38} {str(bvs):<30} {str(ovs):<30}{marker}\n")
    except:
        pass

out.close()
print("Done. Check build/compare_kompl.txt")
