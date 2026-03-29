# -*- coding: utf-8 -*-
"""Compare ALL header fields of our vs бухгалтер report. Run: py -3.13-32 scripts/check_buh_header.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\header_compare.txt", "w", encoding="utf-8")

md = ib.Метаданные.Документы.Найти("ОтчетОРозничныхПродажах")
attrs = []
for i in range(md.Реквизиты.Количество()):
    attrs.append(md.Реквизиты.Получить(i).Имя)

# Бухгалтерский
q1 = ib.NewObject("Query")
q1.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.ОтчетОРозничныхПродажах
ГДЕ Проведен И Склад.Наименование = &С И Комментарий НЕ ПОДОБНО &М УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q1.УстановитьПараметр("С", "АКЗС Витебский")
q1.УстановитьПараметр("М", "%TL|%")
r1 = q1.Выполнить().Выбрать()
r1.Следующий()
buh = r1.Ссылка.ПолучитьОбъект()

# Наш
q2 = ib.NewObject("Query")
q2.Text = """ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.ОтчетОРозничныхПродажах
ГДЕ Комментарий ПОДОБНО &М И НЕ ПометкаУдаления УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
q2.УстановитьПараметр("М", "%TL|СМЕНА%")
r2 = q2.Выполнить().Выбрать()
r2.Следующий()
our = r2.Ссылка.ПолучитьОбъект()

out.write(f"{'Поле':<45} {'Бухгалтер':<35} {'Наш':<35}\n")
out.write("-" * 115 + "\n")

for attr in attrs:
    try:
        bv = getattr(buh, attr)
        ov = getattr(our, attr)
        # Get readable value
        bvs = str(bv) if bv is not None else "None"
        ovs = str(ov) if ov is not None else "None"
        try: bvs = f"{bv.Наименование} ({bv.Код})"
        except:
            try: bvs = bv.Код
            except: pass
        try: ovs = f"{ov.Наименование} ({ov.Код})"
        except:
            try: ovs = ov.Код
            except: pass
        # Check if empty ref
        try:
            if bv.Пустая(): bvs = "<пусто>"
        except: pass
        try:
            if ov.Пустая(): ovs = "<пусто>"
        except: pass

        marker = " <<<" if str(bvs) != str(ovs) else ""
        out.write(f"  {attr:<43} {str(bvs):<35} {str(ovs):<35}{marker}\n")
    except:
        pass

out.close()
print("Done")
