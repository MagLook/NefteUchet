# -*- coding: utf-8 -*-
"""Состояние сопутки и общепита на стенде GIG Base2 за 01-15.05.2026.

Что должно создаваться из ЦБ-канала:
  * Документ.ПоступлениеТоваровУслуг (kind=purchase) — B2B сверка
  * Документ.ОтчетОРозничныхПродажах (kind=retail_sale_sidegoods) — розничная модель
  * Документ.ОтчетПроизводстваЗаСмену (kind=production_release) — общепит
  * Документ.КорректировкаПоступления (kind=return_purchase) — возврат поставщику
  * Документ.ВозвратТоваровОтПокупателя (kind=return_sale) — возврат от покупателя
  * Документ.СписаниеТоваров (kind=ingredients_writeoff) — общепит писания не из ВП

Для каждого типа: сколько штук, проведённых, какие виды операций, типичные проводки.

py -3.13-32 scripts/probe_souputka_obshchepit_state.py
"""
import sys, os, pythoncom, io
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _tl_config as cfg
pythoncom.CoInitialize()
import win32com.client
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

def _val(o,n):
    v=getattr(o,n)
    try: return v() if callable(v) and not hasattr(v,"_oleobj_") else v
    except TypeError: return v
def num(o,n,d=0.0):
    v=_val(o,n)
    try: return float(v or d)
    except: return float(d)
def s(o,n,d=""):
    v=_val(o,n)
    return str(v) if v is not None else d

conn = win32com.client.Dispatch("V83.COMConnector")
ib = cfg.connect(conn)

# Все документы что нас интересуют + признак ЦБ-канала (комментарий ПОДОБНО %TL|%)
DOC_TYPES = [
    ("ПоступлениеТоваровУслуг",       "ПТУ"),
    ("ОтчетОРозничныхПродажах",       "ОРП"),
    ("ОтчетПроизводстваЗаСмену",      "ОПЗС"),
    ("КорректировкаПоступления",      "КПо"),
    ("ВозвратТоваровОтПокупателя",    "ВозТП"),
    ("СписаниеТоваров",               "СписТ"),
]

period = "Д.Дата >= ДАТАВРЕМЯ(2026,5,1) И Д.Дата < ДАТАВРЕМЯ(2026,5,16)"
tl_like = 'ВЫРАЗИТЬ(Д.Комментарий КАК СТРОКА(300)) ПОДОБНО "%TL|%"'

print(f"=== Сопутка/общепит на стенде {cfg.BASE_PATH} за 01-15.05.2026 ===\n")
print(f"{'Тип':<32} {'Всего':>7} {'TL':>4} {'провед':>7} {'не провед':>10}")
print("-" * 72)

for doc_name, short in DOC_TYPES:
    q = ib.NewObject("Query")
    q.Text = f"""ВЫБРАТЬ
        КОЛИЧЕСТВО(*) КАК Всего,
        СУММА(ВЫБОР КОГДА {tl_like} ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК TL,
        СУММА(ВЫБОР КОГДА {tl_like} И Д.Проведен ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Провед,
        СУММА(ВЫБОР КОГДА {tl_like} И НЕ Д.Проведен И НЕ Д.ПометкаУдаления ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Чернов
    ИЗ Документ.{doc_name} КАК Д
    ГДЕ {period}"""
    try:
        r = q.Выполнить().Выбрать(); r.Следующий()
        tot = int(num(r, "Всего")); tl = int(num(r, "TL"))
        pv = int(num(r, "Провед")); cv = int(num(r, "Чернов"))
        print(f"{short + ' (' + doc_name + ')':<32} {tot:>7} {tl:>4} {pv:>7} {cv:>10}")
    except Exception as e:
        print(f"{short:<32} ОШИБКА: {str(e)[:60]}")

print("\n=== Детально по TL-документам (если есть) ===")
for doc_name, short in DOC_TYPES:
    q = ib.NewObject("Query")
    q.Text = f"""ВЫБРАТЬ ПЕРВЫЕ 5
        Д.Номер, Д.Дата, ПРЕДСТАВЛЕНИЕ(Д.ВидОперации) КАК ВО,
        Д.СуммаДокумента, Д.Проведен, Д.Комментарий
    ИЗ Документ.{doc_name} КАК Д
    ГДЕ {period} И {tl_like}
    УПОРЯДОЧИТЬ ПО Д.Дата"""
    try:
        r = q.Выполнить().Выбрать()
        rows = []
        while r.Следующий():
            rows.append((s(r,"Номер"), s(r,"Дата")[:10], s(r,"ВО"),
                        num(r,"СуммаДокумента"), s(r,"Проведен"), s(r,"Комментарий")[:60]))
        if rows:
            print(f"\n## {short}: первые {len(rows)} TL-документов")
            for n,d,vo,sm,pr,cm in rows:
                pflag = "✓" if pr == "True" else "✗"
                print(f"  {pflag} {n:<14} {d}  {sm:>10.2f}  ВО={vo[:25]:<25} {cm}")
    except Exception:
        pass
