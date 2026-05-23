# -*- coding: utf-8 -*-
"""Сравнение проводок TL-ОРП vs бухгалтерский ОРП (созданный вручную)
в БП ГИГ. Цель — понять, как типовая БП формирует 90.01.1 при
СуммаВключаетНДС=Истина (если такие ОРП есть в БД).

py -3.13-32 scripts/probe_orp_vat_compare.py
"""
import io, sys, pythoncom, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _tl_config as cfg
pythoncom.CoInitialize()
import win32com.client
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

def _val(obj, name):
    v = getattr(obj, name)
    try:
        return v() if callable(v) and not hasattr(v, "_oleobj_") else v
    except TypeError:
        return v
def num(o, n, d=0.0):
    v = _val(o, n)
    try: return float(v or d)
    except: return float(d)
def s(o, n, d=""):
    v = _val(o, n)
    return str(v) if v is not None else d

conn = win32com.client.Dispatch("V83.COMConnector")
ib = cfg.connect(conn)

def show_orp_postings(num_doc, dataref):
    """Печатает шапку + проводки конкретного ОРП по ссылке."""
    obj = dataref.ПолучитьОбъект()
    try:
        svn = obj.СуммаВключаетНДС
    except Exception:
        svn = None
    try:
        vo = obj.ВидОперации
    except Exception:
        vo = None
    print(f"\n  ОРП {num_doc}: СуммаВключаетНДС={svn}, ВидОперации={vo}")
    print(f"  СуммаДокумента = {num(obj, 'СуммаДокумента'):.2f}")

    # ТЧ Товары
    q = ib.NewObject("Query")
    q.Text = """ВЫБРАТЬ
        Т.Количество, Т.Цена, Т.Сумма, Т.СтавкаНДС, Т.СуммаНДС
    ИЗ Документ.ОтчетОРозничныхПродажах.Товары КАК Т
    ГДЕ Т.Ссылка = &С"""
    q.УстановитьПараметр("С", dataref)
    rr = q.Выполнить().Выбрать()
    sum_sum = 0.0; sum_nds = 0.0
    while rr.Следующий():
        sum_sum += num(rr, "Сумма")
        sum_nds += num(rr, "СуммаНДС")
    print(f"  Σ ТЧ Сумма = {sum_sum:.2f}, Σ ТЧ СуммаНДС = {sum_nds:.2f}")

    # Проводки
    q2 = ib.NewObject("Query")
    q2.Text = """ВЫБРАТЬ
        ЕСТЬNULL(Прв.СчетДт.Код, "") КАК ДтКод,
        ЕСТЬNULL(Прв.СчетКт.Код, "") КАК КтКод,
        Прв.Сумма КАК Сумма
    ИЗ РегистрБухгалтерии.Хозрасчетный КАК Прв
    ГДЕ Прв.Регистратор = &С"""
    q2.УстановитьПараметр("С", dataref)
    r2 = q2.Выполнить().Выбрать()
    agg = {}
    while r2.Следующий():
        k = (s(r2, "ДтКод"), s(r2, "КтКод"))
        agg[k] = agg.get(k, 0.0) + num(r2, "Сумма")
    s9001 = s9003 = 0.0
    for (dt, kt), sm in sorted(agg.items()):
        print(f"    {dt:>10} {kt:>10}  {sm:>14.2f}")
        if kt == "90.01.1": s9001 += sm
        if kt == "68.02":   s9003 += sm
    sd = num(obj, "СуммаДокумента")
    total = s9001 + s9003
    diff = total - sd
    pct = diff / sd * 100 if sd else 0
    flag = "✅ сходится" if abs(diff) < 0.05 else "❌ расхождение"
    print(f"  Σ(90.01.1+90.03)={total:.2f} vs СуммаДок={sd:.2f}: {diff:+.2f} ({pct:+.2f}%) {flag}")

print(f"=== Подключено к {cfg.BASE_PATH} ===")

# 1) Любой TL-ОРП за май 2026
print("\n== ТЛ-ОРП (наш канал) ==")
q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 2
    Д.Ссылка, Д.Номер, Д.Дата, Д.Комментарий
ИЗ Документ.ОтчетОРозничныхПродажах КАК Д
ГДЕ Д.Проведен И Д.Комментарий ПОДОБНО "TL|%"
    И Д.Дата >= ДАТАВРЕМЯ(2026, 5, 1)
УПОРЯДОЧИТЬ ПО Д.Дата УБЫВ"""
r = q.Выполнить().Выбрать()
while r.Следующий():
    show_orp_postings(s(r, "Номер"), r.Ссылка)

# 2) Любой не-TL ОРП (созданный руками/импортом) за май-апрель 2026
print("\n\n== Не-TL ОРП (для сравнения, если есть) ==")
q2 = ib.NewObject("Query")
q2.Text = """ВЫБРАТЬ ПЕРВЫЕ 3
    Д.Ссылка, Д.Номер, Д.Дата, Д.Комментарий
ИЗ Документ.ОтчетОРозничныхПродажах КАК Д
ГДЕ Д.Проведен И (Д.Комментарий НЕ ПОДОБНО "TL|%" ИЛИ Д.Комментарий = "")
    И Д.Дата >= ДАТАВРЕМЯ(2026, 1, 1)
УПОРЯДОЧИТЬ ПО Д.Дата УБЫВ"""
r2 = q2.Выполнить().Выбрать()
cnt = 0
while r2.Следующий():
    cnt += 1
    show_orp_postings(s(r2, "Номер"), r2.Ссылка)
if cnt == 0:
    print("  (в БД нет ОРП кроме наших TL — сравнить не с чем)")
