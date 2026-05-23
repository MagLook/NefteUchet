# -*- coding: utf-8 -*-
"""Эксперимент: подтверждение варианта B фикса НДС.

На стенде GIG Base2 берём ОДИН TL-ОРП за май 2026, делаем КОПИЮ через
CopyObject (или новый документ с дублированной ТЧ), пересчитываем
Сумма = БЕЗ НДС, проводим и сравниваем проводки с оригиналом.

Эксперимент НЕ модифицирует исходный документ — только создаёт новый
тестовый с пометкой "TL|VAT-TEST|<источник>". Можно безопасно удалить
после эксперимента.

py -3.13-32 scripts/probe_orp_vat_experiment.py
"""
import io, sys, pythoncom, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _tl_config as cfg
pythoncom.CoInitialize()
import win32com.client
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

def _val(o, n):
    v = getattr(o, n)
    try: return v() if callable(v) and not hasattr(v, "_oleobj_") else v
    except TypeError: return v
def num(o, n, d=0.0):
    v = _val(o, n)
    try: return float(v or d)
    except: return float(d)
def s(o, n, d=""):
    v = _val(o, n)
    return str(v) if v is not None else d

conn = win32com.client.Dispatch("V83.COMConnector")
ib = cfg.connect(conn)

# 1) Берём первый TL-ОРП за май 2026 как образец
q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ ПЕРВЫЕ 1
    Д.Ссылка, Д.Номер, Д.СуммаДокумента, Д.Комментарий
ИЗ Документ.ОтчетОРозничныхПродажах КАК Д
ГДЕ Д.Проведен И Д.Комментарий ПОДОБНО "TL|СМЕНА%"
    И Д.Дата >= ДАТАВРЕМЯ(2026, 5, 1)
УПОРЯДОЧИТЬ ПО Д.Дата УБЫВ"""
r = q.Выполнить().Выбрать()
if not r.Следующий():
    print("Нет TL-ОРП за май 2026"); sys.exit(1)
src_ref = r.Ссылка
src_num = s(r, "Номер")
src_sum = num(r, "СуммаДокумента")
print(f"Источник: ОРП {src_num}, СуммаДокумента={src_sum:.2f}")

# 2) Создаём НОВЫЙ ОРП дублированием через копирование шапки + ТЧ
src_obj = src_ref.ПолучитьОбъект()
new_obj = ib.Документы.ОтчетОРозничныхПродажах.СоздатьДокумент()
new_obj.Дата = src_obj.Дата
new_obj.Организация = src_obj.Организация
new_obj.Склад = src_obj.Склад
try: new_obj.СкладОрганизации = src_obj.Склад
except Exception: pass
try: new_obj.ВидОперации = src_obj.ВидОперации
except Exception: pass
try: new_obj.ВалютаДокумента = src_obj.ВалютаДокумента
except Exception: pass
new_obj.КурсДокумента = 1
new_obj.КратностьДокумента = 1
new_obj.СуммаВключаетНДС = True  # эксперимент: явное прямое присваивание
new_obj.Комментарий = "TL|VAT-TEST|варBezNDS|" + src_num
try: new_obj.СчетКасса = src_obj.СчетКасса
except Exception: pass
try: new_obj.СтатьяДвиженияДенежныхСредств = src_obj.СтатьяДвиженияДенежныхСредств
except Exception: pass

# Копируем ТЧ Товары — но Сумма ставим БЕЗ НДС
for src_row in src_obj.Товары:
    nr = new_obj.Товары.Добавить()
    nr.Номенклатура = src_row.Номенклатура
    nr.Количество   = src_row.Количество
    nr.Цена         = src_row.Цена  # цена С НДС, как привыкли
    sum_s_nds = float(src_row.Сумма or 0)
    sum_nds = float(src_row.СуммаНДС or 0)
    nr.Сумма        = round(sum_s_nds - sum_nds, 2)  # ← БЕЗ НДС
    nr.СуммаНДС     = sum_nds
    nr.СтавкаНДС    = src_row.СтавкаНДС
    try: nr.СчетУчета                = src_row.СчетУчета
    except Exception: pass
    try: nr.СчетДоходов              = src_row.СчетДоходов
    except Exception: pass
    try: nr.СчетРасходов             = src_row.СчетРасходов
    except Exception: pass
    try: nr.СчетУчетаНДСПоРеализации = src_row.СчетУчетаНДСПоРеализации
    except Exception: pass

# Копируем ТЧ Оплата чтобы баланс касса/эквайринг сходился
for src_pay in src_obj.Оплата:
    np_ = new_obj.Оплата.Добавить()
    try: np_.ВидОплаты = src_pay.ВидОплаты
    except Exception: pass
    np_.СуммаОплаты = src_pay.СуммаОплаты

# Записываем с проведением (РежимЗаписиДокумента.Проведение через перечисление платформы)
try:
    new_obj.Записать(ib.РежимЗаписиДокумента.Проведение)
except Exception as e:
    print(f"!! Ошибка проведения: {e}")
    try:
        new_obj.Записать(ib.РежимЗаписиДокумента.Запись)
        print("  (записан без проведения)")
    except Exception as e2:
        print(f"!! Даже запись не прошла: {e2}")
        sys.exit(1)

new_ref = new_obj.Ссылка
new_num = s(new_obj, "Номер")
new_sum = num(new_obj, "СуммаДокумента")
print(f"Новый ОРП:  {new_num}, СуммаДокумента={new_sum:.2f}")

# Читаем проводки нового
print("\nПроводки нового ОРП:")
qp = ib.NewObject("Query")
qp.Text = """ВЫБРАТЬ
    ЕСТЬNULL(Прв.СчетДт.Код, "") КАК ДтКод,
    ЕСТЬNULL(Прв.СчетКт.Код, "") КАК КтКод,
    Прв.Сумма КАК Сумма
ИЗ РегистрБухгалтерии.Хозрасчетный КАК Прв
ГДЕ Прв.Регистратор = &С"""
qp.УстановитьПараметр("С", new_ref)
rp = qp.Выполнить().Выбрать()
agg = {}
while rp.Следующий():
    k = (s(rp, "ДтКод"), s(rp, "КтКод"))
    agg[k] = agg.get(k, 0.0) + num(rp, "Сумма")
s9001 = s9003 = 0.0
for (dt, kt), sm in sorted(agg.items()):
    print(f"  {dt:>10} {kt:>10}  {sm:>14.2f}")
    if kt == "90.01.1": s9001 += sm
    if kt == "68.02":   s9003 += sm
total = s9001 + s9003
diff = total - new_sum
print(f"\nΣ(90.01.1+90.03)={total:.2f} vs СуммаДок={new_sum:.2f}: {diff:+.2f}")
if abs(diff) < 0.05:
    print("✅ Эксперимент УСПЕХ: вариант B (Сумма БЕЗ НДС) даёт сходящиеся проводки.")
else:
    print(f"❌ Расхождение остаётся ({diff/new_sum*100:+.2f}%) — нужно копать дальше.")
print(f"\nТестовый документ помечен 'TL|VAT-TEST|варBezNDS|{src_num}', можно удалить.")
