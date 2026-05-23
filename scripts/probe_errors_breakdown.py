# -*- coding: utf-8 -*-
"""Разбор 7 категорий ошибок РегистрСведений.TL_ОшибкиЗагрузки
за период 01-15.05.2026 на стенде GIG Base2.

Для каждой категории — distinct ЗначениеИсточника + примеры + станции
(парсятся из ИсточникUUID = ключа загрузки TL|СМЕНА|sys|station|shift).

py -3.13-32 scripts/probe_errors_breakdown.py
"""
import io, sys, os, pythoncom
from collections import defaultdict, Counter
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _tl_config as cfg
pythoncom.CoInitialize()
import win32com.client
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

def _val(o, n):
    v = getattr(o, n)
    try: return v() if callable(v) and not hasattr(v, "_oleobj_") else v
    except TypeError: return v
def s(o, n, d=""):
    v = _val(o, n)
    return str(v) if v is not None else d

conn = win32com.client.Dispatch("V83.COMConnector")
ib = cfg.connect(conn)

# Получаем все записи за период (КодОшибки через ПРЕДСТАВЛЕНИЕ для перечисления)
q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ
    ПРЕДСТАВЛЕНИЕ(Р.КодОшибки) КАК КодОшибки,
    Р.ЗначениеИсточника КАК ЗначениеИсточника,
    Р.ИсточникUUID КАК Ключ,
    Р.СообщениеОшибки КАК Сообщ,
    Р.ВремяРегистрации КАК Время,
    ПРЕДСТАВЛЕНИЕ(Р.ТипОбъекта) КАК ТипОбъекта,
    Р.Поле КАК Поле
ИЗ РегистрСведений.TL_ОшибкиЗагрузки КАК Р
ГДЕ Р.ВремяРегистрации >= ДАТАВРЕМЯ(2026, 5, 1, 0, 0, 0)
    И Р.ВремяРегистрации <= ДАТАВРЕМЯ(2026, 5, 15, 23, 59, 59)
УПОРЯДОЧИТЬ ПО Р.ВремяРегистрации"""
r = q.Выполнить().Выбрать()

# Категория → station → list of (значение, сообщение, ключ)
by_cat = defaultdict(list)
total = 0
while r.Следующий():
    code = s(r, "КодОшибки")
    val  = s(r, "ЗначениеИсточника")
    klu  = s(r, "Ключ")
    msg  = s(r, "Сообщ")
    field = s(r, "Поле")
    # парс станции из ключа TL|СМЕНА|sys|station|shift
    parts = klu.split("|")
    station = parts[3] if len(parts) >= 4 else "?"
    by_cat[code].append((station, val, msg, klu, field))
    total += 1

print(f"=== TL_ОшибкиЗагрузки за 01-15.05.2026 ===")
print(f"Всего записей: {total}\n")

# Сводка по категориям
print("== Распределение по категориям ==")
for code, items in sorted(by_cat.items(), key=lambda x: -len(x[1])):
    print(f"  {code:<35} {len(items):>4}")
print()

# Детально по каждой
for code in sorted(by_cat, key=lambda c: -len(by_cat[c])):
    items = by_cat[code]
    print(f"\n{'='*80}")
    print(f"== {code} ({len(items)} записей) ==")
    print(f"{'='*80}")

    # 1) Distinct ЗначениеИсточника
    vals = Counter(it[1] for it in items)
    print(f"\n  Distinct ЗначениеИсточника (топ-15):")
    for v, cnt in vals.most_common(15):
        print(f"    [{cnt:>4}]  {v!r}")

    # 2) По станциям
    stations = Counter(it[0] for it in items)
    print(f"\n  По станциям:")
    for st, cnt in stations.most_common():
        print(f"    станция {st}: {cnt}")

    # 3) Distinct Поле
    fields = Counter(it[4] for it in items if it[4])
    if fields:
        print(f"\n  Distinct Поле:")
        for f, cnt in fields.most_common(10):
            print(f"    [{cnt:>4}]  {f!r}")

    # 4) Примеры сообщений (первые 3 уникальных)
    msgs_uniq = []
    seen = set()
    for it in items:
        m = it[2]
        if m and m not in seen:
            seen.add(m); msgs_uniq.append(m)
            if len(msgs_uniq) >= 3: break
    if msgs_uniq:
        print(f"\n  Примеры сообщений:")
        for m in msgs_uniq:
            print(f"    - {m[:180]}")
