# -*- coding: utf-8 -*-
"""Расследование АКЗС Непокоренных (sys=15, station=1) за 01-15.05.2026:
почему 17/29 смен в БП ГИГ vs 29 в STS (-12).

Для каждой STS-смены за период:
  * есть ли соответствующий TL-ОРП в БП?
  * если нет — есть ли запись об ошибке в TL_ОшибкиЗагрузки?
  * если ничего нет — смена просто не пыталась загрузиться (что показывает,
    что цикл загрузчика её пропустил, или dt_open вне периода).

py -3.13-32 scripts/probe_station_1_coverage.py
"""
import io, sys, os, json, urllib.request, pythoncom
from datetime import date, datetime, timedelta
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _tl_config as cfg
pythoncom.CoInitialize()
import win32com.client
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

STATION = 1
SYSTEM = 15
DATE_FROM = "2026-05-01"
DATE_TO = "2026-05-15"

def _val(o, n):
    v = getattr(o, n)
    try: return v() if callable(v) and not hasattr(v, "_oleobj_") else v
    except TypeError: return v
def s(o, n, d=""):
    v = _val(o, n)
    return str(v) if v is not None else d

conn = win32com.client.Dispatch("V83.COMConnector")
ib = cfg.connect(conn)

# 1) STS-смены за период
print(f"=== STS API: смены sys={SYSTEM} station={STATION} за {DATE_FROM}..{DATE_TO} ===")
qs = ib.NewObject("Query")
qs.Text = """ВЫБРАТЬ Значение ИЗ РегистрСведений.TL_Настройки ГДЕ Ключ = &К"""
def get_setting(key):
    qs.УстановитьПараметр("К", key)
    r = qs.Выполнить().Выбрать()
    return s(r, "Значение") if r.Следующий() else ""

base_url = get_setting("URLСервера") or "https://pos.autooplata.ru/tms"
login = get_setting("Логин")
pwd = get_setting("Пароль")
if not login or not pwd:
    print("Нет логина/пароля STS в TL_Настройки"); sys.exit(1)

def login_sts():
    body = {"login": login, "password": pwd, "system_id": SYSTEM,
            "user": {"id": "00000000-0000-0000-0000-000000000000", "name": "System"}}
    req = urllib.request.Request(base_url.rstrip("/") + "/v2/login",
        data=json.dumps(body).encode("utf-8"), method="POST",
        headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=20) as r:
        raw = r.read().decode("utf-8").strip().strip('"')
    if raw.startswith("{"):
        try: j = json.loads(raw); return j.get("token", raw)
        except Exception: pass
    return raw

tok = login_sts()
print(f"Token len={len(tok)}")

# /v1/shifts
def get_shifts():
    url = f"{base_url.rstrip('/')}/v1/shifts?system={SYSTEM}&station={STATION}&date_from={DATE_FROM}&date_to={DATE_TO}"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {tok}"})
    with urllib.request.urlopen(req, timeout=30) as r:
        raw = r.read().decode("utf-8")
    try: data = json.loads(raw)
    except: data = []
    if isinstance(data, dict) and "shifts" in data:
        data = data["shifts"]
    return data if isinstance(data, list) else []

shifts_all = get_shifts()
print(f"STS отдал {len(shifts_all)} смен (за всё время — API игнорирует date_from/date_to)")

# Фильтруем по dt_open в [01.05, 15.05]
shifts = []
for sh in shifts_all:
    dt_open = sh.get("dt_open") or sh.get("date") or sh.get("open_time") or ""
    # ISO формат "2026-05-01T00:00:00"
    if isinstance(dt_open, str) and DATE_FROM <= dt_open[:10] <= DATE_TO:
        shifts.append(sh)
print(f"После фильтра 01-15.05 осталось: {len(shifts)} смен")

sts_keys = []
for sh in shifts:
    shift_num = sh.get("shift") or sh.get("number") or sh.get("num")
    dt_open = sh.get("dt_open") or sh.get("date") or sh.get("open_time")
    sts_keys.append((str(shift_num), dt_open, sh))
    print(f"  shift={shift_num}  dt_open={dt_open}")

# 2) TL-ОРП на стенде за период по станции 1
print(f"\n=== TL-ОРП в БП ГИГ (стенд) за период, станция {STATION} ===")
qo = ib.NewObject("Query")
qo.Text = """ВЫБРАТЬ Д.Номер, Д.Дата, Д.Комментарий
ИЗ Документ.ОтчетОРозничныхПродажах КАК Д
ГДЕ Д.Проведен И Д.Комментарий ПОДОБНО &К
    И Д.Дата >= ДАТАВРЕМЯ(2026, 5, 1)
    И Д.Дата <  ДАТАВРЕМЯ(2026, 5, 16)
УПОРЯДОЧИТЬ ПО Д.Дата"""
qo.УстановитьПараметр("К", f"TL|СМЕНА|{SYSTEM}|{STATION}|%")
ro = qo.Выполнить().Выбрать()
loaded_shifts = set()
while ro.Следующий():
    cm = s(ro, "Комментарий")
    # парс "TL|СМЕНА|15|1|7280 | retail_cash..."
    head = cm.split("|", 5)
    if len(head) >= 5:
        loaded_shifts.add(head[4].split(" ")[0])
    print(f"  {s(ro, 'Номер'):<14} {s(ro, 'Дата'):<25} {cm[:90]}")
print(f"Итого в БП загружено смен: {len(loaded_shifts)}")

# 3) Diff: какие смены STS не загружены
print(f"\n=== STS ↔ БП: пропущенные смены ===")
missing = []
for shift_num, dt_open, raw in sts_keys:
    if shift_num not in loaded_shifts:
        missing.append((shift_num, dt_open, raw))
print(f"Пропущено в БП: {len(missing)} смен")
for shift_num, dt_open, raw in missing[:20]:
    print(f"  shift={shift_num} dt_open={dt_open}")
    # сводка raw
    print(f"    raw keys: {list(raw.keys())[:10]}")

# 4) Для каждой пропущенной — есть ли запись в TL_ОшибкиЗагрузки?
print(f"\n=== Записи об ошибках по ИсточникUUID = ключ смены ===")
if missing:
    keys_to_check = ib.NewObject("Array")
    for sn, _, _ in missing:
        keys_to_check.Add(f"TL|СМЕНА|{SYSTEM}|{STATION}|{sn}")
    qer = ib.NewObject("Query")
    qer.Text = """ВЫБРАТЬ Р.ИсточникUUID, ПРЕДСТАВЛЕНИЕ(Р.КодОшибки) КАК КодОшибки,
        Р.СообщениеОшибки, Р.ВремяРегистрации
    ИЗ РегистрСведений.TL_ОшибкиЗагрузки КАК Р
    ГДЕ Р.ИсточникUUID В (&К)
    УПОРЯДОЧИТЬ ПО Р.ВремяРегистрации"""
    qer.УстановитьПараметр("К", keys_to_check)
    rer = qer.Выполнить().Выбрать()
    by_code = {}
    while rer.Следующий():
        c = s(rer, "КодОшибки")
        by_code.setdefault(c, []).append((s(rer, "ИсточникUUID"), s(rer, "СообщениеОшибки")))
    if not by_code:
        print("  (ни одной записи — смены даже не пытались загрузить)")
    else:
        for c, lst in sorted(by_code.items(), key=lambda x: -len(x[1])):
            print(f"\n  Категория: {c} — {len(lst)} смен")
            for uid, msg in lst[:3]:
                print(f"    [{uid}]")
                print(f"      {msg[:200]}")
        print(f"\n  Итого записей об ошибках: {sum(len(l) for l in by_code.values())}")
