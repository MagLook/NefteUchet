# -*- coding: utf-8 -*-
"""Что отдаёт STS для пропущенных смен 175 (1-минутная ночная),
168 (дневная) vs загруженной 174 на АКЗС Непокоренных (sys=15, station=1).
Цель: подтвердить что пропущенные = sales=пустой."""
import io, sys, os, json, urllib.request, pythoncom
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _tl_config as cfg
pythoncom.CoInitialize()
import win32com.client
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

SYSTEM = 15; STATION = 1
SHIFTS = [(175, "пропущена ночная 1-мин"),
          (168, "пропущена дневная"),
          (174, "ЗАГРУЖЕНА (контроль)")]

conn = win32com.client.Dispatch("V83.COMConnector")
ib = cfg.connect(conn)
qs = ib.NewObject("Query")
qs.Text = """ВЫБРАТЬ Значение ИЗ РегистрСведений.TL_Настройки ГДЕ Ключ = &К"""
def gs(k):
    qs.УстановитьПараметр("К", k)
    r = qs.Выполнить().Выбрать()
    if r.Следующий():
        v = r.Значение
        return v() if callable(v) and not hasattr(v, "_oleobj_") else str(v)
    return ""
base_url = gs("URLСервера") or "https://pos.autooplata.ru/tms"
login = gs("Логин"); pwd = gs("Пароль")

def login_sts():
    body = {"login": login, "password": pwd, "system_id": SYSTEM,
            "user": {"id": "00000000-0000-0000-0000-000000000000", "name": "System"}}
    req = urllib.request.Request(base_url.rstrip("/") + "/v2/login",
        data=json.dumps(body).encode("utf-8"), method="POST",
        headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=20) as r:
        raw = r.read().decode("utf-8").strip().strip('"')
    if raw.startswith("{"):
        try: return json.loads(raw).get("token", raw)
        except: pass
    return raw

tok = login_sts()

for shift, label in SHIFTS:
    print(f"\n{'='*70}\n== shift {shift} — {label} ==\n{'='*70}")
    url = f"{base_url.rstrip('/')}/v1/report/shift_report?system={SYSTEM}&station={STATION}&shift={shift}"
    try:
        req = urllib.request.Request(url, headers={"Authorization": f"Bearer {tok}"})
        with urllib.request.urlopen(req, timeout=30) as r:
            raw = r.read().decode("utf-8")
        data = json.loads(raw)
    except Exception as e:
        print(f"!! ОШИБКА: {e}"); continue
    if isinstance(data, dict):
        for k in ["dt_open", "dt_close", "station", "shift"]:
            print(f"  {k:>10}: {data.get(k)}")
        sales = data.get("sales") or []
        print(f"  sales:  count={len(sales)}")
        if sales:
            for s in sales[:3]:
                print(f"    {s}")
        pumps = data.get("pumps") or []
        print(f"  pumps:  count={len(pumps)}")
        if pumps:
            for p in pumps[:2]:
                print(f"    {p}")
        tanks = data.get("tanks") or []
        print(f"  tanks:  count={len(tanks)}")
        payments = data.get("payments") or []
        print(f"  payments: count={len(payments)}")
    else:
        print(f"  Тип ответа: {type(data).__name__}, raw: {str(data)[:200]}")
