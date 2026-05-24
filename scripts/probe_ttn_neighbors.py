# -*- coding: utf-8 -*-
"""Соседние смены АЗС 208 (6968-6972) — есть ли парные ТТН с +23411 ДТ зим."""
import os, sys, json, ssl, urllib.request, urllib.error
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _tl_config import connect

conn = connect()


def s(v):
    if v is None: return ""
    try: return str(v)
    except Exception: return ""


def setting(key):
    q = conn.NewObject("Запрос")
    q.УстановитьПараметр("К", key)
    q.Текст = "ВЫБРАТЬ ПЕРВЫЕ 1 Рег.Значение КАК З ИЗ РегистрСведений.TL_Настройки КАК Рег ГДЕ Рег.Ключ = &К"
    в = q.Выполнить().Выбрать()
    return s(в.З) if в.Следующий() else ""


URL  = setting("URLСервера")
USER = setting("Логин")
PWD  = setting("Пароль")
ctx = ssl._create_unverified_context()

body = {"login": USER, "password": PWD,
        "user": {"id": "00000000-0000-0000-0000-000000000000", "name": "System"},
        "system_id": 15}
req = urllib.request.Request(f"{URL}/v2/login",
    data=json.dumps(body).encode(),
    headers={"Content-Type": "application/json"}, method="POST")
with urllib.request.urlopen(req, context=ctx, timeout=15) as r:
    token = r.read().decode().strip().strip('"')

print("Смены 6960-6985 АЗС 208 — все ТТН")
print("=" * 90)

for смена in range(6960, 6985):
    try:
        req2 = urllib.request.Request(
            f"{URL}/v1/report/receipts?system=15&station=208&shift={смена}",
            headers={"Authorization": f"Bearer {token}"})
        with urllib.request.urlopen(req2, context=ctx, timeout=15) as r:
            data = json.loads(r.read())
    except urllib.error.HTTPError:
        continue
    except Exception as e:
        continue

    if not isinstance(data, list) or not data:
        continue
    receipts = (data[0].get("shifts") or [{}])[0].get("receipt") or []
    if not receipts:
        continue

    print(f"\nСмена {смена}: {len(receipts)} ТТН")
    for r in receipts:
        ttn = r.get("ttn")
        svc = (r.get("service") or {})
        doc = (r.get("doc") or {})
        vol = doc.get("volume", "")
        amt = doc.get("amount", "")
        dt = r.get("dt", "")
        знак = "🔴" if str(vol).startswith("-") else "🟢"
        print(f"  {знак} ТТН №{ttn:<14} {svc.get('service_code')}/{svc.get('service_name'):<10} vol={vol:<12} amount={amt:<12} dt={dt}")
