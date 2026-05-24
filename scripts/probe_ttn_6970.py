# -*- coding: utf-8 -*-
"""Запросить ТТН по смене 208/6970 из STS — посмотреть RAW JSON."""
import os, sys, json, ssl, urllib.request
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
print(f"OK auth sys=15")

# Получаем поступления (ТТН) для смены 208/6970
endpoints = [
    "/v1/report/receipts?system=15&station=208&shift=6970",
]

for ep in endpoints:
    print(f"\n--- {ep} ---")
    try:
        req2 = urllib.request.Request(f"{URL}{ep}", headers={"Authorization": f"Bearer {token}"})
        with urllib.request.urlopen(req2, context=ctx, timeout=15) as r:
            data = r.read().decode()
        # Покажем красиво
        parsed = json.loads(data)
        # Сводка ВСЕХ ТТН в смене (включая парные)
        if isinstance(parsed, list) and parsed:
            shifts = parsed[0].get("shifts", [])
            if shifts:
                receipts = shifts[0].get("receipt", [])
                print(f"Всего ТТН в смене: {len(receipts)}\n")
                for i, r in enumerate(receipts):
                    ttn = r.get("ttn")
                    tank = r.get("tank")
                    svc = (r.get("service") or {})
                    doc = (r.get("doc") or {})
                    vol = doc.get("volume")
                    amt = doc.get("amount")
                    dens = doc.get("density")
                    dt = r.get("dt")
                    base = (r.get("base") or {}).get("name", "")
                    print(f"  [{i+1}] ТТН №{ttn}  tank={tank}  service={svc.get('service_code')}/{svc.get('service_name')}")
                    print(f"        doc.volume={vol}  doc.amount={amt}  density={dens}  base={base}  dt={dt}")
    except Exception as e:
        print(f"ОШИБКА: {e}")
