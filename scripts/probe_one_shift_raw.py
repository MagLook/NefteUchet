# -*- coding: utf-8 -*-
"""Простой быстрый probe — берёт 5 свежих смен из STS и показывает RAW sales[]."""
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
print(f"OK auth sys=15")

# Список смен по конкретным АЗС за период (нужно station в каждом запросе)
АЗС_СПИСОК = [208, 209, 210, 8, 207]
свежие = []  # [(ст, шифт, close)]

for ст in АЗС_СПИСОК:
    req2 = urllib.request.Request(
        f"{URL}/v1/shifts?system=15&station={ст}",
        headers={"Authorization": f"Bearer {token}"})
    try:
        with urllib.request.urlopen(req2, context=ctx, timeout=30) as r:
            спис = json.loads(r.read())
    except Exception as e:
        print(f"  АЗС {ст}: ошибка {e}")
        continue
    if not isinstance(спис, list): continue
    # Берём 3 самых свежих с этой АЗС
    с_сорт = sorted([с for с in спис if с.get("dt_close")],
                    key=lambda x: x["dt_close"], reverse=True)[:3]
    for с in с_сорт:
        свежие.append((ст, с["shift"], с["dt_close"]))

print(f"Будет проверено {len(свежие)} смен:")
for ст, н, close in свежие:
    print(f"  АЗС {ст}  смена {н}  close={close}")

# Для каждой смены — детальный отчёт
print(f"\n{'='*100}")
for i, (ст, н, close) in enumerate(свежие):
    print(f"\n--- Смена #{i+1}: sys=15 АЗС={ст} смена={н} close={close} ---")
    req3 = urllib.request.Request(
        f"{URL}/v1/report/shift_report?system=15&station={ст}&shift={н}",
        headers={"Authorization": f"Bearer {token}"})
    try:
        with urllib.request.urlopen(req3, context=ctx, timeout=15) as r:
            отчёт = json.loads(r.read())
    except Exception as e:
        print(f"  ОШИБКА: {e}")
        continue
    sales = отчёт.get("sales") or []
    print(f"  sales[]: {len(sales)} записей")
    for j, item in enumerate(sales):
        pay = item.get("pay_type") or {}
        fuel = item.get("fuel") or []
        σ_л = 0.0
        σ_руб = 0.0
        for f in fuel:
            rel = f.get("release") or {}
            try: σ_л += float(rel.get("volume") or 0)
            except: pass
            try: σ_руб += float(rel.get("cost") or 0)
            except: pass
        print(f"    [{j+1}] pay_type='{pay.get('name')}' id={pay.get('id')}  fuel: {σ_л:.2f} л / {σ_руб:.2f} руб")
