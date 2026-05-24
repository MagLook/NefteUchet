# -*- coding: utf-8 -*-
"""Найти все смены последних дней, где в sales[].fuel[].service.service_code
приходят значения отличные от 1/2/3 (стандарт АИ-92/95/ДТ).
"""
import os, sys, json, urllib.request
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _tl_config import connect

# STS API
conn = connect()


def s(v):
    if v is None: return ""
    try: return str(v)
    except Exception: return ""


# Берём настройки STS
def get_setting(key, default=""):
    q = conn.NewObject("Запрос")
    q.УстановитьПараметр("К", key)
    q.Текст = """ВЫБРАТЬ ПЕРВЫЕ 1 Рег.Значение КАК З
                 ИЗ РегистрСведений.TL_НастройкиРасширения КАК Рег
                 ГДЕ Рег.Ключ = &К"""
    в = q.Выполнить().Выбрать()
    return s(в.З) if в.Следующий() else default


URL = get_setting("STS_URL", "https://pos.autooplata.ru/tms")
USER = get_setting("STS_Логин", "OnlinePC_baltop")
PWD = get_setting("STS_Пароль", "")
SYS = "65"  # ГИГ

# Авторизация
import ssl
import urllib.error
ctx = ssl._create_unverified_context()
req = urllib.request.Request(
    f"{URL}/v1/session",
    data=json.dumps({"login": USER, "password": PWD}).encode(),
    headers={"Content-Type": "application/json"},
    method="POST"
)
try:
    with urllib.request.urlopen(req, context=ctx, timeout=15) as r:
        token = json.loads(r.read())["token"]
        print(f"OK авторизация под {USER}")
except Exception as e:
    print(f"ОШИБКА авторизации: {e}")
    sys.exit(1)

# Список смен (последние) по всем АЗС сети 15 — там были ошибки
print("\nЗагружаю свежие смены (по всем АЗС)...")
ВСЕ_АНОМАЛИИ = []
ПЕРИОД_DT_FROM = "2026-05-15"
ПЕРИОД_DT_TO = "2026-05-22"

for sys_id in [15, 65]:
    req2 = urllib.request.Request(
        f"{URL}/v1/shifts/?system={sys_id}&dt_from={ПЕРИОД_DT_FROM}&dt_to={ПЕРИОД_DT_TO}",
        headers={"Authorization": f"Bearer {token}"},
        method="GET"
    )
    try:
        with urllib.request.urlopen(req2, context=ctx, timeout=60) as r:
            смены = json.loads(r.read())
    except Exception as e:
        print(f"  system={sys_id}: ошибка {e}")
        continue
    print(f"  system={sys_id}: смен {len(смены) if isinstance(смены, list) else '?'}")
    if not isinstance(смены, list):
        continue

    # Проверим первые 200 смен (быстро)
    проверено = 0
    for см in смены[:300]:
        ст = см.get("station")
        н = см.get("shift")
        if not ст or not н:
            continue

        req3 = urllib.request.Request(
            f"{URL}/v1/report/shift_report?system={sys_id}&station={ст}&shift={н}",
            headers={"Authorization": f"Bearer {token}"},
            method="GET"
        )
        try:
            with urllib.request.urlopen(req3, context=ctx, timeout=15) as r:
                отчёт = json.loads(r.read())
        except urllib.error.HTTPError as e:
            continue
        except Exception as e:
            continue

        sales = отчёт.get("sales", []) or []
        for s_item in sales:
            fuel = s_item.get("fuel", []) or []
            for f in fuel:
                сервис = f.get("service") or {}
                код = str(сервис.get("service_code", ""))
                имя = сервис.get("service_name", "")
                if код not in ("1", "2", "3"):
                    pay = (s_item.get("pay_type") or {})
                    ВСЕ_АНОМАЛИИ.append({
                        "sys": sys_id, "ст": ст, "см": н,
                        "код": код, "имя": имя,
                        "pay_id": pay.get("id"), "pay_name": pay.get("name"),
                        "литры": (f.get("release") or {}).get("volume"),
                        "сумма": (f.get("release") or {}).get("cost"),
                    })
        проверено += 1
        if проверено >= 250: break  # ограничение

print(f"\nНайдено аномалий: {len(ВСЕ_АНОМАЛИИ)}")
if not ВСЕ_АНОМАЛИИ:
    print("✓ Все service_code = 1/2/3 (АИ-92/АИ-95/ДТ)")
    sys.exit(0)

# Группировка по уникальному (код, имя)
по_коду = {}
for а in ВСЕ_АНОМАЛИИ:
    кл = f"{а['код']} → '{а['имя']}'"
    по_коду.setdefault(кл, []).append(а)

print("\nСводка по уникальным аномальным кодам:")
for кл, лст in по_коду.items():
    print(f"\n  [{кл}]  — {len(лст)} случаев")
    # Примеры
    for а in лст[:5]:
        print(f"    sys={а['sys']} ст={а['ст']} смена={а['см']}  pay={а['pay_id']}/{а['pay_name']}  объём={а['литры']} стоимость={а['сумма']}")
    if len(лст) > 5:
        print(f"    ... и ещё {len(лст)-5}")
