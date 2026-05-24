# -*- coding: utf-8 -*-
"""Посчитать сколько смен за период без розницы (только корпоратив)
и сколько с розницей. По всем АЗС.
"""
import os, sys, json, ssl, urllib.request, urllib.error
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _tl_config import connect

conn = connect()


def s(v):
    if v is None: return ""
    try: return str(v)
    except Exception: return ""


def setting(key, default=""):
    q = conn.NewObject("Запрос")
    q.УстановитьПараметр("К", key)
    q.Текст = """ВЫБРАТЬ ПЕРВЫЕ 1 Рег.Значение КАК З
                 ИЗ РегистрСведений.TL_Настройки КАК Рег
                 ГДЕ Рег.Ключ = &К"""
    в = q.Выполнить().Выбрать()
    return s(в.З) if в.Следующий() else default


URL  = setting("URLСервера", "https://pos.autooplata.ru/tms")
USER = setting("Логин", "")
PWD  = setting("Пароль", "")

ctx = ssl._create_unverified_context()


def auth(sys_id):
    """STS v2/login с system_id"""
    body = {
        "login": USER, "password": PWD,
        "user": {"id": "00000000-0000-0000-0000-000000000000", "name": "System"},
        "system_id": sys_id,
    }
    req = urllib.request.Request(
        f"{URL}/v2/login",
        data=json.dumps(body).encode(),
        headers={"Content-Type": "application/json"},
        method="POST")
    with urllib.request.urlopen(req, context=ctx, timeout=15) as r:
        tok = r.read().decode().strip().strip('"')
    return tok

ПЕР_FROM = "2026-05-15"
ПЕР_TO   = "2026-05-22"

КАНАЛЫ_RETAIL = {"наличные", "наличн", "сбербанк"}  # имена pay_type из STS — упрощённо
КАНАЛЫ_RETAIL_CARD = {"мобил", "сбер"}              # карты которые БП считает retail_card

ИТОГ = {
    "с_розницей": 0,
    "без_розницы_с_корп": 0,
    "пустая": 0,
    "битая": 0,
}
БЕЗ_РОЗНИЦЫ = []  # список (sys, ст, см, что_было)

for sys_id in [15, 65]:
    try:
        token = auth(sys_id)
        print(f"OK авторизация sys={sys_id}")
    except Exception as e:
        print(f"  ОШИБКА авторизации sys={sys_id}: {e}")
        continue
    req2 = urllib.request.Request(
        f"{URL}/v1/shifts/?system={sys_id}&dt_from={ПЕР_FROM}&dt_to={ПЕР_TO}",
        headers={"Authorization": f"Bearer {token}"})
    try:
        with urllib.request.urlopen(req2, context=ctx, timeout=60) as r:
            смены = json.loads(r.read())
    except Exception as e:
        print(f"system={sys_id}: ошибка {e}")
        continue

    if not isinstance(смены, list):
        continue
    print(f"\nsystem={sys_id}: всего смен в API {len(смены)}")

    проверено = 0
    for см in смены:
        ст = см.get("station")
        н  = см.get("shift")
        if not ст or not н:
            continue

        req3 = urllib.request.Request(
            f"{URL}/v1/report/shift_report?system={sys_id}&station={ст}&shift={н}",
            headers={"Authorization": f"Bearer {token}"})
        try:
            with urllib.request.urlopen(req3, context=ctx, timeout=15) as r:
                отчёт = json.loads(r.read())
        except urllib.error.HTTPError as e:
            ИТОГ["битая"] += 1
            continue
        except Exception:
            continue

        sales = отчёт.get("sales") or []
        if not sales:
            ИТОГ["пустая"] += 1
            continue

        есть_наличные = False
        есть_карта = False
        корп_каналы = {}  # имя → литры

        for item in sales:
            pay = (item.get("pay_type") or {})
            имя = (pay.get("name") or "").strip()
            имя_низ = имя.lower()
            fuel = item.get("fuel") or []
            σ_лит = 0
            for f in fuel:
                σ_лит += (f.get("release") or {}).get("volume") or 0

            # классифицируем
            if имя_низ in ("наличные",) or "наличн" in имя_низ:
                if σ_лит > 0:
                    есть_наличные = True
            elif "сбер" in имя_низ or "мобил" in имя_низ:
                if σ_лит > 0:
                    есть_карта = True
            else:
                # корпоратив (талон, балтоп, viacard, инфорком, ведомость, прочие)
                if σ_лит > 0:
                    корп_каналы[имя] = корп_каналы.get(имя, 0) + σ_лит

        есть_розница = есть_наличные or есть_карта

        if есть_розница:
            ИТОГ["с_розницей"] += 1
        elif корп_каналы:
            ИТОГ["без_розницы_с_корп"] += 1
            БЕЗ_РОЗНИЦЫ.append({
                "sys": sys_id, "ст": ст, "см": н,
                "корп": корп_каналы,
            })
        else:
            ИТОГ["пустая"] += 1

        проверено += 1

print("\n" + "=" * 80)
print(f"Период: {ПЕР_FROM} — {ПЕР_TO}, все АЗС (sys 15 + 65)")
print("=" * 80)
print(f"  С розницей:                {ИТОГ['с_розницей']:>5}  ОРП создаётся, всё штатно")
print(f"  Без розницы, есть корп:    {ИТОГ['без_розницы_с_корп']:>5}  ОРП не нужен (только перемещения)")
print(f"  Совсем пустая (0 sales):   {ИТОГ['пустая']:>5}  ничего не отпущено")
print(f"  Битая (HTTP 500):          {ИТОГ['битая']:>5}  серверная ошибка STS")
всего = sum(ИТОГ.values())
print(f"  ВСЕГО:                     {всего:>5}")

if БЕЗ_РОЗНИЦЫ:
    print(f"\nПримеры смен без розницы (первые 15):")
    for r in БЕЗ_РОЗНИЦЫ[:15]:
        корп_сводка = ", ".join(f"{и}={к:.1f}л" for и, к in r["корп"].items())
        print(f"  sys={r['sys']} ст={r['ст']} см={r['см']}  →  {корп_сводка}")
    if len(БЕЗ_РОЗНИЦЫ) > 15:
        print(f"  ... и ещё {len(БЕЗ_РОЗНИЦЫ)-15}")
