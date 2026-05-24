# -*- coding: utf-8 -*-
"""Перебор всех смен за период через STS API: для каждой смены вытаскиваем
sales[] и считаем по pay_type. Цель: найти смены где в нашей логике
получается «нет retail», и посмотреть какие там реально pay_type.
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
                 ИЗ РегистрСведений.TL_Настройки КАК Рег ГДЕ Рег.Ключ = &К"""
    в = q.Выполнить().Выбрать()
    return s(в.З) if в.Следующий() else default


URL  = setting("URLСервера")
USER = setting("Логин")
PWD  = setting("Пароль")
ctx = ssl._create_unverified_context()


def auth(sys_id):
    body = {"login": USER, "password": PWD,
            "user": {"id": "00000000-0000-0000-0000-000000000000", "name": "System"},
            "system_id": sys_id}
    req = urllib.request.Request(f"{URL}/v2/login",
        data=json.dumps(body).encode(),
        headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, context=ctx, timeout=15) as r:
        return r.read().decode().strip().strip('"')


# Маппинг ИмяОплаты → КаналОплаты из TL_МаппингОплат БД
q = conn.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ ОбразецИмени, КаналОплаты, Склад
ИЗ РегистрСведений.TL_МаппингОплат
"""
маппинг = []  # список (паттерн_нрег, канал, склад)
в = q.Выполнить().Выбрать()
while в.Следующий():
    маппинг.append((s(в.ОбразецИмени).lower(), s(в.КаналОплаты), s(в.Склад)))
print(f"Маппинг загружен: {len(маппинг)} записей")
for имя, к, скл in маппинг:
    print(f"  '{имя}' → канал='{к}'  склад='{скл}'")


def найти_канал(имя):
    """Имитация TL_Настройки.НайтиЗаписьМаппинга — поиск по подстроке (НРег)"""
    нр = имя.lower()
    for патт, кан, скл in маппинг:
        if not патт: continue
        if патт in нр:
            return кан, скл
    return None, None


# Свежие смены за неделю
print(f"\nЗапрашиваю смены за 2026-05-15 — 2026-05-22 по обеим сетям...")
все_смены_без_розницы = []
все_смены_с_розницей = 0
смен_проверено = 0
ошибки_api = 0

ПЕР_FROM = "2026-05-15T00:00:00"
ПЕР_TO   = "2026-05-22T23:59:59"

for sys_id in [15, 65]:
    try:
        token = auth(sys_id)
    except Exception as e:
        print(f"  sys={sys_id}: ошибка авторизации {e}")
        continue

    # Список смен
    req = urllib.request.Request(
        f"{URL}/v1/shifts/?system={sys_id}&dt_from={ПЕР_FROM}&dt_to={ПЕР_TO}",
        headers={"Authorization": f"Bearer {token}"})
    try:
        with urllib.request.urlopen(req, context=ctx, timeout=60) as r:
            см_список = json.loads(r.read())
    except Exception as e:
        print(f"  sys={sys_id}: ошибка получения списка смен: {e}")
        continue
    if not isinstance(см_список, list):
        continue

    # Фильтр по периоду
    от_DT = "2026-05-15"
    до_DT = "2026-05-22"
    свежие = [см for см in см_список
              if см.get("dt_close") and от_DT <= см["dt_close"][:10] <= до_DT]
    print(f"\nsys={sys_id}: всего смен {len(см_список)}, в период {len(свежие)}")

    for см in свежие:
        ст = см.get("station")
        н  = см.get("shift")
        if not ст or not н: continue

        req2 = urllib.request.Request(
            f"{URL}/v1/report/shift_report?system={sys_id}&station={ст}&shift={н}",
            headers={"Authorization": f"Bearer {token}"})
        try:
            with urllib.request.urlopen(req2, context=ctx, timeout=15) as r:
                отчёт = json.loads(r.read())
        except urllib.error.HTTPError:
            ошибки_api += 1
            continue
        except Exception:
            continue

        sales = отчёт.get("sales") or []
        смен_проверено += 1

        # Имитация логики ОбработатьСмену: проходим по sales, маппим pay_type → канал
        каналы = {}  # канал → суммарные литры
        неизвестные = []  # pay_type без маппинга
        for item in sales:
            pay = (item.get("pay_type") or {})
            имя = (pay.get("name") or "").strip()
            кан, _скл = найти_канал(имя)
            σ_лит = 0
            for f in item.get("fuel") or []:
                σ_лит += (f.get("release") or {}).get("volume") or 0
            if σ_лит <= 0: continue
            if кан:
                каналы[кан] = каналы.get(кан, 0) + σ_лит
            else:
                неизвестные.append((имя, σ_лит))

        retail = каналы.get("retail_cash", 0) + каналы.get("retail_card", 0)

        if retail > 0:
            все_смены_с_розницей += 1
        else:
            все_смены_без_розницы.append({
                "sys": sys_id, "ст": ст, "см": н,
                "dt_close": см.get("dt_close"),
                "каналы": каналы,
                "неизвестные": неизвестные,
                "всего_sales": len(sales),
            })

print(f"\n{'='*90}")
print(f"Проверено смен: {смен_проверено}  (ошибок API: {ошибки_api})")
print(f"  С розницей:                {все_смены_с_розницей}")
print(f"  Без розницы:               {len(все_смены_без_розницы)}")
print(f"{'='*90}")

if все_смены_без_розницы:
    print(f"\nСМЕНЫ БЕЗ РОЗНИЦЫ — детально (первые 20):")
    for r in все_смены_без_розницы[:20]:
        print(f"\nsys={r['sys']} АЗС={r['ст']} смена={r['см']} close={r['dt_close']} всего sales={r['всего_sales']}")
        if r["каналы"]:
            print(f"  Каналы по маппингу:")
            for к, л in r["каналы"].items():
                print(f"    {к}: {л:.2f} л")
        if r["неизвестные"]:
            print(f"  pay_type БЕЗ маппинга (потеряны):")
            for имя, л in r["неизвестные"]:
                print(f"    '{имя}': {л:.2f} л")
        if not r["каналы"] and not r["неизвестные"]:
            print(f"  ⚠ ВООБЩЕ ПУСТО — sales пуст или volume=0 во всех")
