# -*- coding: utf-8 -*-
"""
Разведка смены 3227 АЗС 210 за 29.04.2026 через STS API.
Запуск: py -3-32 scripts/probe_shift_3227.py
"""
import sys
import json
import urllib.request
import urllib.error
from datetime import date
import win32com.client

BASE_PATH = r"D:\Users\magsp\GIG Base2"
USER = "Гайворонская Татьяна"
PWD = "12345"

STATION = 210
SHIFT = 3227
SHIFT_DATE_FROM = "2026-04-29"
SHIFT_DATE_TO   = "2026-04-30"

def hr(t=""):
    print()
    print("=" * 90)
    if t: print(t)
    print("=" * 90)

def sts_login(base_url, login, pwd, sysid):
    body = {
        "login": login,
        "password": pwd,
        "user": {"id": "00000000-0000-0000-0000-000000000000", "name": "System"},
        "system_id": int(sysid),
    }
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        base_url.rstrip("/") + "/v2/login",
        data=data, method="POST",
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=20) as resp:
        raw = resp.read().decode("utf-8")
    raw = raw.strip().strip('"')
    if raw.startswith("{"):
        try:
            j = json.loads(raw)
            if "token" in j: return j["token"]
        except Exception:
            pass
    return raw

def api_get(base_url, path, token, params=None):
    qs = "&".join(f"{k}={v}" for k, v in (params or {}).items())
    full = base_url.rstrip("/") + path + ("?" + qs if qs else "")
    req = urllib.request.Request(
        full, method="GET",
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            raw = resp.read().decode("utf-8")
    except urllib.error.HTTPError as he:
        body = he.read().decode("utf-8", errors="replace")[:300]
        return {"__error__": f"HTTP {he.code}", "__body__": body}
    if not raw.strip(): return None
    try: return json.loads(raw)
    except json.JSONDecodeError: return raw

def main():
    hr("1) Креды из локальной TL_Настройки")
    com = win32com.client.Dispatch("V83.COMConnector")
    conn = com.Connect(f'File="{BASE_PATH}";Usr="{USER}";Pwd="{PWD}";')

    q = conn.NewObject("Запрос")
    q.Текст = "ВЫБРАТЬ Ключ, Значение ИЗ РегистрСведений.TL_Настройки УПОРЯДОЧИТЬ ПО Ключ"
    sel = q.Выполнить().Выбрать()
    settings = {}
    while sel.Следующий():
        settings[str(sel.Ключ)] = str(sel.Значение)

    url   = settings.get("URLСервера", "https://pos.autooplata.ru/tms")
    login = settings.get("Логин", "")
    pwd   = settings.get("Пароль", "")
    sysid = settings.get("КодСистемы", "15")
    # Если для станции 210 свой код системы — взять его
    sysid_station = settings.get(f"Станция_{STATION}_КодСистемы", sysid)
    print(f"URL          = {url}")
    print(f"Логин        = {login}")
    print(f"Пароль       = {'*' * len(pwd) if pwd else '(пусто)'}")
    print(f"КодСистемы   = {sysid}")
    print(f"АЗС {STATION} sysid = {sysid_station}")

    if not login or not pwd:
        print("ОШИБКА: логин/пароль пустые в TL_Настройки")
        return 2

    hr(f"2) Авторизация STS (system_id={sysid_station})")
    try:
        tok = sts_login(url, login, pwd, sysid_station)
    except Exception as e:
        print(f"Авторизация: ОШИБКА {e}")
        return 3
    if not tok:
        print("Пустой токен")
        return 3
    print(f"Токен: {tok[:40]}... (len={len(tok)})")

    hr(f"3) /v1/shifts station={STATION} {SHIFT_DATE_FROM}..{SHIFT_DATE_TO}")
    params = {
        "system": sysid_station,
        "station": STATION,
        "date_from": SHIFT_DATE_FROM,
        "date_to": SHIFT_DATE_TO,
    }
    r = api_get(url, "/v1/shifts", tok, params)
    print(json.dumps(r, ensure_ascii=False, indent=2)[:3000])

    # Поищем шину shift_id для нашей смены 3227
    shift_id = None
    shift_obj = None
    if isinstance(r, list):
        items = r
    elif isinstance(r, dict) and "shifts" in r:
        items = r["shifts"]
    else:
        items = []
    for s in items:
        if isinstance(s, dict):
            num = s.get("number") or s.get("shift") or s.get("shift_number") or s.get("num")
            if str(num) == str(SHIFT):
                shift_obj = s
                shift_id = s.get("id") or s.get("shift_id") or s.get("uid")
                break
    if shift_obj:
        hr(f"4) Найдена смена {SHIFT} (id={shift_id})")
        print(json.dumps(shift_obj, ensure_ascii=False, indent=2))
    else:
        print(f"\nСмена {SHIFT} не найдена в выдаче /v1/shifts")

    hr(f"5) /v1/report/shift_report для смены {SHIFT}")
    for params in [
        {"system": sysid_station, "station": STATION, "shift": SHIFT},
        {"system": sysid_station, "station": STATION, "shift_number": SHIFT},
        {"system": sysid_station, "station": STATION, "number": SHIFT},
        {"system": sysid_station, "station": STATION, "shift_id": shift_id or SHIFT},
        {"system": sysid_station, "station": STATION,
         "date_from": SHIFT_DATE_FROM, "date_to": SHIFT_DATE_TO},
    ]:
        print(f"\nparams={params}")
        r = api_get(url, "/v1/report/shift_report", tok, params)
        s = json.dumps(r, ensure_ascii=False)
        if "__error__" in s and len(s) < 500:
            print(s)
            continue
        print(s[:2500])
        if "__error__" not in s:
            break

    hr(f"6) /v1/pos/report/sales station={STATION} 29.04..30.04")
    r = api_get(url, "/v1/pos/report/sales", tok,
                {"system": sysid_station, "station": STATION,
                 "date_from": SHIFT_DATE_FROM, "date_to": SHIFT_DATE_TO})
    print(json.dumps(r, ensure_ascii=False, indent=2)[:3000])

    hr(f"7) /v1/transactions station={STATION} 29.04..30.04 (превью)")
    r = api_get(url, "/v1/transactions", tok,
                {"system": sysid_station, "station": STATION,
                 "date_from": SHIFT_DATE_FROM, "date_to": SHIFT_DATE_TO})
    txt = json.dumps(r, ensure_ascii=False)
    print(f"Длина ответа: {len(txt)} симв")
    if isinstance(r, list):
        print(f"Транзакций: {len(r)}")
        if r:
            print("\nПервая транзакция (полный JSON):")
            print(json.dumps(r[0], ensure_ascii=False, indent=2))
            # Сводка по типам оплат
            by_pay = {}
            for tr in r:
                if not isinstance(tr, dict): continue
                pay = tr.get("payment_type") or tr.get("pay_type") or tr.get("paymentType") or tr.get("type") or "?"
                fuel = tr.get("product") or tr.get("nomenclature") or tr.get("fuel") or tr.get("name") or "?"
                amount = float(tr.get("amount") or tr.get("sum") or 0)
                volume = float(tr.get("volume") or tr.get("liters") or tr.get("qty") or 0)
                price  = float(tr.get("price") or 0)
                key = (str(pay), str(fuel))
                acc = by_pay.setdefault(key, {"cnt": 0, "vol": 0.0, "amt": 0.0, "px": set()})
                acc["cnt"] += 1
                acc["vol"] += volume
                acc["amt"] += amount
                if price: acc["px"].add(round(price, 2))
            print("\nСводка транзакций по (вид_оплаты, топливо):")
            for k in sorted(by_pay.keys()):
                a = by_pay[k]
                px = ",".join(str(p) for p in sorted(a["px"]))
                print(f"  {k[0]:20} {k[1]:15} cnt={a['cnt']:>5}  vol={a['vol']:>12.3f}  amt={a['amt']:>14.2f}  px={px}")
    else:
        print(txt[:2000])

if __name__ == "__main__":
    try:
        sys.exit(main() or 0)
    except Exception:
        import traceback
        traceback.print_exc()
        sys.exit(1)
