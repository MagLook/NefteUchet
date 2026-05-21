import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _tl_config import BASE_PATH, USER, PWD
# -*- coding: utf-8 -*-
"""
Разведка АЗС 209 через STS API. Берём 3 свежие смены, смотрим что отдаёт сервер.
Цель: понять почему БП ГИГ показывает "Пустые данные смены".
Запуск: py -3.13-32 scripts/probe_shift_3878_azs209.py
"""
import json
import urllib.request
import urllib.error
import win32com.client

STATION = 209
SHIFTS_TO_PROBE = [3878, 3866, 3867]
SHIFT_DATE_FROM = "2026-05-01"
SHIFT_DATE_TO   = "2026-05-21"


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

    hr(f"3) /v1/info status TO станции {STATION}")
    r = api_get(url, "/v1/info", tok,
                {"system": sysid_station, "station": STATION})
    print(json.dumps(r, ensure_ascii=False, indent=2)[:2000])

    hr(f"4) /v1/shifts station={STATION} {SHIFT_DATE_FROM}..{SHIFT_DATE_TO}")
    r = api_get(url, "/v1/shifts", tok, {
        "system": sysid_station, "station": STATION,
        "date_from": SHIFT_DATE_FROM, "date_to": SHIFT_DATE_TO,
    })
    s = json.dumps(r, ensure_ascii=False)
    print(s[:3000])

    items = []
    if isinstance(r, list): items = r
    elif isinstance(r, dict) and "shifts" in r: items = r["shifts"]

    print(f"\nВсего смен в выдаче за период: {len(items)}")
    for it in items[:8]:
        if isinstance(it, dict):
            num = it.get("number") or it.get("shift") or it.get("shift_number")
            opened = it.get("opened") or it.get("open_date") or it.get("date_from") or "?"
            closed = it.get("closed") or it.get("close_date") or it.get("date_to") or "?"
            print(f"  смена {num}: opened={opened} closed={closed}")

    for SHIFT in SHIFTS_TO_PROBE:
        hr(f"5) shift_report station={STATION} shift={SHIFT}")
        for params in [
            {"system": sysid_station, "station": STATION, "shift": SHIFT},
            {"system": sysid_station, "station": STATION, "shift_number": SHIFT},
        ]:
            r = api_get(url, "/v1/report/shift_report", tok, params)
            s = json.dumps(r, ensure_ascii=False)
            print(f"params={params}")
            print(s[:1500])
            print()
            if "__error__" not in s and r:
                break


if __name__ == "__main__":
    main()
