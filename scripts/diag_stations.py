# -*- coding: utf-8 -*-
"""
Диагностика расхождения: что в TL_Настройки vs что отдаёт STS API.
Запуск: py -3-32 scripts/diag_stations.py
"""
import sys
import json
import urllib.request
import urllib.error
import win32com.client

BASE_PATH = r"D:\Users\magsp\GIG Base2"
USER = "Гайворонская Татьяна"
PWD = "12345"

def hr(t=""):
    print()
    print("=" * 78)
    if t: print(t)
    print("=" * 78)

def main():
    hr("1) COM к 1С — читаем регистр TL_Настройки")
    com = win32com.client.Dispatch("V83.COMConnector")
    conn = com.Connect(f'File="{BASE_PATH}";Usr="{USER}";Pwd="{PWD}";')

    # Прочитаем все настройки
    q = conn.NewObject("Запрос")
    q.Текст = "ВЫБРАТЬ Ключ, Значение ИЗ РегистрСведений.TL_Настройки УПОРЯДОЧИТЬ ПО Ключ"
    sel = q.Выполнить().Выбрать()

    settings = {}
    while sel.Следующий():
        settings[str(sel.Ключ)] = str(sel.Значение)

    # Базовые параметры
    base_keys = ["URLСервера", "Логин", "Пароль", "КодСистемы", "Организация", "ОсновнойСклад"]
    print("\n--- Базовые ---")
    for k in base_keys:
        v = settings.get(k, "(нет)")
        if k == "Пароль" and v != "(нет)":
            v = "*" * len(v)
        print(f"  {k:25} = {v}")

    # Станции
    print("\n--- Станции в регистре ---")
    station_codes = set()
    for k in settings.keys():
        if k.startswith("Станция_"):
            parts = k.split("_")
            if len(parts) >= 3 and parts[1].isdigit():
                station_codes.add(parts[1])
    if not station_codes:
        print("  (нет ни одной станции)")
    for code in sorted(station_codes, key=int):
        name = settings.get(f"Станция_{code}_Наименование", "")
        wh   = settings.get(f"Станция_{code}_Склад", "")
        sysid = settings.get(f"Станция_{code}_КодСистемы", settings.get("КодСистемы", "?"))
        close = settings.get(f"Станция_{code}_ВремяЗакрытия", "")
        print(f"  Код {code:>4}  сеть={sysid:>3}  склад={wh!r:50}  имя={name!r}  закр={close}")

    # Уникальные сети
    networks = set()
    for code in station_codes:
        sysid = settings.get(f"Станция_{code}_КодСистемы", settings.get("КодСистемы", "65"))
        try:
            networks.add(int(sysid))
        except ValueError:
            pass
    if not networks:
        try: networks.add(int(settings.get("КодСистемы", "65")))
        except ValueError: networks.add(65)
    print(f"\n--- Уникальных сетей (system_id): {sorted(networks)}")

    # 2) STS API — прямой запрос
    hr("2) Прямой запрос к STS API — что реально доступно")
    url = settings.get("URLСервера", "https://pos.autooplata.ru/tms")
    login = settings.get("Логин", "")
    pwd_api = settings.get("Пароль", "")
    print(f"URL: {url}")
    print(f"Логин: {login}")

    if not login or not pwd_api:
        print("Логин/Пароль пустые — авторизация невозможна")
        return

    for sysid in sorted(networks):
        print(f"\n--- Сеть system_id={sysid} ---")
        try:
            tok = login_sts(url, login, pwd_api, sysid)
        except Exception as e:
            print(f"  Авторизация: ОШИБКА {e}")
            continue
        if not tok:
            print("  Авторизация: пустой токен")
            continue
        print(f"  Токен: {tok[:30]}... (len={len(tok)})")

        # Тест A: /v1/shifts?station=0 без дат (как делает расширение)
        probe(url, "/v1/shifts", tok, {"system": sysid, "station": 0}, station_codes, "без дат")

        # Тест B: /v1/shifts со сдвинутыми датами
        from datetime import date, timedelta
        today = date.today()
        date_from = (today - timedelta(days=180)).isoformat()
        date_to = today.isoformat()
        for params in [
            {"system": sysid, "station": 0, "date_from": date_from, "date_to": date_to},
            {"system": sysid, "station": 0, "from": date_from, "to": date_to},
            {"system": sysid, "station": 0, "begin": date_from, "end": date_to},
            {"system": sysid, "station": 0, "date_begin": date_from, "date_end": date_to},
        ]:
            ok = probe(url, "/v1/shifts", tok, params, station_codes,
                       f"{params.get('date_from') or params.get('from') or params.get('begin') or params.get('date_begin')}..{params.get('date_to') or params.get('to') or params.get('end') or params.get('date_end')}")
            if ok: break

        # Тест C: список endpoint'ов через Swagger / OpenAPI
        for try_path in [
            "/v1/stations", "/v2/stations", "/v1/system/stations",
            "/v1/station", "/v2/station", "/v1/azs", "/v1/sites",
            "/v1/info", "/v1/me", "/v1/whoami", "/v1/system", "/v1/user",
            "/openapi.json", "/swagger.json", "/api-docs", "/docs",
        ]:
            try:
                r = api_get(url, try_path, tok, {"system": sysid} if "stations" in try_path or "azs" in try_path or "sites" in try_path else None)
                if r is not None:
                    if isinstance(r, str):
                        preview = r[:150].replace("\n", " ")
                    elif isinstance(r, list):
                        preview = f"list[{len(r)}], 1й: {json.dumps(r[0], ensure_ascii=False)[:120]}" if r else "list[0]"
                    else:
                        preview = f"dict keys: {list(r.keys())[:10]}"
                    print(f"  {try_path:30}  OK  {preview}")
            except urllib.error.HTTPError as he:
                if he.code not in (404, 405):
                    print(f"  {try_path:30}  HTTP {he.code}")
            except Exception as e:
                pass

def probe(url, path, tok, params, station_codes, label):
    try:
        r = api_get(url, path, tok, params)
    except Exception as e:
        print(f"  {path} {label}: ОШИБКА {e}")
        return False
    if not isinstance(r, list):
        if isinstance(r, dict) and "shifts" in r and isinstance(r["shifts"], list):
            r = r["shifts"]
        else:
            print(f"  {path} {label}: тип {type(r).__name__}, превью: {str(r)[:200]}")
            return False
    if not r:
        print(f"  {path} {label}: 0 смен")
        return False
    stations = {}
    for s in r:
        st = s.get("station") if isinstance(s, dict) else None
        if st is None: continue
        stations.setdefault(int(st), 0)
        stations[int(st)] += 1
    print(f"  {path} {label}: {len(r)} смен, уникальных станций {len(stations)}")
    for st_code, cnt in sorted(stations.items()):
        in_local = "OK" if str(st_code) in station_codes else "НЕТ В НАСТРОЙКАХ"
        print(f"      station={st_code:>4}  смен={cnt:>4}  {in_local}")
    return True

def login_sts(base_url, login, pwd, sysid):
    body = {
        "login": login,
        "password": pwd,
        "user": {"id": "00000000-0000-0000-0000-000000000000", "name": "System"},
    }
    if sysid:
        body["system_id"] = int(sysid)
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
        headers={"Authorization": f"Bearer {token}"},
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        raw = resp.read().decode("utf-8")
    if not raw.strip(): return None
    try: return json.loads(raw)
    except json.JSONDecodeError: return raw

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        traceback.print_exc()
        sys.exit(1)
