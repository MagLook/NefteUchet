# -*- coding: utf-8 -*-
"""
Диагностика STS API v2 — корректные эндпоинты и параметры.
Запуск: py -3 scripts/diag_v2.py
"""
import json
import urllib.request
import urllib.error
import urllib.parse
from datetime import date, timedelta

URL = "https://pos.autooplata.ru/tms"
LOGIN = "UserApi"
PWD = None  # пароль возьмём через COM из настроек

import win32com.client
def get_pwd():
    com = win32com.client.Dispatch("V83.COMConnector")
    conn = com.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')
    q = conn.NewObject("Запрос")
    q.Текст = '''ВЫБРАТЬ Значение ИЗ РегистрСведений.TL_Настройки ГДЕ Ключ = "Пароль"'''
    sel = q.Выполнить().Выбрать()
    if sel.Следующий():
        return str(sel.Значение)
    return None

def hr(t):
    print()
    print("=" * 78)
    print(t)
    print("=" * 78)

def sts_login(sysid):
    body = {
        "login": LOGIN,
        "password": PWD,
        "user": {"id": "00000000-0000-0000-0000-000000000000", "name": "System"},
    }
    if sysid:
        body["system_id"] = int(sysid)
    req = urllib.request.Request(
        URL + "/v2/login",
        data=json.dumps(body).encode("utf-8"),
        method="POST",
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=20) as resp:
        raw = resp.read().decode("utf-8")
    return raw.strip().strip('"')

def call(token, path, params=None):
    qs = "&".join(f"{k}={urllib.parse.quote(str(v))}" for k, v in (params or {}).items())
    full = URL + path + ("?" + qs if qs else "")
    req = urllib.request.Request(full, method="GET",
                                 headers={"Authorization": f"Bearer {token}"})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            raw = resp.read().decode("utf-8")
        if not raw.strip():
            return ("EMPTY", None, None)
        try:
            return ("OK", json.loads(raw), len(raw))
        except json.JSONDecodeError:
            return ("RAW", raw[:500], len(raw))
    except urllib.error.HTTPError as e:
        body = ""
        try: body = e.read().decode("utf-8", errors="replace")[:300]
        except Exception: pass
        return (f"HTTP {e.code}", body, None)
    except Exception as e:
        return (f"ERR {type(e).__name__}", str(e), None)

def show(label, status, payload, length=None):
    if isinstance(payload, list):
        cnt = len(payload)
        sample = json.dumps(payload[0], ensure_ascii=False)[:300] if cnt else "[]"
        print(f"  {label:50} {status}  len={cnt}  sample: {sample}")
    elif isinstance(payload, dict):
        keys = list(payload.keys())[:15]
        print(f"  {label:50} {status}  keys: {keys}")
        # Выведем целиком, если небольшой
        s = json.dumps(payload, ensure_ascii=False, indent=2)
        if len(s) < 1500:
            for line in s.split("\n"): print(f"    {line}")
    elif payload is None:
        print(f"  {label:50} {status}")
    else:
        print(f"  {label:50} {status}  {str(payload)[:200]}")

def main():
    global PWD
    PWD = get_pwd()
    print(f"Пароль из регистра: {'OK (' + str(len(PWD)) + ' симв.)' if PWD else 'НЕТ'}")

    # Перебор сетей: 65 (известная) + ещё 15 (упоминается в v5.2)
    for sysid in [65, 15]:
        hr(f"СЕТЬ system_id={sysid}")
        try:
            tok = sts_login(sysid)
        except urllib.error.HTTPError as e:
            print(f"  Авторизация: HTTP {e.code} — {e.read()[:200]}")
            continue
        except Exception as e:
            print(f"  Авторизация: ОШИБКА {e}")
            continue
        if not tok or len(tok) < 30:
            print(f"  Авторизация: пустой ответ ({tok!r})")
            continue
        print(f"  Авторизация: OK, токен {len(tok)} симв")

        # 1. /v1/points — список ОС (станций?)
        st, p, L = call(tok, "/v1/points", {"system": sysid})
        show("/v1/points (список ОС)", st, p, L)

        # 2. /v1/info — настройки ОС
        st, p, L = call(tok, "/v1/info", {"system": sysid})
        show("/v1/info (настройки ОС)", st, p, L)

        # 3. /v2/info — улучшенная версия
        st, p, L = call(tok, "/v2/info", {"system": sysid})
        show("/v2/info (настройки ОС2)", st, p, L)

        # 4. /v2/shifts с правильными датами
        today = date.today()
        d_beg = (today - timedelta(days=60)).isoformat()
        d_end = today.isoformat()
        st, p, L = call(tok, "/v2/shifts", {"system": sysid, "station": 0,
                                            "dt_beg": d_beg, "dt_end": d_end})
        show(f"/v2/shifts ({d_beg}..{d_end}, station=0)", st, p, L)

        # 4b. /v2/shifts без station
        st, p, L = call(tok, "/v2/shifts", {"system": sysid,
                                            "dt_beg": d_beg, "dt_end": d_end})
        show(f"/v2/shifts ({d_beg}..{d_end}, без station)", st, p, L)

        # 5. /v1/shifts со всеми вариантами параметров (для полноты)
        for params_set in [
            {"system": sysid, "station": 0},
            {"system": sysid},
        ]:
            st, p, L = call(tok, "/v1/shifts", params_set)
            show(f"/v1/shifts {params_set}", st, p, L)

        # 6. /v1/services
        st, p, L = call(tok, "/v1/services", {"system": sysid})
        show("/v1/services", st, p, L)

        # 7. /v1/tanks
        st, p, L = call(tok, "/v1/tanks", {"system": sysid})
        show("/v1/tanks", st, p, L)

if __name__ == "__main__":
    main()
