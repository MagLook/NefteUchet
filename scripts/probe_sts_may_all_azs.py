import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _tl_config import BASE_PATH, USER, PWD
# -*- coding: utf-8 -*-
"""
Проверка всех смен мая 2026 по всем АЗС через STS API.
Цель: понять масштаб HTTP 500 (серверные ошибки STS) и пустых ответов.
"""
import json
import time
import urllib.request
import urllib.error
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import win32com.client

PERIOD_FROM = "2026-05-01"
PERIOD_TO   = "2026-05-21"
STATIONS    = [5, 8, 208, 209, 210]
PAUSE_SEC   = 0.25
OUT_PATH    = Path(__file__).parent / "probe_sts_may_all_azs_result.json"


def sts_login(base_url, login, pwd, sysid):
    body = {
        "login": login, "password": pwd,
        "user": {"id": "00000000-0000-0000-0000-000000000000", "name": "System"},
        "system_id": int(sysid),
    }
    req = urllib.request.Request(
        base_url.rstrip("/") + "/v2/login",
        data=json.dumps(body).encode("utf-8"),
        method="POST", headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=20) as resp:
        raw = resp.read().decode("utf-8").strip().strip('"')
    if raw.startswith("{"):
        try:
            j = json.loads(raw)
            return j.get("token", raw)
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
        if not raw.strip(): return ("empty", None)
        try: return ("ok", json.loads(raw))
        except json.JSONDecodeError: return ("ok", raw)
    except urllib.error.HTTPError as he:
        body = ""
        try:
            body = he.read().decode("utf-8", errors="replace")[:200]
        except Exception:
            pass
        return (f"http_{he.code}", body)
    except Exception as e:
        return (f"err_{type(e).__name__}", str(e)[:200])


def get_shift_num(item):
    if isinstance(item, dict):
        return item.get("shift") or item.get("number") or item.get("shift_number")
    return None


def get_shift_dates(item):
    if isinstance(item, dict):
        return (item.get("dt_open") or item.get("opened") or "?",
                item.get("dt_close") or item.get("closed") or "?")
    return ("?", "?")


def main():
    # 1. Settings
    com = win32com.client.Dispatch("V83.COMConnector")
    conn = com.Connect(f'File="{BASE_PATH}";Usr="{USER}";Pwd="{PWD}";')
    q = conn.NewObject("Запрос")
    q.Текст = "ВЫБРАТЬ Ключ, Значение ИЗ РегистрСведений.TL_Настройки"
    sel = q.Выполнить().Выбрать()
    settings = {}
    while sel.Следующий():
        settings[str(sel.Ключ)] = str(sel.Значение)

    url   = settings.get("URLСервера", "https://pos.autooplata.ru/tms")
    login = settings.get("Логин", "")
    pwd   = settings.get("Пароль", "")
    sysid_def = settings.get("КодСистемы", "65")

    # System per station
    sysid_by_station = {}
    for st in STATIONS:
        sysid_by_station[st] = settings.get(f"Станция_{st}_КодСистемы", sysid_def)

    print("Настройки:")
    print(f"  URL = {url}")
    print(f"  Логин = {login}")
    print(f"  КодСистемы по умолчанию = {sysid_def}")
    for st, sid in sysid_by_station.items():
        print(f"  АЗС {st}: system_id = {sid}")

    # 2. Login per unique system_id
    tokens = {}
    for sid in set(sysid_by_station.values()):
        try:
            tok = sts_login(url, login, pwd, sid)
            tokens[sid] = tok
            print(f"  Авторизация sysid={sid}: OK (token len={len(tok)})")
        except Exception as e:
            tokens[sid] = None
            print(f"  Авторизация sysid={sid}: FAIL {e}")

    # 3. For each station — list shifts in May 2026
    result = {
        "_meta": {
            "запуск": datetime.now().isoformat(),
            "период": [PERIOD_FROM, PERIOD_TO],
            "settings_sysid": sysid_by_station,
        },
        "по_АЗС": {},
    }

    for station in STATIONS:
        sid = sysid_by_station[station]
        tok = tokens.get(sid)
        info_station = {
            "sysid": sid,
            "total_shifts_in_period": 0,
            "shifts": [],
            "status_summary": defaultdict(int),
        }
        if not tok:
            info_station["error"] = "Нет токена для этого sysid"
            result["по_АЗС"][station] = info_station
            continue

        print(f"\n=== АЗС {station} (sysid={sid}) ===")
        # List shifts
        status, r = api_get(url, "/v1/shifts", tok, {
            "system": sid, "station": station,
            "date_from": PERIOD_FROM, "date_to": PERIOD_TO,
        })
        if status != "ok":
            info_station["shifts_list_status"] = status
            info_station["shifts_list_error"] = r
            result["по_АЗС"][station] = info_station
            print(f"  /v1/shifts: {status}")
            continue

        items = r if isinstance(r, list) else (r.get("shifts", []) if isinstance(r, dict) else [])
        # Filter only May 2026 by dt_open
        may_items = []
        for it in items:
            opened, _ = get_shift_dates(it)
            if isinstance(opened, str) and opened.startswith("2026-05-"):
                may_items.append(it)

        info_station["total_shifts_in_period"] = len(may_items)
        print(f"  Смен в мае 2026: {len(may_items)}")

        # Test report for each May shift
        for it in may_items:
            num = get_shift_num(it)
            opened, closed = get_shift_dates(it)
            if num is None:
                continue
            time.sleep(PAUSE_SEC)
            status, body = api_get(url, "/v1/report/shift_report", tok, {
                "system": sid, "station": station, "shift": num,
            })
            row = {"shift": num, "opened": opened, "closed": closed,
                   "status": status}
            # Brief content check for OK
            if status == "ok":
                psm_total = []
                if isinstance(body, dict):
                    psm = body.get("psm") or {}
                    psm_total = psm.get("total") or []
                    row["releases_count"] = len(psm_total)
                    row["sum_amount"] = round(sum(
                        (s.get("release") or {}).get("amount", 0)
                        for s in psm_total
                    ), 2)
                else:
                    row["body_type"] = type(body).__name__
            elif status.startswith("http_500"):
                row["body_excerpt"] = (body or "")[:150]
            else:
                row["body_excerpt"] = (body or "")[:150]
            info_station["shifts"].append(row)
            info_station["status_summary"][status] += 1
            # Print one line per shift
            extra = ""
            if status == "ok":
                extra = f" | релизов={row.get('releases_count',0)} сумма={row.get('sum_amount',0)}"
            print(f"  смена {num:>5} {opened[:16]:16} → {status}{extra}")

        info_station["status_summary"] = dict(info_station["status_summary"])
        result["по_АЗС"][station] = info_station

    # 4. Summary
    print("\n" + "=" * 80)
    print("СВОДКА")
    print("=" * 80)
    for st, info in result["по_АЗС"].items():
        sumdict = info.get("status_summary", {})
        total = sum(sumdict.values())
        status_str = ", ".join(f"{k}={v}" for k, v in sumdict.items())
        print(f"  АЗС {st} (sysid={info.get('sysid')}): всего {total}  ·  {status_str}")

    OUT_PATH.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str),
                        encoding="utf-8")
    print(f"\nРезультат сохранён: {OUT_PATH.name}")


if __name__ == "__main__":
    main()
