import sys, os, json, urllib.request, urllib.error, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _tl_config import BASE_PATH, USER, PWD
import win32com.client
from collections import defaultdict
from datetime import datetime
from pathlib import Path

# G1 (план luminous-humming-curry): собрать все уникальные pay_type.{id, name}
# по всем АЗС за май 2026. Дать факт-базу для эталона маппинга.

STATIONS = [1, 2, 3, 4, 5, 6, 8, 205, 207, 208, 209, 210]  # все 12 АЗС сети ГИГ
PERIOD_FROM = "2026-05-01"
PERIOD_TO = "2026-05-21"
PAUSE = 0.15

OUT = Path(__file__).parent / "probe_all_paytypes_result.json"


def sts_login(url, login, pwd, sid):
    body = {
        "login": login, "password": pwd,
        "user": {"id": "00000000-0000-0000-0000-000000000000", "name": "System"},
        "system_id": int(sid),
    }
    req = urllib.request.Request(
        url + "/v2/login",
        data=json.dumps(body).encode("utf-8"),
        method="POST",
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=20) as r:
        raw = r.read().decode("utf-8").strip().strip('"')
    try:
        return json.loads(raw).get("token", raw)
    except Exception:
        return raw


def api_get(url, path, tok, params):
    qs = "&".join(f"{k}={v}" for k, v in params.items())
    req = urllib.request.Request(
        url + path + "?" + qs,
        method="GET",
        headers={"Authorization": f"Bearer {tok}"},
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            raw = r.read().decode("utf-8")
        try:
            return json.loads(raw)
        except Exception:
            return raw
    except urllib.error.HTTPError as he:
        return {"__error__": he.code}


def main():
    com = win32com.client.Dispatch("V83.COMConnector")
    conn = com.Connect(f'File="{BASE_PATH}";Usr="{USER}";Pwd="{PWD}";')
    q = conn.NewObject("Запрос")
    q.Текст = "ВЫБРАТЬ Ключ, Значение ИЗ РегистрСведений.TL_Настройки"
    sel = q.Выполнить().Выбрать()
    s = {}
    while sel.Следующий():
        s[str(sel.Ключ)] = str(sel.Значение)
    url, login, pwd, sid_def = s["URLСервера"], s["Логин"], s["Пароль"], s.get("КодСистемы", "65")

    # Сети — 65 (основной) + значение из "Сети" (доп)
    sids = {sid_def}
    if s.get("Сети"):
        for x in s["Сети"].split(","):
            x = x.strip()
            if x:
                sids.add(x)
    # И per-station переопределения
    sids_by_station = {}
    for st in STATIONS:
        sids_by_station[st] = s.get(f"Станция_{st}_КодСистемы", sid_def)
        sids.add(sids_by_station[st])

    tokens = {}
    for sid in sids:
        try:
            tokens[sid] = sts_login(url, login, pwd, sid)
        except Exception as e:
            print(f"sysid={sid}: login FAIL {e}")
            tokens[sid] = None

    # Агрегат:  (pay_type_id, pay_type_name_нижний_регистр) → {count, sum_amt, sum_vol, stations, samples}
    paytypes = defaultdict(lambda: {
        "id_seen": set(),
        "names_seen": set(),
        "count": 0,
        "sum_amt": 0.0,
        "sum_vol": 0.0,
        "stations": set(),
        "fuels": set(),
        "first_seen": None,
    })

    for station in STATIONS:
        sid = sids_by_station[station]
        tok = tokens.get(sid)
        if not tok:
            print(f"АЗС {station} sysid={sid}: нет токена, пропуск")
            continue

        # Список смен
        r = api_get(url, "/v1/shifts", tok, {
            "system": sid, "station": station,
            "date_from": PERIOD_FROM, "date_to": PERIOD_TO,
        })
        if isinstance(r, dict) and "__error__" in r:
            print(f"АЗС {station}: shifts {r}")
            continue
        items = r if isinstance(r, list) else r.get("shifts", []) if isinstance(r, dict) else []
        may = [it for it in items if isinstance(it, dict) and str(it.get("dt_open", "")).startswith("2026-05-")]
        print(f"АЗС {station} (sysid={sid}): смен в мае {len(may)}")

        for it in may:
            num = it.get("shift") or it.get("number")
            if num is None:
                continue
            time.sleep(PAUSE)
            rep = api_get(url, "/v1/report/shift_report", tok, {
                "system": sid, "station": station, "shift": num,
            })
            if not isinstance(rep, dict) or "__error__" in rep:
                continue
            sales = rep.get("sales") or []
            for entry in sales:
                pt = entry.get("pay_type") or {}
                pt_id = pt.get("id")
                pt_name = pt.get("name", "")
                key = pt_name.lower()
                agg = paytypes[key]
                if pt_id is not None:
                    agg["id_seen"].add(pt_id)
                agg["names_seen"].add(pt_name)
                agg["count"] += 1
                agg["stations"].add(station)
                if agg["first_seen"] is None:
                    agg["first_seen"] = f"АЗС {station}/{num}"
                for f in entry.get("fuel") or []:
                    sname = (f.get("service") or {}).get("service_name", "")
                    agg["fuels"].add(sname)
                    rel = f.get("release") or {}
                    agg["sum_amt"] += float(rel.get("cost", 0) or 0)
                    agg["sum_vol"] += float(rel.get("volume", 0) or 0)

    # Печать результата
    print()
    print("=" * 100)
    print("УНИКАЛЬНЫЕ pay_type В STS ЗА МАЙ ПО ВСЕМ 12 АЗС")
    print("=" * 100)
    print(f"{'Имя (исходное)':<25} {'IDs':<10} {'Раз':>6} {'Σ литры':>10} {'Σ ₽':>14} {'АЗС':<15}")
    print("-" * 100)

    # Сортировка по сумме
    sorted_pts = sorted(paytypes.items(), key=lambda x: x[1]["sum_amt"], reverse=True)
    for key, agg in sorted_pts:
        names = " | ".join(sorted(agg["names_seen"]))
        ids = ",".join(str(i) for i in sorted(agg["id_seen"]))
        sts = ",".join(str(st) for st in sorted(agg["stations"]))
        print(f"{names[:24]:<25} {ids:<10} {agg['count']:>6} {agg['sum_vol']:>10.0f} {agg['sum_amt']:>14,.0f} {sts:<15}")

    # Сравнение с TL_МаппингОплат (из настроек)
    print()
    print("=" * 100)
    print("ПРОВЕРКА ПОКРЫТИЯ TL_МаппингОплат")
    print("=" * 100)
    # Загружаем текущий маппинг через COM
    try:
        qm = conn.NewObject("Запрос")
        qm.Текст = """
        ВЫБРАТЬ ОбразецИмени, КаналОплаты, Склад
        ИЗ РегистрСведений.TL_МаппингОплат
        """
        sm = qm.Выполнить().Выбрать()
        mapping = []
        while sm.Следующий():
            mapping.append({
                "pattern": str(sm.ОбразецИмени or ""),
                "channel": str(sm.КаналОплаты or ""),
                "warehouse": str(sm.Склад or ""),
            })
        print(f"Записей в TL_МаппингОплат (на локальном стенде): {len(mapping)}")
    except Exception as e:
        print(f"TL_МаппингОплат недоступен на локальном стенде: {e}")
        print("(на проде маппинг есть, но через COM не достать)")
        mapping = None

    if mapping is not None:
        not_covered = []
        for key, agg in paytypes.items():
            covered = False
            for m in mapping:
                if m["pattern"] and m["pattern"].lower() in key:
                    covered = True
                    break
            if not covered:
                not_covered.append(key)
        if not_covered:
            print(f"\n!! Не покрыто маппингом ({len(not_covered)}):")
            for n in not_covered:
                print(f"  {n!r}")
        else:
            print("\nВсё покрыто маппингом")

    # Сериализация
    out = {
        "_meta": {
            "запуск": datetime.now().isoformat(),
            "период": [PERIOD_FROM, PERIOD_TO],
            "станции": STATIONS,
        },
        "pay_types": {
            key: {
                "ids": sorted(list(agg["id_seen"])),
                "names_seen": sorted(list(agg["names_seen"])),
                "count": agg["count"],
                "sum_amt": round(agg["sum_amt"], 2),
                "sum_vol": round(agg["sum_vol"], 2),
                "stations": sorted(list(agg["stations"])),
                "fuels": sorted(list(agg["fuels"])),
                "first_seen": agg["first_seen"],
            }
            for key, agg in sorted_pts
        },
    }
    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(f"\nСохранено: {OUT.name}")


if __name__ == "__main__":
    main()
