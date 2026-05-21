import sys, os, json, urllib.request, urllib.error, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _tl_config import BASE_PATH, USER, PWD
import win32com.client
from datetime import datetime

STATIONS = [5, 8, 208, 209, 210]
PERIOD_FROM = "2026-05-01"
PERIOD_TO   = "2026-05-21"
PAUSE = 0.2

def sts_login(url, login, pwd, sid):
    body = {"login": login, "password": pwd,
            "user": {"id": "00000000-0000-0000-0000-000000000000", "name": "System"},
            "system_id": int(sid)}
    req = urllib.request.Request(url + "/v2/login",
        data=json.dumps(body).encode("utf-8"), method="POST",
        headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=20) as r:
        raw = r.read().decode("utf-8").strip().strip('"')
    try:
        j = json.loads(raw)
        return j.get("token", raw)
    except Exception:
        return raw

def api_get(url, path, tok, params):
    qs = "&".join(f"{k}={v}" for k,v in params.items())
    req = urllib.request.Request(url + path + "?" + qs, method="GET",
        headers={"Authorization": f"Bearer {tok}"})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            raw = r.read().decode("utf-8")
        try: return json.loads(raw)
        except: return raw
    except urllib.error.HTTPError as he:
        return {"__error__": he.code}

com = win32com.client.Dispatch("V83.COMConnector")
conn = com.Connect(f'File="{BASE_PATH}";Usr="{USER}";Pwd="{PWD}";')
q = conn.NewObject("Запрос")
q.Текст = "ВЫБРАТЬ Ключ, Значение ИЗ РегистрСведений.TL_Настройки"
sel = q.Выполнить().Выбрать()
s = {}
while sel.Следующий(): s[str(sel.Ключ)] = str(sel.Значение)
url, login, pwd, sid_def = s["URLСервера"], s["Логин"], s["Пароль"], s.get("КодСистемы", "65")

tokens = {}
for st in STATIONS:
    sid = s.get(f"Станция_{st}_КодСистемы", sid_def)
    if sid not in tokens:
        tokens[sid] = sts_login(url, login, pwd, sid)

print(f"{'AZS':>4} {'Смена':>6}  {'Дата':<10}  {'АИ-92':>8}  {'АИ-95':>8}  {'ДТ':>8}  {'Σ литры':>9}  {'Σ ₽':>10}")
print("-" * 90)

found = []
total_amt = 0
for station in STATIONS:
    sid = s.get(f"Станция_{station}_КодСистемы", sid_def)
    tok = tokens[sid]
    # список смен мая
    r = api_get(url, "/v1/shifts", tok, {"system": sid, "station": station,
                                         "date_from": PERIOD_FROM, "date_to": PERIOD_TO})
    if isinstance(r, dict) and "__error__" in r:
        continue
    items = r if isinstance(r, list) else r.get("shifts", [])
    may = [it for it in items if isinstance(it, dict) and str(it.get("dt_open","")).startswith("2026-05-")]
    for it in may:
        num = it.get("shift") or it.get("number")
        opened = (it.get("dt_open") or "")[:10]
        time.sleep(PAUSE)
        rep = api_get(url, "/v1/report/shift_report", tok, {"system": sid, "station": station, "shift": num})
        if isinstance(rep, dict) and "__error__" in rep:
            continue
        if not isinstance(rep, dict):
            continue
        # Ищем pay_type "талон" в sales
        sales = rep.get("sales") or []
        for entry in sales:
            pt_name = (entry.get("pay_type") or {}).get("name", "").lower()
            if "талон" not in pt_name:
                continue
            fuels = entry.get("fuel") or []
            by_fuel = {"АИ-92": 0, "АИ-95": 0, "ДТ": 0, "ДРУГОЕ": 0}
            sub_amt = 0
            for f in fuels:
                sname = (f.get("service") or {}).get("service_name","").upper()
                rel = f.get("release") or {}
                vol = float(rel.get("volume", 0) or 0)
                cost = float(rel.get("cost", 0) or 0)
                if "92" in sname: by_fuel["АИ-92"] += vol
                elif "95" in sname: by_fuel["АИ-95"] += vol
                elif "ДТ" in sname: by_fuel["ДТ"] += vol
                else: by_fuel["ДРУГОЕ"] += vol
                sub_amt += cost
            sum_vol = sum(by_fuel.values())
            print(f"{station:>4} {num:>6}  {opened:<10}  {by_fuel['АИ-92']:>8.2f}  {by_fuel['АИ-95']:>8.2f}  {by_fuel['ДТ']:>8.2f}  {sum_vol:>9.2f}  {sub_amt:>10.2f}")
            found.append({"station": station, "shift": num, "date": opened,
                          "vol": sum_vol, "amt": sub_amt, "by_fuel": by_fuel})
            total_amt += sub_amt
            break  # одна строка талонов на смену

print("-" * 90)
print(f"ВСЕГО смен с талонами: {len(found)}")
print(f"СУММА выручки талонами за май: {total_amt:>10.2f} ₽")

# Сохраним
out = os.path.join(os.path.dirname(__file__), "probe_may_talony_result.json")
with open(out, "w", encoding="utf-8") as f:
    json.dump({"найдено": found, "итого_сумма": total_amt}, f, ensure_ascii=False, indent=2)
print(f"\nСохранено: {out}")

# Самая свежая большая смена с талонами — рекомендация для проверки
if found:
    cand = sorted(found, key=lambda x: (x["date"], x["vol"]), reverse=True)
    print("\nСвежие смены для проверки перепроведения:")
    for c in cand[:5]:
        print(f"  АЗС {c['station']:>3} · смена {c['shift']} · {c['date']} · "
              f"{c['vol']:.2f} л · {c['amt']:.2f} ₽")
