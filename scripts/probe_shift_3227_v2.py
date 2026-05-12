# -*- coding: utf-8 -*-
"""
v2: правильные поля транзакций STS + фильтр по shift локально.
quantity=литры, cost=рубли, price=цена/л, amount=кг.
"""
import sys, json, urllib.request, urllib.error, io
import win32com.client

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

BASE_PATH = r"D:\Users\magsp\GIG Base2"
USER = "Гайворонская Татьяна"
PWD = "12345"
STATION = 210
SHIFT = 3227
DF, DT_ = "2026-04-29", "2026-04-30"

def sts_login(url, login, pwd, sysid):
    body = {"login": login, "password": pwd,
            "user": {"id": "00000000-0000-0000-0000-000000000000", "name": "System"},
            "system_id": int(sysid)}
    req = urllib.request.Request(url.rstrip("/")+"/v2/login",
        data=json.dumps(body).encode("utf-8"), method="POST",
        headers={"Content-Type":"application/json"})
    raw = urllib.request.urlopen(req, timeout=20).read().decode("utf-8").strip().strip('"')
    if raw.startswith("{"):
        j = json.loads(raw)
        if "token" in j: return j["token"]
    return raw

def api_get(url, path, token, params=None):
    qs = "&".join(f"{k}={v}" for k,v in (params or {}).items())
    full = url.rstrip("/")+path + ("?"+qs if qs else "")
    req = urllib.request.Request(full, method="GET",
        headers={"Authorization": f"Bearer {token}", "Accept":"application/json"})
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            raw = resp.read().decode("utf-8")
    except urllib.error.HTTPError as he:
        return {"__error__": f"HTTP {he.code}", "__body__": he.read().decode("utf-8","replace")[:200]}
    if not raw.strip(): return None
    try: return json.loads(raw)
    except: return raw

def hr(t=""):
    print("\n" + "="*90)
    if t: print(t)
    print("="*90)

def main():
    com = win32com.client.Dispatch("V83.COMConnector")
    conn = com.Connect(f'File="{BASE_PATH}";Usr="{USER}";Pwd="{PWD}";')
    q = conn.NewObject("Запрос")
    q.Текст = "ВЫБРАТЬ Ключ, Значение ИЗ РегистрСведений.TL_Настройки"
    sel = q.Выполнить().Выбрать()
    s = {}
    while sel.Следующий(): s[str(sel.Ключ)] = str(sel.Значение)

    url, login, pwd = s["URLСервера"], s["Логин"], s["Пароль"]
    sysid = s.get(f"Станция_{STATION}_КодСистемы", s.get("КодСистемы","15"))
    print(f"sysid={sysid} login={login}")
    tok = sts_login(url, login, pwd, sysid)
    print(f"token len={len(tok)}")

    hr(f"A) /v1/report/shift_report shift={SHIFT}")
    r = api_get(url, "/v1/report/shift_report", tok,
                {"system": sysid, "station": STATION, "shift": SHIFT})
    # Сводка по топливам из PSM
    print("\nPSM total (счётчики ТРК):")
    total_psm_q = 0.0
    total_psm_c = 0.0
    if isinstance(r, dict) and "psm" in r and "total" in r["psm"]:
        for it in r["psm"]["total"]:
            name = it["service"]["service_name"]
            q = float(it["release"]["quantity"])
            cost = float(it["release"]["cost"])
            cost_rub = float(it["release"]["amount"])
            print(f"  {name:6} tank={it['tank']}  л={q:>10.2f}  кг={cost:>10.2f}  ₽={cost_rub:>12.2f}")
            total_psm_q += q
            total_psm_c += cost_rub
        print(f"  ИТОГО PSM:                  л={total_psm_q:>10.2f}              ₽={total_psm_c:>12.2f}")

    hr(f"B) /v1/transactions, фильтр shift={SHIFT} локально")
    r = api_get(url, "/v1/transactions", tok,
                {"system": sysid, "station": STATION,
                 "date_from": DF, "date_to": DT_})
    if not isinstance(r, list):
        print("неожиданный тип ответа"); return 1
    print(f"Всего получено: {len(r)} транзакций")
    mine = [t for t in r if isinstance(t,dict) and t.get("shift")==SHIFT]
    print(f"Из них shift={SHIFT}: {len(mine)}")

    # Pay types / fuel name
    fuel_map = {2:"АИ-92", 3:"АИ-95", 5:"ДТ", 6:"ДТ", 4:"АИ-?", 7:"ДТ?"}
    by_pay = {}
    by_pay_fuel = {}
    pay_names = {}
    for t in mine:
        pt = t.get("pay_type") or {}
        pid = pt.get("id"); pname = pt.get("name","?")
        pay_names[pid] = pname
        fuel = t.get("fuel_name") or fuel_map.get(t.get("fuel"),"?")
        try:
            ql = float(t.get("quantity") or 0)   # литры
            cr = float(t.get("cost") or 0)        # ₽
        except: continue
        a = by_pay.setdefault(pid, {"л":0.0, "₽":0.0, "cnt":0, "px":set()})
        a["л"] += ql; a["₽"] += cr; a["cnt"] += 1
        if t.get("price"): a["px"].add(float(t["price"]))
        b = by_pay_fuel.setdefault((pid, fuel), {"л":0.0, "₽":0.0, "cnt":0, "px":set()})
        b["л"] += ql; b["₽"] += cr; b["cnt"] += 1
        if t.get("price"): b["px"].add(float(t["price"]))

    print("\nИтоги транзакций по видам оплат (смена 3227):")
    print(f"  {'id':>3} {'тип оплаты':25} {'чеков':>6} {'литры':>10} {'₽':>14}")
    tot_l = tot_r = 0
    for pid in sorted(by_pay.keys(), key=lambda x:(x is None, x)):
        a = by_pay[pid]
        print(f"  {pid!s:>3} {pay_names.get(pid,'?'):25} {a['cnt']:>6} {a['л']:>10.2f} {a['₽']:>14.2f}")
        tot_l += a["л"]; tot_r += a["₽"]
    print(f"  {'':>3} {'ИТОГО':25} {len(mine):>6} {tot_l:>10.2f} {tot_r:>14.2f}")

    print("\nРазбивка по видам оплат и топливу:")
    print(f"  {'id':>3} {'тип оплаты':22} {'топливо':8} {'чеков':>6} {'литры':>10} {'₽':>12}  цены")
    for (pid, fuel) in sorted(by_pay_fuel.keys(), key=lambda x:(x[0] is None, x[0] or 0, x[1])):
        b = by_pay_fuel[(pid, fuel)]
        px = ",".join(f"{p:.2f}" for p in sorted(b["px"]))
        print(f"  {pid!s:>3} {pay_names.get(pid,'?'):22} {fuel:8} {b['cnt']:>6} {b['л']:>10.2f} {b['₽']:>12.2f}  {px}")

    # Сводка по топливу
    print("\nСводка по топливу (из транзакций):")
    by_fuel = {}
    for (pid, fuel), b in by_pay_fuel.items():
        f = by_fuel.setdefault(fuel, {"л":0.0, "₽":0.0})
        f["л"] += b["л"]; f["₽"] += b["₽"]
    for fuel, b in sorted(by_fuel.items()):
        print(f"  {fuel:8} л={b['л']:>10.2f}  ₽={b['₽']:>12.2f}")

if __name__ == "__main__":
    try: sys.exit(main() or 0)
    except Exception:
        import traceback; traceback.print_exc(); sys.exit(1)
