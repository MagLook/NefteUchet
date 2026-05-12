# -*- coding: utf-8 -*-
"""v3: добираем coupons / pos/coupons / pos/transactions для смены 3227"""
import sys, json, urllib.request, urllib.error, io
import win32com.client
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _tl_config import BASE_PATH, USER, PWD
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
STATION = 210; SHIFT = 3227
DF, DT_ = "2026-04-29", "2026-04-30"

def sts_login(url, login, pwd, sysid):
    body = {"login": login, "password": pwd,
            "user": {"id":"00000000-0000-0000-0000-000000000000","name":"System"},
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
            return resp.read().decode("utf-8")
    except urllib.error.HTTPError as he:
        return f"__HTTP_{he.code}__:{he.read().decode('utf-8','replace')[:200]}"

def hr(t=""):
    print("\n"+"="*90);
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
    tok = sts_login(url, login, pwd, sysid)
    print(f"token len={len(tok)}")

    fuel_name = {1:"АИ-80", 2:"АИ-92", 3:"АИ-95", 4:"АИ-98", 5:"ДТ", 6:"ДТ-евро"}

    def summarize(items, label, num_field_candidates=("shift","shift_number")):
        if not isinstance(items, list):
            print(f"  {label}: тип {type(items).__name__}")
            return
        print(f"  {label}: всего {len(items)}")
        mine = []
        for it in items:
            if not isinstance(it, dict): continue
            for f in num_field_candidates:
                if it.get(f) == SHIFT:
                    mine.append(it); break
        print(f"  из них shift={SHIFT}: {len(mine)}")
        if not mine: return
        # Пример
        print(f"  пример: {json.dumps(mine[0], ensure_ascii=False)[:250]}")
        # Сводка по pay_type x fuel
        agg = {}
        for it in mine:
            pt = it.get("pay_type") or {}
            pid = pt.get("id"); pname = pt.get("name","?")
            fu = it.get("fuel_name") or fuel_name.get(it.get("fuel"),"?")
            try:
                ql = float(it.get("quantity") or 0)
                cr = float(it.get("cost") or it.get("amount") or 0)
            except: continue
            k = (pid, pname, fu)
            a = agg.setdefault(k, {"л":0.0,"₽":0.0,"cnt":0})
            a["л"]+=ql; a["₽"]+=cr; a["cnt"]+=1
        for k in sorted(agg.keys(), key=lambda x:(x[0] is None, x[0] or 0, x[2])):
            a = agg[k]
            print(f"    pay={k[0]} {k[1]:22} {k[2]:8} cnt={a['cnt']:>4} л={a['л']:>10.2f} ₽={a['₽']:>12.2f}")

    for path in ("/v1/coupons", "/v1/coupons_manual", "/v1/pos/coupons", "/v1/pos/transactions"):
        hr(path)
        raw = api_get(url, path, tok, {"system": sysid, "station": STATION,
                                       "date_from": DF, "date_to": DT_})
        if isinstance(raw,str) and raw.startswith("__HTTP"):
            print(raw); continue
        try: items = json.loads(raw)
        except: items = raw
        if isinstance(items, list):
            summarize(items, path)
        elif isinstance(items, dict):
            # обёртка
            for key in ("data","items","coupons","transactions"):
                if key in items and isinstance(items[key], list):
                    summarize(items[key], f"{path}.{key}")
                    break
            else:
                print(f"  dict keys: {list(items.keys())[:10]}")
                print(json.dumps(items, ensure_ascii=False)[:600])
        else:
            print(str(items)[:400])

    # Финальный итог: сложим что нашли по shift=3227
    hr("СВЕРКА: что в STS API о смене 3227 (АЗС 210, 29.04.2026)")
    print("PSM (счётчики ТРК)         298 360,78 ₽ / 4 083,01 л   ← /v1/report/shift_report")
    print("Транзакции shift=3227      198 447,33 ₽ / 2 733,65 л   ← /v1/transactions (90 чеков)")
    print("Сменный отчёт Ledger       288 371,70 ₽ / 4 083,01 л   ← наша HTML")
    print("ОРП в БП ГИ0Г-000727       284 961,70 ₽ / 4 033,01 л   ← минус талоны 50 л / 3 410 ₽")

if __name__ == "__main__":
    try: sys.exit(main() or 0)
    except Exception:
        import traceback; traceback.print_exc(); sys.exit(1)
