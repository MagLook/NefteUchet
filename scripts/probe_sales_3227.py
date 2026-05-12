# -*- coding: utf-8 -*-
"""Проверка поля sales в /v1/report/shift_report для смены 3227 АЗС 210.
Цель: убедиться что талоны в sales имеют volume > 0."""
import sys, json, urllib.request, io
import win32com.client
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

BASE_PATH = r"D:\Users\magsp\GIG Base2"
USER = "Гайворонская Татьяна"; PWD = "12345"

com = win32com.client.Dispatch("V83.COMConnector")
conn = com.Connect(f'File="{BASE_PATH}";Usr="{USER}";Pwd="{PWD}";')
q = conn.NewObject("Запрос")
q.Текст = "ВЫБРАТЬ Ключ, Значение ИЗ РегистрСведений.TL_Настройки"
sel = q.Выполнить().Выбрать()
s = {}
while sel.Следующий(): s[str(sel.Ключ)] = str(sel.Значение)

url, login, pwd = s["URLСервера"], s["Логин"], s["Пароль"]
sysid = s.get("Станция_210_КодСистемы", s.get("КодСистемы","15"))

# login
body = {"login": login, "password": pwd,
        "user": {"id":"00000000-0000-0000-0000-000000000000","name":"System"},
        "system_id": int(sysid)}
req = urllib.request.Request(url.rstrip("/")+"/v2/login",
    data=json.dumps(body).encode("utf-8"), method="POST",
    headers={"Content-Type":"application/json"})
raw = urllib.request.urlopen(req, timeout=20).read().decode("utf-8").strip().strip('"')
tok = json.loads(raw)["token"] if raw.startswith("{") else raw

# shift_report
full = f"{url.rstrip('/')}/v1/report/shift_report?system={sysid}&station=210&shift=3227"
req = urllib.request.Request(full, method="GET",
    headers={"Authorization": f"Bearer {tok}", "Accept":"application/json"})
raw = urllib.request.urlopen(req, timeout=60).read().decode("utf-8")
data = json.loads(raw)

print(f"Корневые ключи: {list(data.keys())}\n")

sales = data.get("sales")
print(f"sales: тип={type(sales).__name__}", f"len={len(sales) if isinstance(sales,list) else '-'}")

if isinstance(sales, list) and sales:
    # Покажу первую запись целиком
    print("\nПервая запись sales:")
    print(json.dumps(sales[0], ensure_ascii=False, indent=2))

    # Сводка по pay_type x fuel: volume, cost
    by = {}
    for it in sales:
        if not isinstance(it, dict): continue
        pt = it.get("pay_type") or {}
        pname = pt.get("name","?")
        for fu in (it.get("fuel") or []):
            srv = fu.get("service") or {}
            sname = srv.get("service_name","?")
            rel = fu.get("release") or {}
            try:
                vol = float(rel.get("volume") or 0)
                cost = float(rel.get("cost") or 0)
            except: continue
            k = (pname, sname)
            a = by.setdefault(k, {"vol":0.0, "cost":0.0})
            a["vol"] += vol; a["cost"] += cost

    print("\nСводка sales по (pay_type, fuel):")
    print(f"  {'pay_type':22} {'fuel':10} {'volume л':>12} {'cost ₽':>14}")
    for k in sorted(by.keys()):
        print(f"  {k[0]:22} {k[1]:10} {by[k]['vol']:>12.2f} {by[k]['cost']:>14.2f}")

    # Особо: талоны
    print("\nТалоны отдельно:")
    found = False
    for it in sales:
        pt = (it.get("pay_type") or {}).get("name","")
        if "талон" in pt.lower() or "voucher" in pt.lower():
            found = True
            print(json.dumps(it, ensure_ascii=False, indent=2))
    if not found: print("  (записей с pay_type='Талоны' нет)")
