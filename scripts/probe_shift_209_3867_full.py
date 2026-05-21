import sys, os, json, urllib.request, urllib.error
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _tl_config import BASE_PATH, USER, PWD
import win32com.client

STATION = 209
SHIFT   = 3867  # известная рабочая смена с продажами

com = win32com.client.Dispatch("V83.COMConnector")
conn = com.Connect(f'File="{BASE_PATH}";Usr="{USER}";Pwd="{PWD}";')
q = conn.NewObject("Запрос")
q.Текст = "ВЫБРАТЬ Ключ, Значение ИЗ РегистрСведений.TL_Настройки"
sel = q.Выполнить().Выбрать()
settings = {}
while sel.Следующий():
    settings[str(sel.Ключ)] = str(sel.Значение)
url, login, pwd = settings["URLСервера"], settings["Логин"], settings["Пароль"]
sid = settings.get(f"Станция_{STATION}_КодСистемы", settings["КодСистемы"])

def sts_login(base_url, login, pwd, sysid):
    body = {"login": login, "password": pwd,
            "user": {"id": "00000000-0000-0000-0000-000000000000", "name": "System"},
            "system_id": int(sysid)}
    req = urllib.request.Request(base_url + "/v2/login",
        data=json.dumps(body).encode("utf-8"), method="POST",
        headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=20) as resp:
        raw = resp.read().decode("utf-8").strip().strip('"')
    if raw.startswith("{"):
        try:
            j = json.loads(raw)
            return j.get("token", raw)
        except Exception: pass
    return raw

def api_get(base_url, path, token, params=None):
    qs = "&".join(f"{k}={v}" for k,v in (params or {}).items())
    full = base_url + path + ("?" + qs if qs else "")
    req = urllib.request.Request(full, method="GET",
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            raw = resp.read().decode("utf-8")
    except urllib.error.HTTPError as he:
        return {"__error__": he.code, "__body__": he.read().decode("utf-8", errors="replace")[:300]}
    if not raw.strip(): return None
    try: return json.loads(raw)
    except json.JSONDecodeError: return raw

tok = sts_login(url, login, pwd, sid)
print(f"sysid={sid}, токен получен (len={len(tok)})\n")

print("=== /v1/report/shift_report — полный ответ ===")
r = api_get(url, "/v1/report/shift_report", tok,
            {"system": sid, "station": STATION, "shift": SHIFT})

# Топ-level ключи
if isinstance(r, dict):
    print(f"Топ-level keys: {list(r.keys())}\n")
    for key, val in r.items():
        if isinstance(val, (dict, list)):
            t = "dict" if isinstance(val, dict) else f"list[{len(val)}]"
            print(f"--- {key}: {t}")
            if isinstance(val, dict):
                print(f"    sub-keys: {list(val.keys())}")
            elif val and isinstance(val[0], dict):
                print(f"    item keys: {list(val[0].keys())}")
        else:
            print(f"--- {key}: {val}")
    print()

# Полный JSON в файл
out = os.path.join(os.path.dirname(__file__), "probe_shift_209_3867_full.json")
with open(out, "w", encoding="utf-8") as f:
    json.dump(r, f, ensure_ascii=False, indent=2)
print(f"Полный JSON сохранён: {out}")

# Поиск ключевых слов «списан», «прочие», «losses», «discharge» в JSON
s = json.dumps(r, ensure_ascii=False).lower()
print("\nПоиск ключевых слов в JSON:")
for kw in ("списан", "прочие", "loss", "discharge", "writeoff", "decom", "технич", "talon", "талон", "прокачк", "техпот"):
    if kw in s:
        # Найти позицию и показать контекст
        idx = s.find(kw)
        ctx = s[max(0,idx-80):idx+120]
        print(f"  [{kw}] {ctx}")

# Также пробуем pos-report-tank_release
print("\n=== /v1/pos/report/tank_release ===")
r2 = api_get(url, "/v1/pos/report/tank_release", tok,
             {"system": sid, "station": STATION, "shift": SHIFT})
print(json.dumps(r2, ensure_ascii=False)[:1500] if isinstance(r2,(dict,list)) else r2)
