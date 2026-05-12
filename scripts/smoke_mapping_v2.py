# -*- coding: utf-8 -*-
"""Smoke v2: проверка регистра TL_МаппингОплат и блокера.
1. Залить дефолтные маппинги.
2. Прогнать ОбработатьСмену для 3227 — должно сработать как раньше.
3. Симулировать неизвестный pay_type → должна быть ошибка-блокер.
"""
import sys, io, json, urllib.request, datetime, win32com.client
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

BASE_PATH = r"D:\Users\magsp\GIG Base2"
USER = "Гайворонская Татьяна"; PWD = "12345"
STATION = 210; SHIFT = 3227

com = win32com.client.Dispatch("V83.COMConnector")
conn = com.Connect(f'File="{BASE_PATH}";Usr="{USER}";Pwd="{PWD}";')

def hr(t):
    print("\n" + "="*90); print(t); print("="*90)

# 1) Залить дефолтные маппинги
hr("1) Заливка TL_МаппингОплат")
defaults = [
    ("наличн",    "retail_cash", 10),
    ("сбербанк",  "retail_card", 20),
    ("безнал",    "retail_card", 30),
    ("viacard",   "cards",       50),
    ("балтоп",    "cards",       51),
    ("baltop",    "cards",       52),
    ("корпоратив","cards",       53),
    ("топливн",   "cards",       54),
    ("агора",     "cards",       55),
    ("яндекс",    "online",      70),
    ("мобил",     "online",      71),
    ("benzuber",  "online",      72),
    ("ведомост",  "ledger",      80),
    ("талон",     "voucher",     90),
    ("купон",     "",            900),
    ("прокач",    "",            901),
]
rs = conn.РегистрыСведений.TL_МаппингОплат
for образец, канал, приор in defaults:
    мн = rs.СоздатьМенеджерЗаписи()
    мн.ОбразецИмени = образец
    мн.Прочитать()
    if мн.Выбран():
        print(f"  УЖЕ: {образец:14} → {канал}")
        continue
    мн.ОбразецИмени = образец
    мн.КаналОплаты = канал
    мн.Приоритет = приор
    мн.Записать()
    print(f"  ЗАПИСАН: {образец:14} → {канал or '(игнор)'}  (P={приор})")

# 2) Удалить старый smoke-ОРП по смене 3227 (чтобы не упереться в "уже загружено")
hr("2) Очистка статусов TL_СтатусыЗагрузки")
sysid = "15"
key = f"TL|СМЕНА|{sysid}|{STATION}|{SHIFT}"
rs2 = conn.РегистрыСведений.TL_СтатусыЗагрузки
for k in (key, key+"|ПКО"):
    мн = rs2.СоздатьМенеджерЗаписи()
    мн.КлючЗагрузки = k
    мн.Прочитать()
    if мн.Выбран():
        мн.Удалить()
        print(f"  Удалён: {k}")

# Также пометить на удаление старые smoke-документы
print("  Пометка старых TL-документов на удаление:")
q = conn.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ Док.Ссылка ИЗ Документ.ОтчетОРозничныхПродажах КАК Док
ГДЕ Док.Комментарий ПОДОБНО &Шаб
"""
q.УстановитьПараметр("Шаб", f"TL|СМЕНА|{sysid}|{STATION}|{SHIFT}%")
sel = q.Выполнить().Выбрать()
while sel.Следующий():
    obj = sel.Ссылка.ПолучитьОбъект()
    obj.УстановитьПометкуУдаления(True)
    print(f"    помечен: {sel.Ссылка}")
for имя_дока in ("ПриходныйКассовыйОрдер", "ПеремещениеТоваров"):
    q.Текст = f"ВЫБРАТЬ Док.Ссылка ИЗ Документ.{имя_дока} КАК Док ГДЕ Док.Комментарий ПОДОБНО &Шаб"
    sel = q.Выполнить().Выбрать()
    while sel.Следующий():
        obj = sel.Ссылка.ПолучитьОбъект()
        obj.УстановитьПометкуУдаления(True)
        print(f"    помечен: {sel.Ссылка}")

# 3) Получить sales и прогнать ОбработатьСмену
hr("3) Прогон ОбработатьСмену на реальных данных смены 3227")
q = conn.NewObject("Запрос")
q.Текст = "ВЫБРАТЬ Ключ, Значение ИЗ РегистрСведений.TL_Настройки"
sel = q.Выполнить().Выбрать()
s = {}
while sel.Следующий(): s[str(sel.Ключ)] = str(sel.Значение)
url, login, pwd = s["URLСервера"], s["Логин"], s["Пароль"]
sysid = s.get(f"Станция_{STATION}_КодСистемы", s.get("КодСистемы","15"))

body = {"login": login, "password": pwd,
        "user": {"id":"00000000-0000-0000-0000-000000000000","name":"System"},
        "system_id": int(sysid)}
req = urllib.request.Request(url.rstrip("/")+"/v2/login",
    data=json.dumps(body).encode("utf-8"), method="POST",
    headers={"Content-Type":"application/json"})
raw = urllib.request.urlopen(req, timeout=20).read().decode("utf-8").strip().strip('"')
tok = json.loads(raw)["token"] if raw.startswith("{") else raw

full = f"{url.rstrip('/')}/v1/report/shift_report?system={sysid}&station={STATION}&shift={SHIFT}"
req = urllib.request.Request(full, method="GET",
    headers={"Authorization": f"Bearer {tok}", "Accept":"application/json"})
data = json.loads(urllib.request.urlopen(req, timeout=60).read().decode("utf-8"))
data["station"] = STATION
data["dt_open"] = "2026-04-29T08:37:29"
data["dt_close"] = "2026-04-30T08:01:24"
json_str = json.dumps(data, ensure_ascii=False)
дата = datetime.datetime(2026,4,29,23,50,0)
результат = conn.TL_СозданиеДокументов.ОбработатьСмену(json_str, key, False, дата)
print(f"  Документов: {результат.Документы.Количество()}")
for i in range(результат.Документы.Количество()):
    print(f"    {i+1}. {результат.Документы.Получить(i)}")
print(f"  Ошибок: {результат.Ошибки.Количество()}")
for i in range(результат.Ошибки.Количество()):
    print(f"    ! {результат.Ошибки.Получить(i)}")

# 4) Тест блокера: подсунуть данные с неизвестным pay_type
hr("4) Тест блокера на неизвестном pay_type")
bad_data = json.loads(json.dumps(data, ensure_ascii=False))
# Добавим запись с неизвестной оплатой
bad_data["sales"].append({
    "pay_type": {"id": 999, "name": "GPN Карта"},
    "fuel": [{"service": {"service_code": 2, "service_name": "АИ-92"},
              "release": {"volume": "10.000", "cost": "635.00", "discount": 0.0}}]
})
key_bad = f"TL|СМЕНА|{sysid}|{STATION}|{SHIFT}|BAD"
результат2 = conn.TL_СозданиеДокументов.ОбработатьСмену(
    json.dumps(bad_data, ensure_ascii=False), key_bad, False, дата)
print(f"  Документов: {результат2.Документы.Количество()} (ожидание: 0)")
print(f"  Ошибок: {результат2.Ошибки.Количество()} (ожидание: 1)")
for i in range(результат2.Ошибки.Количество()):
    print(f"    ! {результат2.Ошибки.Получить(i)}")
