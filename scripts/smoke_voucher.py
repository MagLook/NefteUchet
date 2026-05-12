import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _tl_config import BASE_PATH, USER, PWD
# -*- coding: utf-8 -*-
"""Smoke-фикс для талонов:
1. Создать склад «Талоны» (если нет) в Catalog.Склады
2. Дописать в TL_Настройки три ключа Оплата_voucher_*
3. Прогнать ОбработатьСмену для смены 3227 АЗС 210 (с очисткой регистра статусов)
4. Проверить созданные документы и комментарии
"""
import sys, io, json, urllib.request, win32com.client
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
STATION = 210; SHIFT = 3227

com = win32com.client.Dispatch("V83.COMConnector")
conn = com.Connect(f'File="{BASE_PATH}";Usr="{USER}";Pwd="{PWD}";')

def hr(t):
    print("\n" + "="*90); print(t); print("="*90)

# 1) Склад «Талоны»
hr("1) Склад «Талоны»")
sklad = conn.Справочники.Склады.НайтиПоНаименованию("Талоны", True)
if str(sklad) == "" or sklad is None or bool(getattr(sklad, "Пустая", lambda: False)()):
    obj = conn.Справочники.Склады.СоздатьЭлемент()
    obj.Наименование = "Талоны"
    try: obj.Код = "0Г-Талоны"
    except Exception: pass
    try: obj.ТипСклада = conn.Перечисления.ТипыСкладов.Оптовый
    except Exception: pass
    obj.Записать()
    sklad = obj.Ссылка
    print(f"  СОЗДАН: {sklad}")
else:
    print(f"  УЖЕ ЕСТЬ: {sklad}")

# 2) TL_Настройки — три ключа
hr("2) TL_Настройки — ключи voucher")
rs = conn.РегистрыСведений.TL_Настройки
keys = {
    "Оплата_voucher_Наименование":      "Талоны",
    "Оплата_voucher_Склад":             "Талоны",
    "Оплата_voucher_ТребуетПеремещения": "Да",
}
for k, v in keys.items():
    мн = rs.СоздатьМенеджерЗаписи()
    мн.Ключ = k
    мн.Прочитать()
    if мн.Выбран() and str(мн.Значение) == v:
        print(f"  УЖЕ: {k} = {v}")
        continue
    мн.Ключ = k
    мн.Значение = v
    мн.Записать()
    print(f"  ЗАПИСАН: {k} = {v}")

# 3) Получить сменный отчёт STS, передать в ОбработатьСмену
hr("3) STS shift_report для смены 3227")
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

# В sales есть station? Добавим, иначе ОбработатьСмену не найдёт склад
full = f"{url.rstrip('/')}/v1/report/shift_report?system={sysid}&station={STATION}&shift={SHIFT}"
req = urllib.request.Request(full, method="GET",
    headers={"Authorization": f"Bearer {tok}", "Accept":"application/json"})
data = json.loads(urllib.request.urlopen(req, timeout=60).read().decode("utf-8"))
data["station"] = STATION
data["dt_open"] = "2026-04-29T08:37:29"
data["dt_close"] = "2026-04-30T08:01:24"
json_str = json.dumps(data, ensure_ascii=False)
print(f"  JSON получен: {len(json_str)} симв, sales: {len(data.get('sales',[]))} записей")

# 4) Очистить регистр статусов для нашего ключа, чтобы не упереться в «уже загружено»
hr("4) Очистка TL_РегистрСтатусов для ключа TL|СМЕНА|15|210|3227")
key = f"TL|СМЕНА|{sysid}|{STATION}|{SHIFT}"
try:
    rs2 = conn.РегистрыСведений.TL_СтатусыЗагрузки
    мн = rs2.СоздатьМенеджерЗаписи()
    мн.КлючЗагрузки = key
    мн.Прочитать()
    if мн.Выбран():
        мн.Удалить()
        print(f"  Удалён статус: {key}")
    else:
        print(f"  Статус не найден (норм): {key}")
except Exception as e:
    print(f"  ОШИБКА чистки статуса: {e}")
# Заодно ПКО
try:
    мн = rs2.СоздатьМенеджерЗаписи()
    мн.КлючЗагрузки = key + "|ПКО"
    мн.Прочитать()
    if мн.Выбран():
        мн.Удалить()
        print(f"  Удалён статус ПКО")
except Exception:
    pass

# 5) ОбработатьСмену
hr("5) Запуск TL_СозданиеДокументов.ОбработатьСмену")
import datetime
дата = datetime.datetime(2026,4,29,23,50,0)
результат = conn.TL_СозданиеДокументов.ОбработатьСмену(json_str, key, False, дата)
print(f"  Документов: {результат.Документы.Количество()}")
for i in range(результат.Документы.Количество()):
    док = результат.Документы.Получить(i)
    print(f"    {i+1}. {док}")
print(f"  Ошибок: {результат.Ошибки.Количество()}")
for i in range(результат.Ошибки.Количество()):
    print(f"    ! {результат.Ошибки.Получить(i)}")

# 6) Анализ ОРП
hr("6) Анализ созданного ОРП")
q = conn.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ ПЕРВЫЕ 1 ОРП.Ссылка
ИЗ Документ.ОтчетОРозничныхПродажах КАК ОРП
ГДЕ ОРП.Комментарий ПОДОБНО &Шаб
УПОРЯДОЧИТЬ ПО ОРП.МоментВремени УБЫВ
"""
q.УстановитьПараметр("Шаб", f"TL|СМЕНА|{sysid}|{STATION}|{SHIFT}%")
sel = q.Выполнить().Выбрать()
if sel.Следующий():
    док = sel.Ссылка.ПолучитьОбъект()
    print(f"  Документ: {sel.Ссылка}")
    print(f"  Дата: {док.Дата}  Σ: {float(док.СуммаДокумента):.2f}")
    print(f"  Комментарий: {док.Комментарий}")
    print(f"  Товары ({док.Товары.Количество()} строк):")
    итого_л = итого_р = 0
    for i in range(док.Товары.Количество()):
        r = док.Товары.Получить(i)
        ном = str(r.Номенклатура)
        итого_л += float(r.Количество); итого_р += float(r.Сумма)
        print(f"    {i+1}. {ном:30} л={float(r.Количество):>10.2f}  ц={float(r.Цена):>7.2f}  Σ={float(r.Сумма):>12.2f}")
    print(f"    ИТОГО:                                л={итого_л:>10.2f}                  Σ={итого_р:>12.2f}")
    print(f"  Безналичные оплаты ({док.Оплата.Количество()}):")
    for i in range(док.Оплата.Количество()):
        r = док.Оплата.Получить(i)
        print(f"    {i+1}. {str(r.ВидОплаты):25}  Σ={float(r.СуммаОплаты):>12.2f}")
else:
    print("  ОРП не найден")

# 7) Перемещения
hr("7) Перемещения по этой смене")
q = conn.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ Перем.Ссылка, Перем.СкладПолучатель.Наименование КАК Куда,
    СУММА(ПеремТЧ.Количество) КАК Литры
ИЗ Документ.ПеремещениеТоваров КАК Перем
    ВНУТРЕННЕЕ СОЕДИНЕНИЕ Документ.ПеремещениеТоваров.Товары КАК ПеремТЧ
    ПО ПеремТЧ.Ссылка = Перем.Ссылка
ГДЕ Перем.Комментарий ПОДОБНО &Шаб
СГРУППИРОВАТЬ ПО Перем.Ссылка, Перем.СкладПолучатель.Наименование
"""
q.УстановитьПараметр("Шаб", f"TL|СМЕНА|{sysid}|{STATION}|{SHIFT}%")
sel = q.Выполнить().Выбрать()
n = 0
while sel.Следующий():
    n += 1
    print(f"  {sel.Ссылка}  → {sel.Куда:15}  л={float(sel.Литры):>10.2f}")
if n == 0: print("  (нет перемещений)")
