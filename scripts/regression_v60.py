# -*- coding: utf-8 -*-
"""Регрессия TradeLedger v6.0 — все сценарии маршрутизации после фикса."""
import sys, io, json, urllib.request, datetime, win32com.client
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

BASE_PATH = r"D:\Users\magsp\GIG Base2"
USER = "Гайворонская Татьяна"; PWD = "12345"

com = win32com.client.Dispatch("V83.COMConnector")
conn = com.Connect(f'File="{BASE_PATH}";Usr="{USER}";Pwd="{PWD}";')

ok_count = 0
fail_count = 0
results = []

def hr(t):
    print("\n" + "="*90); print(t); print("="*90)

def check(label, condition, detail=""):
    global ok_count, fail_count
    status = "PASS" if condition else "FAIL"
    mark = "✓" if condition else "✗"
    if condition: ok_count += 1
    else: fail_count += 1
    line = f"  [{status}] {mark} {label}"
    if detail: line += f"  ({detail})"
    print(line)
    results.append((status, label, detail))

def get_settings():
    q = conn.NewObject("Запрос")
    q.Текст = "ВЫБРАТЬ Ключ, Значение ИЗ РегистрСведений.TL_Настройки"
    sel = q.Выполнить().Выбрать()
    s = {}
    while sel.Следующий(): s[str(sel.Ключ)] = str(sel.Значение)
    return s

def get_shift_json(sts_url, sts_login, sts_pwd, sysid, station, shift, dt_open, dt_close):
    body = {"login": sts_login, "password": sts_pwd,
            "user": {"id":"00000000-0000-0000-0000-000000000000","name":"System"},
            "system_id": int(sysid)}
    req = urllib.request.Request(sts_url.rstrip("/")+"/v2/login",
        data=json.dumps(body).encode("utf-8"), method="POST",
        headers={"Content-Type":"application/json"})
    raw = urllib.request.urlopen(req, timeout=20).read().decode("utf-8").strip().strip('"')
    tok = json.loads(raw)["token"] if raw.startswith("{") else raw
    full = f"{sts_url.rstrip('/')}/v1/report/shift_report?system={sysid}&station={station}&shift={shift}"
    req = urllib.request.Request(full, method="GET",
        headers={"Authorization": f"Bearer {tok}", "Accept":"application/json"})
    data = json.loads(urllib.request.urlopen(req, timeout=60).read().decode("utf-8"))
    data["station"] = station
    data["dt_open"] = dt_open
    data["dt_close"] = dt_close
    return data

def clear_status(key):
    rs = conn.РегистрыСведений.TL_СтатусыЗагрузки
    for k in (key, key+"|ПКО"):
        мн = rs.СоздатьМенеджерЗаписи()
        мн.КлючЗагрузки = k
        мн.Прочитать()
        if мн.Выбран(): мн.Удалить()

def get_orp_by_key(key):
    q = conn.NewObject("Запрос")
    q.Текст = ("ВЫБРАТЬ ПЕРВЫЕ 1 ОРП.Ссылка ИЗ Документ.ОтчетОРозничныхПродажах КАК ОРП "
               "ГДЕ ОРП.Комментарий ПОДОБНО &Шаб")
    q.УстановитьПараметр("Шаб", key + "%")
    sel = q.Выполнить().Выбрать()
    if sel.Следующий(): return sel.Ссылка
    return None

def get_перемещения(key):
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
    q.УстановитьПараметр("Шаб", key + "%")
    sel = q.Выполнить().Выбрать()
    rows = []
    while sel.Следующий():
        rows.append((str(sel.Куда), float(sel.Литры)))
    return rows

def помечен_удаление_доки(шаб):
    for имя in ("ОтчетОРозничныхПродажах", "ПриходныйКассовыйОрдер", "ПеремещениеТоваров"):
        q = conn.NewObject("Запрос")
        q.Текст = f"ВЫБРАТЬ Док.Ссылка ИЗ Документ.{имя} КАК Док ГДЕ Док.Комментарий ПОДОБНО &Шаб"
        q.УстановитьПараметр("Шаб", шаб + "%")
        sel = q.Выполнить().Выбрать()
        while sel.Следующий():
            obj = sel.Ссылка.ПолучитьОбъект()
            obj.УстановитьПометкуУдаления(True)

# === ТЕСТЫ ===

hr("ТЕСТ 1: ОпределитьКаналОплаты — каждое имя по маппингу")
cases = [
    ("Наличные",       "retail_cash"),
    ("Сбербанк",       "retail_card"),
    ("VIAcard",        "cards"),
    ("БАЛТОП",         "cards"),
    ("Балтоп",         "cards"),
    ("Корпоративная",  "cards"),
    ("Топливная карта","cards"),
    ("Яндекс.Заправки","online"),
    ("Мобильное приложение", "online"),
    ("Benzuber",       "online"),
    ("Ведомость №1",   "ledger"),
    ("Талоны",         "voucher"),
    ("Купон на сдачу", ""),
    ("Прокачка",       ""),
]
for имя, ожидание in cases:
    канал = conn.TL_СозданиеДокументов.ОпределитьКаналОплаты(имя.lower())
    канал_str = "" if канал is None else str(канал)
    check(f"'{имя}' → '{ожидание}'", канал_str == ожидание, f"получено '{канал_str}'")

# Тест неизвестного имени → должен вернуть Неопределено (или похожее)
канал = conn.TL_СозданиеДокументов.ОпределитьКаналОплаты("GPN-Карта")
check("Неизвестный 'GPN-Карта' → Неопределено/блокер",
      канал is None or канал == "" or канал == "Неопределено" or str(канал) == "None",
      f"получено: {canал if False else type(канал).__name__} = {канал}".replace("canал","канал"))

hr("ТЕСТ 2: НайтиЗаписьМаппинга возвращает Склад")
зап = conn.TL_СозданиеДокументов.НайтиЗаписьМаппинга("viacard")
check("viacard → запись найдена", зап is not None)
if зап is not None:
    check("viacard.КаналОплаты = 'cards'", str(зап.КаналОплаты) == "cards")
    имя_скл = str(зап.Склад.Наименование) if зап.Склад else ""
    check("viacard.Склад = 'Карты'", имя_скл == "Карты", f"получено: '{имя_скл}'")

зап = conn.TL_СозданиеДокументов.НайтиЗаписьМаппинга("талон")
check("талон → запись найдена", зап is not None)
if зап is not None:
    check("талон.КаналОплаты = 'voucher'", str(зап.КаналОплаты) == "voucher")
    имя_скл = str(зап.Склад.Наименование) if зап.Склад else ""
    check("талон.Склад = 'Талоны'", имя_скл == "Талоны", f"получено: '{имя_скл}'")

# === ТЕСТ 3: Полный прогон смены 3227 ===
hr("ТЕСТ 3: Полный прогон смены 3227 АЗС 210")
s = get_settings()
url, login, pwd = s["URLСервера"], s["Логин"], s["Пароль"]
sysid = s.get("Станция_210_КодСистемы", s.get("КодСистемы","15"))

data = get_shift_json(url, login, pwd, sysid, 210, 3227,
                     "2026-04-29T08:37:29", "2026-04-30T08:01:24")
key = f"TL|СМЕНА|{sysid}|210|3227"
помечен_удаление_доки(key)
clear_status(key)

дата = datetime.datetime(2026,4,29,23,50,0)
рез = conn.TL_СозданиеДокументов.ОбработатьСмену(json.dumps(data, ensure_ascii=False), key, False, дата)
errors_3227 = [str(рез.Ошибки.Получить(i)) for i in range(рез.Ошибки.Количество())]
check("Смена 3227: 0 ошибок", len(errors_3227) == 0, "; ".join(errors_3227))
check("Смена 3227: 4 документа", рез.Документы.Количество() == 4,
      f"создано {рез.Документы.Количество()}")

орп = get_orp_by_key(key)
if орп is not None:
    obj = орп.ПолучитьОбъект()
    sum_dok = float(obj.СуммаДокумента)
    check("ОРП.СуммаДокумента = 87 656,67",
          abs(sum_dok - 87656.67) < 0.5, f"получено: {sum_dok:.2f}")
    check("ОРП: 6 строк товаров", obj.Товары.Количество() == 6,
          f"строк: {obj.Товары.Количество()}")
    check("ОРП.Оплата = 60 723,42 (Сбер)",
          obj.Оплата.Количество() == 1 and abs(float(obj.Оплата.Получить(0).СуммаОплаты) - 60723.42) < 0.5,
          f"оплат: {obj.Оплата.Количество()}")
    # Проверка комментария — литры по безналу, суммы по retail.
    # Нормализуем неразрывный пробел (1С форматирует числа NBSP \xa0).
    coмм = str(obj.Комментарий).replace(" ", " ").replace("\xa0", " ")
    check("Комментарий содержит 'voucher: 50,00 л'", "voucher: 50,00 л" in coмм,
          f"комм: {coмм[:120]}")
    check("Комментарий содержит 'cards: 2 740,58 л'", "cards: 2 740,58 л" in coмм,
          f"комм: {coмм[:120]}")
    check("Комментарий содержит 'retail_cash: 26 933,25'",
          "retail_cash: 26 933,25" in coмм)

перем = get_перемещения(key)
скл = {куда: л for куда, л in перем}
check("Перемещение → Карты 2 740,58 л",
      "Карты" in скл and abs(скл.get("Карты", 0) - 2740.58) < 0.5,
      f"скл: {скл}")
check("Перемещение → Талоны 50,00 л",
      "Талоны" in скл and abs(скл.get("Талоны", 0) - 50.0) < 0.1,
      f"скл: {скл}")

# === ТЕСТ 4: Блокер на неизвестном виде оплаты ===
hr("ТЕСТ 4: Блокер на неизвестном виде оплаты")
bad = json.loads(json.dumps(data, ensure_ascii=False))
bad["sales"].append({
    "pay_type": {"id": 999, "name": "GPN Карта"},
    "fuel": [{"service": {"service_code": 2, "service_name": "АИ-92"},
              "release": {"volume": "10.000", "cost": "635.00", "discount": 0.0}}]
})
key_bad = f"TL|СМЕНА|{sysid}|210|3227|BAD"
рез2 = conn.TL_СозданиеДокументов.ОбработатьСмену(json.dumps(bad, ensure_ascii=False), key_bad, False, дата)
check("Неизвестное имя → 0 документов", рез2.Документы.Количество() == 0)
check("Неизвестное имя → 1 ошибка", рез2.Ошибки.Количество() == 1)
if рез2.Ошибки.Количество() > 0:
    err = str(рез2.Ошибки.Получить(0))
    check("Текст ошибки содержит 'GPN Карта'", "GPN Карта" in err)
    check("Текст ошибки содержит подсказку про маппинг", "Маппинг" in err or "маппинг" in err)

# === ТЕСТ 5: Переопределение склада ===
hr("ТЕСТ 5: Переопределение склада на строке маппинга")
тест_склад = conn.Справочники.Склады.НайтиПоНаименованию("БалТоп", True)

мн = conn.РегистрыСведений.TL_МаппингОплат.СоздатьМенеджерЗаписи()
мн.ОбразецИмени = "viacard"
мн.Прочитать()
исх_склад = мн.Склад
мн.Склад = тест_склад
мн.Записать()
check("Переопределён склад viacard → БалТоп", True)

# Прогоняем
key5 = f"TL|СМЕНА|{sysid}|210|3227|OVERRIDE"
clear_status(key5)
помечен_удаление_доки(key5)
рез5 = conn.TL_СозданиеДокументов.ОбработатьСмену(json.dumps(data, ensure_ascii=False), key5, False, дата)
check("Смена с переопр: 0 ошибок", рез5.Ошибки.Количество() == 0,
      "; ".join([str(рез5.Ошибки.Получить(i)) for i in range(рез5.Ошибки.Количество())]))
перем5 = get_перемещения(key5)
скл5 = {куда: л for куда, л in перем5}
check("Перемещение → БалТоп появилось", "БалТоп" in скл5,
      f"склады: {list(скл5.keys())}")
# Литры VIAcard на склад БалТоп = АИ-92 454,45 + АИ-95 35,00 + ДТ 347,01 = 836,46
check("БалТоп = 836,46 л (только VIAcard)",
      "БалТоп" in скл5 and abs(скл5.get("БалТоп", 0) - 836.46) < 0.5,
      f"л={скл5.get('БалТоп', 0):.2f}")
# Карты теперь = только БАЛТОП = 211,60 + 186,40 + 1 506,12 = 1 904,12
check("Карты теперь только БАЛТОП = 1 904,12",
      "Карты" in скл5 and abs(скл5.get("Карты", 0) - 1904.12) < 0.5,
      f"л={скл5.get('Карты', 0):.2f}")

# Откат
мн = conn.РегистрыСведений.TL_МаппингОплат.СоздатьМенеджерЗаписи()
мн.ОбразецИмени = "viacard"
мн.Прочитать()
мн.Склад = исх_склад
мн.Записать()
помечен_удаление_доки(key5)
clear_status(key5)

# === ТЕСТ 6: Игнорируемый канал (купон) ===
hr("ТЕСТ 6: Купоны игнорируются")
бад = json.loads(json.dumps(data, ensure_ascii=False))
бад["sales"].append({
    "pay_type": {"id": 100, "name": "Купон на сдачу"},
    "fuel": [{"service": {"service_code": 2, "service_name": "АИ-92"},
              "release": {"volume": "5.000", "cost": "0.00", "discount": 0.0}}]
})
key6 = f"TL|СМЕНА|{sysid}|210|3227|COUPON"
помечен_удаление_доки(key6)
clear_status(key6)
рез6 = conn.TL_СозданиеДокументов.ОбработатьСмену(json.dumps(бад, ensure_ascii=False), key6, False, дата)
check("Купон не вызывает блокер (есть в маппинге как игнор)",
      рез6.Ошибки.Количество() == 0,
      "; ".join([str(рез6.Ошибки.Получить(i)) for i in range(рез6.Ошибки.Количество())]))
check("Купон не создал лишних документов",
      рез6.Документы.Количество() == 4, f"создано: {рез6.Документы.Количество()}")
перем6 = get_перемещения(key6)
скл6 = {куда: л for куда, л in перем6}
check("Купон не попал в Карты", abs(скл6.get("Карты", 0) - 2740.58) < 0.5,
      f"Карты л={скл6.get('Карты', 0):.2f}")

помечен_удаление_доки(key6)
помечен_удаление_доки(key)

# === ИТОГ ===
hr("ИТОГ")
print(f"  Пройдено: {ok_count}")
print(f"  Провалено: {fail_count}")
print(f"  Всего проверок: {ok_count + fail_count}")
if fail_count > 0:
    print("\nПровалы:")
    for status, label, detail in results:
        if status == "FAIL":
            print(f"  ✗ {label}  ({detail})")
sys.exit(0 if fail_count == 0 else 1)
