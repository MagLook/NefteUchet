# -*- coding: utf-8 -*-
"""
Анализ: создаются ли ПКО при проведении ОРП в БП ГИГ?
Какие проводки делает ПКО (если есть)?

Запуск: py -3-32 scripts/diag_cash_orders.py
"""
import sys
import win32com.client

BASE_PATH = r"D:\Users\magsp\GIG Base2"
USER = "Гайворонская Татьяна"
PWD = "12345"

def hr(t=""):
    print()
    print("=" * 78)
    if t: print(t)
    print("=" * 78)

def safe(v):
    try: return str(v)
    except Exception: return "?"

def main():
    com = win32com.client.Dispatch("V83.COMConnector")
    conn = com.Connect(f'File="{BASE_PATH}";Usr="{USER}";Pwd="{PWD}";')

    hr("1) Список последних ОтчётовОРозничныхПродажах")
    q = conn.NewObject("Запрос")
    q.Текст = '''
    ВЫБРАТЬ ПЕРВЫЕ 20
        Ссылка, Номер, Дата, Организация, Склад, СуммаДокумента, Проведен, Комментарий
    ИЗ Документ.ОтчетОРозничныхПродажах
    УПОРЯДОЧИТЬ ПО Дата УБЫВ
    '''
    sel = q.Выполнить().Выбрать()
    orps = []
    while sel.Следующий():
        orps.append({
            "ref": sel.Ссылка,
            "number": safe(sel.Номер),
            "date": safe(sel.Дата),
            "org": safe(sel.Организация),
            "wh": safe(sel.Склад),
            "sum": safe(sel.СуммаДокумента),
            "posted": bool(sel.Проведен),
            "comment": safe(sel.Комментарий),
        })
    print(f"  Найдено: {len(orps)}")
    for o in orps[:5]:
        print(f"  {o['date']:25} #{o['number']:>10}  {o['sum']:>12}  пров={o['posted']}  склад={o['wh']}")
        if o["comment"]:
            print(f"      ком: {o['comment'][:100]}")

    if not orps:
        print("  Нет ОРП — выйти.")
        return

    # Берём свежий проведённый ОРП
    target = None
    for o in orps:
        if o["posted"]:
            target = o; break
    if target is None:
        target = orps[0]

    hr(f"2) Анализ ОРП {target['number']} от {target['date']}")
    ref = target["ref"]
    obj = ref.ПолучитьОбъект()
    print(f"  Ссылка: {ref}")
    print(f"  Дата: {obj.Дата}")
    print(f"  Организация: {obj.Организация}")
    print(f"  Склад: {obj.Склад}")
    print(f"  СуммаДокумента: {obj.СуммаДокумента}")
    print(f"  Проведен: {obj.Проведен}")

    # Реквизиты которые могут отвечать за ПКО
    interesting = []
    for attr in dir(obj):
        if attr.startswith("_"): continue
        low = attr.lower()
        if any(k in low for k in ["приходный", "кассов", "ордер", "пко", "налич"]):
            try:
                val = getattr(obj, attr)
                interesting.append((attr, str(val)[:120]))
            except Exception:
                pass
    print("\n  Реквизиты с признаками 'кассовый/наличный/ордер':")
    for a, v in interesting:
        print(f"    {a:40} = {v}")

    # ТЧ Оплаты — ключевая для понимания
    hr("3) ТабличнаяЧасть Оплаты")
    try:
        for row in obj.Оплата:
            print(f"  ВидОплаты={safe(row.ВидОплаты):40} Сумма={row.Сумма:>12}")
    except Exception as e:
        print(f"  Нет ТЧ Оплата или ошибка: {e}")

    # ТЧ Товары и/или Чеки
    hr("4) Чеки / Товары — структура")
    for tab_name in ["Чеки", "Товары", "ДенежныеСредстваБезналичные", "БезналичныеОплаты"]:
        try:
            tab = getattr(obj, tab_name)
            cnt = tab.Количество()
            print(f"  ТЧ {tab_name}: {cnt} строк")
            if cnt > 0:
                row = tab.Получить(0)
                cols = []
                for fld in dir(row):
                    if fld.startswith("_"): continue
                    if fld in ("Метаданные", "ВладелецЯчейки"): continue
                    try:
                        v = getattr(row, fld)
                        if not callable(v):
                            cols.append((fld, str(v)[:60]))
                    except Exception:
                        pass
                for k, v in cols[:25]:
                    print(f"    {k:30} = {v}")
        except AttributeError:
            pass
        except Exception as e:
            print(f"  ТЧ {tab_name} ошибка: {e}")

    hr("5) Проводки самого ОРП (регистр Хозрасчётный)")
    q = conn.NewObject("Запрос")
    q.УстановитьПараметр("Регистратор", ref)
    q.Текст = '''
    ВЫБРАТЬ
        Период, СчетДт, СчетКт, Сумма, СуммаНУДт, СуммаНУКт,
        Содержание, ПодразделениеДт, ПодразделениеКт
    ИЗ РегистрБухгалтерии.Хозрасчетный
    ГДЕ Регистратор = &Регистратор
    УПОРЯДОЧИТЬ ПО Период, СчетДт, СчетКт
    '''
    sel = q.Выполнить().Выбрать()
    n = 0
    while sel.Следующий():
        n += 1
        print(f"  Дт {safe(sel.СчетДт):10} Кт {safe(sel.СчетКт):10} = {sel.Сумма:>12}  | {safe(sel.Содержание)[:60]}")
    print(f"  Всего проводок ОРП: {n}")

    hr("6) Поиск кассовых ордеров за дату ОРП по этой организации/складу")
    q = conn.NewObject("Запрос")
    q.УстановитьПараметр("ДатаНач", obj.Дата)
    # Берём день целиком — от 00:00 до 23:59
    from datetime import datetime, time
    d = obj.Дата
    if hasattr(d, "year"):
        # дата 1С — это pywintypes.datetime
        # сравним по дате через 1С-запрос НАЧАЛОПЕРИОДА/КОНЕЦПЕРИОДА
        pass
    q.УстановитьПараметр("Организация", obj.Организация)

    q.Текст = '''
    ВЫБРАТЬ
        Ссылка, Номер, Дата, ВидОперации, СуммаДокумента, ОснованиеДокумента,
        Проведен
    ИЗ Документ.ПриходныйКассовыйОрдер
    ГДЕ Дата >= &ДатаНач
        И Дата <= ДОБАВИТЬКДАТЕ(&ДатаНач, ДЕНЬ, 1)
        И Организация = &Организация
    УПОРЯДОЧИТЬ ПО Дата
    '''
    sel = q.Выполнить().Выбрать()
    pkos = []
    while sel.Следующий():
        pkos.append({
            "ref": sel.Ссылка,
            "number": safe(sel.Номер),
            "date": safe(sel.Дата),
            "vid": safe(sel.ВидОперации),
            "sum": safe(sel.СуммаДокумента),
            "base": safe(sel.ОснованиеДокумента),
            "posted": bool(sel.Проведен),
        })
    print(f"  Найдено ПКО за день: {len(pkos)}")
    for p in pkos:
        print(f"  {p['date']:25} #{p['number']:>10}  {p['sum']:>12}  вид={p['vid'][:30]:30}  пров={p['posted']}")
        print(f"    Основание: {p['base'][:80]}")

    # Если есть ПКО, привязанные к нашему ОРП — посмотрим их проводки
    hr("7) Проводки ПКО, у которых ОснованиеДокумента = наш ОРП")
    matching = [p for p in pkos if str(ref) in p["base"]]
    print(f"  Привязанных к ОРП {target['number']}: {len(matching)}")
    if not matching and pkos:
        print(f"  (показываю проводки первого ПКО за день)")
        matching = pkos[:1]
    for p in matching:
        print(f"\n  --- ПКО {p['number']} от {p['date']}, сумма {p['sum']} ---")
        q = conn.NewObject("Запрос")
        q.УстановитьПараметр("Регистратор", p["ref"])
        q.Текст = '''
        ВЫБРАТЬ Период, СчетДт, СчетКт, Сумма, Содержание
        ИЗ РегистрБухгалтерии.Хозрасчетный
        ГДЕ Регистратор = &Регистратор
        УПОРЯДОЧИТЬ ПО Период, СчетДт, СчетКт
        '''
        sel = q.Выполнить().Выбрать()
        n = 0
        while sel.Следующий():
            n += 1
            print(f"    Дт {safe(sel.СчетДт):10} Кт {safe(sel.СчетКт):10} = {sel.Сумма:>12}  | {safe(sel.Содержание)[:60]}")
        if n == 0:
            print(f"    (проводок нет!)")

    hr("8) То же — РКО (Расходный кассовый ордер) за день")
    q = conn.NewObject("Запрос")
    q.УстановитьПараметр("ДатаНач", obj.Дата)
    q.УстановитьПараметр("Организация", obj.Организация)
    q.Текст = '''
    ВЫБРАТЬ
        Ссылка, Номер, Дата, ВидОперации, СуммаДокумента, ОснованиеДокумента,
        Проведен
    ИЗ Документ.РасходныйКассовыйОрдер
    ГДЕ Дата >= &ДатаНач
        И Дата <= ДОБАВИТЬКДАТЕ(&ДатаНач, ДЕНЬ, 1)
        И Организация = &Организация
    '''
    sel = q.Выполнить().Выбрать()
    rkos = 0
    while sel.Следующий():
        rkos += 1
        print(f"  {safe(sel.Дата):25} #{safe(sel.Номер):>10}  {sel.СуммаДокумента:>12}  вид={safe(sel.ВидОперации)[:30]}")
    print(f"  Всего РКО за день: {rkos}")

    hr("9) Стат: за последние 30 дней — сколько ОРП vs сколько ПКО, есть ли связь")
    q = conn.NewObject("Запрос")
    q.Текст = '''
    ВЫБРАТЬ
        КОЛИЧЕСТВО(ОРП.Ссылка) КАК ВсегоОРП
    ИЗ Документ.ОтчетОРозничныхПродажах ОРП
    ГДЕ ОРП.Дата >= ДОБАВИТЬКДАТЕ(&Сейчас, ДЕНЬ, -30)
    '''
    from pywintypes import Time
    import datetime
    q.УстановитьПараметр("Сейчас", datetime.datetime.now())
    sel = q.Выполнить().Выбрать()
    if sel.Следующий():
        print(f"  ОРП за 30 дней: {sel.ВсегоОРП}")

    q = conn.NewObject("Запрос")
    q.УстановитьПараметр("Сейчас", datetime.datetime.now())
    q.Текст = '''
    ВЫБРАТЬ
        КОЛИЧЕСТВО(ПКО.Ссылка) КАК ВсегоПКО,
        СУММА(ВЫБОР КОГДА ТИПЗНАЧЕНИЯ(ПКО.ОснованиеДокумента) = ТИП(Документ.ОтчетОРозничныхПродажах)
                ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК СвязанСОРП
    ИЗ Документ.ПриходныйКассовыйОрдер ПКО
    ГДЕ ПКО.Дата >= ДОБАВИТЬКДАТЕ(&Сейчас, ДЕНЬ, -30)
    '''
    sel = q.Выполнить().Выбрать()
    if sel.Следующий():
        print(f"  ПКО за 30 дней: {sel.ВсегоПКО}")
        print(f"  ПКО с основанием = ОРП: {sel.СвязанСОРП}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        traceback.print_exc()
        sys.exit(1)
