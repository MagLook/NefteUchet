# -*- coding: utf-8 -*-
"""
Анализ ПеремещенийТоваров за 20.03.2026 (ручной режим, 6 шт)
и 06.04.2026 (TradeLedger, 3 шт). Сравнение по складам/товарам/направлениям.

Запуск: py -3-32 scripts/diag_movements.py
Выход:  scripts/_out_movements.txt (UTF-8)
"""
import datetime
import win32com.client

BASE_PATH = r"D:\Users\magsp\GIG Base2"
USER = "Гайворонская Татьяна"
PWD = "12345"
OUT_PATH = r"D:\Users\magsp\ELSYPLUS\NefteUchet\scripts\_out_movements.txt"

OUT = open(OUT_PATH, "w", encoding="utf-8")
def p(*a): print(*a, file=OUT)
def hr(t=""):
    p()
    p("=" * 78)
    if t: p(t)
    p("=" * 78)

def s(v):
    if v is None: return ""
    sv = str(v)
    if "<COMObject" in sv or "unknown" in sv:
        for attr in ("Представление", "Наименование", "Имя", "Код"):
            try:
                a = getattr(v, attr)
                if a is None: continue
                a_s = str(a)
                if a_s and "<COMObject" not in a_s: return a_s
            except Exception: continue
        try:
            n = getattr(v, "Номер"); d = getattr(v, "Дата")
            if n is not None: return f"{n} от {d}"
        except Exception: pass
        return ""
    return sv

def acc(c):
    try:
        k = str(c.Код)
        if k and "<COMObject" not in k: return k
    except Exception: pass
    return s(c)

def show_movement(conn, ref, with_postings=True):
    obj = ref.ПолучитьОбъект()
    p(f"\n  ── #{s(obj.Номер)} от {obj.Дата}")
    # Атрибуты
    for fld in ["Организация", "СкладОтправитель", "СкладПолучатель", "ВидОперации", "Комментарий", "Ответственный"]:
        try:
            v = getattr(obj, fld)
            sv = s(v)
            if sv: p(f"     {fld:25} = {sv[:80]}")
        except Exception: pass

    # Товары — основная ТЧ
    try:
        товары = obj.Товары
        cnt = товары.Количество()
        p(f"     ТЧ Товары: {cnt} строк")
        sum_qty = 0; sum_money = 0
        # Возьмём столбцы из первой строки
        if cnt > 0:
            row0 = товары.Получить(0)
            cols = []
            for fld in ["Номенклатура", "Количество", "Сумма", "СчетУчета",
                        "СчетУчетаПолучателя", "Партия", "ПартияПолучателя"]:
                if hasattr(row0, fld):
                    cols.append(fld)
            for i in range(min(cnt, 30)):
                row = товары.Получить(i)
                parts = []
                for fld in cols:
                    try:
                        v = getattr(row, fld)
                        if fld in ("Количество", "Сумма") and v is not None:
                            parts.append(f"{fld}={float(v):>10.2f}")
                        else:
                            sv = s(v)
                            if sv:
                                parts.append(f"{fld}={sv[:30]}")
                    except Exception: pass
                p(f"      [{i+1}] " + "  ".join(parts))
                try:
                    sum_qty += float(row.Количество or 0)
                    sum_money += float(row.Сумма or 0)
                except Exception: pass
            if cnt > 30:
                p(f"      ... ещё {cnt-30} строк")
            p(f"      ИТОГО: Кол-во={sum_qty:.2f}  Сумма={sum_money:.2f}")
    except Exception as e:
        p(f"     ТЧ Товары — ошибка: {e}")

    if with_postings:
        # Проводки
        q = conn.NewObject("Запрос")
        q.УстановитьПараметр("Регистратор", ref)
        q.Текст = '''
        ВЫБРАТЬ СчетДт, СчетКт, Сумма, Содержание
        ИЗ РегистрБухгалтерии.Хозрасчетный
        ГДЕ Регистратор = &Регистратор
        УПОРЯДОЧИТЬ ПО НомерСтроки
        '''
        sel = q.Выполнить().Выбрать()
        n = 0
        while sel.Следующий():
            n += 1
            p(f"      Дт {acc(sel.СчетДт):>8} Кт {acc(sel.СчетКт):>8}  Σ={sel.Сумма:>12.2f}  | {s(sel.Содержание)[:45]}")
        if n == 0:
            p(f"      (проводок нет)")

def list_for_day(conn, day_label, dt_beg, dt_end):
    hr(f"{day_label}: ПеремещениеТоваров")
    q = conn.NewObject("Запрос")
    q.УстановитьПараметр("Д1", dt_beg)
    q.УстановитьПараметр("Д2", dt_end)
    q.Текст = '''
    ВЫБРАТЬ Ссылка, Номер, Дата, СкладОтправитель, СкладПолучатель, ВидОперации,
            Комментарий, Проведен
    ИЗ Документ.ПеремещениеТоваров
    ГДЕ Дата МЕЖДУ &Д1 И &Д2
    УПОРЯДОЧИТЬ ПО Дата, Номер
    '''
    sel = q.Выполнить().Выбрать()
    docs = []
    while sel.Следующий():
        docs.append(sel.Ссылка)
        p(f"  #{s(sel.Номер):>15}  {sel.Дата}  {s(sel.СкладОтправитель)[:25]:25} → {s(sel.СкладПолучатель)[:25]:25}  ком={s(sel.Комментарий)[:35]}  пров={sel.Проведен}")
    p(f"  Итого: {len(docs)}")
    return docs

def main():
    com = win32com.client.Dispatch("V83.COMConnector")
    conn = com.Connect(f'File="{BASE_PATH}";Usr="{USER}";Pwd="{PWD}";')

    # Метаданные ПеремещенияТоваров — посмотрим какие реквизиты есть
    md = conn.Метаданные.НайтиПоПолномуИмени("Документ.ПеремещениеТоваров")
    hr("0) Реквизиты Документ.ПеремещениеТоваров")
    for rq in md.Реквизиты:
        try: t = str(rq.Тип)
        except Exception: t = ""
        # Только ключевые
        if any(k in rq.Имя.lower() for k in ["склад", "вид", "опер", "сум", "комм", "осн", "ответ"]):
            p(f"    {rq.Имя:35} тип={t[:80]}")
    p(f"  ТабличныеЧасти: {[t.Имя for t in md.ТабличныеЧасти]}")

    # Перемещения за 20.03.2026 (ручной)
    docs_mar = list_for_day(conn, "1) 20.03.2026 (ручной режим)",
                            datetime.datetime(2026, 3, 20, 0, 0),
                            datetime.datetime(2026, 3, 20, 23, 59, 59))

    hr("2) ДЕТАЛИ всех 6 мартовских перемещений")
    for ref in docs_mar:
        show_movement(conn, ref, with_postings=True)

    # Перемещения за 06.04.2026 (TradeLedger)
    docs_apr = list_for_day(conn, "3) 06.04.2026 (TradeLedger)",
                            datetime.datetime(2026, 4, 6, 0, 0),
                            datetime.datetime(2026, 4, 6, 23, 59, 59))

    hr("4) ДЕТАЛИ перемещений TradeLedger 06.04.2026")
    for ref in docs_apr:
        show_movement(conn, ref, with_postings=True)

    # Свод по складам за весь март — куда/откуда что движется
    hr("5) СВОД по перемещениям за март 2026 — направления склад→склад")
    q = conn.NewObject("Запрос")
    q.УстановитьПараметр("Д1", datetime.datetime(2026, 3, 1))
    q.УстановитьПараметр("Д2", datetime.datetime(2026, 4, 1))
    q.Текст = '''
    ВЫБРАТЬ
        ПРЕДСТАВЛЕНИЕ(СкладОтправитель) КАК Откуда,
        ПРЕДСТАВЛЕНИЕ(СкладПолучатель) КАК Куда,
        КОЛИЧЕСТВО(*) КАК Кол
    ИЗ Документ.ПеремещениеТоваров
    ГДЕ Дата МЕЖДУ &Д1 И &Д2 И Проведен
    СГРУППИРОВАТЬ ПО ПРЕДСТАВЛЕНИЕ(СкладОтправитель), ПРЕДСТАВЛЕНИЕ(СкладПолучатель)
    УПОРЯДОЧИТЬ ПО Кол УБЫВ
    '''
    sel = q.Выполнить().Выбрать()
    p(f"  {'Откуда':30} → {'Куда':30}  кол")
    while sel.Следующий():
        p(f"  {(sel.Откуда or '')[:30]:30} → {(sel.Куда or '')[:30]:30}  {sel.Кол:>4}")

    hr("6) СВОД по перемещениям за апрель 2026 (период TradeLedger)")
    q = conn.NewObject("Запрос")
    q.УстановитьПараметр("Д1", datetime.datetime(2026, 4, 1))
    q.УстановитьПараметр("Д2", datetime.datetime(2026, 5, 1))
    q.Текст = '''
    ВЫБРАТЬ
        ПРЕДСТАВЛЕНИЕ(СкладОтправитель) КАК Откуда,
        ПРЕДСТАВЛЕНИЕ(СкладПолучатель) КАК Куда,
        КОЛИЧЕСТВО(*) КАК Кол
    ИЗ Документ.ПеремещениеТоваров
    ГДЕ Дата МЕЖДУ &Д1 И &Д2 И Проведен
    СГРУППИРОВАТЬ ПО ПРЕДСТАВЛЕНИЕ(СкладОтправитель), ПРЕДСТАВЛЕНИЕ(СкладПолучатель)
    УПОРЯДОЧИТЬ ПО Кол УБЫВ
    '''
    sel = q.Выполнить().Выбрать()
    p(f"  {'Откуда':30} → {'Куда':30}  кол")
    while sel.Следующий():
        p(f"  {(sel.Откуда or '')[:30]:30} → {(sel.Куда or '')[:30]:30}  {sel.Кол:>4}")

if __name__ == "__main__":
    try:
        main()
    finally:
        OUT.close()
        print(f"OK: {OUT_PATH}")
