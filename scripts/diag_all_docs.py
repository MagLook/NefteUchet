# -*- coding: utf-8 -*-
"""
Сканер: за день/АЗС какие документы создаёт бухгалтер вручную.
Период — март (когда был ручной режим), сравнение с апрелем (после TradeLedger).

Запуск: py -3-32 scripts/diag_all_docs.py
Выход:  scripts/_out_all_docs.txt (UTF-8)
"""
import datetime
import win32com.client

BASE_PATH = r"D:\Users\magsp\GIG Base2"
USER = "Гайворонская Татьяна"
PWD = "12345"
OUT_PATH = r"D:\Users\magsp\ELSYPLUS\NefteUchet\scripts\_out_all_docs.txt"

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

def main():
    com = win32com.client.Dispatch("V83.COMConnector")
    conn = com.Connect(f'File="{BASE_PATH}";Usr="{USER}";Pwd="{PWD}";')

    # === Шаг 1: Перебор ВСЕХ типов документов и подсчёт за две даты ===
    # Период ручного режима: 20.03.2026
    # Период после TradeLedger: 06.04.2026
    DATE_MANUAL = "20260320"  # 1С формат ггггммдд
    DATE_AUTO   = "20260406"

    hr(f"1) Сканирование ВСЕХ документов: {DATE_MANUAL} (ручной) vs {DATE_AUTO} (TradeLedger)")

    # Собираем имена всех документов
    doc_names = []
    for d in conn.Метаданные.Документы:
        doc_names.append(d.Имя)
    p(f"  Всего типов документов в конфигурации: {len(doc_names)}")

    counts = {}
    for name in doc_names:
        try:
            q = conn.NewObject("Запрос")
            q.УстановитьПараметр("Д1Нач", datetime.datetime(2026, 3, 20, 0, 0))
            q.УстановитьПараметр("Д1Кон", datetime.datetime(2026, 3, 20, 23, 59, 59))
            q.УстановитьПараметр("Д2Нач", datetime.datetime(2026, 4, 6, 0, 0))
            q.УстановитьПараметр("Д2Кон", datetime.datetime(2026, 4, 6, 23, 59, 59))
            q.Текст = f'''
            ВЫБРАТЬ
                СУММА(ВЫБОР КОГДА Д.Дата МЕЖДУ &Д1Нач И &Д1Кон ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Март20,
                СУММА(ВЫБОР КОГДА Д.Дата МЕЖДУ &Д2Нач И &Д2Кон ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Апрель06
            ИЗ Документ.{name} КАК Д
            ГДЕ Д.Дата МЕЖДУ &Д1Нач И &Д2Кон
            '''
            sel = q.Выполнить().Выбрать()
            if sel.Следующий():
                m = sel.Март20 or 0
                a = sel.Апрель06 or 0
                if m > 0 or a > 0:
                    counts[name] = (m, a)
        except Exception:
            pass

    p(f"\n  Типов документов с записями в эти дни: {len(counts)}")
    p(f"  {'Документ':45} {'20.03':>8} {'06.04':>8}  Δ")
    for name in sorted(counts.keys()):
        m, a = counts[name]
        delta = ""
        if m > 0 and a == 0: delta = "← только МАРТ (ручная)"
        if m == 0 and a > 0: delta = "← только АПРЕЛЬ (TradeLedger)"
        p(f"  {name:45} {m:>8} {a:>8}  {delta}")

    # === Шаг 2: Детально по 20.03.2026 — что именно сделано ===
    hr("2) ДЕТАЛЬНО по 20.03.2026 — все документы, которые проводила бухгалтер")

    interesting = [n for n, (m, _) in counts.items() if m > 0]
    for name in interesting:
        try:
            q = conn.NewObject("Запрос")
            q.УстановитьПараметр("Д1", datetime.datetime(2026, 3, 20, 0, 0))
            q.УстановитьПараметр("Д2", datetime.datetime(2026, 3, 20, 23, 59, 59))
            # Определим ключевые реквизиты, чтобы вытащить
            md = conn.Метаданные.НайтиПоПолномуИмени(f"Документ.{name}")
            req_names = set()
            for rq in md.Реквизиты:
                req_names.add(rq.Имя)

            extra = []
            for f in ["Номер", "Дата", "СуммаДокумента", "Организация", "Склад",
                      "Контрагент", "ВидОперации", "Проведен", "Комментарий"]:
                if f in req_names or f in ("Номер", "Дата", "Проведен"):
                    extra.append(f)

            q.Текст = f"""
            ВЫБРАТЬ {','.join(extra)}, Ссылка
            ИЗ Документ.{name}
            ГДЕ Дата МЕЖДУ &Д1 И &Д2
            УПОРЯДОЧИТЬ ПО Дата
            """
            sel = q.Выполнить().Выбрать()
            rows = []
            while sel.Следующий():
                row = {f: getattr(sel, f) for f in extra}
                row["Ссылка"] = sel.Ссылка
                rows.append(row)
            if not rows: continue

            p(f"\n--- {name}  ({len(rows)} шт)")
            for r in rows:
                # компактный вывод
                parts = []
                if "Номер" in r: parts.append(f"#{s(r['Номер'])}")
                if "СуммаДокумента" in r and r['СуммаДокумента'] is not None:
                    parts.append(f"Σ={r['СуммаДокумента']:.2f}")
                for f in ["ВидОперации", "Контрагент", "Склад", "Комментарий"]:
                    if f in r and s(r[f]):
                        parts.append(f"{f}={s(r[f])[:35]}")
                pr = r.get("Проведен")
                if pr is not None: parts.append(f"пров={pr}")
                p(f"   {' | '.join(parts)}")

                # Проводки документа
                q2 = conn.NewObject("Запрос")
                q2.УстановитьПараметр("Регистратор", r["Ссылка"])
                q2.Текст = """
                ВЫБРАТЬ СчетДт, СчетКт, Сумма, Содержание
                ИЗ РегистрБухгалтерии.Хозрасчетный
                ГДЕ Регистратор = &Регистратор
                УПОРЯДОЧИТЬ ПО НомерСтроки
                """
                sel2 = q2.Выполнить().Выбрать()
                k = 0; total_dt50 = 0; total_kt50 = 0
                lines = []
                while sel2.Следующий():
                    k += 1
                    da = acc(sel2.СчетДт); ka = acc(sel2.СчетКт)
                    if da.startswith("50"): total_dt50 += sel2.Сумма or 0
                    if ka.startswith("50"): total_kt50 += sel2.Сумма or 0
                    lines.append(f"      Дт {da:>8} Кт {ka:>8}  Σ={sel2.Сумма:>10.2f}  | {s(sel2.Содержание)[:45]}")
                if k == 0:
                    p(f"      (проводок нет)")
                elif k <= 8:
                    for ln in lines: p(ln)
                else:
                    # Обзор + только сч.50
                    p(f"      [показаны только касса 50.x; всего {k} проводок]")
                    for ln in lines:
                        if " 50" in ln: p(ln)

        except Exception as e:
            p(f"  {name}: ошибка {e}")

    # === Шаг 3: Связь по структуре подчинённости — какие документы на основании одного ОРП ===
    hr("3) Структура подчинённости — что создано на основании ОРП за март")
    q = conn.NewObject("Запрос")
    q.УстановитьПараметр("Д1", datetime.datetime(2026, 3, 20, 0, 0))
    q.УстановитьПараметр("Д2", datetime.datetime(2026, 3, 20, 23, 59, 59))
    q.Текст = """
    ВЫБРАТЬ Ссылка, Номер, СуммаДокумента, Склад
    ИЗ Документ.ОтчетОРозничныхПродажах
    ГДЕ Дата МЕЖДУ &Д1 И &Д2
    """
    sel = q.Выполнить().Выбрать()
    while sel.Следующий():
        ref = sel.Ссылка
        p(f"\n  ОРП #{s(sel.Номер)} ({s(sel.Склад)}, Σ={sel.СуммаДокумента:.2f})")
        # Запрос всех документов где этот ОРП фигурирует как реквизит-ссылка
        # Самый верный способ — пройтись по регистру ДокументыНа... но у БП 3.0
        # типичный регистр сведений «СтруктураПодчинённости» отсутствует, ищем через
        # реквизиты-Документы у потенциальных типов
        for child_name in ["ПриходныйКассовыйОрдер", "СчетФактураВыданныйНаАвансы",
                           "СчетФактураВыданный", "ОперацияБух",
                           "ПоступлениеНаРасчетныйСчет"]:
            try:
                q2 = conn.NewObject("Запрос")
                q2.УстановитьПараметр("ОРП", ref)
                # Перебрать реквизиты и собрать ВЫБОР по совпадению
                md = conn.Метаданные.НайтиПоПолномуИмени(f"Документ.{child_name}")
                if md is None: continue
                # Собираем имена ссылочных реквизитов, которые могут хранить ссылку
                ref_attrs = []
                for rq in md.Реквизиты:
                    ttype = ""
                    try: ttype = str(rq.Тип)
                    except Exception: pass
                    if "ОтчетОРозничныхПродажах" in ttype or "Документ" in rq.Имя.lower() \
                            or "основ" in rq.Имя.lower():
                        ref_attrs.append(rq.Имя)
                if not ref_attrs: continue
                where = " ИЛИ ".join([f"Д.{a} = &ОРП" for a in ref_attrs])
                q2.Текст = f"""
                ВЫБРАТЬ Д.Ссылка, Д.Номер, Д.СуммаДокумента
                ИЗ Документ.{child_name} КАК Д
                ГДЕ ({where})
                """
                sel2 = q2.Выполнить().Выбрать()
                while sel2.Следующий():
                    p(f"      ↳ {child_name} #{s(sel2.Номер)}  Σ={sel2.СуммаДокумента or 0:.2f}")
            except Exception as e:
                pass

if __name__ == "__main__":
    try:
        main()
    finally:
        OUT.close()
        print(f"OK: {OUT_PATH}")
