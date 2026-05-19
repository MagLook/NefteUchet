# -*- coding: utf-8 -*-
"""
Probe: исследование структуры данных от АЗС 208 в ЦБ ЭЛСИ.АЗК в разрезе смен.

Цель — понять как организованы:
  1. Сменные отчёты (документы смены)
  2. Документы реализации (ОРП, ОтчётКассы) — как привязаны к смене
  3. Документы поступления (ПТУ) — как привязаны к смене
  4. Документы общепита (выпуск, списание сырья)

Подключение — к локальной TsB_Live_20260429 (там РИБ-копия данных с 208).

Запуск:
    py -3-32 D:\\Users\\magsp\\ELSYPLUS\\NefteUchet\\scripts\\probe_smen_struktura.py
"""
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")


def connect_tsb():
    here = os.path.dirname(os.path.abspath(__file__))
    secrets_path = os.path.join(
        here, "..", "..", "ElsyPlusMSN_TsB", "scripts", "secrets.json"
    )
    with open(os.path.abspath(secrets_path), encoding="utf-8") as f:
        s = json.load(f)
    cs = f'File="{s["tsb_base"]}";Usr="{s["tsb_user"]}";Pwd="{s["tsb_pwd"]}";'
    import win32com.client
    return win32com.client.Dispatch("V83.COMConnector").Connect(cs)


def section(title):
    print()
    print("=" * 78)
    print(f"  {title}")
    print("=" * 78)


def list_typed_documents(conn, pattern_list):
    """Найти типы документов в Метаданные.Документы по паттернам в имени."""
    found = []
    for d in conn.Метаданные.Документы:
        name = str(d.Имя)
        if any(p.lower() in name.lower() for p in pattern_list):
            found.append(name)
    return found


def doc_structure(conn, doc_name):
    """Вернуть структуру документа: реквизиты + табличные части."""
    try:
        meta = getattr(conn.Метаданные.Документы, doc_name)
    except Exception as e:
        return {"error": str(e)}
    info = {"name": doc_name, "Реквизиты": [], "ТабличныеЧасти": {}}
    try:
        for r in meta.Реквизиты:
            info["Реквизиты"].append({
                "имя": str(r.Имя),
                "тип": str(r.Тип),
            })
    except Exception:
        pass
    try:
        for tc in meta.ТабличныеЧасти:
            cols = []
            for col in tc.Реквизиты:
                cols.append(str(col.Имя))
            info["ТабличныеЧасти"][str(tc.Имя)] = cols
    except Exception:
        pass
    return info


def count_documents(conn, doc_name, period_from, period_to, kod_azs=None):
    """Посчитать документы за период (опционально фильтр по станции)."""
    q_text = (
        f"ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол "
        f"ИЗ Документ.{doc_name} КАК Д "
        f"ГДЕ Д.Дата МЕЖДУ &ДатаС И &ДатаПо "
    )
    if kod_azs:
        q_text += " И ВЫРАЗИТЬ(Д.Склад.Код КАК ЧИСЛО(10)) = &КодАЗС "
    q = conn.NewObject("Запрос", q_text)
    q.УстановитьПараметр("ДатаС", period_from)
    q.УстановитьПараметр("ДатаПо", period_to)
    if kod_azs:
        q.УстановитьПараметр("КодАЗС", kod_azs)
    try:
        sel = q.Выполнить().Выбрать()
        if sel.Следующий():
            return int(sel.Кол)
    except Exception as e:
        return f"ERROR: {e}"
    return 0


def sample_documents(conn, doc_name, period_from, period_to, limit=5):
    """Примеры документов за период — Номер, Дата, Склад, Сумма."""
    q_text = (
        f"ВЫБРАТЬ ПЕРВЫЕ {limit} "
        f"Д.Номер КАК Номер, Д.Дата КАК Дата, "
        f"Д.Склад.Наименование КАК СкладИмя, "
        f"Д.Склад.Код КАК СкладКод, "
        f"Д.СуммаДокумента КАК Сумма, "
        f"Д.Проведен КАК Проведен "
        f"ИЗ Документ.{doc_name} КАК Д "
        f"ГДЕ Д.Дата МЕЖДУ &ДатаС И &ДатаПо "
        f"УПОРЯДОЧИТЬ ПО Д.Дата УБЫВ"
    )
    q = conn.NewObject("Запрос", q_text)
    q.УстановитьПараметр("ДатаС", period_from)
    q.УстановитьПараметр("ДатаПо", period_to)
    rows = []
    try:
        sel = q.Выполнить().Выбрать()
        while sel.Следующий():
            rows.append({
                "Номер": str(sel.Номер).strip(),
                "Дата": str(sel.Дата),
                "СкладКод": str(sel.СкладКод),
                "Склад": str(sel.СкладИмя),
                "Сумма": float(sel.Сумма) if sel.Сумма else 0,
                "Проведен": bool(sel.Проведен),
            })
    except Exception as e:
        rows.append({"error": str(e)})
    return rows


def main():
    print("Подключение к ЦБ ЭЛСИ.АЗК (TsB_Live_20260429)...")
    conn = connect_tsb()
    print("OK")

    # Период: апрель-май 2026 на 208
    date_from = conn.NewObject("XMLReader")  # placeholder
    # У 1С даты передаются как python datetime
    from datetime import datetime
    df = datetime(2026, 4, 28)
    dt = datetime(2026, 5, 2, 23, 59, 59)
    kod_208 = 208

    section("1. Поиск типов документов 'смены / отчёты' в метаданных")
    smen_types = list_typed_documents(
        conn,
        ["Смен", "ОтчётС", "ОтчетС", "ОтчётКасс", "ОтчетКасс", "ТМЦАЗС"]
    )
    for t in smen_types:
        print(f"  → {t}")

    section("2. Поиск документов реализации")
    sale_types = list_typed_documents(
        conn,
        ["Реализ", "ОтчетОРозн", "ОтчётОРозн", "ОтчётККМ", "ПродажТов",
         "ЧекККМ", "Розн"]
    )
    for t in sale_types:
        print(f"  → {t}")

    section("3. Поиск документов поступления")
    purch_types = list_typed_documents(
        conn,
        ["Поступление", "ПриходТМЦ", "Приходная", "ПТУ"]
    )
    for t in purch_types:
        print(f"  → {t}")

    section("4. Поиск документов общепита")
    food_types = list_typed_documents(
        conn,
        ["Выпуск", "Производст", "ТТК", "Калькуляц", "ТребованиеНаклад"]
    )
    for t in food_types:
        print(f"  → {t}")

    section("5. Структура ТОП документов смен")
    for t in smen_types[:5]:
        print()
        print(f"--- {t} ---")
        info = doc_structure(conn, t)
        for r in info.get("Реквизиты", [])[:30]:
            print(f"    {r['имя']:30s} : {r['тип']}")
        for tc_name, cols in info.get("ТабличныеЧасти", {}).items():
            print(f"    ТЧ {tc_name}: {', '.join(cols[:10])}")

    section("6. Счётчики документов на АЗС 208 за период 28.04 - 02.05.2026")
    for t in smen_types[:3] + sale_types[:3] + purch_types[:3] + food_types[:3]:
        c = count_documents(conn, t, df, dt, kod_208)
        print(f"  {t:40s} → {c}")

    section("7. Примеры документов смен (за апрель 2026)")
    for t in smen_types[:3]:
        print()
        print(f"--- {t} ---")
        for row in sample_documents(conn, t, df, dt, limit=3):
            print(f"    {row}")

    section("8. Примеры ОРП/реализаций")
    for t in sale_types[:3]:
        print()
        print(f"--- {t} ---")
        for row in sample_documents(conn, t, df, dt, limit=3):
            print(f"    {row}")

    section("9. Примеры поступлений")
    for t in purch_types[:3]:
        print()
        print(f"--- {t} ---")
        for row in sample_documents(conn, t, df, dt, limit=3):
            print(f"    {row}")

    print()
    print("=" * 78)
    print("Готово. Используйте выводы для проектирования packet_per_shift.")
    print("=" * 78)


if __name__ == "__main__":
    main()
