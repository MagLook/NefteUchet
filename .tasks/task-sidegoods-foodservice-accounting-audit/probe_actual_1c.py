# -*- coding: utf-8 -*-
import ast
import datetime
import json
import pathlib
import sys
import traceback

import win32com.client


ROOT = pathlib.Path(r"D:\Users\magsp\ELSYPLUS\Ledger")
TASK = ROOT / "ext-bp" / ".tasks" / "task-sidegoods-foodservice-accounting-audit"
CONNECTION_SOURCE = ROOT / "ext-cb-export" / "scripts" / "_tl_pipeline_com.py"
START = datetime.datetime(2026, 6, 11)
END = datetime.datetime(2026, 8, 10)


def connection_string(name):
    tree = ast.parse(CONNECTION_SOURCE.read_text(encoding="utf-8-sig"))
    for node in tree.body:
        if not isinstance(node, ast.Assign) or len(node.targets) != 1:
            continue
        target = node.targets[0]
        if isinstance(target, ast.Name) and target.id == name:
            return ast.literal_eval(node.value)
    raise RuntimeError(f"В {CONNECTION_SOURCE} не найден {name}")


def as_text(conn, value):
    if value is None:
        return ""
    try:
        return str(conn.String(value)).strip()
    except Exception:
        return str(value).strip()


def as_int(value):
    try:
        return int(value or 0)
    except Exception:
        return 0


def as_float(value):
    try:
        return float(value or 0)
    except Exception:
        return 0.0


def query(conn, text, params=None):
    request = conn.NewObject("Запрос", text)
    for key, value in (params or {}).items():
        request.УстановитьПараметр(key, value)
    return request.Выполнить().Выбрать()


def one_row(conn, text, params=None):
    selection = query(conn, text, params)
    if not selection.Следующий():
        return None
    return selection


def metadata_exists(conn, collection, name):
    try:
        return getattr(conn.Метаданные, collection).Найти(name) is not None
    except Exception:
        return False


def safe_section(result, name, callback):
    try:
        result[name] = callback()
    except Exception as error:
        result[name] = {"error": str(error), "trace": traceback.format_exc().splitlines()[-1]}


def bp_doc_summary(conn, doc_name):
    if not metadata_exists(conn, "Документы", doc_name):
        return {"exists": False}
    row = one_row(conn, f"""
        ВЫБРАТЬ
            КОЛИЧЕСТВО(*) КАК Всего,
            СУММА(ВЫБОР КОГДА Д.Проведен ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Проведено,
            СУММА(ВЫБОР КОГДА НЕ Д.Проведен И НЕ Д.ПометкаУдаления ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Черновики,
            СУММА(ВЫБОР КОГДА Д.ПометкаУдаления ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Помечено,
            МИНИМУМ(Д.Дата) КАК ПерваяДата,
            МАКСИМУМ(Д.Дата) КАК ПоследняяДата
        ИЗ Документ.{doc_name} КАК Д
        ГДЕ Д.Дата >= &Начало И Д.Дата < &Конец
            И Д.Комментарий ПОДОБНО "%TL|%"
            И Д.Комментарий ПОДОБНО "%АЗС208%"
    """, {"Начало": START, "Конец": END})
    return {
        "exists": True,
        "total": as_int(row.Всего),
        "posted": as_int(row.Проведено),
        "draft": as_int(row.Черновики),
        "marked": as_int(row.Помечено),
        "first_date": as_text(conn, row.ПерваяДата),
        "last_date": as_text(conn, row.ПоследняяДата),
    }


def bp_postings(conn, doc_name):
    if not metadata_exists(conn, "Документы", doc_name):
        return []
    selection = query(conn, f"""
        ВЫБРАТЬ
            Х.СчетДт КАК Дт,
            Х.СчетКт КАК Кт,
            КОЛИЧЕСТВО(*) КАК Движений,
            СУММА(Х.Сумма) КАК Сумма
        ИЗ РегистрБухгалтерии.Хозрасчетный КАК Х
        ГДЕ Х.Период >= &Начало И Х.Период < &Конец
            И Х.Регистратор ССЫЛКА Документ.{doc_name}
            И ВЫРАЗИТЬ(Х.Регистратор КАК Документ.{doc_name}).Комментарий ПОДОБНО "%TL|%"
            И ВЫРАЗИТЬ(Х.Регистратор КАК Документ.{doc_name}).Комментарий ПОДОБНО "%АЗС208%"
        СГРУППИРОВАТЬ ПО Х.СчетДт, Х.СчетКт
    """, {"Начало": START, "Конец": END})
    rows = []
    while selection.Следующий():
        rows.append({
            "debit": as_text(conn, selection.Дт),
            "credit": as_text(conn, selection.Кт),
            "entries": as_int(selection.Движений),
            "amount": round(as_float(selection.Сумма), 2),
        })
    return rows


def bp_latest_docs(conn, doc_name):
    if not metadata_exists(conn, "Документы", doc_name):
        return []
    selection = query(conn, f"""
        ВЫБРАТЬ ПЕРВЫЕ 5
            ПРЕДСТАВЛЕНИЕ(Д.Ссылка) КАК Документ,
            Д.Дата КАК Дата,
            Д.Проведен КАК Проведен,
            Д.ПометкаУдаления КАК ПометкаУдаления
        ИЗ Документ.{doc_name} КАК Д
        ГДЕ Д.Дата >= &Начало И Д.Дата < &Конец
            И Д.Комментарий ПОДОБНО "%TL|%"
            И Д.Комментарий ПОДОБНО "%АЗС208%"
        УПОРЯДОЧИТЬ ПО Д.Дата УБЫВ
    """, {"Начало": START, "Конец": END})
    rows = []
    while selection.Следующий():
        rows.append({
            "document": as_text(conn, selection.Документ),
            "date": as_text(conn, selection.Дата),
            "posted": bool(selection.Проведен),
            "marked": bool(selection.ПометкаУдаления),
        })
    return rows


def bp_extensions(conn):
    rows = []
    extensions = conn.РасширенияКонфигурации.Получить()
    for index in range(extensions.Количество()):
        extension = extensions.Получить(index)
        row = {}
        for source, target in (
            ("Имя", "name"),
            ("Версия", "version"),
            ("Активно", "active"),
            ("БезопасныйРежим", "safe_mode"),
            ("ИмяФайла", "file_name"),
            ("UID", "uid"),
            ("ХешСумма", "hash"),
        ):
            try:
                value = getattr(extension, source)
                row[target] = value if isinstance(value, (bool, int, float)) else as_text(conn, value)
            except Exception:
                pass
        rows.append(row)
        if row.get("name") == "TradeLedger":
            installed = TASK / "installed_TradeLedger.cfe"
            extension.ПолучитьДанные().Записать(str(installed))
            row["dump"] = str(installed)
    return rows


def bp_network_returns(conn):
    result = {}
    for doc_name in ("КорректировкаПоступления", "ВозвратТоваровПоставщику"):
        if not metadata_exists(conn, "Документы", doc_name):
            result[doc_name] = {"exists": False}
            continue
        row = one_row(conn, f"""
            ВЫБРАТЬ
                КОЛИЧЕСТВО(*) КАК Всего,
                СУММА(ВЫБОР КОГДА Д.Проведен ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Проведено,
                СУММА(ВЫБОР КОГДА НЕ Д.Проведен И НЕ Д.ПометкаУдаления ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Черновики,
                МАКСИМУМ(Д.Дата) КАК ПоследняяДата
            ИЗ Документ.{doc_name} КАК Д
            ГДЕ Д.Дата >= &Начало И Д.Дата < &Конец
                И Д.Комментарий ПОДОБНО "%TL|ЦБ|%"
        """, {"Начало": datetime.datetime(2026, 1, 1), "Конец": END})
        result[doc_name] = {
            "exists": True,
            "total": as_int(row.Всего),
            "posted": as_int(row.Проведено),
            "draft": as_int(row.Черновики),
            "last_date": as_text(conn, row.ПоследняяДата),
        }
    return result


def bp_settings(conn):
    keys = {
        "foodservice_model": ("Общепит_МодельУчёта", "B"),
        "foodservice_completion": ("Общепит_КомплектацияБлюд", "1"),
        "separate_accounting": ("ОбщепитВключён", "0"),
        "force_vat_22": ("Розница_ФорсНДС22Аварийный", "0"),
        "protect_posted": ("Защита_НеПерезаписыватьПроведённые", "1"),
    }
    return {
        name: as_text(conn, conn.TL_Настройки.ПолучитьЗначение(key, default))
        for name, (key, default) in keys.items()
    }


def bp_negative_4102(conn):
    account = conn.ПланыСчетов.Хозрасчетный.НайтиПоКоду("41.02")
    warehouses = conn.NewObject("Массив")
    selection = query(conn, """
        ВЫБРАТЬ С.Ссылка КАК Ссылка
        ИЗ Справочник.Склады КАК С
        ГДЕ С.Наименование ПОДОБНО "%208%" И НЕ С.ПометкаУдаления
    """)
    while selection.Следующий():
        warehouses.Добавить(selection.Ссылка)
    selection = query(conn, """
        ВЫБРАТЬ
            Ост.Субконто1 КАК Номенклатура,
            Ост.КоличествоОстаток КАК Количество,
            Ост.СуммаОстаток КАК Сумма
        ИЗ РегистрБухгалтерии.Хозрасчетный.Остатки(, Счет = &Счет, , Субконто3 В (&Склады)) КАК Ост
        ГДЕ Ост.КоличествоОстаток < 0 ИЛИ Ост.СуммаОстаток < 0
    """, {"Счет": account, "Склады": warehouses})
    rows = []
    while selection.Следующий():
        rows.append({
            "item": as_text(conn, selection.Номенклатура),
            "qty": round(as_float(selection.Количество), 6),
            "amount": round(as_float(selection.Сумма), 2),
        })
    rows.sort(key=lambda row: (row["amount"], row["qty"]))
    return {"count": len(rows), "total_amount": round(sum(row["amount"] for row in rows), 2), "items": rows[:30]}


def run_bp():
    print("Подключение к БП ГИГ; это один read-only сеанс и он может занять около 6 минут...", flush=True)
    conn = win32com.client.Dispatch("V83.COMConnector").Connect(connection_string("CS_BP"))
    result = {
        "database": "BP GIG production",
        "captured_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "period": {"from": START.isoformat(), "to": END.isoformat()},
    }
    safe_section(result, "extensions", lambda: bp_extensions(conn))
    safe_section(result, "settings", lambda: bp_settings(conn))
    docs = {}
    postings = {}
    latest = {}
    for doc_name in (
        "ПоступлениеТоваровУслуг",
        "КорректировкаПоступления",
        "ВозвратТоваровПоставщику",
        "ОтчетОРозничныхПродажах",
        "КомплектацияНоменклатуры",
        "ОтчетПроизводстваЗаСмену",
        "СписаниеТоваров",
        "ПриходныйКассовыйОрдер",
    ):
        try:
            docs[doc_name] = bp_doc_summary(conn, doc_name)
        except Exception as error:
            docs[doc_name] = {"error": str(error), "trace": traceback.format_exc().splitlines()[-1]}
        try:
            latest[doc_name] = bp_latest_docs(conn, doc_name)
        except Exception as error:
            latest[doc_name] = {"error": str(error)}
        try:
            postings[doc_name] = bp_postings(conn, doc_name)
        except Exception as error:
            postings[doc_name] = {"error": str(error), "trace": traceback.format_exc().splitlines()[-1]}
    result["documents_station_208_tl"] = docs
    result["latest_station_208_tl"] = latest
    result["postings_station_208_tl"] = postings
    safe_section(result, "returns_all_stations_tl", lambda: bp_network_returns(conn))
    safe_section(result, "negative_41_02_station_208", lambda: bp_negative_4102(conn))
    return result


def cb_doc_summary(conn, doc_name):
    if not metadata_exists(conn, "Документы", doc_name):
        return {"exists": False}
    row = one_row(conn, f"""
        ВЫБРАТЬ
            КОЛИЧЕСТВО(*) КАК Всего,
            СУММА(ВЫБОР КОГДА Д.Проведен ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Проведено,
            СУММА(ВЫБОР КОГДА НЕ Д.Проведен И НЕ Д.ПометкаУдаления ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Черновики,
            СУММА(ВЫБОР КОГДА Д.ПометкаУдаления ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Помечено,
            МИНИМУМ(Д.Дата) КАК ПерваяДата,
            МАКСИМУМ(Д.Дата) КАК ПоследняяДата
        ИЗ Документ.{doc_name} КАК Д
        ГДЕ Д.Дата >= &Начало И Д.Дата < &Конец
            И Д.Склад.Наименование ПОДОБНО "%208%"
    """, {"Начало": START, "Конец": END})
    return {
        "exists": True,
        "total": as_int(row.Всего),
        "posted": as_int(row.Проведено),
        "draft": as_int(row.Черновики),
        "marked": as_int(row.Помечено),
        "first_date": as_text(conn, row.ПерваяДата),
        "last_date": as_text(conn, row.ПоследняяДата),
    }


def cb_line_summary(conn, doc_name, table_name, extra_fields=""):
    if not metadata_exists(conn, "Документы", doc_name):
        return {"exists": False}
    fields = "," + extra_fields if extra_fields else ""
    selection = query(conn, f"""
        ВЫБРАТЬ
            КОЛИЧЕСТВО(*) КАК Строк,
            КОЛИЧЕСТВО(РАЗЛИЧНЫЕ Т.Ссылка) КАК Документов
            {fields}
        ИЗ Документ.{doc_name}.{table_name} КАК Т
        ГДЕ Т.Ссылка.Дата >= &Начало И Т.Ссылка.Дата < &Конец
            И Т.Ссылка.Склад.Наименование ПОДОБНО "%208%"
            И НЕ Т.Ссылка.ПометкаУдаления
    """, {"Начало": START, "Конец": END})
    if not selection.Следующий():
        return {"rows": 0, "documents": 0}
    result = {"rows": as_int(selection.Строк), "documents": as_int(selection.Документов)}
    if extra_fields:
        for name in ("Количество", "Сумма"):
            try:
                result[name.lower()] = round(as_float(getattr(selection, name)), 6 if name == "Количество" else 2)
            except Exception:
                pass
    return result


def cb_dish_coverage(conn):
    sales = query(conn, """
        ВЫБРАТЬ
            Т.Номенклатура КАК Номенклатура,
            Т.Номенклатура.ВидНоменклатуры КАК ВидНоменклатуры,
            СУММА(Т.Количество) КАК Количество
        ИЗ Документ.ОтчетОРозничныхПродажах.Товары КАК Т
        ГДЕ Т.Ссылка.Дата >= &Начало И Т.Ссылка.Дата < &Конец
            И Т.Ссылка.Склад.Наименование ПОДОБНО "%208%"
            И НЕ Т.Ссылка.ПометкаУдаления
        СГРУППИРОВАТЬ ПО Т.Номенклатура, Т.Номенклатура.ВидНоменклатуры
    """, {"Начало": START, "Конец": END})
    sold = {}
    while sales.Следующий():
        if as_text(conn, sales.ВидНоменклатуры) != "Набор - комплект":
            continue
        key = as_text(conn, sales.Номенклатура)
        sold[key] = {"dish": key, "sold_qty": round(as_float(sales.Количество), 6)}

    releases = query(conn, """
        ВЫБРАТЬ
            Т.Номенклатура КАК Номенклатура,
            СУММА(Т.Количество) КАК Количество,
            СУММА(ВЫБОР КОГДА Т.ТТК = ЗНАЧЕНИЕ(Документ.ТТК.ПустаяСсылка) ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК СтрокБезТТК
        ИЗ Документ.ВыпускПродукции.Товары КАК Т
        ГДЕ Т.Ссылка.Дата >= &Начало И Т.Ссылка.Дата < &Конец
            И Т.Ссылка.Склад.Наименование ПОДОБНО "%208%"
            И НЕ Т.Ссылка.ПометкаУдаления
        СГРУППИРОВАТЬ ПО Т.Номенклатура
    """, {"Начало": START, "Конец": END})
    released = {}
    while releases.Следующий():
        key = as_text(conn, releases.Номенклатура)
        released[key] = {
            "dish": key,
            "released_qty": round(as_float(releases.Количество), 6),
            "rows_without_recipe": as_int(releases.СтрокБезТТК),
        }
    missing_release = [value for key, value in sold.items() if key not in released]
    missing_recipe = [value for value in released.values() if value["rows_without_recipe"] > 0]
    return {
        "sold_dishes": len(sold),
        "released_dishes": len(released),
        "sold_without_release": missing_release,
        "release_rows_without_recipe": missing_recipe,
    }


def cb_purchase_skips(conn):
    result = {}
    for table_name in ("Услуги", "ВозвратнаяТара"):
        try:
            row = one_row(conn, f"""
                ВЫБРАТЬ
                    КОЛИЧЕСТВО(*) КАК Строк,
                    КОЛИЧЕСТВО(РАЗЛИЧНЫЕ Т.Ссылка) КАК Документов
                ИЗ Документ.ПоступлениеТоваровУслугНаАЗК.{table_name} КАК Т
                ГДЕ Т.Ссылка.Дата >= &Начало И Т.Ссылка.Дата < &Конец
                    И Т.Ссылка.Склад.Наименование ПОДОБНО "%208%"
                    И НЕ Т.Ссылка.ПометкаУдаления
            """, {"Начало": START, "Конец": END})
            result[table_name] = {"rows": as_int(row.Строк), "documents": as_int(row.Документов)}
        except Exception as error:
            result[table_name] = {"error": str(error)}
    return result


def cb_network_returns(conn):
    result = {}
    for doc_name in ("ВозвратТоваровПоставщикуНаАЗК", "ВозвратТоваровПоставщику"):
        if not metadata_exists(conn, "Документы", doc_name):
            result[doc_name] = {"exists": False}
            continue
        row = one_row(conn, f"""
            ВЫБРАТЬ
                КОЛИЧЕСТВО(*) КАК Всего,
                СУММА(ВЫБОР КОГДА Д.Проведен ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Проведено,
                МАКСИМУМ(Д.Дата) КАК ПоследняяДата
            ИЗ Документ.{doc_name} КАК Д
            ГДЕ Д.Дата >= &Начало И Д.Дата < &Конец
                И НЕ Д.ПометкаУдаления
        """, {"Начало": datetime.datetime(2026, 1, 1), "Конец": END})
        result[doc_name] = {
            "exists": True,
            "total": as_int(row.Всего),
            "posted": as_int(row.Проведено),
            "last_date": as_text(conn, row.ПоследняяДата),
        }
    return result


def run_cb():
    print("Подключение к центральной базе; read-only...", flush=True)
    conn = win32com.client.Dispatch("V83.COMConnector").Connect(connection_string("CS_CB"))
    result = {
        "database": "Elsy AZK central production",
        "captured_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "period": {"from": START.isoformat(), "to": END.isoformat()},
    }
    docs = {}
    for doc_name in (
        "ПоступлениеТоваровУслугНаАЗК",
        "ВозвратТоваровПоставщикуНаАЗК",
        "ВозвратТоваровПоставщику",
        "ОтчетОРозничныхПродажах",
        "ВыпускПродукции",
        "ТТК",
    ):
        try:
            docs[doc_name] = cb_doc_summary(conn, doc_name)
        except Exception as error:
            docs[doc_name] = {"error": str(error), "trace": traceback.format_exc().splitlines()[-1]}
    result["documents_station_208"] = docs
    safe_section(result, "purchase_goods", lambda: cb_line_summary(
        conn, "ПоступлениеТоваровУслугНаАЗК", "Товары", "СУММА(Т.Количество) КАК Количество, СУММА(Т.Сумма) КАК Сумма"))
    safe_section(result, "retail_goods", lambda: cb_line_summary(
        conn, "ОтчетОРозничныхПродажах", "Товары", "СУММА(Т.Количество) КАК Количество, СУММА(Т.Сумма) КАК Сумма"))
    safe_section(result, "retail_returns", lambda: cb_line_summary(
        conn, "ОтчетОРозничныхПродажах", "ВозвращенныеТовары", "СУММА(Т.Количество) КАК Количество, СУММА(Т.Сумма) КАК Сумма"))
    safe_section(result, "purchase_documents_skipped_by_exporter", lambda: cb_purchase_skips(conn))
    safe_section(result, "returns_all_stations", lambda: cb_network_returns(conn))
    safe_section(result, "foodservice_dish_coverage", lambda: cb_dish_coverage(conn))
    return result


def main():
    if len(sys.argv) != 2 or sys.argv[1] not in ("cb", "bp"):
        raise SystemExit("Использование: py -3.13-32 probe_actual_1c.py cb|bp")
    mode = sys.argv[1]
    result = run_cb() if mode == "cb" else run_bp()
    output = TASK / f"actual_{mode}_audit.json"
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(f"Готово: {output}", flush=True)


if __name__ == "__main__":
    main()
