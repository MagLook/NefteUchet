# -*- coding: utf-8 -*-
"""
Мониторинг очереди JSON-пакетов в OwnCloud-каталоге (или D:\\TL_BP_Export\\).

Используется для отладки канала ЦБ → БП ГИГ:
- Показать список пакетов
- По каждому — извлечь шапку (Тип, ИсточникUUID, Хеш, ВремяВыгрузки, Документов)
- Проверить статус: загружен в БП ГИГ? (по TL_СоответствиеИсточников)

Запуск:
  py -3-32 D:\\Users\\magsp\\ELSYPLUS\\NefteUchet\\scripts\\check_owncloud_queue.py
  py -3-32 ... --dir D:\\Users\\magsp\\owncloud\\gig_packages
  py -3-32 ... --only-pending  (только незагруженные)
"""
import argparse
import datetime as dt
import json
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

try:
    import win32com.client
except ImportError:
    win32com = None

DEFAULT_DIR = Path(r"D:\TL_BP_Export")

BP_BASE = r"D:\Users\magsp\GIG Base2"
BP_USER = "Гайворонская Татьяна"
BP_PWD = "12345"


def собрать_пакеты(каталог):
    """Список JSON-пакетов в каталоге с краткой статистикой."""
    if not каталог.exists():
        print(f"  Каталог {каталог} не существует")
        return []
    пакеты = []
    for файл in sorted(каталог.glob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True):
        try:
            with open(файл, encoding="utf-8-sig") as f:
                pkg = json.load(f)
        except Exception as e:
            пакеты.append({"файл": файл, "ошибка": str(e)})
            continue
        документы = pkg.get("Документы", [])
        # Извлечь основные поля
        пакеты.append({
            "файл": файл,
            "ИдентификаторПакета": pkg.get("ИдентификаторПакета", "?"),
            "Источник": pkg.get("Источник", "?"),
            "ВремяВыгрузки": pkg.get("ВремяВыгрузки", "?"),
            "ВерсияФормата": pkg.get("ВерсияФормата", "?"),
            "ДокументовВ Пакете": len(документы),
            "Документы": [
                {
                    "Тип": d.get("Тип"),
                    "ИсточникUUID": d.get("ИсточникUUID", "?"),
                    "Номер": d.get("Номер", "?"),
                    "Хеш": d.get("Хеш", "?"),
                }
                for d in документы
            ],
            "размер": файл.stat().st_size,
            "изменён": dt.datetime.fromtimestamp(файл.stat().st_mtime).isoformat(),
        })
    return пакеты


def проверить_статус_в_бп(uuids_docs):
    """Проверить какие UUID документов уже загружены в БП ГИГ (TL_СоответствиеИсточников)."""
    if win32com is None or not uuids_docs:
        return {}
    try:
        connector = win32com.client.Dispatch("V83.COMConnector")
        cs = f'File="{BP_BASE}";Usr="{BP_USER}";Pwd="{BP_PWD}";'
        conn = connector.Connect(cs)
    except Exception as e:
        print(f"  [WARN] Не удалось подключиться к БП ГИГ: {e}")
        return {}

    q = conn.NewObject("Запрос")
    список_uuid = list(set(uuids_docs))
    плэйсхолдер = ",".join([f'"{u}"' for u in список_uuid])
    q.Текст = (
        f"ВЫБРАТЬ ИсточникUUID, ХешПоследнегоПакета, ДатаПоследнейЗагрузки "
        f"ИЗ РегистрСведений.TL_СоответствиеИсточников "
        f"ГДЕ ИсточникUUID В ({плэйсхолдер})"
    )
    try:
        v = q.Выполнить().Выбрать()
    except Exception as e:
        print(f"  [WARN] Запрос статуса упал: {e}")
        return {}
    статус = {}
    while v.Следующий():
        статус[str(v.ИсточникUUID)] = {
            "хеш": str(v.ХешПоследнегоПакета),
            "дата": str(v.ДатаПоследнейЗагрузки),
        }
    return статус


def main():
    парсер = argparse.ArgumentParser()
    парсер.add_argument("--dir", default=str(DEFAULT_DIR), help="Каталог с пакетами")
    парсер.add_argument("--only-pending", action="store_true",
                        help="Только незагруженные в БП ГИГ")
    парсер.add_argument("--no-bp", action="store_true",
                        help="Не подключаться к БП ГИГ (быстрее, без проверки статуса)")
    args = парсер.parse_args()

    каталог = Path(args.dir)
    print(f"Каталог: {каталог}")
    print(f"Время:   {dt.datetime.now():%Y-%m-%d %H:%M:%S}")

    пакеты = собрать_пакеты(каталог)
    if not пакеты:
        print("  Пусто.")
        return 0

    print(f"\nВсего пакетов: {len(пакеты)}")

    # Собрать все UUID документов для одной проверки в БП
    все_uuid = []
    for p in пакеты:
        if "ошибка" in p:
            continue
        for d in p["Документы"]:
            uuid = d.get("ИсточникUUID")
            if uuid and uuid != "?":
                все_uuid.append(uuid)

    статус = {}
    if not args.no_bp and все_uuid:
        print(f"\nПроверка статуса в БП ГИГ (TL_СоответствиеИсточников)...")
        статус = проверить_статус_в_бп(все_uuid)
        print(f"  Найдено {len(статус)} из {len(set(все_uuid))} уникальных UUID")

    # Вывод
    print(f"\n{'='*78}")
    for p in пакеты:
        if "ошибка" in p:
            print(f"  [ERR] {p['файл'].name}: {p['ошибка']}")
            continue

        # Определить статус всех документов пакета
        статусы_док = []
        for d in p["Документы"]:
            uuid = d.get("ИсточникUUID")
            if uuid in статус:
                статусы_док.append("loaded")
            else:
                статусы_док.append("pending")

        общий_статус = "loaded" if all(s == "loaded" for s in статусы_док) else "pending"

        if args.only_pending and общий_статус == "loaded":
            continue

        emoji = "✓" if общий_статус == "loaded" else "○"
        print(f"\n  {emoji} {p['файл'].name}")
        print(f"     Время выгрузки: {p['ВремяВыгрузки']}")
        print(f"     Источник:       {p['Источник']}")
        print(f"     Документов:     {p['ДокументовВ Пакете']}")
        for i, d in enumerate(p["Документы"]):
            uuid = d.get("ИсточникUUID", "?")
            статус_док = "loaded" if uuid in статус else "pending"
            print(f"       [{i+1}] {d['Тип']:<25} {статус_док:<8} (UUID {uuid[:8]}...)")
            if uuid in статус and d.get("Хеш") and d["Хеш"] != "?":
                совпадение_хеш = d["Хеш"] == статус[uuid]["хеш"]
                ярлык = "same" if совпадение_хеш else "DIFF (модификация!)"
                print(f"           Хеш: {ярлык}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
