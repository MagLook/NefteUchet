"""Срез 4 / return_purchase — smoke в БП ГИГ на реальных UUID.

Берёт из БП ГИГ:
  - Организация ГИГ (ИНН 7839440090)
  - Контрагент ТК МЕГАПОЛИС АО (уже сопоставлен с UUID из Срез 1)
  - Договор Мегаполис (уже создан)
  - Склад АЗС Выборг (208) (уже сопоставлен)
  - Любую сопоставленную SKU из загруженного ОРП

Формирует пакет kind=return_purchase, грузит, проверяет ВозвратТоваровПоставщику.
"""
import datetime as dt
import hashlib
import json
import sys
import uuid as uuidlib
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass
import win32com.client

BP_BASE = r"D:\Users\magsp\GIG Base2"
BP_USER = "Гайворонская Татьяна"
BP_PWD = "12345"

OUT_DIR = Path(r"D:\TL_BP_Export")


def get_uuid(conn, тип_перечисления, sku_только=False):
    """Получить UUID источника по типу из TL_СоответствиеИсточников.
    Возвращает первый попавшийся."""
    q = conn.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ ПЕРВЫЕ 1 ИсточникUUID ИЗ РегистрСведений.TL_СоответствиеИсточников "
        "ГДЕ Тип = &Т"
    )
    q.УстановитьПараметр("Т", тип_перечисления)
    v = q.Выполнить().Выбрать()
    return str(v.ИсточникUUID) if v.Следующий() else None


def main():
    conn = win32com.client.Dispatch("V83.COMConnector").Connect(
        f'File="{BP_BASE}";Usr="{BP_USER}";Pwd="{BP_PWD}";'
    )
    ENUM = conn.Перечисления.TL_ТипОбъектаИсточника

    print("[1] Подбор UUID из TL_СоответствиеИсточников...")
    org_uuid = get_uuid(conn, ENUM.Организация)
    contr_uuid = get_uuid(conn, ENUM.Контрагент)
    dog_uuid = get_uuid(conn, ENUM.Договор)
    sklad_uuid = get_uuid(conn, ENUM.Склад)
    nom_uuid = get_uuid(conn, ENUM.Номенклатура)
    print(f"  Орг:       {org_uuid}")
    print(f"  Контр:     {contr_uuid}")
    print(f"  Договор:   {dog_uuid}")
    print(f"  Склад:     {sklad_uuid}")
    print(f"  Номенкл:   {nom_uuid}")
    if not all([org_uuid, contr_uuid, sklad_uuid, nom_uuid]):
        print("  FAIL: не хватает UUID для синтетики")
        return 1

    print("\n[2] Формирование пакета return_purchase...")
    источник = str(uuidlib.uuid4())
    package_id = str(uuidlib.uuid4())
    док = {
        "Тип": "return_purchase",
        "ИсточникUUID": источник,
        "Номер": "СИНТЕТ-RP-1",
        "Дата": dt.datetime.now().isoformat(),
        "Проведен": True,
        "ПометкаУдаления": False,
        "Организация": org_uuid,
        "Контрагент": contr_uuid,
        "ДоговорКонтрагента": dog_uuid,
        "Склад": sklad_uuid,
        "СуммаДокумента": 500.00,
        "Товары": [
            {
                "НомерСтроки": 1,
                "Номенклатура": nom_uuid,
                "Количество": 5.0,
                "Цена": 100.00,
                "Сумма": 500.00,
                "СтавкаНДС": "НДС22",
                "СуммаНДС": 90.16,
            }
        ],
    }
    canonical = json.dumps(док, ensure_ascii=False, sort_keys=True)
    док["Хеш"] = hashlib.sha256(canonical.encode("utf-8")).hexdigest()

    пакет = {
        "ВерсияФормата": "1",
        "ИдентификаторПакета": package_id,
        "Источник": "СИНТЕТ Срез 4 RP",
        "ВремяВыгрузки": dt.datetime.now().isoformat(),
        "НСИ": [],
        "Документы": [док],
    }
    путь = OUT_DIR / f"tl_synth_rp_{package_id}.json"
    with open(путь, "w", encoding="utf-8") as f:
        json.dump(пакет, f, ensure_ascii=False, indent=2)
    print(f"  Сохранён: {путь}")

    print("\n[3] Загрузка в БП ГИГ...")
    read_result = conn.TL_HTTPКлиентЦБ.ПрочитатьПакетИзФайла(str(путь))
    if read_result.Ошибка:
        print(f"  FAIL чтения: {read_result.Ошибка}")
        return 1
    res = conn.TL_СопуткаСервис.ОбработатьПакет(read_result.Пакет, False)  # без проведения для дебага
    print(f"  Создано: {res.Документы.Количество()}, Ошибок: {res.Ошибки.Количество()}")
    for i in range(res.Ошибки.Количество()):
        print(f"    [{i+1}] {res.Ошибки.Получить(i)}")

    print("\n[4] Проверка ВозвратТоваровПоставщику в БП ГИГ...")
    q = conn.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка, СуммаДокумента, Проведен "
        "ИЗ Документ.ВозвратТоваровПоставщику "
        "ГДЕ ВЫРАЗИТЬ(Комментарий КАК СТРОКА(200)) ПОДОБНО &Ш"
    )
    q.УстановитьПараметр("Ш", f"TL|ЦБ|{package_id}|%")
    v = q.Выполнить().Выбрать()
    if not v.Следующий():
        print("  FAIL: документ не найден")
        return 1
    ссс = v.Ссылка
    print(f"  Найден, СуммаДокумента={float(v.СуммаДокумента):.2f}, Проведен={v.Проведен}")

    # Проводки
    print("\n  Проводки Хозрасчётный:")
    q2 = conn.NewObject("Запрос")
    q2.Текст = (
        "ВЫБРАТЬ СчетДт.Код КАК Дт, СчетКт.Код КАК Кт, СУММА(Сумма) КАК С, КОЛИЧЕСТВО(*) КАК К "
        "ИЗ РегистрБухгалтерии.Хозрасчетный ГДЕ Регистратор = &Р "
        "СГРУППИРОВАТЬ ПО СчетДт.Код, СчетКт.Код"
    )
    q2.УстановитьПараметр("Р", ссс)
    v2 = q2.Выполнить().Выбрать()
    проводок = 0
    while v2.Следующий():
        print(f"    Дт {v2.Дт} Кт {v2.Кт}: {int(v2.К)} строк, {float(v2.С):.2f} ₽")
        проводок += 1
    print(f"\n  ИТОГ: {'✓ УСПЕХ' if проводок > 0 and v.Проведен else '⚠ нет проводок'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
