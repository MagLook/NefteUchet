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
    fallback_nom_uuid = get_uuid(conn, ENUM.Номенклатура)  # любая сопоставленная SKU
    # ПервичнаяПТУ — берём через JOIN с реальной ПТУ
    q = conn.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ ПЕРВЫЕ 1 С.ИсточникUUID, С.СсылкаОбъект "
        "ИЗ РегистрСведений.TL_СоответствиеИсточников КАК С "
        "  ВНУТРЕННЕЕ СОЕДИНЕНИЕ Документ.ПоступлениеТоваровУслуг КАК П "
        "  ПО С.СсылкаОбъект = П.Ссылка "
        "ГДЕ С.Тип = &Т"
    )
    q.УстановитьПараметр("Т", ENUM.Поступление)
    тз = q.Выполнить().Выгрузить()
    исх_пту_ссылка = None
    исх_пту_uuid = None
    if тз.Количество() > 0:
        row = тз.Получить(0)
        исх_пту_uuid = str(row.ИсточникUUID)
        исх_пту_ссылка = row.СсылкаОбъект

    # Возьмём первую сопоставленную SKU и её строку в исходной ПТУ
    nom_uuid = None
    qty_orig = price_orig = sum_orig = vat_orig = 0
    vat_name = "БезНДС"
    nom_ссылка = None
    if исх_пту_ссылка:
        # Найдём все строки ПТУ + проверим какие SKU сопоставлены
        q.Текст = (
            "ВЫБРАТЬ Т.Номенклатура КАК Н, Т.Количество КАК Кол, "
            "  Т.Цена КАК Цн, Т.Сумма КАК Сум, Т.СтавкаНДС КАК Нд, Т.СуммаНДС КАК СНд "
            "ИЗ Документ.ПоступлениеТоваровУслуг.Товары КАК Т ГДЕ Т.Ссылка = &С"
        )
        q.УстановитьПараметр("С", исх_пту_ссылка)
        тз = q.Выполнить().Выгрузить()
        # Берём UUID->Ссылка из TL_СоответствиеИсточников
        q2 = conn.NewObject("Запрос")
        q2.Текст = (
            "ВЫБРАТЬ ИсточникUUID, СсылкаОбъект ИЗ РегистрСведений.TL_СоответствиеИсточников "
            "ГДЕ Тип = &Т"
        )
        q2.УстановитьПараметр("Т", ENUM.Номенклатура)
        тз_ном = q2.Выполнить().Выгрузить()
        # Карта Ссылка→UUID
        ссылка_к_uuid = {}
        for j in range(тз_ном.Количество()):
            r = тз_ном.Получить(j)
            ссылка_к_uuid[str(conn.XMLСтрока(r.СсылкаОбъект))] = str(r.ИсточникUUID)
        # Перебираем строки ПТУ
        for i in range(тз.Количество()):
            row = тз.Получить(i)
            ном_key = str(conn.XMLСтрока(row.Н))
            if ном_key in ссылка_к_uuid:
                nom_uuid = ссылка_к_uuid[ном_key]
                qty_orig = float(row.Кол)
                price_orig = float(row.Цн)
                sum_orig = float(row.Сум)
                vat_orig = float(row.СНд or 0)
                vat_name = str(conn.XMLСтрока(row.Нд)) if row.Нд else "БезНДС"
                break

    # Fallback на любую сопоставленную SKU + синтетические значения если в ПТУ не нашли
    if not nom_uuid:
        nom_uuid = fallback_nom_uuid
        qty_orig = 5.0
        price_orig = 100.0
        sum_orig = 500.0
        vat_orig = 90.16
        vat_name = "НДС22"

    print(f"  Орг:       {org_uuid}")
    print(f"  Контр:     {contr_uuid}")
    print(f"  Договор:   {dog_uuid}")
    print(f"  Склад:     {sklad_uuid}")
    print(f"  ПервичнПТУ:{исх_пту_uuid}")
    print(f"  Номенкл:   {nom_uuid}")
    if not all([org_uuid, contr_uuid, sklad_uuid, nom_uuid, исх_пту_uuid]):
        print("  FAIL: не хватает UUID для синтетики")
        return 1

    print("\n[2] Формирование пакета return_purchase (через КорректировкуПоступления)...")
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
        "ПервичнаяПТУ_UUID": исх_пту_uuid,
        "СуммаДокумента": sum_orig,
        "Товары": [
            {
                "НомерСтроки": 1,
                "Номенклатура": nom_uuid,
                "Количество": qty_orig,  # сколько было в ПТУ
                "Цена": price_orig,
                "Сумма": sum_orig,
                "СтавкаНДС": vat_name,
                "СуммаНДС": vat_orig,
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
    res = conn.TL_СопуткаСервис.ОбработатьПакет(read_result.Пакет, True)
    print(f"  Создано: {res.Документы.Количество()}, Ошибок: {res.Ошибки.Количество()}")
    for i in range(res.Ошибки.Количество()):
        print(f"    [{i+1}] {res.Ошибки.Получить(i)}")

    print("\n[4] Проверка КорректировкаПоступления в БП ГИГ...")
    q = conn.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка, СуммаДокумента, Проведен "
        "ИЗ Документ.КорректировкаПоступления "
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
