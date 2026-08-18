from pathlib import Path
import hashlib
import json
import re
import unicodedata
import uuid
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parents[2]
XML = ROOT / "xml-v4"
MD = "{http://v8.1c.ru/8.3/MDClasses}"
DUMP = "{http://v8.1c.ru/8.3/xcf/dumpinfo}"


def _bsl_function(text: str, name: str) -> str:
    return text.split(f"Функция {name}", 1)[1].split("КонецФункции", 1)[0]


def _canonical_hash(value: object) -> str:
    data = json.dumps(
        value, ensure_ascii=False, sort_keys=True, separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(data).hexdigest()


def test_receiver_sources_and_mirrors_are_complete():
    receiver = (ROOT / "src/CommonModules/TL_ПриемникV3.bsl").read_text("utf-8")
    receiver_mirror = (
        XML / "CommonModules/TL_ПриемникV3/Ext/Module.bsl"
    ).read_text("utf-8")
    service = (ROOT / "src/CommonModules/TL_СопуткаСервис.bsl").read_text("utf-8")
    service_mirror = (
        XML / "CommonModules/TL_СопуткаСервис/Ext/Module.bsl"
    ).read_text("utf-8")
    mapping = (ROOT / "src/CommonModules/TL_МаппингЦБ.bsl").read_text("utf-8")
    mapping_mirror = (
        XML / "CommonModules/TL_МаппингЦБ/Ext/Module.bsl"
    ).read_text("utf-8")

    assert receiver == receiver_mirror
    assert service == service_mirror
    assert mapping == mapping_mirror
    assert "TL_ПриемникV3.ОбработатьПакет(Пакет)" in service
    assert "TL_ПриемникV3.ПроверитьLegacyРазрешен(Пакет)" in service
    assert "ОбработатьАдаптированныйV3ВТранзакции" in service
    assert "business_hash_mismatch" in receiver
    assert "source_hash_mismatch" in receiver
    assert "blocked_mapping: payment" in receiver
    assert "component_count_mismatch" in receiver
    assert "target-orp-v1" in receiver
    assert "target-assembly-v1" in receiver
    assert "РазрешитьNeedsReview" in receiver
    assert "explicit_resolution_required" in receiver
    assert "НайтиВидОплатыРозницаБезСоздания" in receiver
    assert "НайтиВидОплатыРозницаБезСоздания" in mapping
    assert "ПринятоЧастично" not in receiver


def test_receiver_bsl_blocks_are_balanced():
    for path in (ROOT / "src/CommonModules/TL_ПриемникV3.bsl",):
        text = path.read_text("utf-8")
        assert len(re.findall(r"(?m)^[ \t]*Функция[ \t]", text)) == len(
            re.findall(r"(?m)^[ \t]*КонецФункции[ \t]*;?[ \t]*$", text)
        )
        assert len(re.findall(r"(?m)^[ \t]*Процедура[ \t]", text)) == len(
            re.findall(r"(?m)^[ \t]*КонецПроцедуры[ \t]*;?[ \t]*$", text)
        )
        assert len(re.findall(r"(?m)^[ \t]*Попытка[ \t]*$", text)) == len(
            re.findall(r"(?m)^[ \t]*КонецПопытки[ \t]*;?[ \t]*$", text)
        )
        assert text.count("#Область") == text.count("#КонецОбласти")


def test_v3_metadata_is_registered_and_well_formed():
    objects = [
        ("CommonModules", "TL_ПриемникV3"),
        ("InformationRegisters", "TL_ДоверенныеПолитики"),
        ("InformationRegisters", "TL_ТекущиеРевизии"),
        ("InformationRegisters", "TL_КомпонентыГруппы"),
        ("InformationRegisters", "TL_ПопыткиПриема"),
    ]
    for folder, name in objects:
        ET.parse(XML / folder / f"{name}.xml")

    configuration = ET.parse(XML / "Configuration.xml").getroot()
    values = {node.text for node in configuration.iter() if node.text}
    assert {name for _, name in objects} <= values

    dump = ET.parse(XML / "ConfigDumpInfo.xml").getroot()
    names = {node.attrib["name"] for node in dump.iter(f"{DUMP}Metadata")}
    assert "CommonModule.TL_ПриемникV3.Module" in names
    for _, name in objects[1:]:
        assert f"InformationRegister.{name}" in names
    assert "InformationRegister.TL_ПопыткиПриема.Resource.НомерПопытки" in names
    assert "InformationRegister.TL_ТекущиеРевизии.Resource.TransportProducer" in names


def test_receiver_metadata_is_external_connection_and_any_ib_ref():
    module = ET.parse(XML / "CommonModules/TL_ПриемникV3.xml").getroot()
    assert module.find(f".//{MD}ExternalConnection").text == "true"
    assert module.find(f".//{MD}Server").text == "true"
    assert module.find(f".//{MD}Privileged").text == "false"

    components = (
        XML / "InformationRegisters/TL_КомпонентыГруппы.xml"
    ).read_text("utf-8-sig")
    assert "cfg:AnyIBRef" in components
    assert "cfg:AnyRef" not in components


def test_v3_statuses_and_copy_map_are_present():
    enum = ET.parse(XML / "Enums/TL_СтатусПакета.xml").getroot()
    names = {node.text for node in enum.iter(f"{MD}Name")}
    assert {"ТребуетПроверки", "ЗаблокированоМаппингом"} <= names
    dev = (ROOT / "scripts/dev.ps1").read_text("utf-8-sig")
    assert "TL_ПриемникV3.bsl" in dev
    assert "[switch]$NoInit" in dev
    assert "if ($NoInit)" in dev
    assert "Инициализация пропущена (-NoInit)" in dev


def test_source_types_were_not_cross_contaminated():
    service = (ROOT / "src/CommonModules/TL_СопуткаСервис.bsl").read_text("utf-8")
    purchase = service.split("Процедура ОбработатьPurchase", 1)[1].split(
        "КонецПроцедуры", 1
    )[0]
    retail = service.split("Процедура ОбработатьRetailSale", 1)[1].split(
        "КонецПроцедуры", 1
    )[0]
    assert "TL_ТипОбъектаИсточника.Поступление" in purchase
    assert "TL_ТипОбъектаИсточника.ОтчётОРозничныхПродажах" not in purchase
    assert "TL_ТипОбъектаИсточника.ОтчётОРозничныхПродажах" in retail


def test_identity_golden_vectors_and_mutations_are_guarded_before_state():
    shift = "2f8f67da-419f-5d18-a552-7e67066cb929"
    company = "11111111-2222-3333-4444-555555555555"
    station = "208"
    business_date = "2026-08-17"
    business_key = _canonical_hash({
        "BusinessShiftID": shift, "CompanyID": company, "StationID": station,
    })
    packet_id = str(uuid.uuid5(
        uuid.UUID("6ba7b811-9dad-11d1-80b4-00c04fd430c8"),
        "urn:elsyplus:ledger:accounting-packet:v1:" + shift,
    ))
    edge_alias = _canonical_hash({
        "algorithm": "business-shift-alias-v1", "business_date": business_date,
        "company_id": company, "internal_shift_no": "208-000777",
        "station_id": station,
    })
    common_alias = _canonical_hash({
        "algorithm": "business-shift-common-alias-v1", "business_date": business_date,
        "company_id": company, "ose": "OSE-208-777", "station_id": station,
    })
    assert business_key == "d099a89d7c6db617a6632d5b0229d71473f9e2f3a855b5a0d65b7078f455874d"
    assert packet_id == "bb950501-9578-5556-aef0-675e6fcdea2d"
    assert edge_alias == "9d6c7cf4ca2e2982100999bfdde84635fa5add5a4636492e112a8fef3570f98d"
    assert common_alias == "c0b114e6c614acd22e68d85db55a490b0cdc0b30384cbfd72b501ea9edf82024"

    for field, changed in (
        ("BusinessShiftID", shift[:-1] + "8"),
        ("CompanyID", company[:-1] + "6"),
        ("StationID", "209"),
    ):
        identity = {
            "BusinessShiftID": shift, "CompanyID": company, "StationID": station,
        }
        identity[field] = changed
        assert _canonical_hash(identity) != business_key
    assert str(uuid.uuid5(
        uuid.UUID("6ba7b811-9dad-11d1-80b4-00c04fd430c8"),
        "urn:elsyplus:ledger:accounting-packet:v1:" + shift[:-1] + "8",
    )) != packet_id

    receiver = (ROOT / "src/CommonModules/TL_ПриемникV3.bsl").read_text("utf-8")
    identity_guard = _bsl_function(receiver, "ПроверитьИдентичностьПакета")
    packet_guard = _bsl_function(receiver, "ПроверитьПакет")
    for token in (
        "business_key_mismatch", "packet_uuid_mismatch", "alias_hash_mismatch",
        "alias_algorithm_unknown", "alias_duplicate_or_unsorted",
        "business-shift-alias-v1", "business-shift-common-alias-v1",
    ):
        assert token in identity_guard or token in receiver
    assert "6ba7b8119dad11d180b400c04fd430c8" in receiver
    assert "urn:elsyplus:ledger:accounting-packet:v1:" in receiver
    assert "ХешФункция.SHA1" in receiver
    assert "Неопредело" not in receiver
    assert packet_guard.index("ПроверитьИдентичностьПакета") < packet_guard.index(
        "ПолучитьДовереннуюПолитику"
    ) < packet_guard.index("ПроверитьРевизию")
    assert "Тип(" not in receiver


def test_nfc_contract_gate_and_decomposed_golden_vector():
    decomposed = "Cafe\u0301"
    composed = "Café"
    assert decomposed != composed
    assert unicodedata.normalize("NFC", decomposed) == composed
    assert _canonical_hash({"Name": decomposed}) != _canonical_hash({"Name": composed})
    assert _canonical_hash({
        "Name": unicodedata.normalize("NFC", decomposed),
    }) == _canonical_hash({"Name": composed})

    receiver = (ROOT / "src/CommonModules/TL_ПриемникV3.bsl").read_text("utf-8")
    identity_guard = _bsl_function(receiver, "ПроверитьИдентичностьПакета")
    declaration_guard = _bsl_function(receiver, "ПроверитьДекларациюNFCV3")
    packet_guard = _bsl_function(receiver, "ПроверитьПакет")
    assert 'Пакет, "UnicodeNormalization", "")) = "NFC"' in declaration_guard
    assert "unicode_normalization_required" in identity_guard
    assert identity_guard.index("ПроверитьДекларациюNFCV3") < identity_guard.index(
        "ПроверитьТочныеПоляJSON"
    )
    assert "UnicodeNormalization" in identity_guard
    assert "UnicodeNormalization" in packet_guard
    assert packet_guard.index("ПроверитьДекларациюNFCV3") < packet_guard.index(
        "Обязательные ="
    )
    assert "НормализоватьСтроку" not in receiver


def test_null_shift_close_uses_business_date_without_clock_fallback():
    receiver = (ROOT / "src/CommonModules/TL_ПриемникV3.bsl").read_text("utf-8")
    date_fallback = _bsl_function(receiver, "ПолучитьДатуДокументовV3")
    adapter = _bsl_function(receiver, "АдаптироватьПакет")
    documents = _bsl_function(receiver, "АдаптироватьДокументы")

    assert 'Смена, "ЗакрытаВ", Неопределено' in date_fallback
    assert "Если ЗначениеЗаполнено(ЗакрытаВ)" in date_fallback
    assert 'Пакет, "BusinessDate", ""' in date_fallback
    assert 'Смена, "ОСЭ", ""' in date_fallback
    assert "ЭтоКаноническаяДатаV3(BusinessDate)" in date_fallback
    assert "Возврат BusinessDate" in date_fallback
    assert "ТекущаяДата" not in date_fallback
    assert "ТекущаяДатаСеанса" not in date_fallback
    assert "ДатаДокумента = ПолучитьДатуДокументовV3(Пакет)" in adapter
    assert 'Смена.Вставить("Закрытие", ДатаДокумента)' in adapter
    assert "АдаптироватьДокументы(Пакет, Политика, ДатаДокумента)" in adapter
    assert 'Retail.Вставить("Дата", ДатаДокумента)' in documents
    assert 'СтрокаRetail.Вставить("ДатаРеализации", ДатаДокумента)' in receiver


def test_repeat_live_validation_and_atomic_exact_nsi_path():
    receiver = (ROOT / "src/CommonModules/TL_ПриемникV3.bsl").read_text("utf-8")
    service = (ROOT / "src/CommonModules/TL_СопуткаСервис.bsl").read_text("utf-8")
    receive = _bsl_function(receiver, "ОбработатьПакет")
    writer = _bsl_function(service, "ОбработатьАдаптированныйV3ВТранзакции")
    adapter = _bsl_function(receiver, "АдаптироватьПакет")
    preflight = _bsl_function(receiver, "ПроверитьПолнотуИСостав")

    assert receive.index("НачатьТранзакцию") < receive.index(
        "Повторная.ИдемпотентныйПовтор"
    ) < receive.index("ПроверитьСуществующиеКомпоненты")
    for token in (
        "target_hash_conflict", "posted_component", "deleted_component",
        "target_component_missing", "component_source_hash_conflict",
    ):
        assert token in receiver
    assert "ОжидаемыеКлючи" in receiver
    assert "НРег(Строка(Выборка.TargetHash))" not in receiver
    assert "НРег(Строка(Выборка.SourceHash))" not in receiver
    revision_rule = _bsl_function(receiver, "ПроверитьПравилоРевизии")
    assert "НРег(ТекущийХеш)" not in revision_rule
    assert 'Итог.Вставить("НСИ", TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON' in adapter
    assert "ОбработатьЭлементНСИ" in writer
    assert writer.index("ОбработатьЭлементНСИ") < writer.index("ОбработатьРецепт") < writer.index(
        "ОбработатьRetailSale"
    )
    assert "НачатьТранзакцию" not in writer
    assert "ЗафиксироватьТранзакцию" not in writer
    assert "TL_МаппингЦБ.ПолучитьНоменклатуруПоUUID" not in preflight
    assert "ПроверитьНСИПакета" in preflight
    assert '"cash"' not in preflight and '"cashless"' not in preflight
    assert "ВызватьИсключение" in writer
