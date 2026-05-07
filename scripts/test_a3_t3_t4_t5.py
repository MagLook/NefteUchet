# -*- coding: utf-8 -*-
"""
Тесты T3/T4/T5 чек-листа A3 (02_TradeLedger_cfe_проект.md §6).

T3: пакет с неизвестной номенклатурой → ошибка НСИ_НеНайдена,
    документ не создаётся, статус ПринятоЧастично.

T4: пакет с ПометкаУдаления=true для существующего документа →
    пометка переносится через УстановитьПометкуУдаления.

T5: пакет с Проведен=false → документ создаётся непроведённым.

Запуск: py -3-32 D:\\Users\\magsp\\ELSYPLUS\\NefteUchet\\scripts\\test_a3_t3_t4_t5.py
"""
import sys
import json
import uuid
import datetime
import tempfile
import os

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

import win32com.client

BASE = r'D:\Users\magsp\GIG Base2'
USER = 'Гайворонская Татьяна'
PWD = '12345'
CONNSTR = f'File="{BASE}";Usr="{USER}";Pwd="{PWD}";'

print('=' * 76)
print('Тесты T3 / T4 / T5 — чек-лист A3 (02_TradeLedger_cfe_проект.md §6)')
print('=' * 76)

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect(CONNSTR)
print(f'\nПодключение OK. Конфигурация: {conn.Метаданные.Версия}\n')

ENUM = conn.Перечисления.TL_ТипОбъектаИсточника


def find_one(text, params):
    q = conn.NewObject('Запрос')
    q.Текст = text
    for k, v in params.items():
        q.УстановитьПараметр(k, v)
    sel = q.Выполнить().Выбрать()
    if sel.Следующий():
        return sel.Ссылка
    return None


def write_match(тип, src_uuid, ref):
    мз = conn.РегистрыСведений.TL_СоответствиеИсточников.СоздатьМенеджерЗаписи()
    мз.Тип = тип
    мз.ИсточникUUID = src_uuid
    мз.СсылкаОбъект = ref
    мз.ДатаПоследнейЗагрузки = datetime.datetime.now()
    мз.Записать()


# Базовая инфраструктура: организация, контрагент, склад, номенклатура
print('Подготовка базовой инфраструктуры...')
org = find_one(
    "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.Организации "
    "ГДЕ ИНН = &ИНН И НЕ ПометкаУдаления",
    {'ИНН': '7839440090'}  # ГАЗИНВЕСТГРУПП
)
contr = find_one(
    "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.Контрагенты "
    "ГДЕ ИНН = &ИНН И НЕ ПометкаУдаления",
    {'ИНН': '5003052454'}  # Мегаполис
)
warehouse = find_one(
    "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.Склады "
    "ГДЕ Наименование ПОДОБНО &Имя И НЕ ПометкаУдаления",
    {'Имя': '%АЗС 208%'}
)
nom = find_one(
    "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.Номенклатура "
    "ГДЕ НЕ ПометкаУдаления И НЕ ЭтоГруппа",
    {}
)
print(f'  Организация:  {org}')
print(f'  Контрагент:   {contr}')
print(f'  Склад:        {warehouse}')
print(f'  Номенклатура: {nom}')


def make_package(doc_uuid, package_id, *, vat_amount=166.67, sum_amount=1000.00,
                 quantity=10, price=100.00, проведен=True, пометка_удаления=False,
                 use_unknown_nom=False, reuse_nsi=None):
    """Сгенерить JSON-пакет с одним документом purchase.

    reuse_nsi: dict с ключами 'org', 'contr', 'wh', 'nom' — UUID для повторного
    использования (нужно для T4 шаг 2, где ИсточникUUID документа тот же,
    но и НСИ-UUID должны совпадать чтобы приёмник нашёл существующие соответствия).
    """
    if reuse_nsi:
        src_org = reuse_nsi['org']
        src_contr = reuse_nsi['contr']
        src_wh = reuse_nsi['wh']
        src_nom = reuse_nsi['nom']
    else:
        src_org = str(uuid.uuid4())
        src_contr = str(uuid.uuid4())
        src_wh = str(uuid.uuid4())
        src_nom = str(uuid.uuid4())  # этот UUID будет в JSON
    # Записать соответствия (если не используем unknown — иначе не записываем номенклатуру)
    if not write_match.get_done(src_org):
        write_match(ENUM.Организация, src_org, org)
    if not write_match.get_done(src_contr):
        write_match(ENUM.Контрагент, src_contr, contr)
    if not write_match.get_done(src_wh):
        write_match(ENUM.Склад, src_wh, warehouse)
    if not use_unknown_nom:
        write_match(ENUM.Номенклатура, src_nom, nom)
    # else: src_nom не записан → приёмник его не найдёт → ошибка НСИ_НеНайдена

    doc = {
        'Тип': 'purchase',
        'ИсточникUUID': doc_uuid,
        'Номер': f'TEST-{package_id[:6]}',
        'Дата': '2026-04-28T10:00:00+03:00',
        'Проведен': проведен,
        'ПометкаУдаления': пометка_удаления,
        'Организация': src_org,
        'Контрагент': src_contr,
        'Склад': src_wh,
        'ВидОперации': 'ОтПоставщика',
        'СуммаДокумента': sum_amount,
        'СуммаНДС': vat_amount,
        'ВалютаДокумента': 'RUB',
        'СуммаВключаетНДС': True,
        'НДСНеВыделять': False,
        'НДСВключенВСтоимость': False,
        'Товары': [{
            'НомерСтроки': 1,
            'Номенклатура': src_nom,
            'Количество': quantity,
            'Единица': 'шт',
            'Цена': price,
            'Сумма': sum_amount,
            'СтавкаНДС': 'НДС20',
            'СуммаНДС': vat_amount,
        }],
    }

    return {
        'ВерсияФормата': '1',
        'ВремяВыгрузки': '2026-05-07T22:00:00+03:00',
        'ИдентификаторПакета': package_id,
        'Источник': 'ЭЛСИ.АЗК ЦБ (тест)',
        'Пакет': {
            'ОрганизацияЦБ': src_org,
            'Склад': src_wh,
            'ПериодС': '2026-04-28',
            'ПериодПо': '2026-04-28',
        },
        'НСИ': [
            {'Тип': 'Организация', 'ИсточникUUID': src_org, 'Наименование': 'ГАЗИНВЕСТГРУПП', 'ИНН': '7839440090', 'ПометкаУдаления': False},
            {'Тип': 'Контрагент', 'ИсточникUUID': src_contr, 'Наименование': 'Мегаполис', 'ИНН': '5003052454', 'ПометкаУдаления': False},
            {'Тип': 'Склад', 'ИсточникUUID': src_wh, 'Наименование': 'АЗС 208', 'ПометкаУдаления': False},
            {'Тип': 'Номенклатура', 'ИсточникUUID': src_nom, 'Наименование': 'Тестовый SKU', 'СтавкаНДС': 'НДС20', 'Единица': 'шт', 'ПометкаУдаления': False},
        ],
        'Документы': [doc],
    }


# Простой кэш чтобы не писать одни и те же соответствия несколько раз
_match_done = set()
def _get_done(uid):
    return uid in _match_done
def _set_done(uid):
    _match_done.add(uid)
write_match.get_done = _get_done

# Обернём write_match чтобы запоминать
_orig_write = write_match
def _wrap(тип, src_uuid, ref):
    if src_uuid in _match_done:
        return
    _orig_write(тип, src_uuid, ref)
    _match_done.add(src_uuid)
write_match = _wrap
write_match.get_done = _get_done


def run_package(pkg, label):
    """Записать пакет во временный файл и прогнать через ОбработатьПакет."""
    tmp = os.path.join(tempfile.gettempdir(), f'test_{label}_{pkg["ИдентификаторПакета"]}.json')
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(pkg, f, ensure_ascii=False)
    rr = conn.TL_HTTPКлиентЦБ.ПрочитатьПакетИзФайла(tmp)
    if rr.Ошибка:
        return None, f'FAIL чтения: {rr.Ошибка}', None
    proc = conn.TL_СопуткаСервис.ОбработатьПакет(rr.Пакет, True)
    return proc, None, tmp


def find_doc_by_comment(package_id):
    q = conn.NewObject('Запрос')
    q.Текст = (
        "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.ПоступлениеТоваровУслуг "
        "ГДЕ ВЫРАЗИТЬ(Комментарий КАК СТРОКА(200)) ПОДОБНО &Шаблон"
    )
    q.УстановитьПараметр('Шаблон', f'TL|ЦБ|{package_id}|%')
    sel = q.Выполнить().Выбрать()
    return sel.Ссылка if sel.Следующий() else None


# ============================================================================
# T3: пакет с неизвестной номенклатурой
# ============================================================================
print('\n' + '=' * 76)
print('T3: пакет с неизвестной номенклатурой')
print('=' * 76)

pkg_id_t3 = str(uuid.uuid4())
doc_uid_t3 = str(uuid.uuid4())
pkg_t3 = make_package(doc_uid_t3, pkg_id_t3, use_unknown_nom=True)
proc, err, _ = run_package(pkg_t3, 't3')
if err:
    print(err); sys.exit(1)

print(f'  Статус:        {proc.Статус}')
print(f'  Создано:       {proc.Документы.Количество()} (ожидалось 0)')
print(f'  Ошибок:        {proc.Ошибки.Количество()} (ожидалось ≥1)')
err_msgs = [proc.Ошибки.Получить(i) for i in range(proc.Ошибки.Количество())]
for i, msg in enumerate(err_msgs[:5], 1):
    print(f'    [{i}] {msg}')

# Документ не должен быть создан
doc = find_doc_by_comment(pkg_id_t3)
print(f'  Документ найден по комментарию: {"да (ОШИБКА!)" if doc else "нет (правильно)"}')

# Запись в TL_ОшибкиЗагрузки
q = conn.NewObject('Запрос')
q.Текст = (
    "ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК N ИЗ РегистрСведений.TL_ОшибкиЗагрузки "
    "ГДЕ ИдентификаторПакета = &ИдПакета И КодОшибки = &Код"
)
q.УстановитьПараметр('ИдПакета', pkg_id_t3)
q.УстановитьПараметр('Код', conn.Перечисления.TL_КодОшибкиЗагрузки.НСИ_НеНайдена)
sel = q.Выполнить().Выбрать()
sel.Следующий()
print(f'  TL_ОшибкиЗагрузки записей с НСИ_НеНайдена: {sel.N}')

t3_ok = (
    proc.Документы.Количество() == 0
    and proc.Ошибки.Количество() >= 1
    and doc is None
    and sel.N >= 1
)
print(f'  T3: {"✓ PASS" if t3_ok else "✗ FAIL"}')

# ============================================================================
# T4: пакет с ПометкаУдаления=true для существующего документа
# ============================================================================
print('\n' + '=' * 76)
print('T4: ПометкаУдаления=true переносится')
print('=' * 76)

# Шаг 1: загрузить пакет нормально (создаётся документ)
pkg_id_t4a = str(uuid.uuid4())
doc_uid_t4 = str(uuid.uuid4())  # будем использовать тот же UUID документа в обоих пакетах
pkg_t4a = make_package(doc_uid_t4, pkg_id_t4a)
proc, err, _ = run_package(pkg_t4a, 't4a')
if err:
    print(err); sys.exit(1)
print(f'  Шаг 1 (создание): создано {proc.Документы.Количество()}, ошибок {proc.Ошибки.Количество()}')

# Найти документ
doc_t4 = find_doc_by_comment(pkg_id_t4a)
if not doc_t4:
    print('  FAIL: документ T4 не создан на шаге 1')
    sys.exit(1)
obj_before = doc_t4.ПолучитьОбъект()
print(f'  Документ создан: {doc_t4}')
print(f'  ПометкаУдаления до: {obj_before.ПометкаУдаления}')

# Шаг 2: тот же ИсточникUUID документа + те же UUID НСИ + ПометкаУдаления=true
pkg_id_t4b = str(uuid.uuid4())
nsi_t4 = {
    'org': pkg_t4a['Документы'][0]['Организация'],
    'contr': pkg_t4a['Документы'][0]['Контрагент'],
    'wh': pkg_t4a['Документы'][0]['Склад'],
    'nom': pkg_t4a['Документы'][0]['Товары'][0]['Номенклатура'],
}
pkg_t4b = make_package(doc_uid_t4, pkg_id_t4b, пометка_удаления=True, reuse_nsi=nsi_t4)
proc, err, _ = run_package(pkg_t4b, 't4b')
if err:
    print(err); sys.exit(1)
print(f'  Шаг 2 (пометка):  создано {proc.Документы.Количество()}, ошибок {proc.Ошибки.Количество()}')
for i in range(proc.Ошибки.Количество()):
    print(f'    err[{i+1}]: {proc.Ошибки.Получить(i)}')

# Перезапросить документ
doc_t4_after = find_doc_by_comment(pkg_id_t4b)
if not doc_t4_after:
    # Возможно комментарий не обновился, ищем по тому же ИсточникUUID через регистр соответствия
    ссылка = conn.TL_МаппингЦБ.НайтиСоответствие(ENUM.Поступление, doc_uid_t4)
    doc_t4_after = ссылка
obj_after = doc_t4_after.ПолучитьОбъект() if doc_t4_after else None
помечен = obj_after.ПометкаУдаления if obj_after else None
print(f'  Документ найден после: {doc_t4_after}')
print(f'  ПометкаУдаления после: {помечен}')

t4_ok = (помечен is True)
print(f'  T4: {"✓ PASS" if t4_ok else "✗ FAIL"}')

# ============================================================================
# T5: Проведен=false → документ создаётся непроведённым
# ============================================================================
print('\n' + '=' * 76)
print('T5: Проведен=false → документ непроведён')
print('=' * 76)

pkg_id_t5 = str(uuid.uuid4())
doc_uid_t5 = str(uuid.uuid4())
pkg_t5 = make_package(doc_uid_t5, pkg_id_t5, проведен=False)
proc, err, _ = run_package(pkg_t5, 't5')
if err:
    print(err); sys.exit(1)
print(f'  Создано: {proc.Документы.Количество()}, ошибок {proc.Ошибки.Количество()}')

doc_t5 = find_doc_by_comment(pkg_id_t5)
if not doc_t5:
    print('  FAIL: документ T5 не найден')
    sys.exit(1)
obj_t5 = doc_t5.ПолучитьОбъект()
проведен_фактически = obj_t5.Проведен
print(f'  Проведен фактически: {проведен_фактически} (ожидалось False)')

t5_ok = (проведен_фактически is False)
print(f'  T5: {"✓ PASS" if t5_ok else "✗ FAIL"}')

# ============================================================================
# Резюме
# ============================================================================
print('\n' + '=' * 76)
print('РЕЗЮМЕ')
print('=' * 76)
print(f'  T3 (НСИ_НеНайдена):        {"✓ PASS" if t3_ok else "✗ FAIL"}')
print(f'  T4 (ПометкаУдаления):       {"✓ PASS" if t4_ok else "✗ FAIL"}')
print(f'  T5 (Проведен=false):        {"✓ PASS" if t5_ok else "✗ FAIL"}')
all_ok = t3_ok and t4_ok and t5_ok
print('=' * 76)
print(f'ИТОГ: {"✓ ВСЕ ТЕСТЫ ПРОЙДЕНЫ" if all_ok else "⚠ ЕСТЬ ПРОВАЛЫ"}')
print('=' * 76)
sys.exit(0 if all_ok else 1)
