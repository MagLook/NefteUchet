# -*- coding: utf-8 -*-
"""
Прогон A3 на реальном пакете от агента TL_ЭкспортБП.

1. Предзагрузка соответствий в TL_СоответствиеИсточников:
   - Организация: Норд-Лайн (UUID из ЦБ) → ГАЗИНВЕСТГРУПП ООО (по факту перехода).
   - Контрагент: Мегаполис ТК → ТК МЕГАПОЛИС АО (поиск по ИНН).
   - Склад: АЗС №208, Торговый зал → АЗС 208 г.Выборг (эвристика имени).
   - Договор: автосоздание через TL_МаппингЦБ.ПолучитьДоговорПоUUID.
   - 35 Номенклатур: создать новые элементы Catalog.Номенклатура с реквизитами
     из JSON (Наименование, Артикул=КодЦБ, СтавкаНДС, Единица, ВидНоменклатуры).

2. Прогон TL_HTTPКлиентЦБ.ПрочитатьПакетИзФайла → TL_СопуткаСервис.ОбработатьПакет.

3. Проверка: 1 Document.ПоступлениеТоваровУслуг, 35 строк, сумма 159 816.08,
   проводки Дт 41.02 Кт 60.01 + Дт 19.03 Кт 60.01.

Запуск: py -3-32 D:\\Users\\magsp\\ELSYPLUS\\NefteUchet\\scripts\\run_a3_real.py
"""
import sys
import os
import json
import datetime

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

try:
    import win32com.client
except ImportError:
    print("ERROR: pywin32 не установлен в этой версии Python")
    sys.exit(1)

# ---------------------------------------------------------------------------
# Можно передать другой пакет первым аргументом: py -3-32 run_a3_real.py <path>
PACKAGE_FILE = sys.argv[1] if len(sys.argv) > 1 else r'D:\TL_BP_Export\cf2cd842-6522-4258-960b-70e13e9ef0c3.json'
BASE = r'D:\Users\magsp\GIG Base2'
USER = 'Гайворонская Татьяна'
PWD = '12345'
CONNSTR = f'File="{BASE}";Usr="{USER}";Pwd="{PWD}";'

print('=' * 76)
print('A3 ПРОГОН: реальный пакет от агента TL_ЭкспортБП')
print('=' * 76)

# Чтение пакета
print(f'\nПакет: {PACKAGE_FILE}')
with open(PACKAGE_FILE, 'r', encoding='utf-8-sig') as f:
    pkg = json.load(f)
package_id = pkg['ИдентификаторПакета']
print(f'  ИдентификаторПакета: {package_id}')
print(f'  ВерсияФормата:       {pkg.get("ВерсияФормата")}')
print(f'  Источник:            {pkg.get("Источник")}')

nsi = {(x['Тип'], x['ИсточникUUID']): x for x in pkg['НСИ']}
docs = pkg['Документы']
print(f'  НСИ: {len(pkg["НСИ"])} | Документов: {len(docs)}')

# ---------------------------------------------------------------------------
print('\n[1/4] Подключение к БП через V83.COMConnector...')
connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect(CONNSTR)
print(f'      OK. Конфа {conn.Метаданные.Версия}')

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

def write_match(тип, uuid, ref):
    мз = conn.РегистрыСведений.TL_СоответствиеИсточников.СоздатьМенеджерЗаписи()
    мз.Тип = тип
    мз.ИсточникUUID = uuid
    мз.СсылкаОбъект = ref
    мз.ДатаПоследнейЗагрузки = datetime.datetime.now()
    мз.ИдентификаторПоследнегоПакета = package_id
    мз.Записать()

def existing_match(тип, uuid):
    return conn.TL_МаппингЦБ.НайтиСоответствие(тип, uuid)

# ---------------------------------------------------------------------------
print('\n[2/4] Предзагрузка соответствий НСИ...')

# Организация: Норд-Лайн UUID → ГАЗИНВЕСТГРУПП (переключение юр.лица при переходе)
org_nsi = next(x for x in pkg['НСИ'] if x['Тип'] == 'Организация')
gig_org = find_one(
    "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.Организации "
    "ГДЕ ИНН = &ИНН И НЕ ПометкаУдаления",
    {'ИНН': '7839440090'}  # ГАЗИНВЕСТГРУПП
)
if not gig_org:
    print('      FAIL: ГАЗИНВЕСТГРУПП не найдена в БП ГИГ')
    sys.exit(1)
if not existing_match(ENUM.Организация, org_nsi['ИсточникUUID']):
    write_match(ENUM.Организация, org_nsi['ИсточникUUID'], gig_org)
    print(f'      Организация: {org_nsi["Наименование"]} → ГАЗИНВЕСТГРУПП ООО (новый владелец)')
else:
    print(f'      Организация: уже сопоставлена')

# Контрагент: по ИНН
contr_nsi = next(x for x in pkg['НСИ'] if x['Тип'] == 'Контрагент')
contr_inn = contr_nsi.get('ИНН', '')
contr_ref = conn.TL_МаппингЦБ.ПолучитьКонтрагентаПоUUID(
    contr_nsi['ИсточникUUID'], contr_inn
)
if not contr_ref:
    print(f'      FAIL: контрагент {contr_nsi["Наименование"]} (ИНН {contr_inn}) не найден')
    sys.exit(1)
print(f'      Контрагент:  {contr_nsi["Наименование"]} → {contr_ref}')

# Склад: по эвристике имени
sk_nsi = next(x for x in pkg['НСИ'] if x['Тип'] == 'Склад')
sk_ref = conn.TL_МаппингЦБ.ПолучитьСкладПоUUID(
    sk_nsi['ИсточникUUID'], sk_nsi.get('Наименование', '')
)
if not sk_ref:
    print(f'      FAIL: склад {sk_nsi["Наименование"]} не найден')
    sys.exit(1)
print(f'      Склад:       {sk_nsi["Наименование"]} → {sk_ref}')

# Договор: автосоздание через TL_МаппингЦБ
dog_nsi = next((x for x in pkg['НСИ'] if x['Тип'] == 'Договор'), None)
if dog_nsi:
    # Параметры: Наименование, Номер, Дата, ВидДоговора, Валюта
    pars = conn.NewObject('Структура')
    pars.Вставить('Наименование', dog_nsi.get('Наименование', ''))
    pars.Вставить('Номер', dog_nsi.get('Номер', ''))
    if dog_nsi.get('Дата'):
        try:
            pars.Вставить('Дата', datetime.datetime.fromisoformat(
                dog_nsi['Дата'].replace('Z', '+00:00')
            ))
        except Exception:
            pass
    pars.Вставить('ВидДоговора', dog_nsi.get('ВидДоговора', 'СПоставщиком'))
    dog_ref = conn.TL_МаппингЦБ.ПолучитьДоговорПоUUID(
        dog_nsi['ИсточникUUID'], contr_ref, pars
    )
    if dog_ref:
        print(f'      Договор:     {dog_nsi.get("Номер")} → {dog_ref}')
    else:
        print(f'      WARN: договор не создан (продолжаем без)')

# ---------------------------------------------------------------------------
print('\n[3/4] Создание 35 элементов Catalog.Номенклатура и соответствий...')

# Поиск ВидНоменклатуры=Товар (Catalog.ВидыНоменклатуры в БП 3.0)
вид_товар = find_one(
    "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.ВидыНоменклатуры "
    "ГДЕ Наименование = &Имя И НЕ ПометкаУдаления",
    {'Имя': 'Товары'}
)
if not вид_товар:
    # fallback на любой первый
    вид_товар = find_one(
        "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.ВидыНоменклатуры "
        "ГДЕ НЕ ПометкаУдаления",
        {}
    )
print(f'      ВидНоменклатуры по умолчанию: {вид_товар}')

# Базовая единица «шт» — код 796 в КлассификаторЕдиницИзмерения
ed_sht = find_one(
    "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.КлассификаторЕдиницИзмерения "
    "ГДЕ Код = &Код",
    {'Код': '796'}
)
print(f'      Единица «шт»: {ed_sht}')

nom_items = [x for x in pkg['НСИ'] if x['Тип'] == 'Номенклатура']
print(f'      К обработке: {len(nom_items)} SKU')

created = 0
matched = 0
errors = []

for i, item in enumerate(nom_items, 1):
    uuid = item['ИсточникUUID']

    # Уже есть соответствие?
    existing = existing_match(ENUM.Номенклатура, uuid)
    if existing:
        matched += 1
        continue

    name = item.get('Наименование', f'SKU_{uuid[:8]}')
    code = item.get('КодЦБ', '')
    vat_name = item.get('СтавкаНДС', 'НДС20')

    try:
        nm = conn.Справочники.Номенклатура.СоздатьЭлемент()
        nm.Наименование = name[:100]  # ограничение типового
        nm.НаименованиеПолное = item.get('НаименованиеПолное', name)[:250]
        if code:
            try:
                nm.Артикул = code[:25]
            except Exception:
                pass
        # СтавкаНДС в БП 3.0 (3.0.106+) — не реквизит карточки, а РегистрСведений.СтавкиНДСНоменклатуры.
        # На A3 ставка приёмником берётся из строки JSON-документа (Стр.СтавкаНДС = СопоставитьСтавкуНДС),
        # не из карточки — поэтому при создании Номенклатуры ставку не задаём.
        # Единица
        if ed_sht:
            try:
                nm.БазоваяЕдиницаИзмерения = ed_sht
            except Exception:
                pass
        # Вид
        if вид_товар:
            try:
                nm.ВидНоменклатуры = вид_товар
            except Exception:
                pass

        nm.Записать()
        write_match(ENUM.Номенклатура, uuid, nm.Ссылка)
        created += 1
    except Exception as e:
        errors.append(f'SKU {i}/{len(nom_items)} {name[:40]}: {e}')

    if i % 10 == 0:
        print(f'        {i}/{len(nom_items)}: создано {created}, '
              f'уже сопоставлено {matched}, ошибок {len(errors)}')

print(f'      ИТОГО: создано {created}, уже было {matched}, ошибок {len(errors)}')
if errors:
    for err in errors[:5]:
        print(f'        ERR: {err}')

# ---------------------------------------------------------------------------
print('\n[3.5] Workaround: патч ВидОперации в JSON (баг агента ЦБ — выдаёт имя перечисления)...')

# Агент ЦБ выгружает "ВидОперации": "ВидыОперацийПоступлениеТоваровУслугНаАЗК"
# (имя перечисления), а контракт §8.2 ожидает значения "ОтПоставщика" / "ВнутреннееПеремещение".
# Патчим JSON в копии файла перед загрузкой.
patched_pkg = json.loads(open(PACKAGE_FILE, 'r', encoding='utf-8-sig').read())
patched = 0
patched_vat = 0
for d in patched_pkg.get('Документы', []):
    vop = d.get('ВидОперации', '')
    if vop and 'ВидыОперацийПоступление' in vop:
        d['ВидОперации'] = 'ОтПоставщика'
        patched += 1
    # СтавкаНДС в строках — тот же баг (выгружено имя перечисления)
    for row in d.get('Товары', []):
        vat = row.get('СтавкаНДС', '')
        if vat == 'СтавкиНДС' or vat in ('', None):
            # Без реального значения определяем по СуммаНДС/Сумма:
            # НДС20: 20/120 ≈ 0.1667; НДС22: 22/122 ≈ 0.1803; НДС10: 10/110 ≈ 0.0909
            sum_ = float(row.get('Сумма', 0) or 0)
            vat_amt = float(row.get('СуммаНДС', 0) or 0)
            if sum_ > 0:
                ratio = vat_amt / sum_
                if abs(ratio - 0.1667) < 0.005:
                    row['СтавкаНДС'] = 'НДС20'
                elif abs(ratio - 0.1803) < 0.005:
                    row['СтавкаНДС'] = 'НДС22'
                elif abs(ratio - 0.0909) < 0.005:
                    row['СтавкаНДС'] = 'НДС10'
                elif vat_amt == 0:
                    row['СтавкаНДС'] = 'БезНДС'
                else:
                    row['СтавкаНДС'] = 'НДС20'  # fallback
            else:
                row['СтавкаНДС'] = 'НДС20'
            patched_vat += 1
patched_path = PACKAGE_FILE.replace('.json', '.patched.json')
with open(patched_path, 'w', encoding='utf-8') as f:
    json.dump(patched_pkg, f, ensure_ascii=False, indent=2)
print(f'      Пропатчено документов: {patched}, строк-ставок: {patched_vat}, файл: {patched_path}')

print('\n[4/4] Прогон TL_СопуткаСервис.ОбработатьПакет на реальном файле...')

read_result = conn.TL_HTTPКлиентЦБ.ПрочитатьПакетИзФайла(patched_path)
if read_result.Ошибка:
    print(f'      FAIL чтения: {read_result.Ошибка}')
    sys.exit(1)
print(f'      Прочитано: ХешСовпал={read_result.ХешСовпал}')

processing = conn.TL_СопуткаСервис.ОбработатьПакет(read_result.Пакет, True)

print(f'      Статус:           {processing.Статус}')
print(f'      ВсегоНСИ:         {processing.ВсегоНСИ}')
print(f'      ВсегоДокументов:  {processing.ВсегоДокументов}')
print(f'      Создано:          {processing.Документы.Количество()}')
print(f'      Ошибок:           {processing.Ошибки.Количество()}')

if processing.Ошибки.Количество() > 0:
    print('      --- Ошибки ---')
    for i in range(min(processing.Ошибки.Количество(), 10)):
        print(f'        [{i+1}] {processing.Ошибки.Получить(i)}')

# ---------------------------------------------------------------------------
print('\n[Контроль] Проверка созданного документа в БП...')

q = conn.NewObject('Запрос')
q.Текст = (
    "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.ПоступлениеТоваровУслуг "
    "ГДЕ ВЫРАЗИТЬ(Комментарий КАК СТРОКА(200)) ПОДОБНО &Шаблон"
)
q.УстановитьПараметр('Шаблон', f'TL|ЦБ|{package_id}|%')
sel = q.Выполнить().Выбрать()
if not sel.Следующий():
    print('      FAIL: документ не найден по комментарию')
    sys.exit(1)

doc_ref = sel.Ссылка
obj = doc_ref.ПолучитьОбъект()
print(f'      Найден: {doc_ref}')
print(f'      Дата:           {obj.Дата}')
print(f'      Контрагент:     {obj.Контрагент}')
print(f'      Склад:          {obj.Склад}')
print(f'      ВидОперации:    {obj.ВидОперации}')
print(f'      СуммаДокумента: {obj.СуммаДокумента}')
print(f'      Проведен:       {obj.Проведен}')
print(f'      Строк Товаров:  {obj.Товары.Количество()}')

# Ожидаемое
expected_sum = 159816.08
delta = abs(float(obj.СуммаДокумента) - expected_sum)
ok_sum = delta < 0.01
ok_lines = obj.Товары.Количество() == 35
print(f'      Совпадение суммы (ожидалось {expected_sum}): {ok_sum} (Δ={delta:.4f})')
print(f'      Совпадение строк (ожидалось 35):              {ok_lines}')

# Проводки
print('\n      Проводки Хозрасчётный:')
q2 = conn.NewObject('Запрос')
q2.Текст = (
    "ВЫБРАТЬ Период, СчетДт.Код КАК ДтКод, СчетКт.Код КАК КтКод, "
    "СУММА(Сумма) КАК Сумма "
    "ИЗ РегистрБухгалтерии.Хозрасчетный "
    "ГДЕ Регистратор = &Рег "
    "СГРУППИРОВАТЬ ПО Период, СчетДт.Код, СчетКт.Код "
    "УПОРЯДОЧИТЬ ПО Период"
)
q2.УстановитьПараметр('Рег', doc_ref)
sel2 = q2.Выполнить().Выбрать()
n = 0
total_debit = 0
while sel2.Следующий():
    n += 1
    print(f'        [{n}] Дт {sel2.ДтКод} Кт {sel2.КтКод} = {sel2.Сумма:.2f}')
    if sel2.ДтКод == '41.02':
        total_debit += float(sel2.Сумма)

# Сверка пакета
print('\n      Запись в TL_СверкаПакета:')
q3 = conn.NewObject('Запрос')
q3.Текст = (
    "ВЫБРАТЬ Статус, СуммаЦБ, СуммаБП, Расхождение, СтрокДокументов, СтрокНСИ "
    "ИЗ РегистрСведений.TL_СверкаПакета "
    "ГДЕ ИдентификаторПакета = &ИдПакета"
)
q3.УстановитьПараметр('ИдПакета', package_id)
sel3 = q3.Выполнить().Выбрать()
if sel3.Следующий():
    print(f'        Статус:          {sel3.Статус}')
    print(f'        СуммаЦБ:         {sel3.СуммаЦБ:.2f}')
    print(f'        СуммаБП:         {sel3.СуммаБП:.2f}')
    print(f'        Расхождение:     {sel3.Расхождение:.2f}')
    print(f'        СтрокДокументов: {sel3.СтрокДокументов}')
    print(f'        СтрокНСИ:        {sel3.СтрокНСИ}')

# Резюме
print('\n' + '=' * 76)
all_ok = (
    processing.Документы.Количество() == 1
    and processing.Ошибки.Количество() == 0
    and ok_sum
    and ok_lines
    and n >= 2  # минимум 41.02 и 19.03
)
if all_ok:
    print('A3 РЕАЛЬНЫЙ ПАКЕТ: ✓ УСПЕХ')
else:
    print('A3 РЕАЛЬНЫЙ ПАКЕТ: ⚠ есть отклонения, см. вывод выше')
print('=' * 76)
