# -*- coding: utf-8 -*-
"""Полная функциональная проверка TradeLedger: настройки, регистр, документы, проводки."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

PASS = 0
FAIL = 0
WARN = 0

def ok(msg):
    global PASS
    PASS += 1
    print(f'  [OK] {msg}')

def fail(msg):
    global FAIL
    FAIL += 1
    print(f'  [FAIL] {msg}')

def warn(msg):
    global WARN
    WARN += 1
    print(f'  [WARN] {msg}')

# =============================================================================
print('=' * 80)
print('1. НАСТРОЙКИ (TL_Настройки)')
print('=' * 80)

critical_keys = {
    'URLСервера': 'https://pos.autooplata.ru/tms',
    'Логин': 'UserApi',
    'КодСистемы': '65',
    'Организация': 'ГАЗИНВЕСТГРУПП ООО',
    'ОсновнойСклад': 'Основной склад',
    'СтавкаНДС': '22',
    'Станция_5_Наименование': 'АКЗС Витебский',
    'Станция_5_Склад': 'АКЗС Витебский',
}

q = conn.NewObject('Query')
q.Text = 'SELECT Ключ, Значение FROM InformationRegister.TL_Настройки ORDER BY Ключ'
r = q.Execute().Choose()
settings = {}
while r.Next():
    settings[str(r.Ключ)] = str(r.Значение)

for key, expected in critical_keys.items():
    actual = settings.get(key, '')
    if actual == expected:
        ok(f'{key} = {actual}')
    elif actual:
        warn(f'{key} = "{actual}" (ожидали "{expected}")')
    else:
        fail(f'{key} — ПУСТО!')

pwd = settings.get('Пароль', '')
if pwd:
    ok(f'Пароль = ***')
else:
    fail('Пароль — ПУСТО!')

print(f'\nВсего настроек: {len(settings)}')

# =============================================================================
print()
print('=' * 80)
print('2. РЕГИСТР СТАТУСОВ (TL_СтатусыЗагрузки)')
print('=' * 80)

q2 = conn.NewObject('Query')
q2.Text = 'SELECT КлючЗагрузки, Статус, ДатаЗагрузки, ТипДокумента FROM InformationRegister.TL_СтатусыЗагрузки ORDER BY ДатаЗагрузки DESC'
r2 = q2.Execute().Choose()
statuses = []
while r2.Next():
    statuses.append({
        'key': str(r2.КлючЗагрузки),
        'status': str(r2.Статус),
        'date': str(r2.ДатаЗагрузки)[:16],
        'type': str(r2.ТипДокумента),
    })

print(f'Всего записей: {len(statuses)}')

# Проверка: нет ли дубликатов ключей
keys = [s['key'] for s in statuses]
dupes = [k for k in set(keys) if keys.count(k) > 1]
if dupes:
    fail(f'Дубликаты ключей: {dupes}')
else:
    ok('Нет дубликатов ключей')

# Статистика
loaded = sum(1 for s in statuses if s['status'] == 'Загружен')
posted = sum(1 for s in statuses if s['status'] == 'Проведён')
other = sum(1 for s in statuses if s['status'] not in ('Загружен', 'Проведён'))
print(f'  Загружен: {loaded}, Проведён: {posted}, Другое: {other}')

# Проверка целостности пакетов: для каждого базового ключа должны быть дочерние
base_keys = set()
child_keys = set()
for s in statuses:
    k = s['key']
    if '|ПЕРЕМ' in k or '|КОМПЛ' in k or '|online|' in k or '|cards|' in k:
        child_keys.add(k)
        # Определить базовый
        for suffix in ['|ПЕРЕМ', '|КОМПЛ', '|online|', '|cards|', '|ledger|']:
            idx = k.find(suffix)
            if idx > 0:
                base_keys.add(k[:idx])
                break
    elif k.startswith('TL|СМЕНА|') or k.startswith('TL|ТТН|'):
        base_keys.add(k)

orphan_children = []
for ck in child_keys:
    has_parent = False
    for bk in base_keys:
        if ck.startswith(bk + '|'):
            has_parent = True
            break
    if not has_parent:
        orphan_children.append(ck)

if orphan_children:
    warn(f'Сироты (дочерние без базового): {len(orphan_children)}')
    for oc in orphan_children[:5]:
        print(f'    {oc}')
else:
    ok('Все дочерние документы имеют базовый ключ')

# =============================================================================
print()
print('=' * 80)
print('3. ДОКУМЕНТЫ 1С (с комментарием TL|)')
print('=' * 80)

doc_stats = {}
for doc_type in ['ОтчетОРозничныхПродажах', 'ПеремещениеТоваров', 'КомплектацияНоменклатуры']:
    q3 = conn.NewObject('Query')
    q3.Text = f'SELECT COUNT(*) AS Всего, SUM(ВЫБОР КОГДА Проведен ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) AS Проведено, SUM(ВЫБОР КОГДА ПометкаУдаления ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) AS Удалено FROM Документ.{doc_type} WHERE Комментарий LIKE "%TL|%"'
    r3 = q3.Execute().Choose()
    if r3.Next():
        total = int(str(r3.Всего))
        posted_d = int(str(r3.Проведено or 0))
        deleted = int(str(r3.Удалено or 0))
        active = total - deleted
        doc_stats[doc_type] = {'total': total, 'posted': posted_d, 'deleted': deleted, 'active': active}
        print(f'  {doc_type}: всего={total}, проведено={posted_d}, удалено={deleted}, активно={active}')

# =============================================================================
print()
print('=' * 80)
print('4. ПРОВОДКИ (регистр бухгалтерии)')
print('=' * 80)

# Считаем проводки по проведённым TL| документам
for doc_type in ['ОтчетОРозничныхПродажах', 'ПеремещениеТоваров', 'КомплектацияНоменклатуры']:
    q4 = conn.NewObject('Query')
    q4.Text = f'''SELECT COUNT(РАЗЛИЧНЫЕ Регистратор) AS ДокСПроводками
    FROM РегистрБухгалтерии.Хозрасчетный
    WHERE Регистратор ССЫЛКА Документ.{doc_type}'''
    try:
        r4 = q4.Execute().Choose()
        if r4.Next():
            with_postings = int(str(r4.ДокСПроводками or 0))
            posted_d = doc_stats.get(doc_type, {}).get('posted', 0)
            if posted_d > 0 and with_postings == 0:
                fail(f'{doc_type}: {posted_d} проведённых, но 0 с проводками!')
            elif posted_d > 0:
                if with_postings >= posted_d:
                    ok(f'{doc_type}: {with_postings} документов с проводками из {posted_d} проведённых')
                else:
                    warn(f'{doc_type}: {with_postings} с проводками из {posted_d} проведённых')
    except:
        pass

# =============================================================================
print()
print('=' * 80)
print('5. ПРОВЕРКА МАСКИ ДОЧЕРНИХ (BUG-6 fix)')
print('=' * 80)

# Проверяем что маска | работает правильно
q5 = conn.NewObject('Query')
q5.Text = '''SELECT КлючЗагрузки FROM InformationRegister.TL_СтатусыЗагрузки
WHERE КлючЗагрузки LIKE "%|СМЕНА|65|5|9%"
ORDER BY КлючЗагрузки'''
r5 = q5.Execute().Choose()
keys_9x = []
while r5.Next():
    keys_9x.append(str(r5.КлючЗагрузки))

# Проверяем: нет ли пересечения между 91 и 910
for k in keys_9x:
    for k2 in keys_9x:
        if k != k2 and k2.startswith(k) and not k2.startswith(k + '|'):
            fail(f'Маска конфликт: "{k}" → "{k2}"')

if keys_9x:
    ok(f'Ключи смен 9x: {len(keys_9x)} — конфликтов нет')

# =============================================================================
print()
print('=' * 80)
print('6. API ПОДКЛЮЧЕНИЕ')
print('=' * 80)

try:
    url = conn.TL_Настройки.ПолучитьЗначение('URLСервера', '')
    login = conn.TL_Настройки.ПолучитьЗначение('Логин', '')
    pwd_val = conn.TL_Настройки.ПолучитьЗначение('Пароль', '')
    code = conn.TL_Настройки.ПолучитьЗначение('КодСистемы', '0')

    if url and login and pwd_val and code != '0':
        ok(f'Параметры API заполнены: {url}, login={login}, system={code}')
    else:
        fail(f'Параметры API неполные: url={bool(url)}, login={bool(login)}, pwd={bool(pwd_val)}, code={code}')

    # Попробуем авторизоваться
    try:
        token = conn.TL_ApiКлиент.Авторизация(url, login, pwd_val, int(code))
        if token:
            ok(f'Авторизация STS OK, токен получен ({len(str(token))} символов)')
        else:
            fail('Авторизация вернула пустой токен')
    except Exception as e:
        fail(f'Ошибка авторизации: {e}')
except Exception as e:
    fail(f'Ошибка доступа к TL_Настройки: {e}')

# =============================================================================
print()
print('=' * 80)
print(f'ИТОГО: OK={PASS}, FAIL={FAIL}, WARN={WARN}')
print('=' * 80)
