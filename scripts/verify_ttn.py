# -*- coding: utf-8 -*-
"""Проверка цепочки ТТН 3310."""
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

def qval(sel, name):
    """Безопасно получить значение поля запроса."""
    try:
        v = getattr(sel, name, None)
        if callable(v):
            return None
        return v
    except:
        return None

def qfloat(sel, name):
    v = qval(sel, name)
    if v is None:
        return 0.0
    try:
        return float(v)
    except:
        return 0.0

def qstr(sel, name):
    v = qval(sel, name)
    if v is None:
        return ''
    return str(v)

print('=' * 80)
print('ПРОВЕРКА ЦЕПОЧКИ ТТН 3310 (АКЗС Витебский)')
print('=' * 80)

# === 1. Перемещение тонн ===
print()
print('--- ШАГ 1: Перемещение тонн (Основной склад → АКЗС Витебский) ---')

q = conn.NewObject('Query')
q.Text = '''
SELECT
    Док.Представление AS Документ,
    Док.Проведен AS Проведен,
    Док.СкладОтправитель.Наименование AS Откуда,
    Док.СкладПолучатель.Наименование AS Куда
FROM Документ.ПеремещениеТоваров AS Док
WHERE Док.Комментарий LIKE "%TL|ТТН|65|5|3310|ПЕРЕМ%"
'''
res = q.Execute().Choose()
while res.Next():
    st = 'ПРОВЕДЁН' if res.Проведен else 'не проведён'
    print(f'  {qstr(res, "Документ")} | {st}')
    print(f'  Откуда: {qstr(res, "Откуда")} → Куда: {qstr(res, "Куда")}')

# Проводки
q1 = conn.NewObject('Query')
q1.Text = '''
SELECT
    СчетДт.Код AS Дт, СчетКт.Код AS Кт,
    Сумма AS С, КоличествоДт AS К
FROM РегистрБухгалтерии.Хозрасчетный
WHERE Регистратор.Комментарий LIKE "%TL|ТТН|65|5|3310|ПЕРЕМ%"
ORDER BY НомерСтроки
'''
print('  Проводки:')
r1 = q1.Execute().Choose()
while r1.Next():
    dt = qstr(r1, "Дт").strip()
    kt = qstr(r1, "Кт").strip()
    s = qfloat(r1, "С")
    k = qfloat(r1, "К")
    ok = 'OK — тонны 41.01 → 41.01' if dt == '41.01' and kt == '41.01' else f'ВНИМАНИЕ: {dt} → {kt}'
    print(f'    Дт {dt} ← Кт {kt} | {k:,.3f} т | {s:>12,.2f} р | {ok}')

# === 2. Комплектация ===
print()
print('--- ШАГ 2: Комплектация (тонны → литры) ---')

q2 = conn.NewObject('Query')
q2.Text = '''
SELECT
    Док.Представление AS Документ,
    Док.Проведен AS Проведен,
    Док.Комментарий AS Комм
FROM Документ.КомплектацияНоменклатуры AS Док
WHERE Док.Комментарий LIKE "%TL|ТТН|65|5|3310|КОМПЛ%"
'''
r2 = q2.Execute().Choose()
while r2.Next():
    st = 'ПРОВЕДЁН' if r2.Проведен else 'не проведён'
    print(f'  {qstr(r2, "Документ")} | {st}')
    print(f'  {qstr(r2, "Комм")}')

# Проводки комплектации
q2p = conn.NewObject('Query')
q2p.Text = '''
SELECT
    СчетДт.Код AS Дт, СчетКт.Код AS Кт,
    Сумма AS С, КоличествоДт AS КД, КоличествоКт AS КК
FROM РегистрБухгалтерии.Хозрасчетный
WHERE Регистратор.Комментарий LIKE "%TL|ТТН|65|5|3310|КОМПЛ%"
ORDER BY НомерСтроки
'''
print('  Проводки:')
r2p = q2p.Execute().Choose()
while r2p.Next():
    dt = qstr(r2p, "Дт").strip()
    kt = qstr(r2p, "Кт").strip()
    s = qfloat(r2p, "С")
    kd = qfloat(r2p, "КД")
    kk = qfloat(r2p, "КК")

    if dt == '41.02' and kt == '41.01':
        density = kk / kd if kd > 0 else 0
        print(f'    Дт {dt} (литры={kd:,.3f}) ← Кт {kt} (тонны={kk:,.3f}) | {s:>12,.2f} р')
        print(f'      ✓ Комплектация: {kk:,.3f} т → {kd:,.3f} л | плотность = {density:.4f}')
    elif dt == '41.01' and kt == '41.02':
        density = kd / kk if kk > 0 else 0
        print(f'    Дт {dt} (тонны={kd:,.3f}) ← Кт {kt} (литры={kk:,.3f}) | {s:>12,.2f} р')
        print(f'      Обратная проводка: {kd:,.3f} т ← {kk:,.3f} л | плотность = {density:.4f}')
    else:
        print(f'    Дт {dt} (кол={kd:,.3f}) ← Кт {kt} (кол={kk:,.3f}) | {s:>12,.2f} р')

# === 3. Сверка тоннажа ===
print()
print('=' * 80)
print('СВЕРКА:')

q_p = conn.NewObject('Query')
q_p.Text = 'SELECT КоличествоДт AS К, Сумма AS С FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор.Комментарий LIKE "%TL|ТТН|65|5|3310|ПЕРЕМ%" AND СчетДт.Код = "41.01"'
rp = q_p.Execute().Choose()
pt = 0; ps = 0
while rp.Next():
    pt += qfloat(rp, "К"); ps += qfloat(rp, "С")

q_k = conn.NewObject('Query')
q_k.Text = 'SELECT КоличествоКт AS К, Сумма AS С FROM РегистрБухгалтерии.Хозрасчетный WHERE Регистратор.Комментарий LIKE "%TL|ТТН|65|5|3310|КОМПЛ%" AND СчетКт.Код = "41.01"'
rk = q_k.Execute().Choose()
kt_val = 0; ks = 0
while rk.Next():
    kt_val += qfloat(rk, "К"); ks += qfloat(rk, "С")

print(f'  Перемещено на АЗС:    {pt:,.3f} т | {ps:>12,.2f} р')
print(f'  Скомплектовано (расход): {kt_val:,.3f} т | {ks:>12,.2f} р')
if abs(pt - kt_val) < 0.001:
    print(f'  ✓ Тоннаж совпадает — полный цикл пройден')
else:
    print(f'  ✗ Расхождение: {pt - kt_val:,.3f} т')

print()
print('СХЕМА:')
print('  Основной склад (41.01) → [Перемещение] → АКЗС Витебский (41.01) → [Комплектация] → АКЗС Витебский (41.02)')
print('=' * 80)
