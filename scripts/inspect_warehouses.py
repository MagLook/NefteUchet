# -*- coding: utf-8 -*-
"""Существующие склады АЗС: их тип/вид/реквизиты."""
import win32com.client

BASE = r"D:\Users\magsp\GIG Base2"
USER = "Гайворонская Татьяна"
PWD = "12345"
OUT_PATH = r"D:\Users\magsp\ELSYPLUS\NefteUchet\scripts\_out_warehouses.txt"
OUT = open(OUT_PATH, "w", encoding="utf-8")

def p(*a): print(*a, file=OUT)

def s(v):
    if v is None: return ""
    sv = str(v)
    if "<COMObject" in sv or "unknown" in sv:
        for attr in ("Наименование", "Имя", "Код", "Представление"):
            try:
                a = getattr(v, attr)
                if a is None: continue
                a_s = str(a)
                if a_s and "<COMObject" not in a_s: return a_s
            except Exception: continue
        return ""
    return sv

com = win32com.client.Dispatch("V83.COMConnector")
conn = com.Connect(f'File="{BASE}";Usr="{USER}";Pwd="{PWD}";')

# 1. Все склады АЗС
p("=" * 78)
p("Существующие склады АЗС")
p("=" * 78)
Q = conn.NewObject("Запрос")
Q.Текст = '''
ВЫБРАТЬ Ссылка
ИЗ Справочник.Склады
ГДЕ НЕ ПометкаУдаления
    И (Наименование ПОДОБНО "%АЗС%"
       ИЛИ Наименование ПОДОБНО "%АКЗС%"
       ИЛИ Наименование ПОДОБНО "%АГЗС%")
УПОРЯДОЧИТЬ ПО Наименование
'''
Sel = Q.Выполнить().Выбрать()
n = 0
while Sel.Следующий():
    n += 1
    obj = Sel.Ссылка.ПолучитьОбъект()
    p(f"  {n:>2}. {s(obj.Наименование):<30}")
    # Все реквизиты вкратце
    md = conn.Метаданные.Справочники.Склады
    for rq in md.Реквизиты:
        try:
            v = getattr(obj, rq.Имя)
            sv = s(v)
            if sv and sv != "0" and sv != "False":
                p(f"        {rq.Имя:30} = {sv[:60]}")
        except Exception: pass
p(f"  Всего: {n}")

_skip = '''
# 2. АКЗС Витебский — все реквизиты целиком
p()
p("=" * 78)
p("АКЗС Витебский — все заполненные реквизиты (как пример)")
p("=" * 78)
Скл = conn.Справочники.Склады.НайтиПоНаименованию("АКЗС Витебский", True)
if not Скл.Пустая():
    obj = Скл.ПолучитьОбъект()
    md = conn.Метаданные.Справочники.Склады
    for rq in md.Реквизиты:
        try:
            v = getattr(obj, rq.Имя)
            sv = s(v)
            if sv:
                p(f"  {rq.Имя:35} = {sv[:80]}")
        except Exception: pass

'''
# 3. Все типы складов в Перечислении ТипыСкладов (если есть)
p()
p("=" * 78)
p("Возможные типы складов")
p("=" * 78)
for имя_перечисл in ["ТипыСкладов", "ВидыСкладов"]:
    try:
        Q = conn.NewObject("Запрос")
        Q.Текст = f"ВЫБРАТЬ Ссылка ИЗ Перечисление.{имя_перечисл}"
        Sel = Q.Выполнить().Выбрать()
        p(f"  Перечисление.{имя_перечисл}:")
        while Sel.Следующий():
            p(f"    {s(Sel.Ссылка)}")
    except Exception as e:
        p(f"  Перечисление.{имя_перечисл}: НЕТ ({e})")

OUT.close()
print(f"OK: {OUT_PATH}")
