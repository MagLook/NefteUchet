# -*- coding: utf-8 -*-
"""Детальная проверка состояния после загрузки пользователем."""
import win32com.client
import re

OUT = open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\scripts\_out_state2.txt", "w", encoding="utf-8")
def p(*a): print(*a, file=OUT)

def s(v):
    if v is None: return ""
    sv = str(v)
    if "<COMObject" in sv or "unknown" in sv:
        for attr in ("Представление","Наименование","Имя","Код","Номер"):
            try:
                a = getattr(v, attr)
                if a is None: continue
                a_s = str(a)
                if a_s and "<COMObject" not in a_s: return a_s
            except Exception: continue
        return ""
    return sv

com = win32com.client.Dispatch("V83.COMConnector")
conn = com.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

p("=" * 78)
p("Все TL-документы по типам и проведению")
p("=" * 78)
для_типов = ["ОтчетОРозничныхПродажах","ПеремещениеТоваров","КомплектацияНоменклатуры","ПриходныйКассовыйОрдер"]
все = {}
для_ст = {}
for тип in для_типов:
    Q = conn.NewObject("Запрос")
    Q.УстановитьПараметр("Ш", "TL|%")
    Q.Текст = 'ВЫБРАТЬ Ссылка, Номер, Дата, Комментарий, Проведен ИЗ Документ.' + тип + ' ГДЕ ВЫРАЗИТЬ(Комментарий КАК СТРОКА(200)) ПОДОБНО &Ш УПОРЯДОЧИТЬ ПО Дата'
    Sel = Q.Выполнить().Выбрать()
    список = []
    while Sel.Следующий():
        список.append({
            "ссылка": Sel.Ссылка,
            "номер": str(Sel.Номер),
            "дата": Sel.Дата,
            "комм": str(Sel.Комментарий or "").strip(),
            "провед": bool(Sel.Проведен),
        })
    все[тип] = список
    провед = sum(1 for x in список if x["провед"])
    p(f"  {тип:30}: всего={len(список):>3}  провед={провед:>3}")

p()
p("=" * 78)
p("Непроведённые документы — детально")
p("=" * 78)
for тип, список in все.items():
    непровед = [x for x in список if not x["провед"]]
    if непровед:
        p(f"\n{тип} ({len(непровед)} непроведённых):")
        for d in непровед:
            p(f"  #{d['номер']}  {d['дата']}  ком={d['комм'][:80]}")

p()
p("=" * 78)
p("Группировка по базовому ключу (TL|СМЕНА|сеть|стн|номер)")
p("=" * 78)
по_ключу = {}
for тип, список in все.items():
    for d in список:
        m = re.match(r'^(TL\|СМЕНА\|\d+\|\d+\|\d+)', d["комм"])
        if m:
            ключ = m.group(1)
        else:
            m = re.match(r'^(TL\|ТТН\|\d+\|\d+\|\d+(?:\|\d+)?)', d["комм"])
            ключ = m.group(1) if m else d["комм"][:30]
        по_ключу.setdefault(ключ, {"типы": {}}).setdefault("типы", {}).setdefault(тип, []).append(d)

for ключ in sorted(по_ключу.keys()):
    д = по_ключу[ключ]["типы"]
    итого = sum(len(v) for v in д.values())
    провед_всего = sum(1 for доки in д.values() for x in доки if x["провед"])
    статус = "✓" if итого == провед_всего else f"⚠ {итого-провед_всего} непров"
    p(f"\n{ключ}  ({итого} документов, {статус}):")
    for тип, доки in д.items():
        for d in доки:
            знак = "✓" if d["провед"] else "✗"
            p(f"  {знак} {тип:30} #{d['номер']:<14} {d['дата']}")

OUT.close()
print("OK: scripts/_out_state2.txt")
