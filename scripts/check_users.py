# -*- coding: utf-8 -*-
"""Анализ прав расширения у каждого пользователя."""
import win32com.client

OUT = open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\scripts\_out_users.txt", "w", encoding="utf-8")
def p(*a): print(*a, file=OUT)

com = win32com.client.Dispatch("V83.COMConnector")
conn = com.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')

p("=" * 78)
p("Пользователи и роль TL_ОсновнаяРоль")
p("=" * 78)
p(f"  {'Имя':<35}  ролей  тип        TL")

МП = conn.ПользователиИнформационнойБазы.ПолучитьПользователей()
бухгалтеры = []
для_назначения = []
for У in МП:
    Имя = str(У.Имя)
    Полное = str(У.ПолноеИмя)
    КолРолей = 0
    есть_админ = False
    есть_TL = False
    try:
        for р in У.Роли:
            КолРолей += 1
            ИмяРоли = str(р.Имя)
            if 'ПолныеПрава' in ИмяРоли or 'Администрат' in ИмяРоли:
                есть_админ = True
            if ИмяРоли.startswith('TL_'):
                есть_TL = True
    except Exception: pass

    если_бухгалтер = (50 < КолРолей < 300 and not есть_админ)
    тип = "АДМИН" if есть_админ else ("БУХГАЛТЕР" if если_бухгалтер else "другое")
    TL = "ДА" if есть_TL else "НЕТ"
    p(f"  {Имя:<35}  {КолРолей:>5}  {тип:<10}  {TL}")
    if если_бухгалтер:
        бухгалтеры.append((Имя, Полное, КолРолей, есть_TL))
    if (not есть_TL) and КолРолей > 0:
        для_назначения.append(Имя)

p()
p(f"Найдено {len(бухгалтеры)} бухгалтеров")

OUT.close()
print("OK")
