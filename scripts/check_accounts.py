import pythoncom
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch('V83.COMConnector')
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

q = ib.NewObject('Query')
q.Text = (
    'SELECT TOP 10 '
    '  Номенклатура.Наименование AS Наим,'
    '  СчетУчетаБУ.Код AS СчетБУ '
    'FROM РегистрСведений.СчетаУчетаНоменклатуры '
    'WHERE Номенклатура.Наименование LIKE "%92%" '
    '   OR Номенклатура.Наименование LIKE "%95%" '
    '   OR Номенклатура.Наименование LIKE "%ДТ%" '
    '   OR Номенклатура.Наименование LIKE "%Дизель%"'
)

try:
    res = q.Execute().Choose()
    found = 0
    while res.Next():
        print(f"  {res.Наим} -> счёт {res.СчетБУ}")
        found += 1
    if found == 0:
        print("НЕТ записей в СчетаУчетаНоменклатуры для топлива")
        print("Нужно настроить: АИ-92(т)->41.01, АИ-92(л)->41.02 и т.д.")
except Exception as e:
    print(f"Ошибка запроса: {e}")
    print("Попробуем просто список номенклатуры...")
    q2 = ib.NewObject('Query')
    q2.Text = (
        'SELECT TOP 10 Наименование '
        'FROM Справочник.Номенклатура '
        'WHERE Наименование LIKE "%92%" '
        '   OR Наименование LIKE "%ДТ%"'
    )
    res2 = q2.Execute().Choose()
    while res2.Next():
        print(f"  Номенклатура: {res2.Наименование}")
