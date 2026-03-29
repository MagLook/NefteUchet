# -*- coding: utf-8 -*-
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')
connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')
if not conn.TL_Настройки.НастройкиЗаполнены():
    conn.TL_Настройки.ИнициализироватьПоУмолчаниюГИГ()
    print('  Настройки инициализированы')
else:
    print('  Настройки уже заполнены')
