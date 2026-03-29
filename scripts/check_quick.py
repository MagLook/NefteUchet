# -*- coding: utf-8 -*-
import sys, win32com.client
sys.stdout.reconfigure(encoding='utf-8')
connector = win32com.client.Dispatch('V83.COMConnector')
conn = connector.Connect('File="D:\\Users\\magsp\\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345";')
q = conn.NewObject('Query')
q.Text = 'SELECT Ключ, Значение FROM InformationRegister.TL_Настройки WHERE Ключ IN ("URLСервера","Логин","КодСистемы") ORDER BY Ключ'
r = q.Execute().Choose()
while r.Next():
    print(str(r.Ключ) + ' = ' + str(r.Значение))
