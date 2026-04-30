# -*- coding: utf-8 -*-
import sys, datetime

try:
    import pythoncom
    pythoncom.CoInitialize()
    import win32com.client
    connector = win32com.client.Dispatch("V83.COMConnector")
    conn = connector.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="' + '\u0413\u0430\u0439\u0432\u043e\u0440\u043e\u043d\u0441\u043a\u0430\u044f \u0422\u0430\u0442\u044c\u044f\u043d\u0430' + r'";Pwd="12345"')
    print("OK\n")

    d1 = datetime.datetime(2026, 4, 1)
    d2 = datetime.datetime(2026, 4, 30, 23, 59, 59)

    q = conn.NewObject("Query")
    q.Text = (
        '\u0412\u042b\u0411\u0420\u0410\u0422\u042c '  # ВЫБРАТЬ
        '  \u0422\u043e\u0432.\u041a\u043e\u043b\u0438\u0447\u0435\u0441\u0442\u0432\u043e \u041a\u0410\u041a Kol, '  # Тов.Количество КАК Kol
        '  \u0422\u043e\u0432.\u0426\u0435\u043d\u0430 \u041a\u0410\u041a Cena, '  # Тов.Цена КАК Cena
        '  \u0422\u043e\u0432.\u0421\u0443\u043c\u043c\u0430 \u041a\u0410\u041a Summa, '  # Тов.Сумма КАК Summa
        '  \u041f\u0420\u0415\u0414\u0421\u0422\u0410\u0412\u041b\u0415\u041d\u0418\u0415(\u0422\u043e\u0432.\u041d\u043e\u043c\u0435\u043d\u043a\u043b\u0430\u0442\u0443\u0440\u0430) \u041a\u0410\u041a Nom, '  # ПРЕДСТАВЛЕНИЕ(Тов.Номенклатура) КАК Nom
        '  \u0414\u043e\u043a.\u041d\u043e\u043c\u0435\u0440 \u041a\u0410\u041a DocNum, '  # Док.Номер КАК DocNum
        '  \u0414\u043e\u043a.\u0414\u0430\u0442\u0430 \u041a\u0410\u041a DocDate '  # Док.Дата КАК DocDate
        '\u0418\u0417 '  # ИЗ
        '  \u0414\u043e\u043a\u0443\u043c\u0435\u043d\u0442.\u041e\u0442\u0447\u0435\u0442\u041e\u0420\u043e\u0437\u043d\u0438\u0447\u043d\u044b\u0445\u041f\u0440\u043e\u0434\u0430\u0436\u0430\u0445.\u0422\u043e\u0432\u0430\u0440\u044b \u041a\u0410\u041a \u0422\u043e\u0432 '  # Документ.ОтчетОРозничныхПродажах.Товары КАК Тов
        '  \u041b\u0415\u0412\u041e\u0415 \u0421\u041e\u0415\u0414\u0418\u041d\u0415\u041d\u0418\u0415 \u0414\u043e\u043a\u0443\u043c\u0435\u043d\u0442.\u041e\u0442\u0447\u0435\u0442\u041e\u0420\u043e\u0437\u043d\u0438\u0447\u043d\u044b\u0445\u041f\u0440\u043e\u0434\u0430\u0436\u0430\u0445 \u041a\u0410\u041a \u0414\u043e\u043a '  # ЛЕВОЕ СОЕДИНЕНИЕ ... КАК Док
        '    \u041f\u041e \u0414\u043e\u043a.\u0421\u0441\u044b\u043b\u043a\u0430 = \u0422\u043e\u0432.\u0421\u0441\u044b\u043b\u043a\u0430 '  # ПО Док.Ссылка = Тов.Ссылка
        '\u0413\u0414\u0415 '  # ГДЕ
        '  \u0414\u043e\u043a.\u0414\u0430\u0442\u0430 >= &D1 '
        '  \u0418 \u0414\u043e\u043a.\u0414\u0430\u0442\u0430 <= &D2 '
        '  \u0418 \u0414\u043e\u043a.\u041f\u0440\u043e\u0432\u0435\u0434\u0435\u043d '  # И Док.Проведен
        '  \u0418 \u0414\u043e\u043a.\u041a\u043e\u043c\u043c\u0435\u043d\u0442\u0430\u0440\u0438\u0439 \u041f\u041e\u0414\u041e\u0411\u041d\u041e "%TL|%" '
        '\u0423\u041f\u041e\u0420\u042f\u0414\u041e\u0427\u0418\u0422\u042c \u041f\u041e \u0414\u043e\u043a.\u0414\u0430\u0442\u0430'
    )
    q.SetParameter("D1", d1)
    q.SetParameter("D2", d2)
    sel = q.Execute().Select()

    print(f"{'Doc':>12} {'Date':>12} {'Nom':>20} {'Qty':>12} {'Price':>12} {'Sum':>12} {'Check':>12} {'Diff':>8}")
    print("-" * 100)
    while sel.Next():
        qty = float(sel.Kol)
        price = float(sel.Cena)
        summ = float(sel.Summa)
        check = round(qty * price, 2)
        diff = round(summ - check, 2)
        print(f"{str(sel.DocNum):>12} {str(sel.DocDate)[:10]:>12} {str(sel.Nom):>20} {qty:>12.4f} {price:>12.4f} {summ:>12.2f} {check:>12.2f} {diff:>8.2f}")

except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()
