# -*- coding: utf-8 -*-
import sys, datetime

try:
    import pythoncom
    pythoncom.CoInitialize()
    import win32com.client
    connector = win32com.client.Dispatch("V83.COMConnector")
    conn = connector.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="' + chr(1043) + chr(1072) + chr(1081) + chr(1074) + chr(1086) + chr(1088) + chr(1086) + chr(1085) + chr(1089) + chr(1082) + chr(1072) + chr(1103) + ' ' + chr(1058) + chr(1072) + chr(1090) + chr(1100) + chr(1103) + chr(1085) + chr(1072) + r'";Pwd="12345"')
    print("OK: connected\n")

    d1 = datetime.datetime(2026, 4, 1)
    d2 = datetime.datetime(2026, 4, 30, 23, 59, 59)

    # --- 1. Отчёты о розничных продажах: заголовки + товары ---
    q = conn.NewObject("Query")
    q.Text = (
        "SELECT "
        "  Doc.Ref AS Ref, Doc.Date AS Dt, Doc.Number AS Num, Doc.Posted AS Posted, "
        "  Doc.Comment AS Comment, "
        "  Items.LineNumber AS LN, "
        "  PRESENTATION(Items.Nomenclature) AS Nom, "
        "  Items.Quantity AS Qty, Items.Price AS Price, Items.Sum AS Summ "
        "FROM "
        "  Document.RetailSalesReport.Goods AS Items "
        "  LEFT JOIN Document.RetailSalesReport AS Doc ON Doc.Ref = Items.Ref "
        "WHERE "
        "  Doc.Date >= &D1 AND Doc.Date <= &D2 "
        "  AND Doc.Comment LIKE " + '"%TL|%"'
        " ORDER BY Doc.Date, Items.LineNumber"
    )
    q.SetParameter("D1", d1)
    q.SetParameter("D2", d2)

    # Fallback: русские имена
    try:
        res = q.Execute()
    except:
        q2 = conn.NewObject("Query")
        # Используем запрос через табличную часть напрямую
        qt = (
            'ВЫБРАТЬ '
            '  Док.Ссылка КАК Ссыл, Док.Дата КАК Дт, Док.Номер КАК Ном, '
            '  Док.Проведен КАК Пров, Док.Комментарий КАК Комм, '
            '  Тов.НомерСтроки КАК НС, '
            '  ПРЕДСТАВЛЕНИЕ(Тов.Номенклатура) КАК Номенкл, '
            '  Тов.Количество КАК Кол, Тов.Цена КАК Цена, Тов.Сумма КАК Сумма '
            'ИЗ '
            '  Документ.ОтчетОРозничныхПродажах.Товары КАК Тов '
            '  ЛЕВОЕ СОЕДИНЕНИЕ Документ.ОтчетОРозничныхПродажах КАК Док '
            '    ПО Док.Ссылка = Тов.Ссылка '
            'ГДЕ '
            '  Док.Дата >= &Д1 И Док.Дата <= &Д2 '
            '  И Док.Комментарий ПОДОБНО "%TL|%" '
            'УПОРЯДОЧИТЬ ПО Док.Дата, Тов.НомерСтроки'
        )
        q2.Text = qt
        q2.SetParameter("\u04141", d1)
        q2.SetParameter("\u04142", d2)
        res = q2.Execute()

    sel = res.Select()
    print("=== OTCHETY O ROZNICHNYKH PRODAZHAKH (April) ===")
    prev = ""
    while sel.Next():
        try:
            num = sel.Ном
            dt = sel.Дт
            posted = sel.Пров
            comment = sel.Комм
            ln = sel.НС
            nom = sel.Номенкл
            qty = sel.Кол
            price = sel.Цена
            summ = sel.Сумма
        except:
            num = sel.Num
            dt = sel.Dt
            posted = sel.Posted
            comment = sel.Comment
            ln = sel.LN
            nom = sel.Nom
            qty = sel.Qty
            price = sel.Price
            summ = sel.Summ

        key = f"{num}|{dt}"
        if key != prev:
            print(f"\n--- #{num} | {dt} | Posted={posted} ---")
            print(f"  {str(comment)[:90]}")
            prev = key
        print(f"  [{ln}] {nom}")
        print(f"      Qty={qty}  Price={price}  Sum={summ}")

except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()
