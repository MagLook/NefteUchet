# -*- coding: utf-8 -*-
"""Удалить тестовый ОРП с комментарием TL|VAT-TEST|*. Запустить один раз."""
import io, sys, pythoncom, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _tl_config as cfg
pythoncom.CoInitialize()
import win32com.client
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

conn = win32com.client.Dispatch("V83.COMConnector")
ib = cfg.connect(conn)
q = ib.NewObject("Query")
q.Text = """ВЫБРАТЬ Д.Ссылка, Д.Номер
ИЗ Документ.ОтчетОРозничныхПродажах КАК Д
ГДЕ Д.Комментарий ПОДОБНО "TL|VAT-TEST|%" """
r = q.Выполнить().Выбрать()
n = 0
while r.Следующий():
    obj = r.Ссылка.ПолучитьОбъект()
    # сначала отмена проведения
    try: obj.Записать(ib.РежимЗаписиДокумента.ОтменаПроведения)
    except Exception: pass
    obj.Удалить()
    n += 1
    print(f"Удалён ОРП {r.Номер}")
print(f"Итого удалено: {n}")
