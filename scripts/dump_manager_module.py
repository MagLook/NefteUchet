# -*- coding: utf-8 -*-
"""Dump ОтчетОРозничныхПродажах МодульМенеджера around line 3893. Run: py -3.13-32 scripts/dump_manager_module.py"""
import pythoncom, io
pythoncom.CoInitialize()
import win32com.client

conn = win32com.client.Dispatch("V83.COMConnector")
ib = conn.Connect(r'File="D:\Users\magsp\GIG Base2";Usr="Гайворонская Татьяна";Pwd="12345"')

out = io.open(r"D:\Users\magsp\ELSYPLUS\NefteUchet\build\manager_module.txt", "w", encoding="utf-8")

# Get module text via metadata
md = ib.Метаданные.Документы.Найти("ОтчетОРозничныхПродажах")
if md:
    # МодульМенеджера
    try:
        mod_text = md.МодульМенеджера.Текст if hasattr(md, 'МодульМенеджера') else None
        if mod_text:
            lines = mod_text.split("\n")
            out.write(f"Всего строк: {len(lines)}\n\n")
            # Lines around 3893
            start = max(0, 3880)
            end = min(len(lines), 3910)
            for i in range(start, end):
                marker = " >>>" if i+1 == 3893 else "    "
                out.write(f"{i+1:5}{marker} {lines[i]}\n")
        else:
            out.write("МодульМенеджера.Текст не доступен\n")
    except Exception as e:
        out.write(f"Ошибка: {e}\n")

    # Also try МодульОбъекта
    try:
        mod_obj = md.МодульОбъекта
        if mod_obj:
            text = mod_obj.Текст if hasattr(mod_obj, 'Текст') else None
            if text:
                lines2 = text.split("\n")
                out.write(f"\n\nМодульОбъекта: {len(lines2)} строк\n")
                start2 = max(0, 3880)
                end2 = min(len(lines2), 3910)
                for i in range(start2, end2):
                    marker = " >>>" if i+1 == 3893 else "    "
                    out.write(f"{i+1:5}{marker} {lines2[i]}\n")
    except Exception as e:
        out.write(f"МодульОбъекта ошибка: {e}\n")
else:
    out.write("Документ не найден в метаданных\n")

out.close()
print("Done")
