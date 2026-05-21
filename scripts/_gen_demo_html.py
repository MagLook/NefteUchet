# -*- coding: utf-8 -*-
"""Сгенерировать демо-HTML детального просмотра пакета смены через
TL_HTMLГенератор.СформироватьДеталиПакетаСмены — для показа заказчику
на встрече. Берёт реальный пакет из C:\\TL_BP_Export\\ который уже
обработан и есть в TL_СверкаПакета (значит есть НСИ-кэш в БП).
"""
import os, sys, glob
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _tl_config import connect

conn = connect()

# Возьмём 2 удачных пакета (ПринятоЦеликом) и один с общепитом
КАТАЛОГ = r"C:\TL_BP_Export"
ВЫХОД = r"D:\Users\magsp\OneDrive\GIG Ledger\Проекты_перехода\АЗС_208_Сопутка_Общепит\demo_html"
os.makedirs(ВЫХОД, exist_ok=True)

# Найдём 3 свежих больших пакета
все = glob.glob(os.path.join(КАТАЛОГ, "*.json"))
print(f"Найдено {len(все)} json пакетов в {КАТАЛОГ}")

# По одному пакету на смену — берём свежий по mtime
import re
по_сменам = {}
for ф in все:
    м = re.search(r"смена-(\d+)", os.path.basename(ф))
    if м:
        ключ = м.group(1)
        if ключ not in по_сменам or os.path.getmtime(ф) > os.path.getmtime(по_сменам[ключ]):
            по_сменам[ключ] = ф

# Сортируем по дате смены (она в имени АЗС208_2026-MM-DD_...)
def дата_имя(п):
    м = re.search(r"(\d{4}-\d{2}-\d{2})", os.path.basename(п))
    return м.group(1) if м else ""

файлы = sorted(по_сменам.values(), key=дата_имя, reverse=True)
print(f"Уникальных смен: {len(файлы)}")
print(f"Свежие 6:")
for ф in файлы[:6]:
    размер = os.path.getsize(ф) / 1024
    print(f"  {os.path.basename(ф)[:80]}... ({размер:.1f} KB)")

# Берём 4 свежие смены (по одной каждая)
сделано = 0
for ф in файлы[:6]:
    try:
        имя_файла = os.path.basename(ф)
        # Имя для заголовка
        заголовок = f"Пакет {имя_файла}"
        # Вызываем функцию TL_HTMLГенератор
        html = conn.TL_HTMLГенератор.СформироватьДеталиПакетаСмены(ф, заголовок)
        if not html or len(html) < 1000:
            print(f"  {имя_файла}: HTML пустой/слишком короткий, пропуск")
            continue
        # Сохраняем
        выход_имя = имя_файла.replace(".json", ".html")
        путь_выхода = os.path.join(ВЫХОД, выход_имя)
        with open(путь_выхода, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"  ✓ {имя_файла} → {выход_имя} ({len(html)/1024:.1f} KB HTML)")
        сделано += 1
        if сделано >= 4:
            break
    except Exception as e:
        print(f"  ✗ {os.path.basename(ф)}: {type(e).__name__}: {str(e)[:80]}")

print(f"\nГотово: {сделано} HTML-файлов в {ВЫХОД}")
