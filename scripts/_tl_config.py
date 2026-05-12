# -*- coding: utf-8 -*-
"""Конфигурация подключения к 1С для probe/smoke/regression скриптов.

Переопределяется через переменные окружения:
- TL_BASE         путь к информационной базе 1С (по умолчанию: D:\\Users\\magsp\\GIG Base2)
- TL_USER         имя пользователя 1С
- TL_PWD          пароль 1С
- TL_PLATFORM_EXE путь к 1cv8.exe
- TL_EXT          имя расширения (TradeLedger)

На MAG_2025 переменные обычно не выставлены → используется дефолт.
На 1c-dev-01 переменные выставляются через PowerShell profile (.gig-ledger.env.ps1).
"""
import os

BASE_PATH    = os.environ.get("TL_BASE",  r"D:\Users\magsp\GIG Base2")
USER         = os.environ.get("TL_USER",  "Гайворонская Татьяна")
PWD          = os.environ.get("TL_PWD",   "12345")
PLATFORM_EXE = os.environ.get("TL_PLATFORM_EXE",
                              r"C:\Program Files (x86)\1cv8\8.3.27.2074\bin\1cv8.exe")
EXT_NAME     = os.environ.get("TL_EXT",   "TradeLedger")

def connect(connector=None):
    """Удобная функция подключения через COM. Возвращает Connection."""
    if connector is None:
        import win32com.client
        connector = win32com.client.Dispatch("V83.COMConnector")
    return connector.Connect(f'File="{BASE_PATH}";Usr="{USER}";Pwd="{PWD}";')

if __name__ == "__main__":
    import sys, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    print(f"BASE_PATH    = {BASE_PATH}")
    print(f"USER         = {USER}")
    print(f"PWD          = {'*' * len(PWD)}")
    print(f"PLATFORM_EXE = {PLATFORM_EXE}")
    print(f"EXT_NAME     = {EXT_NAME}")
