param(
    [switch]$NoCopy,
    [switch]$DumpCfe,
    [switch]$Clean
)

[Console]::InputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# === Пути (можно переопределить через env: TL_PLATFORM_EXE, TL_BASE, TL_EXT, TL_USER, TL_PWD) ===
$platform = if ($env:TL_PLATFORM_EXE) { $env:TL_PLATFORM_EXE } else { "C:\Program Files (x86)\1cv8\8.3.27.2074\bin\1cv8.exe" }
$base     = if ($env:TL_BASE)         { $env:TL_BASE }         else { "D:\Users\magsp\GIG Base2" }
$ext      = if ($env:TL_EXT)          { $env:TL_EXT }          else { "TradeLedger" }
$user     = if ($env:TL_USER)         { $env:TL_USER }         else { "Гайворонская Татьяна" }
$pwd      = if ($env:TL_PWD)          { $env:TL_PWD }          else { "12345" }

$ScriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir
$xmlPath     = Join-Path $ProjectRoot 'xml-v4'
$srcCommon   = Join-Path $ProjectRoot 'src\CommonModules'
$srcDP       = Join-Path $ProjectRoot 'src\DataProcessors'
$BuildDir    = Join-Path $ProjectRoot 'build'

Write-Host "=== TradeLedger dev ===" -ForegroundColor Cyan
Write-Host "База:       $base"
Write-Host "Платформа:  $platform"
Write-Host "XML:        $xmlPath"
Write-Host "Расширение: $ext"
Write-Host ""

# === Шаг 1: Копирование .bsl из src/ в xml-v4/ ===
if (-not $NoCopy) {
    Write-Host "Шаг 1: Копирование .bsl из src/ в xml-v4/..." -ForegroundColor Yellow

    $CopyMap = @(
        @{ Src = (Join-Path $srcCommon 'TL_ApiКлиент.bsl');           Dst = (Join-Path $xmlPath 'CommonModules\TL_ApiКлиент\Ext\Module.bsl');           Name = 'TL_ApiКлиент' },
        @{ Src = (Join-Path $srcCommon 'TL_Настройки.bsl');           Dst = (Join-Path $xmlPath 'CommonModules\TL_Настройки\Ext\Module.bsl');           Name = 'TL_Настройки' },
        @{ Src = (Join-Path $srcCommon 'TL_СозданиеДокументов.bsl');  Dst = (Join-Path $xmlPath 'CommonModules\TL_СозданиеДокументов\Ext\Module.bsl');  Name = 'TL_СозданиеДокументов' },
        @{ Src = (Join-Path $srcCommon 'TL_HTMLГенератор.bsl');       Dst = (Join-Path $xmlPath 'CommonModules\TL_HTMLГенератор\Ext\Module.bsl');       Name = 'TL_HTMLГенератор' },
        @{ Src = (Join-Path $srcCommon 'TL_ПомощьКонтент.bsl');       Dst = (Join-Path $xmlPath 'CommonModules\TL_ПомощьКонтент\Ext\Module.bsl');       Name = 'TL_ПомощьКонтент' },
        @{ Src = (Join-Path $srcCommon 'TL_Маппинг.bsl');             Dst = (Join-Path $xmlPath 'CommonModules\TL_Маппинг\Ext\Module.bsl');             Name = 'TL_Маппинг' },
        @{ Src = (Join-Path $srcCommon 'TL_РегистрСтатусов.bsl');     Dst = (Join-Path $xmlPath 'CommonModules\TL_РегистрСтатусов\Ext\Module.bsl');     Name = 'TL_РегистрСтатусов' },
        @{ Src = (Join-Path $srcCommon 'TL_HTTPКлиентЦБ.bsl');         Dst = (Join-Path $xmlPath 'CommonModules\TL_HTTPКлиентЦБ\Ext\Module.bsl');         Name = 'TL_HTTPКлиентЦБ' },
        @{ Src = (Join-Path $srcCommon 'TL_МаппингЦБ.bsl');            Dst = (Join-Path $xmlPath 'CommonModules\TL_МаппингЦБ\Ext\Module.bsl');            Name = 'TL_МаппингЦБ' },
        @{ Src = (Join-Path $srcCommon 'TL_Сверка.bsl');               Dst = (Join-Path $xmlPath 'CommonModules\TL_Сверка\Ext\Module.bsl');               Name = 'TL_Сверка' },
        @{ Src = (Join-Path $srcCommon 'TL_СопуткаСервис.bsl');        Dst = (Join-Path $xmlPath 'CommonModules\TL_СопуткаСервис\Ext\Module.bsl');        Name = 'TL_СопуткаСервис' },
        @{ Src = (Join-Path $srcDP 'TL_Загрузка\Forms\Форма\Module.bsl');              Dst = (Join-Path $xmlPath 'DataProcessors\TL_Загрузка\Forms\Форма\Ext\Form\Module.bsl');              Name = 'TL_Загрузка форма' },
        @{ Src = (Join-Path $srcDP 'TL_Загрузка\Forms\ФормаДетали\Module.bsl');         Dst = (Join-Path $xmlPath 'DataProcessors\TL_Загрузка\Forms\ФормаДетали\Ext\Form\Module.bsl');         Name = 'TL_Загрузка ФормаДетали' },
        @{ Src = (Join-Path $srcDP 'TL_НастройкаРасширения\Forms\Форма\Module.bsl');   Dst = (Join-Path $xmlPath 'DataProcessors\TL_НастройкаРасширения\Forms\Форма\Ext\Form\Module.bsl');   Name = 'TL_НастройкаРасширения форма' }
    )

    $copied = 0
    foreach ($item in $CopyMap) {
        if (-not (Test-Path $item.Src)) {
            Write-Warning "  НЕТ: $($item.Name) — $($item.Src)"
            continue
        }
        $dstDir = Split-Path -Parent $item.Dst
        if (-not (Test-Path $dstDir)) {
            Write-Warning "  НЕТ целевой папки: $dstDir"
            continue
        }
        Copy-Item $item.Src $item.Dst -Force
        Write-Host "  OK: $($item.Name)" -ForegroundColor Green
        $copied++
    }
    Write-Host "  Скопировано: $copied/$($CopyMap.Count)" -ForegroundColor Cyan
    Write-Host ""
} else {
    Write-Host "Шаг 1: Пропущен (-NoCopy)" -ForegroundColor Gray
    Write-Host ""
}

# === Папка build ===
if (-not (Test-Path $BuildDir)) {
    New-Item -ItemType Directory -Path $BuildDir | Out-Null
}

# === Шаг 1.5 (при -Clean): Удалить расширение из базы ===
if ($Clean) {
    Write-Host "Шаг 1.5: Выгрузка текущего расширения из базы..." -ForegroundColor Magenta
    $dumpDir = Join-Path $BuildDir 'xml-dump-live'
    if (Test-Path $dumpDir) { Remove-Item $dumpDir -Recurse -Force }
    New-Item -ItemType Directory -Path $dumpDir | Out-Null

    $dumpLogPath = Join-Path $BuildDir 'dev_dump.log'
    $dumpArgs = @(
        "DESIGNER",
        "/F", "`"$base`"",
        "/N", "`"$user`"",
        "/P", "`"$pwd`"",
        "/DumpConfigToFiles", "`"$dumpDir`"",
        "-Extension", $ext,
        "/Out", "`"$dumpLogPath`"",
        "/DisableStartupDialogs",
        "/DisableStartupMessages"
    )
    $proc = Start-Process -FilePath $platform -ArgumentList $dumpArgs -Wait -PassThru -NoNewWindow
    if ($proc.ExitCode -eq 0) {
        Write-Host "  Выгружено в $dumpDir" -ForegroundColor Green
    } else {
        Write-Host "  Выгрузка не удалась (код $($proc.ExitCode)) — продолжаем" -ForegroundColor Yellow
    }
    Write-Host ""
}

# === Шаг 2: LoadConfigFromFiles + UpdateDBCfg ===
Write-Host "Шаг 2: Загрузка в 1С и обновление БД..." -ForegroundColor Yellow

$designerArgs = @(
    "DESIGNER",
    "/F", "`"$base`"",
    "/N", "`"$user`"",
    "/P", "`"$pwd`"",
    "/LoadConfigFromFiles", "`"$xmlPath`"",
    "-Extension", $ext,
    "/UpdateDBCfg",
    "/Out", "`"$(Join-Path $BuildDir 'dev_load.log')`"",
    "/DisableStartupDialogs",
    "/DisableStartupMessages"
)

Write-Host "  $platform DESIGNER /F ... /N `"$user`" /LoadConfigFromFiles ... -Extension $ext /UpdateDBCfg" -ForegroundColor Gray
$proc = Start-Process -FilePath $platform -ArgumentList $designerArgs -Wait -PassThru -NoNewWindow

$logPath = Join-Path $BuildDir 'dev_load.log'
if ($proc.ExitCode -ne 0) {
    Write-Host "  ОШИБКА: код возврата $($proc.ExitCode)" -ForegroundColor Red
    if (Test-Path $logPath) {
        Write-Host "  Лог:" -ForegroundColor Red
        [System.IO.File]::ReadAllText($logPath, [System.Text.Encoding]::GetEncoding(1251)) | Write-Host
    }
    exit 1
}
Write-Host "  Загрузка + обновление БД -> OK" -ForegroundColor Green
if (Test-Path $logPath) {
    $logText = [System.IO.File]::ReadAllText($logPath, [System.Text.Encoding]::GetEncoding(1251)).Trim()
    if ($logText) { Write-Host "  Лог: $logText" -ForegroundColor Gray }
}
Write-Host ""

# === Шаг 2.5: Инициализация настроек через COM (py -3-32) ===
Write-Host "Шаг 2.5: Инициализация настроек ГИГ через COM..." -ForegroundColor Yellow
$initPy = Join-Path $ScriptDir 'init_defaults.py'
if (-not (Test-Path $initPy)) {
    # Создаём скрипт инициализации
    @'
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
'@ | Set-Content -Path $initPy -Encoding UTF8
}
$pyProc = Start-Process -FilePath 'py' -ArgumentList @('-3-32', $initPy) -Wait -PassThru -NoNewWindow
if ($pyProc.ExitCode -eq 0) {
    Write-Host "  Инициализация -> OK" -ForegroundColor Green
} else {
    Write-Host "  Инициализация -> ошибка (не критично)" -ForegroundColor Yellow
}
Write-Host ""

# === Шаг 3 (опционально): Выгрузка .cfe ===
if ($DumpCfe) {
    if (-not (Test-Path $BuildDir)) {
        New-Item -ItemType Directory -Path $BuildDir | Out-Null
    }
    $cfeFile = Join-Path $BuildDir "$ext.cfe"

    Write-Host "Шаг 3: Выгрузка .cfe..." -ForegroundColor Yellow
    $dumpArgs = @(
        "DESIGNER",
        "/F", "`"$base`"",
        "/N", "`"$user`"",
        "/P", "`"$pwd`"",
        "/DumpCfg", "`"$cfeFile`"",
        "-Extension", $ext
    )
    $proc = Start-Process -FilePath $platform -ArgumentList $dumpArgs -Wait -PassThru -NoNewWindow
    if ($proc.ExitCode -ne 0) {
        Write-Host "  ОШИБКА выгрузки .cfe: код $($proc.ExitCode)" -ForegroundColor Red
        exit 1
    }
    $size = [math]::Round((Get-Item $cfeFile).Length / 1KB, 1)
    Write-Host "  $cfeFile ($size KB) -> OK" -ForegroundColor Green
    Write-Host ""
}

Write-Host "=== Готово ===" -ForegroundColor Cyan
