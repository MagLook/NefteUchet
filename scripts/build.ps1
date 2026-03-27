param(
    [string]$PlatformExe = '',
    [string]$ConnectionString = '',
    [string]$ExtensionName = 'TradeLedger',
    [string]$XmlPath = '',
    [switch]$LoadOnly,
    [switch]$UpdateDB
)

[Console]::InputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# === Определение путей ===
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir
$BuildDir = Join-Path $ProjectRoot 'build'
$srcCommon = Join-Path $ProjectRoot 'src\CommonModules'
$srcDP = Join-Path $ProjectRoot 'src\DataProcessors'

if (-not $XmlPath) {
    $XmlPath = Join-Path $ProjectRoot 'xml-v4'
}

# === Поиск платформы 1С ===
if (-not $PlatformExe) {
    $SearchPaths = @(
        'C:\Program Files\1cv8\*\bin\1cv8.exe',
        'C:\Program Files (x86)\1cv8\*\bin\1cv8.exe',
        'C:\Program Files (x86)\1cv8t\*\bin\1cv8.exe'
    )
    foreach ($pattern in $SearchPaths) {
        $found = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue |
                 Sort-Object FullName -Descending |
                 Select-Object -First 1
        if ($found) {
            $PlatformExe = $found.FullName
            break
        }
    }
    if (-not $PlatformExe) {
        Write-Error "Платформа 1С не найдена. Укажите -PlatformExe"
        exit 1
    }
}

Write-Host "Платформа: $PlatformExe" -ForegroundColor Cyan
Write-Host "XML:       $XmlPath" -ForegroundColor Cyan
Write-Host "Расширение: $ExtensionName" -ForegroundColor Cyan

# === Копирование .bsl из src/ в xml/ ===
Write-Host "`nКопирование .bsl модулей из src/ в xml/..." -ForegroundColor Yellow

if (-not (Test-Path $XmlPath)) {
    Write-Error "XML-путь не найден: $XmlPath"
    exit 1
}

$CopyMap = @(
    @{
        Source = Join-Path $srcCommon 'TL_ApiКлиент.bsl'
        Target = Join-Path $XmlPath 'CommonModules\TL_ApiКлиент\Ext\Module.bsl'
        Label = 'TL_ApiКлиент.bsl'
    },
    @{
        Source = Join-Path $srcCommon 'TL_Настройки.bsl'
        Target = Join-Path $XmlPath 'CommonModules\TL_Настройки\Ext\Module.bsl'
        Label = 'TL_Настройки.bsl'
    },
    @{
        Source = Join-Path $srcCommon 'TL_СозданиеДокументов.bsl'
        Target = Join-Path $XmlPath 'CommonModules\TL_СозданиеДокументов\Ext\Module.bsl'
        Label = 'TL_СозданиеДокументов.bsl'
    },
    @{
        Source = Join-Path $srcCommon 'TL_HTMLГенератор.bsl'
        Target = Join-Path $XmlPath 'CommonModules\TL_HTMLГенератор\Ext\Module.bsl'
        Label = 'TL_HTMLГенератор.bsl'
    },
    @{
        Source = Join-Path $srcCommon 'TL_Маппинг.bsl'
        Target = Join-Path $XmlPath 'CommonModules\TL_Маппинг\Ext\Module.bsl'
        Label = 'TL_Маппинг.bsl'
    },
    @{
        Source = Join-Path $srcCommon 'TL_РегистрСтатусов.bsl'
        Target = Join-Path $XmlPath 'CommonModules\TL_РегистрСтатусов\Ext\Module.bsl'
        Label = 'TL_РегистрСтатусов.bsl'
    },
    @{
        Source = Join-Path $srcDP 'TL_Загрузка\Forms\Форма\Module.bsl'
        Target = Join-Path $XmlPath 'DataProcessors\TL_Загрузка\Forms\Форма\Ext\Form\Module.bsl'
        Label = 'TL_Загрузка форма'
    },
    @{
        Source = Join-Path $srcDP 'TL_НастройкаРасширения\Forms\Форма\Module.bsl'
        Target = Join-Path $XmlPath 'DataProcessors\TL_НастройкаРасширения\Forms\Форма\Ext\Form\Module.bsl'
        Label = 'TL_НастройкаРасширения форма'
    }
)

foreach ($Item in $CopyMap) {
    if (-not (Test-Path $Item.Source)) {
        Write-Warning "  Исходник не найден: $($Item.Source)"
        continue
    }

    $TargetDir = Split-Path -Parent $Item.Target
    if (-not (Test-Path $TargetDir)) {
        Write-Warning "  Целевая папка не найдена: $TargetDir"
        continue
    }

    Copy-Item $Item.Source $Item.Target -Force
    Write-Host "  $($Item.Label) -> OK"
}

# === Проверка строки подключения ===
if (-not $ConnectionString) {
    Write-Host "`nСтрока подключения не указана." -ForegroundColor Red
    Write-Host "Укажите -ConnectionString, например:" -ForegroundColor Yellow
    Write-Host '  .\build.ps1 -ConnectionString "/F D:\Bases\TestBP30"' -ForegroundColor Gray
    Write-Host '  .\build.ps1 -ConnectionString "/S localhost\TestBP30"' -ForegroundColor Gray
    Write-Host "`nБез строки подключения доступна только копирование .bsl."
    exit 0
}

$baseArgs = @('DESIGNER', $ConnectionString)

# === Создание папки build ===
if (-not (Test-Path $BuildDir)) {
    New-Item -ItemType Directory -Path $BuildDir | Out-Null
}

# === Шаг 1: Загрузка XML в расширение ===
Write-Host "`nШаг 1: Загрузка XML в расширение..." -ForegroundColor Green
$loadArgs = $baseArgs + @('/LoadConfigFromFiles', $XmlPath, '-Extension', $ExtensionName)
Write-Host "  $PlatformExe $($loadArgs -join ' ')" -ForegroundColor Gray

$proc = Start-Process -FilePath $PlatformExe -ArgumentList $loadArgs -Wait -PassThru -NoNewWindow
if ($proc.ExitCode -ne 0) {
    Write-Error "Ошибка загрузки XML (код $($proc.ExitCode))"
    exit 1
}
Write-Host "  Загрузка XML -> OK" -ForegroundColor Green

# === Шаг 2: Обновление БД (если указано) ===
if ($UpdateDB) {
    Write-Host "`nШаг 2: Обновление конфигурации БД..." -ForegroundColor Green
    $updateArgs = $baseArgs + @('/UpdateDBCfg', '-Extension', $ExtensionName)
    $proc = Start-Process -FilePath $PlatformExe -ArgumentList $updateArgs -Wait -PassThru -NoNewWindow
    if ($proc.ExitCode -ne 0) {
        Write-Error "Ошибка обновления БД (код $($proc.ExitCode))"
        exit 1
    }
    Write-Host "  Обновление БД -> OK" -ForegroundColor Green
}

# === Шаг 3: Выгрузка .cfe (если не LoadOnly) ===
if (-not $LoadOnly) {
    $cfeFile = Join-Path $BuildDir "$ExtensionName.cfe"
    Write-Host "`nШаг 3: Выгрузка .cfe..." -ForegroundColor Green
    $dumpArgs = $baseArgs + @('/DumpCfg', $cfeFile, '-Extension', $ExtensionName)
    $proc = Start-Process -FilePath $PlatformExe -ArgumentList $dumpArgs -Wait -PassThru -NoNewWindow
    if ($proc.ExitCode -ne 0) {
        Write-Error "Ошибка выгрузки .cfe (код $($proc.ExitCode))"
        exit 1
    }
    $size = (Get-Item $cfeFile).Length / 1KB
    Write-Host "  Выгрузка -> $cfeFile ($([math]::Round($size, 1)) KB)" -ForegroundColor Green
}

Write-Host "`nГотово!" -ForegroundColor Cyan
