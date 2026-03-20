[Console]::InputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

param(
    [string]$PlatformExe = '',
    [string]$ConnectionString = '',
    [string]$ExtensionName = 'НУ_НефтеУчёт',
    [switch]$LoadOnly,
    [switch]$UpdateDB
)

# === Определение путей ===
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir
$XmlPath = Join-Path $ProjectRoot 'xml'
$BuildDir = Join-Path $ProjectRoot 'build'

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

$srcCommon = Join-Path $ProjectRoot 'src\CommonModules'
$srcDP = Join-Path $ProjectRoot 'src\DataProcessors'

if (Test-Path "$srcCommon\НУ_ApiКлиент.bsl") {
    Copy-Item "$srcCommon\НУ_ApiКлиент.bsl" "$XmlPath\CommonModules\НУ_ApiКлиент\Ext\Module.bsl" -Force
    Write-Host "  НУ_ApiКлиент.bsl -> OK"
}
if (Test-Path "$srcCommon\НУ_СозданиеДокументов.bsl") {
    Copy-Item "$srcCommon\НУ_СозданиеДокументов.bsl" "$XmlPath\CommonModules\НУ_СозданиеДокументов\Ext\Module.bsl" -Force
    Write-Host "  НУ_СозданиеДокументов.bsl -> OK"
}
if (Test-Path "$srcDP\НУ_Загрузка\Forms\Форма\Module.bsl") {
    Copy-Item "$srcDP\НУ_Загрузка\Forms\Форма\Module.bsl" "$XmlPath\DataProcessors\НУ_Загрузка\Forms\Форма\Ext\Form\Module.bsl" -Force
    Write-Host "  НУ_Загрузка форма -> OK"
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
