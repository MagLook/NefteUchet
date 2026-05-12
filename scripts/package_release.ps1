# -*- coding: utf-8 -*-
# Скрипт упаковки релиза v6.0.0
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Version = "6.0.0"
$Root = "D:\Users\magsp\ELSYPLUS\NefteUchet"
$Dst = Join-Path $Root "release\v$Version"

if (-not (Test-Path $Dst)) {
    New-Item -ItemType Directory -Path $Dst -Force | Out-Null
}

# Копируем .cfe
$Cfe = Join-Path $Root "build\TradeLedger.cfe"
if (-not (Test-Path $Cfe)) {
    Write-Host "ERROR: $Cfe не найден" -ForegroundColor Red
    exit 1
}
Copy-Item $Cfe $Dst -Force
$CfeOut = Join-Path $Dst "TradeLedger.cfe"
$Size = [math]::Round((Get-Item $CfeOut).Length / 1KB, 1)
Write-Host "OK: TradeLedger.cfe ($Size KB)" -ForegroundColor Green

# Покажем содержимое релиза
Write-Host "`nСодержимое $Dst :"
Get-ChildItem $Dst | Format-Table Name, Length, LastWriteTime
