param()

[Console]::InputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$updateScript = Join-Path $scriptDir 'update_gig.ps1'

Write-Host "export_update_import.ps1 переведён в режим совместимости." -ForegroundColor Yellow
Write-Host "Используется актуальный сценарий update_gig.ps1 -> TradeLedger/xml-v4." -ForegroundColor Yellow

if (-not (Test-Path $updateScript)) {
    Write-Error "Не найден update_gig.ps1: $updateScript"
    exit 1
}

& $updateScript
exit $LASTEXITCODE
