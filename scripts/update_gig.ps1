[Console]::InputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# === Совместимость: делегирует в dev.ps1 ===
Write-Host "update_gig.ps1 -> dev.ps1" -ForegroundColor Yellow

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$devScript = Join-Path $scriptDir 'dev.ps1'

if (-not (Test-Path $devScript)) {
    Write-Error "Не найден dev.ps1: $devScript"
    exit 1
}

& $devScript
exit $LASTEXITCODE
