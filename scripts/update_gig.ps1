[Console]::InputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$platform = "C:\Program Files (x86)\1cv8\8.3.27.2074\bin\1cv8.exe"
$base = "D:\Users\magsp\GIG Base2"
$xml = "D:\Users\magsp\ELSYPLUS\NefteUchet\xml-v4"
$ext = "TradeLedger"
$user = "Гайворонская Татьяна"
$pwd = "12345"

Write-Host "Загрузка XML в расширение $ext..." -ForegroundColor Cyan
Write-Host "База: $base"
Write-Host "XML:  $xml"

$args = @(
    "DESIGNER",
    "/F", $base,
    "/N", $user,
    "/P", $pwd,
    "/LoadConfigFromFiles", $xml,
    "-Extension", $ext,
    "/UpdateDBCfg"
)

$proc = Start-Process -FilePath $platform -ArgumentList $args -Wait -PassThru -NoNewWindow
Write-Host "Код возврата: $($proc.ExitCode)" -ForegroundColor $(if ($proc.ExitCode -eq 0) {"Green"} else {"Red"})
