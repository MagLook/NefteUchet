$ErrorActionPreference = "SilentlyContinue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$root = "D:\Users\magsp\ELSYPLUS\NefteUchet\src"

function Get-ModuleName($file) {
    $dir = $file.DirectoryName
    $name = $file.BaseName

    if ($dir -match 'CommonModules$') {
        return $name
    }
    if ($dir -match 'CommonModules\\([^\\]+)') {
        return $matches[1]
    }
    if ($dir -match 'DataProcessors\\([^\\]+)(\\Forms\\([^\\]+))?') {
        if ($matches[3]) {
            return "DP." + $matches[1] + ".Form." + $matches[3]
        }
        return "DP." + $matches[1]
    }
    if ($dir -match 'Documents\\([^\\]+)') {
        return "Doc." + $matches[1]
    }
    if ($dir -match 'Catalogs\\([^\\]+)') {
        return "Cat." + $matches[1]
    }
    return $name
}

$defined = @{}
foreach ($file in Get-ChildItem -Path $root -Recurse -Filter "*.bsl") {
    $moduleName = Get-ModuleName $file
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    $content = [System.Text.Encoding]::UTF8.GetString($bytes)
    $lines = $content -split "`r?`n"
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ($line -match '^\s*(Процедура|Функция)\s+([\p{L}_]\w*)') {
            $name = $matches[2]
            $key = $moduleName + "." + $name
            if (-not $defined.ContainsKey($key)) {
                $defined[$key] = $true
            }
        }
    }
}

$calls = @{}
foreach ($file in Get-ChildItem -Path $root -Recurse -Filter "*.bsl") {
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    $content = [System.Text.Encoding]::UTF8.GetString($bytes)
    $lines = $content -split "`r?`n"
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ($line -match '^\s*//') { continue }
        $regex = [regex]::Matches($line, '(TL_[\p{L}_]\w*)\.([\p{L}_]\w*)\s*\(')
        foreach ($m in $regex) {
            $modulez = $m.Groups[1].Value
            $method = $m.Groups[2].Value
            $key = $modulez + "." + $method
            if (-not $calls.ContainsKey($key)) {
                $calls[$key] = @()
            }
            $calls[$key] += [PSCustomObject]@{
                File = $file.FullName
                Line = $i + 1
                Text = $line.Trim()
            }
        }
    }
}

$badCalls = 0
foreach ($key in $calls.Keys | Sort-Object) {
    if (-not $defined.ContainsKey($key)) {
        $badCalls++
        Write-Host ""
        Write-Host ("BROKEN: " + $key)
        foreach ($loc in $calls[$key]) {
            $relPath = $loc.File.Substring($root.Length + 1)
            Write-Host ("  " + $relPath + ":" + $loc.Line + "  " + $loc.Text)
        }
    }
}
Write-Host ""
Write-Host ("=== Total broken calls: " + $badCalls + " ===")
