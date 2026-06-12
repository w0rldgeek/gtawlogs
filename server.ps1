# ============================================================
#  Локальный сервер для приложения «Абоненты»
#  Запускается встроенным в Windows PowerShell — ничего ставить не нужно.
#  Открывает приложение на http://localhost (безопасный контекст:
#  работают доступ к файлу-базе и установка приложения).
# ============================================================
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

$mime = @{
    '.html'='text/html; charset=utf-8'; '.htm'='text/html; charset=utf-8';
    '.js'='text/javascript; charset=utf-8'; '.mjs'='text/javascript; charset=utf-8';
    '.css'='text/css; charset=utf-8'; '.json'='application/json; charset=utf-8';
    '.webmanifest'='application/manifest+json; charset=utf-8';
    '.png'='image/png'; '.jpg'='image/jpeg'; '.jpeg'='image/jpeg';
    '.svg'='image/svg+xml'; '.ico'='image/x-icon';
    '.xlsx'='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    '.woff2'='font/woff2'; '.woff'='font/woff'; '.txt'='text/plain; charset=utf-8';
    '.map'='application/json; charset=utf-8'
}

# Поднимаем TCP-сервер на localhost (не требует прав администратора)
$listener = $null
foreach ($port in 8765, 8080, 8000, 3000) {
    try {
        $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $port)
        $listener.Start()
        break
    } catch { $listener = $null }
}
if (-not $listener) {
    Write-Host "Не удалось занять порт. Закройте другие серверы и повторите." -ForegroundColor Red
    Read-Host "Нажмите Enter для выхода"; exit 1
}

$url = "http://localhost:$port/"
Write-Host ""
Write-Host "  Приложение запущено: $url" -ForegroundColor Green
Write-Host "  Это окно — сервер. Закрытие останавливает приложение." -ForegroundColor DarkGray
Write-Host ""

# Открываем в ОТДЕЛЬНОМ ОКНЕ как приложение (режим --app у Edge/Chrome:
# без вкладок и адресной строки — выглядит как десктоп-программа).
$dataDir = Join-Path $env:LOCALAPPDATA 'AbonentyApp'
$appArgs = @("--app=$url", "--window-size=1200,820", "--user-data-dir=`"$dataDir`"")
$browsers = @(
    "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
)
$launched = $false
foreach ($b in $browsers) {
    if (Test-Path $b) {
        Start-Process -FilePath $b -ArgumentList $appArgs
        $launched = $true
        break
    }
}
if (-not $launched) {
    # Запасной вариант: открыть в браузере по умолчанию
    Start-Process $url
}

$rootFull = [System.IO.Path]::GetFullPath($root)

while ($true) {
    $client = $listener.AcceptTcpClient()
    try {
        $stream = $client.GetStream()
        $stream.ReadTimeout = 5000
        $reader = New-Object System.IO.StreamReader($stream)
        $requestLine = $reader.ReadLine()
        if (-not $requestLine) { $client.Close(); continue }
        while (($line = $reader.ReadLine()) -ne $null -and $line -ne '') { }

        $path = (($requestLine -split ' ')[1] -split '\?')[0]
        if ($path -eq '/' -or [string]::IsNullOrEmpty($path)) { $path = '/index.html' }
        $path = [System.Uri]::UnescapeDataString($path)

        $rel = $path.TrimStart('/').Replace('/', '\')
        $file = [System.IO.Path]::GetFullPath((Join-Path $root $rel))

        if ($file.StartsWith($rootFull) -and (Test-Path $file -PathType Leaf)) {
            $bytes = [System.IO.File]::ReadAllBytes($file)
            $ext = [System.IO.Path]::GetExtension($file).ToLower()
            $ct = $mime[$ext]; if (-not $ct) { $ct = 'application/octet-stream' }
            $head = "HTTP/1.1 200 OK`r`nContent-Type: $ct`r`nContent-Length: $($bytes.Length)`r`nCache-Control: no-cache`r`nConnection: close`r`n`r`n"
        } else {
            $bytes = [System.Text.Encoding]::UTF8.GetBytes('404 Not Found')
            $head = "HTTP/1.1 404 Not Found`r`nContent-Type: text/plain; charset=utf-8`r`nContent-Length: $($bytes.Length)`r`nConnection: close`r`n`r`n"
        }
        $hb = [System.Text.Encoding]::ASCII.GetBytes($head)
        $stream.Write($hb, 0, $hb.Length)
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Flush()
    } catch {
    } finally {
        $client.Close()
    }
}
