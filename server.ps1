# Local web server for the "Abonenty" app.
# Opens the app in its own window (Edge/Chrome --app mode).
# Pure ASCII on purpose, so encoding never breaks parsing.

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProgressPreference = 'SilentlyContinue'

$mime = @{
    '.html'='text/html; charset=utf-8'; '.htm'='text/html; charset=utf-8';
    '.js'='text/javascript; charset=utf-8'; '.mjs'='text/javascript; charset=utf-8';
    '.css'='text/css; charset=utf-8'; '.json'='application/json; charset=utf-8';
    '.webmanifest'='application/manifest+json; charset=utf-8';
    '.png'='image/png'; '.jpg'='image/jpeg'; '.jpeg'='image/jpeg';
    '.svg'='image/svg+xml'; '.ico'='image/x-icon';
    '.xlsx'='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    '.woff2'='font/woff2'; '.woff'='font/woff'; '.txt'='text/plain; charset=utf-8'
}

try {
    Write-Host ""
    Write-Host "  Starting Abonenty server..." -ForegroundColor Cyan

    $listener = $null
    $port = 0
    foreach ($p in 8765,8080,8000,3000,5500) {
        try {
            $l = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $p)
            $l.Start()
            $listener = $l; $port = $p; break
        } catch { $listener = $null }
    }
    if (-not $listener) {
        Write-Host "  ERROR: no free port (8765/8080/8000/3000/5500 busy)." -ForegroundColor Red
        Read-Host "  Press Enter to exit"; return
    }

    $url = "http://localhost:$port/"
    Write-Host "  Running: $url" -ForegroundColor Green
    Write-Host "  This window IS the server. Closing it stops the app." -ForegroundColor DarkGray
    Write-Host ""

    # Open app in its OWN window (no tabs / no address bar)
    $browsers = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
    )
    $opened = $false
    foreach ($b in $browsers) {
        if (Test-Path $b) {
            Start-Process -FilePath $b -ArgumentList "--app=$url", "--window-size=1180,820"
            $opened = $true; break
        }
    }
    if (-not $opened) {
        Write-Host "  Edge/Chrome not found - opening in default browser." -ForegroundColor Yellow
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
            if ($requestLine) {
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
                    $head = "HTTP/1.1 404 Not Found`r`nContent-Type: text/plain`r`nContent-Length: $($bytes.Length)`r`nConnection: close`r`n`r`n"
                }
                $hb = [System.Text.Encoding]::ASCII.GetBytes($head)
                $stream.Write($hb, 0, $hb.Length)
                $stream.Write($bytes, 0, $bytes.Length)
                $stream.Flush()
            }
        } catch {
        } finally {
            $client.Close()
        }
    }
} catch {
    Write-Host ""
    Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Read-Host "  Press Enter to exit"
}
