@echo off
title Abonenty - local server
echo.
echo   Starting the "Abonenty" app...
echo   A browser window will open automatically.
echo   Keep this window open while you work.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0server.ps1"
echo.
echo   Server stopped.
pause
