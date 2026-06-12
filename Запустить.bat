@echo off
title Abonenty
rem Run the server in THIS window so any error stays visible.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0server.ps1"
echo.
echo If no app window opened, screenshot any red text above and send it.
pause
