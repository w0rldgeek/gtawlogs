@echo off
rem Запуск приложения «Абоненты» в отдельном окне.
rem Сервер стартует свёрнутым; приложение открывается как десктоп-окно.
start "Abonenty server" /min powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0server.ps1"
exit
