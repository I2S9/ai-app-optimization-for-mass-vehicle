@echo off
title WGHT - Database (BD)
cd /d "%~dp0"
echo.
echo Demarrage du serveur WGHT...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run-bd-server.ps1"
set ERR=%ERRORLEVEL%
if %ERR% NEQ 0 (
    echo.
    echo ERREUR - le serveur na pas demarre. Lisez le message rouge ci-dessus.
    echo.
)
pause
