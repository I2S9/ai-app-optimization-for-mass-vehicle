@echo off
title WGHT - Supabase (cloud, port 8000)
cd /d "%~dp0"
echo.
echo Demarrage de l'application WGHT avec Supabase (persistance cloud)...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run-supabase-server.ps1"
set ERR=%ERRORLEVEL%
if %ERR% NEQ 0 (
    echo.
    echo ERREUR - le serveur n'a pas demarre. Lisez le message rouge ci-dessus.
    echo.
)
pause
