@echo off
title WGHT - API Python + Supabase (port 8000)
cd /d "%~dp0.."
echo.
echo Demarrage de l'API Python (FastAPI + Supabase)...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\run-supabase-server.ps1"
set ERR=%ERRORLEVEL%
if %ERR% NEQ 0 (
    echo.
    echo ERREUR - l'API n'a pas demarre. Lisez le message ci-dessus.
    echo.
)
pause
