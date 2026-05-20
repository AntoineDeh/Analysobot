@echo off
echo Arret de SYREQ...
taskkill /f /im ollama.exe >nul 2>&1
taskkill /f /im python.exe >nul 2>&1
echo SYREQ arrete.
timeout /t 2 /nobreak >nul
