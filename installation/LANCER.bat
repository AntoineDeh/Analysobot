@echo off
setlocal enabledelayedexpansion

set "INST=%~dp0"
set "ROOT=%~dp0..\"

REM --- Python : venv en priorite ---
set "PYTHON="
if exist "%ROOT%.venv\Scripts\python.exe" set "PYTHON=%ROOT%.venv\Scripts\python.exe"
if not defined PYTHON if exist "%INST%python\python.exe" set "PYTHON=%INST%python\python.exe"
if not defined PYTHON for /f "tokens=*" %%p in ('where python 2^>nul') do if not defined PYTHON set "PYTHON=%%p"

if not defined PYTHON (
    echo Python introuvable. Installer Python ou executer SETUP_USB.bat.
    pause & exit /b 1
)

REM --- Ollama : portable puis systeme ---
set "OLLAMA="
set "PORTABLE_OLLAMA=0"
if exist "%INST%ollama\ollama.exe" (
    set "OLLAMA=%INST%ollama\ollama.exe"
    set "PORTABLE_OLLAMA=1"
)
if not defined OLLAMA if exist "%LOCALAPPDATA%\Programs\Ollama\ollama.exe" set "OLLAMA=%LOCALAPPDATA%\Programs\Ollama\ollama.exe"
if not defined OLLAMA if exist "%LOCALAPPDATA%\Ollama\ollama.exe" set "OLLAMA=%LOCALAPPDATA%\Ollama\ollama.exe"
if not defined OLLAMA for /f "tokens=*" %%o in ('where ollama 2^>nul') do if not defined OLLAMA set "OLLAMA=%%o"

if not defined OLLAMA (
    echo Ollama introuvable. Installer Ollama ou executer SETUP_USB.bat.
    pause & exit /b 1
)

REM --- Demarrage Ollama portable ---
if "%PORTABLE_OLLAMA%"=="1" (
    set OLLAMA_MODELS=%INST%ollama_models
    start /min "Ollama" "!OLLAMA!" serve
    timeout /t 5 /nobreak >nul
)

REM --- Lancer Flask ---
echo.
echo  Ouvrir dans Chrome : http://localhost:5000
echo  Ctrl+C pour arreter.
echo.

cd /d "%ROOT%dev"
"!PYTHON!" analyse_exigences.py

if "%PORTABLE_OLLAMA%"=="1" taskkill /f /im ollama.exe >nul 2>&1
endlocal
