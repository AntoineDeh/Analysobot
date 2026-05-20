@echo off
title SYREQ — Preparation cle USB
setlocal enabledelayedexpansion

set "INST=%~dp0"
set "ROOT=%~dp0..\"

cls
echo.
echo  ====================================================
echo   SYREQ — Preparation de la cle USB
echo   Executer CE SCRIPT sur votre machine principale
echo   (celle ou Python et Ollama sont installes)
echo  ====================================================
echo.
echo  Ce script va :
echo    1. Installer Python embeddable dans python\
echo    2. Installer les packages Flask/Ollama/etc.
echo    3. Copier ollama.exe dans ollama\
echo    4. Copier les modeles dans ollama_models\
echo.
echo  Appuyer sur une touche pour continuer...
pause >nul

REM =====================================================
REM ETAPE 1 : Verifier Python embeddable
REM =====================================================
echo.
echo [1/4] Verification de Python embeddable...

if not exist "%INST%python\python.exe" (
    echo.
    echo  Python embeddable non trouve dans installation\python\
    echo.
    echo  Telecharger python-3.12.x-embed-amd64.zip depuis :
    echo    https://www.python.org/downloads/windows/
    echo    (section "Windows embeddable package (64-bit)")
    echo.
    echo  Extraire le contenu du ZIP dans : %INST%python\
    echo.
    echo  Relancer ce script ensuite.
    pause & exit /b 1
)
echo  [OK] Python embeddable trouve.

REM --- Activer site-packages dans le python embeddable ---
for /f "delims=" %%f in ('dir /b /a-d "%INST%python\python*._pth" 2^>nul') do set "PTH=%%f"
if defined PTH (
    powershell -Command ^
        "$f='%INST%python\!PTH!'; $c=Get-Content $f; " ^
        "$c=$c -replace '#import site','import site'; " ^
        "if ($c -notmatch 'Lib\\\\site-packages') { $c += 'Lib\site-packages' }; " ^
        "Set-Content $f $c"
    echo  [OK] site-packages active dans !PTH!
)

REM =====================================================
REM ETAPE 2 : Installer les packages Python
REM =====================================================
echo.
echo [2/4] Installation des packages Python...

if not exist "%INST%python\Scripts\pip.exe" (
    echo  Installation de pip...
    "%INST%python\python.exe" -c "import urllib.request; urllib.request.urlretrieve('https://bootstrap.pypa.io/get-pip.py','%TEMP%\get-pip.py')"
    "%INST%python\python.exe" "%TEMP%\get-pip.py" --no-warn-script-location
    del "%TEMP%\get-pip.py" >nul 2>&1
)

"%INST%python\Scripts\pip.exe" install ^
    flask ^
    ollama ^
    python-docx ^
    openpyxl ^
    pypdf ^
    json-repair ^
    --target="%INST%python\Lib\site-packages" ^
    --no-warn-script-location ^
    --quiet

if errorlevel 1 (
    echo  [ERREUR] Installation des packages echouee.
    echo  Verifier la connexion internet et relancer.
    pause & exit /b 1
)
echo  [OK] Packages installes.

REM =====================================================
REM ETAPE 3 : Copie d'Ollama
REM =====================================================
echo.
echo [3/4] Copie d'Ollama...

if not exist "%INST%ollama" mkdir "%INST%ollama"

set "OLLAMA_SRC="
if exist "%LOCALAPPDATA%\Programs\Ollama\ollama.exe"   set "OLLAMA_SRC=%LOCALAPPDATA%\Programs\Ollama\ollama.exe"
if exist "%LOCALAPPDATA%\Ollama\ollama.exe"            set "OLLAMA_SRC=%LOCALAPPDATA%\Ollama\ollama.exe"
if exist "C:\Program Files\Ollama\ollama.exe"          set "OLLAMA_SRC=C:\Program Files\Ollama\ollama.exe"

if defined OLLAMA_SRC (
    copy /y "!OLLAMA_SRC!" "%INST%ollama\ollama.exe" >nul
    echo  [OK] ollama.exe copie depuis : !OLLAMA_SRC!
) else (
    echo  [ATTENTION] ollama.exe non trouve automatiquement.
    echo  Copier manuellement ollama.exe dans : %INST%ollama\
    echo.
    echo  Ou telecharger ollama-windows-amd64.zip depuis :
    echo    https://github.com/ollama/ollama/releases
    echo  Et placer ollama.exe dans : %INST%ollama\
)

REM --- Copier les DLLs CUDA si disponibles (pour GPU) ---
set "OLLAMA_DIR="
if exist "%LOCALAPPDATA%\Programs\Ollama" set "OLLAMA_DIR=%LOCALAPPDATA%\Programs\Ollama"
if exist "%LOCALAPPDATA%\Ollama"          set "OLLAMA_DIR=%LOCALAPPDATA%\Ollama"
if defined OLLAMA_DIR (
    xcopy /E /I /Y /Q "!OLLAMA_DIR!\lib" "%INST%ollama\lib\" >nul 2>&1
    echo  [OK] Librairies Ollama copiees.
)

REM =====================================================
REM ETAPE 4 : Copie des modeles Ollama
REM =====================================================
echo.
echo [4/4] Copie des modeles Ollama (peut durer plusieurs minutes)...

if not exist "%INST%ollama_models" mkdir "%INST%ollama_models"

set "MODELS_SRC=%USERPROFILE%\.ollama\models"
if exist "!MODELS_SRC!" (
    echo  Source : !MODELS_SRC!
    echo  Destination : %INST%ollama_models\
    echo  Copie en cours...
    xcopy /E /I /Y /Q "!MODELS_SRC!" "%INST%ollama_models\" >nul
    echo  [OK] Modeles copies.
) else (
    echo  [ATTENTION] Aucun modele trouve dans %USERPROFILE%\.ollama\models
    echo  Installer un modele d'abord : ollama pull mistral
    echo  Puis relancer ce script.
)

REM =====================================================
REM Résumé
REM =====================================================
echo.
echo  ====================================================
echo   Preparation terminee !
echo.
echo   Structure de la cle USB :
echo    %ROOT%
echo    ├── LANCER.bat              ← Double-clic pour demarrer
echo    ├── dev\                    ← Code source
echo    └── installation\
echo        ├── python\             ← Python portable
echo        ├── ollama\             ← Ollama portable
echo        └── ollama_models\      ← Modeles IA
echo.
echo   Pour tester : double-clic sur LANCER.bat a la racine
echo  ====================================================
echo.
pause
endlocal
