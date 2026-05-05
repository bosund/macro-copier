@echo off
REM VBA Macro Copier - Nuitka Build Script
REM Kompilerer macro_copier.py til standalone executable

setlocal enabledelayedexpansion

echo.
echo ========================================
echo VBA Macro Copier - Nuitka Build Script
echo =========================================
echo.

REM 1. Check Python Installation
echo [1/4] Tjekker Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo   X Python ikke fundet!
    exit /b 1
)
for /f "tokens=*" %%i in ('python --version 2^>^&1') do set pythonVersion=%%i
echo   * Python found: %pythonVersion%

REM 2. Install/Update Nuitka og afhængigheder
echo [2/4] Installerer Nuitka og afhængigheder...
pip install --upgrade nuitka ttkbootstrap >nul 2>&1
if errorlevel 1 (
    echo   X Fejl ved installation af afhængigheder
    exit /b 1
)
echo   * Nuitka og afhængigheder installeret

REM 3. Compile med Nuitka
echo [3/4] Kompilerer med Nuitka onefile...
echo   Dette kan tage 1-3 minutter...

if exist "build" rmdir /s /q "build" 2>nul
if exist "dist" rmdir /s /q "dist" 2>nul

python -m nuitka ^
    --onefile ^
    --output-dir=dist ^
    --enable-plugin=tk-inter ^
    --windows-console-mode=disable ^
    --windows-icon-from-ico=icon.ico ^
    macro_copier.py

if errorlevel 1 (
    echo   X Fejl ved kompilering
    exit /b 1
)
echo   * Kompilering gennemfort!

REM 4. Test af executable
echo [4/4] Tester executable...

set exePath=dist\vba_macro_copier.exe

if not exist "%exePath%" (
    echo   X Executable ikke fundet!
    exit /b 1
)

for %%A in ("%exePath%") do set fileSize=%%~zA
set /a fileSizeMB=%fileSize% / 1048576

echo   * Executable oprettet: %exePath%
echo   * Filstørrelse: %fileSizeMB% MB

echo   Starter applikation for test...
timeout /t 1 /nobreak >nul
start "" /w "%exePath%" >nul 2>&1

echo   * Applikation startede uden fejl

echo.
echo ========================================
echo ^* Build gennemfort!
echo Executable: %exePath%
echo ========================================
echo.
