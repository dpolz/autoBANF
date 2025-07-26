@echo off
echo ====================================
echo AutoBANF Setup Script
echo ====================================
echo.
echo Dieses Script richtet die Python-Umgebung für AutoBANF ein.
echo.

REM Prüfen ob Python verfügbar ist
python --version >nul 2>&1
if errorlevel 1 (
    echo FEHLER: Python ist nicht installiert oder nicht im PATH verfügbar!
    echo Bitte installieren Sie Python 3.8 oder höher von python.org
    echo.
    pause
    exit /b 1
)

echo Python gefunden:
python --version
echo.

REM Altes venv löschen falls vorhanden
if exist venv (
    echo Lösche vorhandenes Virtual Environment...
    rmdir /s /q venv
)

REM Neues venv erstellen
echo Erstelle neues Virtual Environment...
python -m venv venv
if errorlevel 1 (
    echo FEHLER: Virtual Environment konnte nicht erstellt werden!
    pause
    exit /b 1
)

REM venv aktivieren
echo Aktiviere Virtual Environment...
call venv\Scripts\activate.bat

REM Pip upgraden
echo Upgrade pip...
python -m pip install --upgrade pip

REM Requirements installieren
echo Installiere Python-Pakete...
pip install playwright pandas openpyxl cryptography
if errorlevel 1 (
    echo FEHLER: Python-Pakete konnten nicht installiert werden!
    pause
    exit /b 1
)

REM Playwright Browser installieren
echo Installiere Playwright Browser (Chromium)...
playwright install chromium
if errorlevel 1 (
    echo WARNUNG: Playwright Browser konnte nicht installiert werden!
    echo Das Programm funktioniert möglicherweise nicht korrekt.
)

echo.
echo ====================================
echo Setup erfolgreich abgeschlossen!
echo ====================================
echo.
echo Sie können jetzt AutoBANF verwenden:
echo   python autoBANF.py templates\mouser.xlsx
echo.
echo Hinweise:
echo - Bei der ersten Ausführung werden Sie nach SAP-Anmeldedaten gefragt
echo - Diese werden sicher verschlüsselt im .credentials Ordner gespeichert
echo - Screenshots werden im .temp Ordner erstellt
echo.
pause