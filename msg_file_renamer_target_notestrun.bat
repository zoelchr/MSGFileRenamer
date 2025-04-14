@echo off
setlocal enabledelayedexpansion

:: =============================================
:: Konfiguration – HIER Defaultpfad festlegen
:: =============================================
set "DEFAULT_ZIELPFAD=D:\@MyWorkLocal\lgl-email-beispiele"

:: Verzeichnis der Batch-Datei ermitteln
set "SCRIPT_DIR=%~dp0"
set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"  :: letztes Backslash entfernen

:: =======================================================
:: 1. Benutzerfrage: Konfiguration des Aurufs prüfen?
:: =======================================================
echo =====================================================
echo   *** Bitte Konfiguration prüfen ***
echo       - !!!KEIN!!! Testlauf (File-Änderung!!!)
echo       - keine PDF-Erzeugung
echo       - Zielverzeichnis für MSG-Suche erforderlich
echo         oder Default-Zielverzeichnis DEFAULT_ZIELPFAD
echo =====================================================
echo.
set /p ANTWORT="Soll diese Konfiguration verwender werden? (J/N)":

if /i "%ANTWORT%"=="J" (
    echo.
) else (
    echo.
    echo "Der Vorgang wird abgebrochen."
    goto :ENDE
)

:: =============================================
:: 1. Benutzerfrage: Standardpfad verwenden?
:: =============================================
echo ===============================================
echo   *** Zielverzeichnis-Auswahl ***
echo ===============================================
echo.
echo Standard-Zielverzeichnis:
echo   %DEFAULT_ZIELPFAD%
echo.

set /p ANTWORT="Möchtest du diesen Pfad verwenden?" (J/N):

if /i "%ANTWORT%"=="J" (
    set "ZIELPFAD=%DEFAULT_ZIELPFAD%"
) else (
    echo.
    set /p ZIELPFAD="Bitte gib den gewünschten Zielpfad ein:"
)

:: =============================================
:: 2. Zielpfad prüfen
:: =============================================
if not exist "%ZIELPFAD%\" (
    echo.
    echo ❌ Fehler: Das Verzeichnis "%ZIELPFAD%" existiert NICHT.
    echo Der Vorgang wird abgebrochen.
    goto :ENDE
)

:: =============================================
:: 3. Verarbeitung starten
:: =============================================
echo.
echo ✅ Das Zielverzeichnis existiert.
echo Starte Verarbeitung im Verzeichnis:
echo   "%ZIELPFAD%"
echo.

:: Aktiviert die virtuelle Python-Umgebung
call "venv\Scripts\activate"

echo python msg_file_renamer.py --no_test_run --search_directory "%ZIELPFAD%" --excel_log_directory "./"
python "msg_file_renamer.py" --search_directory "%ZIELPFAD%" --no_test_run --excel_log_directory "./"

:: =============================================
:: 4. Hinweis auf Logdatei im Batch-Verzeichnis
:: =============================================
echo -----------------------------------------------
echo Die Verarbeitung wurde abgeschlossen.
echo Die Excel-Logdatei wurde im Verzeichnis der Batch-Datei abgelegt:
echo   "%SCRIPT_DIR%"
echo -----------------------------------------------

:ENDE
echo.
pause