@echo off
setlocal enabledelayedexpansion

:: ============================================================================
:: Anleitung:
:: Dieses Batch-Skript konfiguriert und startet das msg_file_renamer Modul.
:: Es werden verschiedene Optionen über die Kommandozeile definiert, die darüber
:: entscheiden, ob ein Testlauf durchgeführt wird, ob Dateidaten angepasst werden,
:: ob PDF-Dateien aus MSG-Dateien erzeugt und ggf. überschrieben werden und ob der
:: Suchpfad rekursiv abgearbeitet wird.
::
:: Mögliche Argumente:
::   - -ntr / --no_test_run:       Testlauf aktivieren (keine echten File-Änderungen)
::   - -fd / --set_filedate:        Dateidatum anpassen (False = Dateidatum nicht ändern)
::   - -sd / --search_directory:    Such-Verzeichnis für MSG-Dateien
::   - -spn / --no_shorten_path_name: Kein Kürzen langer Pfadnamen
::   - -pdf / --generate_pdf:       PDF-Erzeugung aus MSG-Dateien aktivieren
::   - -opdf / --overwrite_pdf:     Bereits vorhandene PDF-Dateien überschreiben (wenn aktiviert)
::   - -rs / --recursive_search:    Rekursive Suche im Zielverzeichnis (optional)
::
:: Die einzutragenden Einstellungen können während des Skriptablaufs interaktiv
:: festgelegt oder der Default-Wert verwendet werden.
:: ============================================================================

:: =============================================
:: Konfiguration ? HIER Defaultpfad festlegen
:: =============================================
set "DEFAULT_ZIELPFAD=D:\@MyWorkLocal\lgl-email-beispiele"

:: =============================================
:: Konfiguration ? Defaultwerte vorbelegen
:: =============================================
set "NOTESTLAUF="
set "SETFILEDATE="
set "GENERATEPDF="
set "OVERWRITEPDF="
set "NOSHORTENPATHNAME="

:: Verzeichnis der Batch-Datei ermitteln
set "SCRIPT_DIR=%~dp0"
set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"  :: letztes Backslash entfernen

:: =======================================================
:: 1. Benutzerfrage: Default-Konfiguration verwenden
:: =======================================================
echo =====================================================
echo   *** Bitte Konfiguration prüfen ***
echo       - Testlauf (Keine File-Änderung!)
echo       - Keine rekursive Suche
echo       - Dateinamen werden gekürzt
echo       - Keine Erzeugung von PDF-Dateien
echo       - Zielverzeichnis für MSG-Suche erforderlich
echo         oder Default-Zielverzeichnis DEFAULT_ZIELPFAD
echo =====================================================
echo.

:ABFRAGE1
set /p ANTWORT="Soll die Default-Konfiguration verwendet oder abgebrochen werden? (J/N/A)":
if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else if /i "%ANTWORT%"=="J" (
    echo.
    echo Du hast mit JA geantwortet.
    rem Abfrage der Konfiguration überspringen...
    goto :ZIELPFADPRE
) else if /i "%ANTWORT%"=="N" (
    echo.
    rem Weiter mit dem Skript...
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ungültige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE1
)

:: =======================================================
:: 2. Benutzerfrage: Testlauf?
:: =======================================================
echo =====================================================
echo   *** Testlauf? ***
echo =====================================================
echo.

:Abfrage2
set /p ANTWORT="Möchtest du einen Testlauf durchführen? (J/N/A)"
if /i "%ANTWORT%"=="N" (
    set "NOTESTLAUF=--no_test_run"
    echo .
    echo Folgendes Argument wird übernommen: !NOTESTLAUF!
    rem Weiter mit dem Skript...
) else if /i "%ANTWORT%"=="J" (
    echo.
    echo Es wird ein Testlauf ohne Änderungen an Dateien durchgeführt.
    rem Einige Abfragen können übersprungen werden...
    goto :Abfrage6pre
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ungültige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE2
)

:: =======================================================
:: 3. Benutzerfrage: Filedate?
:: =======================================================
echo.
echo =====================================================
echo   *** Filedate? ***
echo =====================================================
echo.

:Abfrage3
set /p ANTWORT="Möchtest du das Dateidatum anpassen? (J/N/A)"
if /i "%ANTWORT%"=="J" (
    set "SETFILEDATE=--set_filedate"
    echo .
    echo Folgendes Argument wird übernommen: !SETFILEDATE!
    rem Weiter mit dem Skript...
) else if /i "%ANTWORT%"=="N" (
    echo.
    echo Das Filedatum der Dateien wird nicht geändert.
    rem Weiter mit dem Skript...
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ungültige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE3
)

:: =======================================================
:: 4. Benutzerfrage: PDF-Erzeugung?
:: =======================================================
echo.
echo =====================================================
echo   *** PDF-Erzeugung? ***
echo =====================================================
echo.

:Abfrage4
set /p ANTWORT="Möchtest du zusätzliche PDF-Dateien erzeugen? (J/N/A)"
if /i "%ANTWORT%"=="J" (
    set "GENERATEPDF=--generate_pdf"
    echo .
    echo Folgendes Argument wird übernommen: !GENERATEPDF!
    rem Weiter mit dem Skript...
) else if /i "%ANTWORT%"=="N" (
    echo.
    echo Es werden keine zusätzliche PDF-Dateien erzeugt.
    rem Die nächste Abfrage ist dann überflüssig
    goto :Abfrage6pre
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ungültige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE4
)

:: =======================================================
:: 5. Benutzerfrage: PDF-Dateien überschreiben?
:: =======================================================
echo.
echo =====================================================
echo   *** PDF-Dateien überschreiben? ***
echo =====================================================
echo.

:Abfrage5
set /p ANTWORT="Möchtest du bestehende PDF-Dateien überschreien? (J/N/A)"
if /i "%ANTWORT%"=="J" (
    set "OVERWRITEPDF=--overwrite_pdf"
    echo .
    echo Folgendes Argument wird übernommen: !OVERWRITEPDF!
    rem Weiter mit dem Skript...
) else if /i "%ANTWORT%"=="N" (
    echo.
    echo Es werden keine zusätzliche PDF-Dateien erzeugt.
    rem Weiter mit dem Skript...
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ungültige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE5
)

:Abfrage6pre
:: =======================================================
:: 6. Benutzerfrage: Dateinamen kürzen?
:: =======================================================
echo.
echo =====================================================
echo   *** Dateinamen kürzen? ***
echo =====================================================
echo.

:Abfrage6
set /p ANTWORT="Möchtest du die Dateinamen kürzen? (J/N/A)"
if /i "%ANTWORT%"=="N" (
    set "NOSHORTENPATHNAME=--no_shorten_path_name"
    echo .
    echo Folgendes Argument wird übernommen: !NOSHORTENPATHNAME!
    rem Weiter mit dem Skript...
) else if /i "%ANTWORT%"=="J" (
    echo.
    echo Es werden Datenamen gekürzt.
    rem Weiter mit dem Skript...
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ungültige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE6
)

:: =======================================================
:: 7. Benutzerfrage: Rekursive Suche?
:: =======================================================
echo.
echo =====================================================
echo   *** Rekursive Suche? ***
echo =====================================================
echo.

:Abfrage7
set /p ANTWORT="Möchtest du das Verzeichnis rekursiv durchsuchen? (J/N/A)"
if /i "%ANTWORT%"=="J" (
    set "RECURSIVESEARCH=--recursive_search"
    echo .
    echo Folgendes Argument wird übernommen: !RECURSIVESEARCH!
    rem Weiter mit dem Skript...
) else if /i "%ANTWORT%"=="N" (
    echo.
    echo Es erfolgt keine rekursive Suche.
    rem Weiter mit dem Skript...
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ungültige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE7
)

:ZIELPFADPRE
:: =============================================
:: 8. Benutzerfrage: Standardpfad verwenden?
:: =============================================
echo.
echo ===============================================
echo   *** Zielverzeichnis-Auswahl ***
echo ===============================================
echo.
echo Standard-Zielverzeichnis:
echo   %DEFAULT_ZIELPFAD%
echo.

:ZIELPFAD
set /p ANTWORT="Möchtest du diesen Pfad verwenden? (J/N/A)"

if /i "%ANTWORT%"=="J" (
    set "ZIELPFAD=%DEFAULT_ZIELPFAD%"
    echo.
    rem Weiter mit dem Skript...
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else if /i "%ANTWORT%"=="N" (
    echo.
    set /p ZIELPFAD="Bitte gib den gewünschten Zielpfad ein:"
    rem Weiter mit dem Skript...
) else (
    echo.
    echo Ungültige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ZIELPFAD
)

:: =============================================
:: 9. Zielpfad prüfen
:: =============================================
if not exist "%ZIELPFAD%\" (
    echo.
    echo ? Fehler: Das Verzeichnis "%ZIELPFAD%" existiert NICHT.
    echo Der Vorgang wird abgebrochen.
    goto :ENDE
)

:: =============================================
:: 10. Verarbeitung starten
:: =============================================
echo.
echo ? Das Zielverzeichnis existiert.
echo.
echo Starte Verarbeitung im Verzeichnis:
echo   "%ZIELPFAD%"
echo.

:: Aktiviert die virtuelle Python-Umgebung
call "venv\Scripts\activate"

echo python msg_file_renamer.py %NOTESTLAUF% %SETFILEDATE% %GENERATEPDF% %OVERWRITEPDF% %RECURSIVESEARCH% --search_directory "%ZIELPFAD%" --excel_log_directory "./"
python "msg_file_renamer.py" %NOTESTLAUF% %SETFILEDATE% %GENERATEPDF% %OVERWRITEPDF% %RECURSIVESEARCH% --search_directory "%ZIELPFAD%" --excel_log_directory "./"

:: =============================================
:: 11. Hinweis auf Logdatei im Batch-Verzeichnis
:: =============================================
echo -----------------------------------------------
echo Die Verarbeitung wurde abgeschlossen.
echo Die Excel-Logdatei wurde im Verzeichnis der Batch-Datei abgelegt:
echo   "%SCRIPT_DIR%"
echo -----------------------------------------------

:ENDE
echo.
pause