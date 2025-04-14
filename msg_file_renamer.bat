@echo off
setlocal enabledelayedexpansion

:: ============================================================================
:: Anleitung:
:: Dieses Batch-Skript konfiguriert und startet das msg_file_renamer Modul.
:: Es werden verschiedene Optionen �ber die Kommandozeile definiert, die dar�ber
:: entscheiden, ob ein Testlauf durchgef�hrt wird, ob Dateidaten angepasst werden,
:: ob PDF-Dateien aus MSG-Dateien erzeugt und ggf. �berschrieben werden und ob der
:: Suchpfad rekursiv abgearbeitet wird.
::
:: M�gliche Argumente:
::   - -ntr / --no_test_run:       Testlauf aktivieren (keine echten File-�nderungen)
::   - -fd / --set_filedate:        Dateidatum anpassen (False = Dateidatum nicht �ndern)
::   - -sd / --search_directory:    Such-Verzeichnis f�r MSG-Dateien
::   - -spn / --no_shorten_path_name: Kein K�rzen langer Pfadnamen
::   - -pdf / --generate_pdf:       PDF-Erzeugung aus MSG-Dateien aktivieren
::   - -opdf / --overwrite_pdf:     Bereits vorhandene PDF-Dateien �berschreiben (wenn aktiviert)
::   - -rs / --recursive_search:    Rekursive Suche im Zielverzeichnis (optional)
::
:: Die einzutragenden Einstellungen k�nnen w�hrend des Skriptablaufs interaktiv
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
echo   *** Bitte Konfiguration pr�fen ***
echo       - Testlauf (Keine File-�nderung!)
echo       - Keine rekursive Suche
echo       - Dateinamen werden gek�rzt
echo       - Keine Erzeugung von PDF-Dateien
echo       - Zielverzeichnis f�r MSG-Suche erforderlich
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
    rem Abfrage der Konfiguration �berspringen...
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
    echo Ung�ltige Eingabe. Bitte J, N oder A eingeben.
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
set /p ANTWORT="M�chtest du einen Testlauf durchf�hren? (J/N/A)"
if /i "%ANTWORT%"=="N" (
    set "NOTESTLAUF=--no_test_run"
    echo .
    echo Folgendes Argument wird �bernommen: !NOTESTLAUF!
    rem Weiter mit dem Skript...
) else if /i "%ANTWORT%"=="J" (
    echo.
    echo Es wird ein Testlauf ohne �nderungen an Dateien durchgef�hrt.
    rem Einige Abfragen k�nnen �bersprungen werden...
    goto :Abfrage6pre
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ung�ltige Eingabe. Bitte J, N oder A eingeben.
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
set /p ANTWORT="M�chtest du das Dateidatum anpassen? (J/N/A)"
if /i "%ANTWORT%"=="J" (
    set "SETFILEDATE=--set_filedate"
    echo .
    echo Folgendes Argument wird �bernommen: !SETFILEDATE!
    rem Weiter mit dem Skript...
) else if /i "%ANTWORT%"=="N" (
    echo.
    echo Das Filedatum der Dateien wird nicht ge�ndert.
    rem Weiter mit dem Skript...
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ung�ltige Eingabe. Bitte J, N oder A eingeben.
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
set /p ANTWORT="M�chtest du zus�tzliche PDF-Dateien erzeugen? (J/N/A)"
if /i "%ANTWORT%"=="J" (
    set "GENERATEPDF=--generate_pdf"
    echo .
    echo Folgendes Argument wird �bernommen: !GENERATEPDF!
    rem Weiter mit dem Skript...
) else if /i "%ANTWORT%"=="N" (
    echo.
    echo Es werden keine zus�tzliche PDF-Dateien erzeugt.
    rem Die n�chste Abfrage ist dann �berfl�ssig
    goto :Abfrage6pre
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ung�ltige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE4
)

:: =======================================================
:: 5. Benutzerfrage: PDF-Dateien �berschreiben?
:: =======================================================
echo.
echo =====================================================
echo   *** PDF-Dateien �berschreiben? ***
echo =====================================================
echo.

:Abfrage5
set /p ANTWORT="M�chtest du bestehende PDF-Dateien �berschreien? (J/N/A)"
if /i "%ANTWORT%"=="J" (
    set "OVERWRITEPDF=--overwrite_pdf"
    echo .
    echo Folgendes Argument wird �bernommen: !OVERWRITEPDF!
    rem Weiter mit dem Skript...
) else if /i "%ANTWORT%"=="N" (
    echo.
    echo Es werden keine zus�tzliche PDF-Dateien erzeugt.
    rem Weiter mit dem Skript...
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ung�ltige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE5
)

:Abfrage6pre
:: =======================================================
:: 6. Benutzerfrage: Dateinamen k�rzen?
:: =======================================================
echo.
echo =====================================================
echo   *** Dateinamen k�rzen? ***
echo =====================================================
echo.

:Abfrage6
set /p ANTWORT="M�chtest du die Dateinamen k�rzen? (J/N/A)"
if /i "%ANTWORT%"=="N" (
    set "NOSHORTENPATHNAME=--no_shorten_path_name"
    echo .
    echo Folgendes Argument wird �bernommen: !NOSHORTENPATHNAME!
    rem Weiter mit dem Skript...
) else if /i "%ANTWORT%"=="J" (
    echo.
    echo Es werden Datenamen gek�rzt.
    rem Weiter mit dem Skript...
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ung�ltige Eingabe. Bitte J, N oder A eingeben.
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
set /p ANTWORT="M�chtest du das Verzeichnis rekursiv durchsuchen? (J/N/A)"
if /i "%ANTWORT%"=="J" (
    set "RECURSIVESEARCH=--recursive_search"
    echo .
    echo Folgendes Argument wird �bernommen: !RECURSIVESEARCH!
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
    echo Ung�ltige Eingabe. Bitte J, N oder A eingeben.
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
set /p ANTWORT="M�chtest du diesen Pfad verwenden? (J/N/A)"

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
    set /p ZIELPFAD="Bitte gib den gew�nschten Zielpfad ein:"
    rem Weiter mit dem Skript...
) else (
    echo.
    echo Ung�ltige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ZIELPFAD
)

:: =============================================
:: 9. Zielpfad pr�fen
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