@echo off
REM chcp 1252>nul
chcp 850
setlocal enabledelayedexpansion

REM ============================================================================
REM Anleitung:
REM Dieses Batch-Skript konfiguriert und startet das msg_file_renamer Modul.
REM Es werden verschiedene Optionen über die Kommandozeile definiert, die darueber
REM entscheiden, ob ein Testlauf durchgeführt wird, ob Dateidaten angepasst werden,
REM ob PDF-Dateien aus MSG-Dateien erzeugt und ggf. überschrieben werden und ob der
REM Suchpfad rekursiv abgearbeitet wird.
REM
REM Mögliche Argumente:
REM   -ntr / --no_test_run:          Testlauf aktivieren (keine echten File-Aenderungen)
REM   -fd / --set_filedate:          Dateidatum anpassen (False = Dateidatum nicht ändern)
REM   -sd / --search_directory:      Such-Verzeichnis für MSG-Dateien
REM   -spn / --no_shorten_path_name: Kein Kuerzen langer Pfadnamen
REM   -pdf / --generate_pdf:         PDF-Erzeugung aus MSG-Dateien aktivieren
REM   -opdf / --overwrite_pdf:       Bereits vorhandene PDF-Dateien überschreiben (wenn aktiviert)
REM   -rs / --recursive_search:      Rekursive Suche im Zielverzeichnis (optional)
REM   -ucf / --use_knownsender_file: CSV-Datei mit Liste den bekannten Sender nutzen (wenn aktiviert)
REM   -cf / --knownsender_file:      CSV-Datei mit Liste der bekannten Absender (Default: ".\config\known_senders_private.csv")
REM
REM Die einzutragenden Einstellungen können während des Skriptablaufs interaktiv
REM festgelegt oder der Default-Wert verwendet werden.
REM
REM Kurze Erlaeuterung zum Bau des Release:
REM
REM Hier der Befehl: pyinstaller --onefile --console msg_file_renamer.py
REM
REM --onefile: Diese Option sorgt dafür, dass PyInstaller dein Python-Skript zu einer einzigen .exe-Datei zusammenpackt.
REM --console: Diese Option sorgt dafür, dass die Ausgabe in der Konsole erfolgt.
REM --add-data: Diese Option sorgt dafür, dass zusaetzliche Dateien in die .exe-Datei eingebunden werden.
REM
REM ============================================================================

REM =============================================
REM Konfigurationen
REM =============================================
set "DEFAULT_ZIELPFAD=.\tests\functional\testdir"
set PYTHONIOENCODING=utf-8
set "NOTESTLAUF="
set "SETFILEDATE="
set "GENERATEPDF="
set "OVERWRITEPDF="
set "NOSHORTENPATHNAME="
set "RECURSIVESEARCH="
set "USEKNOWNSENDERFILE=--use_knownsender_file"
set "KNOWNSENDERFILE=.\config\known_senders.csv"

REM Verzeichnis der Batch-Datei ermitteln
set "SCRIPT_DIR=%~dp0"
set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"  REM letztes Backslash entfernen
echo Das Skript wurde im Verzeichnis "%SCRIPT_DIR%" gestartet.

:ABFRAGE_DEFAULT_KONFIG_1
echo =====================================================
echo   *** Bitte Konfiguration pruefen ***
echo       - Testlauf (Keine File-Aenderung!)
echo       - Keine rekursive Suche
echo       - Dateinamen werden gekuerzt
echo       - Keine Erzeugung von PDF-Dateien
echo       - CSV-Datei mit Liste bekannter Absender nutzen?
echo       - CSV-Datei mit Liste der bekannten Absender
echo         (Default: ".\config\known_senders_private.csv")
echo       - Zielverzeichnis fuer MSG-Suche erforderlich
echo         oder Default-Zielverzeichnis DEFAULT_ZIELPFAD
echo =====================================================
echo.
:ABFRAGE_DEFAULT_KONFIG_2
set /p ANTWORT="Soll die Default-Konfiguration verwendet oder abgebrochen werden? (J/N/A) "
if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else if /i "%ANTWORT%"=="J" (
    echo.
    echo Du hast mit JA geantwortet.
    rem Abfrage der Konfiguration teilweise überspringen...
    goto :ABFRAGE_STANDARDPFAD_1
) else if /i "%ANTWORT%"=="N" (
    echo.
    rem Weiter mit dem Skript...
    goto :ABFRAGE_DEFAULT_KONFIG_3
) else (
    echo.
    echo Ungueltige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE_DEFAULT_KONFIG_2
)
:ABFRAGE_DEFAULT_KONFIG_3

:ABFRAGE_TESTLAUF_1
echo =====================================================
echo   *** Testlauf? ***
echo =====================================================
echo.
:ABFRAGE_TESTLAUF_2
set /p ANTWORT="Moechtest du einen Testlauf durchfuehren? (J/N/A) "
if /i "%ANTWORT%"=="N" (
    set "NOTESTLAUF=--no_test_run"
    echo.
    echo Folgendes Argument wird uebernommen: !NOTESTLAUF!
    rem Weiter mit dem Skript...
    goto :ABFRAGE_TESTLAUF_3
) else if /i "%ANTWORT%"=="J" (
    echo.
    echo Es wird ein Testlauf ohne Aenderungen an Dateien durchgefuehrt.
    goto :ABFRAGE_DATEINAMEN_KUERZEN_1
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ungueltige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE_TESTLAUF_2
)
:ABFRAGE_TESTLAUF_3

:ABFRAGE_DATEINAMEN_KUERZEN_1
echo.
echo =====================================================
echo   *** Dateinamen kuerzen? ***
echo =====================================================
echo.
:ABFRAGE_DATEINAMEN_KUERZEN_2
set /p ANTWORT="Moechtest du die Dateinamen kuerzen? (J/N/A) "
if /i "%ANTWORT%"=="N" (
    set "NOSHORTENPATHNAME=--no_shorten_path_name"
    echo.
    rem Weiter mit dem Skript...
    goto :ABFRAGE_DATEINAMEN_KUERZEN_3
) else if /i "%ANTWORT%"=="J" (
    echo.
    echo Es werden Datenamen gekuerzt.
    rem Weiter mit dem Skript...
    goto :ABFRAGE_DATEINAMEN_KUERZEN_3
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ungueltige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE_DATEINAMEN_KUERZEN_2
)
:ABFRAGE_DATEINAMEN_KUERZEN_3

:ABFRAGE_REKURSIVE_SUCHE_1
echo.
echo =====================================================
echo   *** Rekursive Suche? ***
echo =====================================================
echo.
:ABFRAGE_REKURSIVE_SUCHE_2
set /p ANTWORT="Moechtest du das Verzeichnis rekursiv durchsuchen? (J/N/A) "
if /i "%ANTWORT%"=="J" (
    set "RECURSIVESEARCH=--recursive_search"
    echo.
    echo Folgendes Argument wird uebernommen: !RECURSIVESEARCH!
    rem Weiter mit dem Skript...
    goto :ABFRAGE_REKURSIVE_SUCHE_3
) else if /i "%ANTWORT%"=="N" (
    echo.
    echo Es erfolgt keine rekursive Suche.
    rem Weiter mit dem Skript...
    goto :ABFRAGE_REKURSIVE_SUCHE_3
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ungueltige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE_REKURSIVE_SUCHE_2
)
:ABFRAGE_REKURSIVE_SUCHE_3

REM Die folgenden Abfragen nur durchführen, wenn es sich nicht um einen Testlauf handelt
if /i "%NOTESTLAUF%"=="" (
    echo Only a Test Run
    goto :ABFRAGE_STANDARDPFAD_1
)

:ABFRAGE_FILEDATE_1
echo.
echo =====================================================
echo   *** Filedate? ***
echo =====================================================
echo.
:ABFRAGE_FILEDATE_2
set /p ANTWORT="Moechtest du das Dateidatum anpassen? (J/N/A) "
if /i "%ANTWORT%"=="J" (
    set "SETFILEDATE=--set_filedate"
    echo.
    echo Folgendes Argument wird uebernommen: !SETFILEDATE!
    rem Weiter mit dem Skript...
    goto :ABFRAGE_FILEDATE_3
) else if /i "%ANTWORT%"=="N" (
    echo.
    echo Das Filedatum der Dateien wird nicht geaendert.
    rem Weiter mit dem Skript...
    goto :ABFRAGE_FILEDATE_3
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ungueltige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE_FILEDATE_2
)
:ABFRAGE_FILEDATE_3

:ABFRAGE_PDF_ERZEUGUNG_1
echo.
echo =====================================================
echo   *** PDF-Erzeugung? ***
echo =====================================================
echo.
:ABFRAGE_PDF_ERZEUGUNG_2
set /p ANTWORT="Moechtest du zusaetzliche PDF-Dateien erzeugen? (J/N/A) "
if /i "%ANTWORT%"=="J" (
    set "GENERATEPDF=--generate_pdf"
    echo.
    echo Folgendes Argument wird übernommen: !GENERATEPDF!
    rem Weiter mit dem Skript...
    goto :ABFRAGE_PDF_ERZEUGUNG_3
) else if /i "%ANTWORT%"=="N" (
    echo.
    echo Es werden keine zusaetzliche PDF-Dateien erzeugt.
    rem Die nächste Abfrage ist dann überflüssig
    goto :ABFRAGE_PDF_UEBERSCHREIBEN_3
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ungueltige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE_PDF_ERZEUGUNG_2
)
:ABFRAGE_PDF_ERZEUGUNG_3

:ABFRAGE_PDF_UEBERSCHREIBEN_1
echo.
echo =====================================================
echo   *** PDF-Dateien ueberschreiben? ***
echo =====================================================
echo.
:ABFRAGE_PDF_UEBERSCHREIBEN_2
set /p ANTWORT="Moechtest du bestehende PDF-Dateien ueberschreiben? (J/N/A) "
if /i "%ANTWORT%"=="J" (
    set "OVERWRITEPDF=--overwrite_pdf"
    echo.
    echo Folgendes Argument wird uebernommen: !OVERWRITEPDF!
    rem Weiter mit dem Skript...
    goto :ABFRAGE_PDF_UEBERSCHREIBEN_3
) else if /i "%ANTWORT%"=="N" (
    echo.
    echo Es werden keine zusaetzliche PDF-Dateien erzeugt.
    rem Weiter mit dem Skript...
    goto :ABFRAGE_PDF_UEBERSCHREIBEN_3
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ungueltige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE_PDF_UEBERSCHREIBEN_2
)
:ABFRAGE_PDF_UEBERSCHREIBEN_3

:ABFRAGE_STANDARDPFAD_1
echo.
echo ===============================================
echo   *** Zielverzeichnis-Auswahl ***
echo ===============================================
echo.
echo Standard-Zielverzeichnis:
echo   %DEFAULT_ZIELPFAD%
echo.
:ABFRAGE_STANDARDPFAD_2
set /p ANTWORT="Moechtest du diesen Pfad verwenden? (J/N/A) "
if /i "%ANTWORT%"=="J" (
    set "ZIELPFAD=%DEFAULT_ZIELPFAD%"
    echo.
    rem Weiter mit dem Skript...
    goto :START
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else if /i "%ANTWORT%"=="N" (
    echo.
    rem Weiter mit Zielpfadeingabe...
    goto :ABFRAGE_ZIELPFAD_1
) else (
    echo.
    echo Ungueltige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE_STANDARDPFAD_2
)
:ABFRAGE_STANDARDPFAD_3

:ABFRAGE_ZIELPFAD_1
echo.
echo ===============================================
echo   *** Zielverzeichnis-Validierung ***
echo ===============================================
echo.
:ABFRAGE_ZIELPFAD_2
echo Bitte gib den gewuenschten Zielpfad (Windows-konforme Pfadangabe) ein.
echo Soll ein Laufwerk als Zielpfad genutzt werden, bitte ohne Backslash am Ende eingeben (z.B. C:)
set /p ZIELPFAD="Zielpfad "
for %%A in ("%ZIELPFAD%\") do (
      if "%%~dA"=="" (
          echo.
          echo Fehler: Ungültige Eingabe für den Zielpfad!
          goto :ENDE
      )
)
if not exist "%ZIELPFAD%\" (
    echo.
    echo ? Fehler: Das Verzeichnis "%ZIELPFAD%" existiert NICHT.
    echo Der Vorgang wird abgebrochen.
    goto :ENDE
) else (
    echo.
    echo Das Zielverzeichnis für Suche nach MSG-Dateien existiert:
    echo   "%ZIELPFAD%"
    echo.
)
:ABFRAGE_ZIELPFAD_3

:START
echo ===============================================
echo   *** Laufzeitumgebung/Environment pruefen ***
echo ===============================================
echo.
set "FULL_PATH=%~dp0"

REM Entferne den letzten Backslash
set "FULL_PATH=%FULL_PATH:~0,-1%"

REM Extrahiere das letzte Verzeichnis
for %%A in ("%FULL_PATH%") do set "LAST_DIR=%%~nxA"

REM Setze die Variable ENVIRONMENT basierend auf dem Verzeichnis
if "%LAST_DIR%"=="MSGFileRenamer" (
    set "ENVIRONMENT=Entwicklungsumgebung"
) else (
    set "ENVIRONMENT=Produktivbetrieb"
)
:: Angepasste Ausgabe je nach Modus
echo Das Programm wurde in der %ENVIRONMENT% gestartet.

echo.
echo ==========================================================
echo   *** Verarbeitung Starten, abhaengig von ENVIRONMENT ***
echo ==========================================================
echo.

if "%ENVIRONMENT%"=="Entwicklungsumgebung" (
    echo *** Entwicklungsumgebung ***
    echo.
    echo Aktiviere die virtuelle Python-Umgebung
    call "venv\Scripts\activate"
    echo python msg_file_renamer.py %NOTESTLAUF% %SETFILEDATE% %GENERATEPDF% %OVERWRITEPDF% %RECURSIVESEARCH% --search_directory "%ZIELPFAD%" --excel_log_directory "./" %USEKNOWNSENDERFILE% --knownsender_file "%KNOWNSENDERFILE%" --debug_mode --max_console_output
    python "msg_file_renamer.py" %NOTESTLAUF% %SETFILEDATE% %GENERATEPDF% %OVERWRITEPDF% %RECURSIVESEARCH% --search_directory "%ZIELPFAD%" --excel_log_directory "./" %USEKNOWNSENDERFILE% --knownsender_file "%KNOWNSENDERFILE%" --debug_mode --max_console_output
) else (
    echo *** Produktivbetrieb ***
    echo.
    echo msg_file_renamer.exe %NOTESTLAUF% %SETFILEDATE% %GENERATEPDF% %OVERWRITEPDF% %RECURSIVESEARCH% --search_directory "%ZIELPFAD%" --excel_log_directory "./" %USEKNOWNSENDERFILE% --knownsender_file "%KNOWNSENDERFILE%" --debug_mode --max_console_output
    msg_file_renamer.exe %NOTESTLAUF% %SETFILEDATE% %GENERATEPDF% %OVERWRITEPDF% %RECURSIVESEARCH% --search_directory "%ZIELPFAD%" --excel_log_directory "./" %USEKNOWNSENDERFILE% --knownsender_file "%KNOWNSENDERFILE%" --debug_mode --max_console_output
)

:ENDE
pause