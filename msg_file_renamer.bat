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
::   -ntr / --no_test_run:          Testlauf aktivieren (keine echten File-Änderungen)
::   -fd / --set_filedate:          Dateidatum anpassen (False = Dateidatum nicht ändern)
::   -sd / --search_directory:      Such-Verzeichnis für MSG-Dateien
::   -spn / --no_shorten_path_name: Kein Kürzen langer Pfadnamen
::   -pdf / --generate_pdf:         PDF-Erzeugung aus MSG-Dateien aktivieren
::   -opdf / --overwrite_pdf:       Bereits vorhandene PDF-Dateien überschreiben (wenn aktiviert)
::   -rs / --recursive_search:      Rekursive Suche im Zielverzeichnis (optional)
::   -ucf / --use_knownsender_file: CSV-Datei mit Liste den bekannten Sender nutzen (wenn aktiviert)
::   -cf / --knownsender_file:      CSV-Datei mit Liste der bekannten Absender (Default: ".\config\known_senders_private.csv")
::
:: Die einzutragenden Einstellungen können während des Skriptablaufs interaktiv
:: festgelegt oder der Default-Wert verwendet werden.
:: ============================================================================
::
:: Die einzutragenden Einstellungen können während des Skriptablaufs interaktiv
:: festgelegt oder der Default-Wert verwendet werden.
:: ============================================================================

:: ============================================================
:: Prüfe Python-Installation und installiere falls erforderlich
:: ============================================================
echo =====================================================
echo   *** Prüfe Python-Installation ***
echo =====================================================

REM === Konfiguration ===
set PYTHON_VERSION=3.11.8
set PYTHON_INSTALLER=python-%PYTHON_VERSION%-amd64.exe
set DOWNLOAD_URL=https://www.python.org/ftp/python/%PYTHON_VERSION%/%PYTHON_INSTALLER%
set INSTALL_DIR=%USERPROFILE%\AppData\Local\Programs\Python\Python%PYTHON_VERSION:~0,1%%PYTHON_VERSION:~2,2%

REM === 1. Prüfen ob Python vorhanden ist ===
where python >nul 2>nul
if %errorlevel%==0 (
    echo Python ist bereits installiert.
) else (
    echo Python ist nicht installiert. Starte Installation...

    REM === 2. Installer downloaden ===
    powershell -Command "Invoke-WebRequest -Uri '%DOWNLOAD_URL%' -OutFile '%PYTHON_INSTALLER%'"

    REM === 3. Installation ohne Adminrechte ===
    .\%PYTHON_INSTALLER% /quiet InstallAllUsers=0 PrependPath=1 Include_test=0

    REM Kurze Wartezeit, damit Installation abgeschlossen ist
    timeout /t 10 >nul

    REM === 4. Prüfen ob Installation erfolgreich war ===
    if exist "%INSTALL_DIR%\python.exe" (
        echo Python erfolgreich installiert in %INSTALL_DIR%
        set "PATH=%INSTALL_DIR%;%PATH%"
    ) else (
        echo Fehler bei der Installation von Python.
        pause
        exit /b 1
    )
)
echo.
echo.

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
set "RECURSIVESEARCH="
set "USEKNOWNSENDERFILE=--use_knownsender_file"
set "KNOWNSENDERFILE=.\config\known_senders_private.csv"

:: Verzeichnis der Batch-Datei ermitteln
set "SCRIPT_DIR=%~dp0"
set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"  :: letztes Backslash entfernen

:ABFRAGE_DEFAULT_KONFIG_1
:: =======================================================
::    Benutzerfrage: Default-Konfiguration verwenden
:: =======================================================
echo =====================================================
echo   *** Bitte Konfiguration prüfen ***
echo       - Testlauf (Keine File-Änderung!)
echo       - Keine rekursive Suche
echo       - Dateinamen werden gekürzt
echo       - Keine Erzeugung von PDF-Dateien
echo       - CSV-Datei mit Liste bekannter Absender nutzen?
echo       - CSV-Datei mit Liste der bekannten Absender
echo         (Default: ".\config\known_senders_private.csv")
echo       - Zielverzeichnis für MSG-Suche erforderlich
echo         oder Default-Zielverzeichnis DEFAULT_ZIELPFAD
echo =====================================================
echo.
:ABFRAGE_DEFAULT_KONFIG_2
set /p ANTWORT="Soll die Default-Konfiguration verwendet oder abgebrochen werden? (J/N/A)"
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
    echo Ungültige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE_DEFAULT_KONFIG_2
)
:ABFRAGE_DEFAULT_KONFIG_3

:ABFRAGE_TESTLAUF_1
:: =======================================================
::    Benutzerfrage: Testlauf?
:: =======================================================
echo =====================================================
echo   *** Testlauf? ***
echo =====================================================
echo.
:ABFRAGE_TESTLAUF_2
set /p ANTWORT="Möchtest du einen Testlauf durchführen? (J/N/A)"
if /i "%ANTWORT%"=="N" (
    set "NOTESTLAUF=--no_test_run"
    echo.
    echo Folgendes Argument wird übernommen: !NOTESTLAUF!
    rem Weiter mit dem Skript...
    goto :ABFRAGE_TESTLAUF_3
) else if /i "%ANTWORT%"=="J" (
    echo.
    echo Es wird ein Testlauf ohne Änderungen an Dateien durchgeführt.
    goto :ABFRAGE_DATEINAMEN_KUERZEN_1
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ungültige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE_TESTLAUF_2
)
:ABFRAGE_TESTLAUF_3

:ABFRAGE_DATEINAMEN_KUERZEN_1
:: =======================================================
:: Benutzerfrage: Dateinamen kürzen?
:: =======================================================
echo.
echo =====================================================
echo   *** Dateinamen kürzen? ***
echo =====================================================
echo.
:ABFRAGE_DATEINAMEN_KUERZEN_2
set /p ANTWORT="Möchtest du die Dateinamen kürzen? (J/N/A)"
if /i "%ANTWORT%"=="N" (
    set "NOSHORTENPATHNAME=--no_shorten_path_name"
    echo.
    rem Weiter mit dem Skript...
    goto :ABFRAGE_DATEINAMEN_KUERZEN_3
) else if /i "%ANTWORT%"=="J" (
    echo.
    echo Es werden Datenamen gekürzt.
    rem Weiter mit dem Skript...
    goto :ABFRAGE_DATEINAMEN_KUERZEN_3
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ungültige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE_DATEINAMEN_KUERZEN_2
)
:ABFRAGE_DATEINAMEN_KUERZEN_3

:ABFRAGE_REKURSIVE_SUCHE_1
:: =======================================================
:: Benutzerfrage: Rekursive Suche?
:: =======================================================
echo.
echo =====================================================
echo   *** Rekursive Suche? ***
echo =====================================================
echo.
:ABFRAGE_REKURSIVE_SUCHE_2
set /p ANTWORT="Möchtest du das Verzeichnis rekursiv durchsuchen? (J/N/A)"
if /i "%ANTWORT%"=="J" (
    set "RECURSIVESEARCH=--recursive_search"
    echo.
    echo Folgendes Argument wird übernommen: !RECURSIVESEARCH!
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
    echo Ungültige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE_REKURSIVE_SUCHE_2
)
:ABFRAGE_REKURSIVE_SUCHE_3

:: Die folgenden Abfragen nur durchführen, wenn es sich nicht um einen Testlauf handelt
if /i "%NOTESTLAUF%"=="" (
    print No Test Run
    goto :ABFRAGE_STANDARDPFAD_1
)

:ABFRAGE_FILEDATE_1
:: =======================================================
::   Benutzerfrage: Filedate?
:: =======================================================
echo.
echo =====================================================
echo   *** Filedate? ***
echo =====================================================
echo.
:ABFRAGE_FILEDATE_2
set /p ANTWORT="Möchtest du das Dateidatum anpassen? (J/N/A)"
if /i "%ANTWORT%"=="J" (
    set "SETFILEDATE=--set_filedate"
    echo.
    echo Folgendes Argument wird übernommen: !SETFILEDATE!
    rem Weiter mit dem Skript...
    goto :ABFRAGE_FILEDATE_3
) else if /i "%ANTWORT%"=="N" (
    echo.
    echo Das Filedatum der Dateien wird nicht geändert.
    rem Weiter mit dem Skript...
    goto :ABFRAGE_FILEDATE_3
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ungültige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE_FILEDATE_2
)
:ABFRAGE_FILEDATE_3

:ABFRAGE_PDF_ERZEUGUNG_1
:: =======================================================
::   Benutzerfrage: PDF-Erzeugung?
:: =======================================================
echo.
echo =====================================================
echo   *** PDF-Erzeugung? ***
echo =====================================================
echo.
:ABFRAGE_PDF_ERZEUGUNG_2
set /p ANTWORT="Möchtest du zusätzliche PDF-Dateien erzeugen? (J/N/A)"
if /i "%ANTWORT%"=="J" (
    set "GENERATEPDF=--generate_pdf"
    echo.
    echo Folgendes Argument wird übernommen: !GENERATEPDF!
    rem Weiter mit dem Skript...
    goto :ABFRAGE_PDF_ERZEUGUNG_3
) else if /i "%ANTWORT%"=="N" (
    echo.
    echo Es werden keine zusätzliche PDF-Dateien erzeugt.
    rem Die nächste Abfrage ist dann überflüssig
    goto :ABFRAGE_PYTHON_UMGEBUNG_1
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ungültige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE_PDF_ERZEUGUNG_2
)
:ABFRAGE_PDF_ERZEUGUNG_3

:ABFRAGE_PDF_UEBERSCHREIBEN_1
:: =======================================================
:: Benutzerfrage: PDF-Dateien überschreiben?
:: =======================================================
echo.
echo =====================================================
echo   *** PDF-Dateien überschreiben? ***
echo =====================================================
echo.
:ABFRAGE_PDF_UEBERSCHREIBEN_2
set /p ANTWORT="Möchtest du bestehende PDF-Dateien überschreiben? (J/N/A)"
if /i "%ANTWORT%"=="J" (
    set "OVERWRITEPDF=--overwrite_pdf"
    echo.
    echo Folgendes Argument wird übernommen: !OVERWRITEPDF!
    rem Weiter mit dem Skript...
    goto :ABFRAGE_PDF_UEBERSCHREIBEN_3
) else if /i "%ANTWORT%"=="N" (
    echo.
    echo Es werden keine zusätzliche PDF-Dateien erzeugt.
    rem Weiter mit dem Skript...
    goto :ABFRAGE_PDF_UEBERSCHREIBEN_3
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ungültige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE_PDF_UEBERSCHREIBEN_2
)
:ABFRAGE_PDF_UEBERSCHREIBEN_3

:ABFRAGE_PYTHON_UMGEBUNG_1
:: =============================================
:: Benutzerfrage: Virtuelle Umgebung?
:: =============================================
echo.
echo ===============================================
echo   *** Abfrage-Virtuelle_Umgebung ***
echo ===============================================
echo.
:ABFRAGE_PYTHON_UMGEBUNG_2
set /p ANTWORT="Virtuelle Python-Umgebung nutzen? (J/N/A)"
if /i "%ANTWORT%"=="J" (
    :: Aktiviert die virtuelle Python-Umgebung
    call "venv\Scripts\activate"
    echo.
    rem Weiter mit dem Skript...
    goto :ABFRAGE_PYTHON_UMGEBUNG_3
) else if /i "%ANTWORT%"=="N" (
    echo.
    rem Weiter mit dem Skript...
    goto :ABFRAGE_PYTHON_UMGEBUNG_3
) else if /i "%ANTWORT%"=="A" (
    echo.
    echo Der Vorgang wird abgebrochen.
    rem Sprung zum Ende des Skripts...
    goto :ENDE
) else (
    echo.
    echo Ungültige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE_PYTHON_UMGEBUNG_2
)
:ABFRAGE_PYTHON_UMGEBUNG_3

:ABFRAGE_STANDARDPFAD_1
:: =============================================
:: Benutzerfrage: Standardpfad verwenden?
:: =============================================
echo.
echo ===============================================
echo   *** Zielverzeichnis-Auswahl ***
echo ===============================================
echo.
echo Standard-Zielverzeichnis:
echo   %DEFAULT_ZIELPFAD%
echo.
:ABFRAGE_STANDARDPFAD_2
set /p ANTWORT="Möchtest du diesen Pfad verwenden? (J/N/A)"
if /i "%ANTWORT%"=="J" (
    set "ZIELPFAD=%DEFAULT_ZIELPFAD%"
    echo.
    :: Aktiviert die virtuelle Python-Umgebung
    call "venv\Scripts\activate"
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
    echo Ungültige Eingabe. Bitte J, N oder A eingeben.
    echo.
    goto :ABFRAGE_STANDARDPFAD_2
)
:ABFRAGE_STANDARDPFAD_3

:ABFRAGE_ZIELPFAD_1
:: =============================================
:: Zielpfad prüfen
:: =============================================
echo.
echo ===============================================
echo   *** Zielverzeichnis-Validierung ***
echo ===============================================
echo.
:ABFRAGE_ZIELPFAD_2
set /p ZIELPFAD="Bitte gib den gewünschten Zielpfad ein: (Windows-konforme Pfadangabe)"
if not exist "%ZIELPFAD%\" (
    echo.
    echo ? Fehler: Das Verzeichnis "%ZIELPFAD%" existiert NICHT.
    echo Der Vorgang wird abgebrochen.
    goto :ENDE
) else (
    echo.
    echo ? Das Zielverzeichnis existiert.
    echo.
    echo Zielverzeichnis für Suche nach MSG-Dateien:
    echo   "%ZIELPFAD%"
    echo.
)
:ABFRAGE_ZIELPFAD_3

:START
:: =============================================
:: Verarbeitung starten
:: =============================================
echo.
echo ===============================================
echo   *** Verarbeitung Starten ***
echo ===============================================
echo.
echo python msg_file_renamer.py %NOTESTLAUF% %SETFILEDATE% %GENERATEPDF% %OVERWRITEPDF% %RECURSIVESEARCH% --search_directory "%ZIELPFAD%" --excel_log_directory "./" %USEKNOWNSENDERFILE% --knownsender_file "%KNOWNSENDERFILE%"
python "msg_file_renamer.py" %NOTESTLAUF% %SETFILEDATE% %GENERATEPDF% %OVERWRITEPDF% %RECURSIVESEARCH% --search_directory "%ZIELPFAD%" --excel_log_directory "./" %USEKNOWNSENDERFILE% --knownsender_file "%KNOWNSENDERFILE%"
echo.

:ENDE
pause