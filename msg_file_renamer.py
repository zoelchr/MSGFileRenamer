# -*- coding: utf-8 -*-
"""
msg_file_renamer.py

Dieses Modul bietet Funktionen für das Umbenennen von MSG-Dateien basierend auf den in den Dateien enthaltenen Metadaten.
Es stellt sicher, dass die Dateien konsistent benannt werden und speichert wichtige Informationen über den Umbenennungsprozess in einem Log.
Ziel ist es, MSG-Dateien (z.B. E-Mail-Dateien) nach einem definierten Schema umzubenennen, um sie leichter zu organisieren und zu finden.

Funktionen und Prozesse:
- Extraktion von Metadaten:
  - Liest relevante Metadaten aus MSG-Dateien aus, wie Betreff, Absender, Empfänger und Datum.
  - Erkennt und behandelt Sonderzeichen, um valide Dateinamen zu generieren.
- Dateiumbenennung: (Optional)
  - Generiert neue Dateinamen basierend auf Metadaten und einem vordefinierten Schema.
  - Überprüft, ob Dateien ohne Konflikte im Zielverzeichnis gespeichert werden können.
- Änderung des Dateierstellungsdatums: (Optional)
  - Setzt das Erstellungsdatum auf das Versanddatum der MSG-Datei
  - Setzt das Änderungsdatum auf das Versanddatum der MSG-Datei
- Generierung von PDF-Dateien: (Optional)
  - Generiert eine PDF-Datei aus der MSG-Datei.
  - Überschreibt vorhandene PDF-Dateien, falls gewünscht.
- Kürzung von Dateinamen: (Optional)
  - Optional: Kürzt Dateinamen, die länger als 255 Zeichen sind, um Probleme mit dem Dateisystem zu vermeiden.
- Fehlerbehandlung:
  - Erkennt problematische Dateien (z.B. beschädigte MSG-Dateien) und protokolliert entsprechende Fehler.
  - Behandelt bestehende Dateien mit demselben Namen, um Namenskonflikte zu vermeiden.
- Protokollierung:
  - Alle Änderungen werden auch im Testmode in einer Exceldatei abgespeichert
  - Führt detaillierte Aufzeichnungen über umbenannte Dateien sowie über auftretende Probleme während des Prozesses.

Spezielle Parameter zur Konfiguration
- TARGET_DIRECTORY: Verzeichnis, in dem nach MSG-Dateien gesucht und (optional) auch bearbeitet werden sollen.
- SOURCE_DIRECTORY_TEST_DATA: Verzeichnis, in dem Testdaten gespeichert sind. Diese Testdaten bleiben immer unverändert.
- TARGET_DIRECTORY_TEST_DATA: Verzeichnis, in dem während eines Programmlaufs die Testdaten kopiert und dann auch verändert werden.
- LOG_FILE_PATH: Pfad zur Log-Datei.
- LOG_LEVEL: Gibt das gewünschte Log-Level an (z. B. DEBUG, INFO, WARNING, ERROR).
- MAX_LENGTH_SUBJECT: Maximale Länge des Betreffs, bevor er gekürzt wird.
- MAX_LENGTH_SENDERLIST: Maximale Länge der Empfängerliste, bevor sie gekürzt wird.
- MAX_BODY_LENGTH: Maximale Länge des Nachrichtentextes, bevor er gekürzt wird.
- MAX_FILENAME_LENGTH: Maximale Länge des Dateinamens, bevor er gekürzt wird.
- MAX_PATH_LENGTH: Maximale Länge des Dateipfads, bevor er gekürzt wird.
- MAX_PDF_LENGTH: Maximale Länge des PDF-Inhalts, bevor er gekürzt wird.

Kommandozeilenargumente:
--search_directory <Zielpfad>
    Gibt das Verzeichnis an, in dem nach MSG-Dateien gesucht wird.
    Wird kein Pfad angegeben, wird ein Standardverzeichnis verwendet.
--excel_log_directory <Zielpfad>
    Gibt das Verzeichnis an, in dem die Excel-Log-Datei erstellt wird.
    Wird kein Pfad angegeben, wird das gleiche Verzeichnis verwendet, wo auch die Python-Datei liegt.
--no_test_run
    Wenn dieses Flag gesetzt wird, wird der Testmodus deaktiviert,
    d.h. die Dateien werden tatsächlich umbenannt bzw. verändert.
    (Standard: Testmodus aktiv, d. h. keine echten Dateiänderungen.)
--set_filedate
    Wenn dieses Flag gesetzt wird, wird das Erstellungs- und das Änderungsdatum der MSG-Datei auf das
    Datum des Versandzeitpunkts gesetzt.
    (Standard: False)
--no_shorten_path_name
    Wenn dieses Flag gesetzt ist, wird der Pfad nicht gekürzt, wenn er zu lang ist.
    (Standard: False)
--generate_pdf
    Wenn dieses Flag gesetzt ist, wird aus der MSG-Datei eine PDF-Datei erstellt.
    (Standard: False)
--overwrite_pdf
    Gibt an, ob vorhandene PDF-Dateien überschrieben werden sollen.
    Bei True wird eine ggf. schon vorhandene, gleichnamige PDF-Datei überschrieben,
    bei False bleibt die bestehende PDF-Datei erhalten.
--max_console_output
    Reduzierte Ausgabe des Vorgangs auf der Console

Kommandozeilenargumente speziell zu Testzwecken
--debug_mode
    Wenn dieses Flag gesetzt ist, dann wird der Debug-Mode aktiviert.
    (Standard: False)
--no_test_run
    Wenn dieses Flag gesetzt ist, dann wird der Testmodus deaktiviert.
    (Standard: False)
--init_testdata
    Wenn dieses Flag gesetzt ist, dann werden alle Dateien aus dem Verzeichnis SOURCE_DIRECTORY_TEST_DATA in das Verzeichnis TARGET_DIRECTORY_TEST_DATA kopiert.
--debug_log_directory <Zielpfad>
    Gibt den Dateinamen an, in den die Log-Nachrichten geschrieben werden sollen.
    Wird kein Pfad angegeben, wird das gleiche Verzeichnis verwendet, wo auch die Python-Datei liegt.

Verwendung:
Das Modul kann verwendet werden, um MSG-Dateien schnell und effizient umzubenennen sowie deren Organisation zu verbessern.
Es ist besonders nützlich, um große Mengen an E-Mail-Dateien nach einheitlichen Kriterien zu strukturieren.
Der Testmodus ist ideal, um den Prozess vor der endgültigen Ausführung zu validieren.

Hinweise:
- Stellen Sie sicher, dass alle Abhängigkeiten (zum Beispiel für das Lesen von MSG-Dateien und die PDF-Erzeugung) installiert sind.
- Passen Sie das Namensschema, die optionalen Attribute und die Konfiguration an Ihre Anforderungen an, bevor Sie das Modul produktiv einsetzen.

Dieses Modul benötigt einige zusätzliche Python-Pakete, die für bestimmte Funktionen essenziell sind:

- extract_msg:
    Zum Auslesen von Outlook-.msg-Dateien, inkl. Betreff, Body, Anhängen und Metadaten.

- pandas:
    Für die Verarbeitung strukturierter Daten, insbesondere zum Einlesen und Pflegen
    von CSV-Dateien wie Log-Dateien oder Listen bekannter Absender.

- pywin32:
    Ermöglicht Zugriff auf Windows-spezifische Systemfunktionen:
    - pywintypes: Umwandlung und Verarbeitung von Datei-Zeitstempeln
    - win32file: Zugriff auf FileHandles sowie Lesen/Setzen von Datei-Zeitstempeln
    - win32con: Auslesen von Datei-Attributen (z.B. versteckt, schreibgeschützt)

- fpdf2:
    Zur Erstellung von PDF-Dokumenten für strukturierte Ausgaben oder Berichte.

- openpyxl:
    Für das Lesen und Schreiben von Excel-Dateien im .xlsx-Format (ab Excel 2007).
"""
import os
import datetime
import argparse
import sys
import importlib.util
from pathlib import Path

from modules.msg_generate_new_filename import generate_new_msg_filename
from utils.file_handling import rename_file, test_file_access, FileAccessStatus, set_file_creation_date, set_file_modification_date, FileOperationResult
from modules.msg_handling import log_entry_neu, create_log_file_neu
from utils.excel_handling import clean_old_excel_files
from utils.testset_preparation import prepare_test_directory
from utils.pdf_generation import generate_pdf_from_msg
from config import SOURCE_DIRECTORY_TEST_DATA, TARGET_DIRECTORY_TEST_DATA, MAX_PATH_LENGTH, DEBUG_LEVEL, LOG_FILE_DIRECTORY, MAX_EXCEL_LOG_FILE_COUNT, ENV_LIST_OF_KNOWN_SENDERS

#import optimierter Logger
from logger import initialize_logger, clean_logs_and_initialize, DEBUG_LEVEL_TEXT, prog_log_file_path

# Initialisierung im Hauptprogramm
clean_logs_and_initialize()

# In der Log-Datei wird als Quelle der Modulname "__main__" verwendet
app_logger = initialize_logger(__name__)
app_logger.debug("Debug-Logging im Modul 'main' aktiviert.")

os.system('chcp 65001  > nul')  # Setzt Konsole auf UTF-8

# # Verzeichnisse für die Tests definieren
# # SOURCE_DIRECTORY_TEST_DATA = r'.\data\sample_files\testset-short-longpath'
# SOURCE_DIRECTORY_TEST_DATA = r'.\data\sample_files\testset-public'
# # SOURCE_DIRECTORY_TEST_DATA = r'.\data\sample_files\testset-short'
# TARGET_DIRECTORY_TEST_DATA = r'.\tests\functional\testdir'
# TARGET_DIRECTORY = ""
# LIST_OF_KNOWN_SENDERS = r'.\config\known_senders.csv' # Liste der bekannten Email-Absender aus einer CSV-Datei
#
# # Maximal zulässige Pfadlänge für Windows 11
# MAX_PATH_LENGTH = 260

def check_module_installed(module_name: str, install_text: str):
    """
    Überprüft, ob ein bestimmtes Python-Modul installiert ist.

    Diese Funktion versucht, das angegebene Modul zu importieren.
    Wenn der Import fehlschlägt, wird eine Fehlermeldung ausgegeben und das Programm beendet.

    Parameter:
    - module_name (str): Der Name des zu überprüfenden Moduls.

    Verwendung:
    Diese Funktion sollte zu Beginn eines Programms aufgerufen werden, um sicherzustellen,
    dass alle erforderlichen Abhängigkeiten vorhanden sind, bevor der Hauptteil des Programms ausgeführt wird.
    """
    if importlib.util.find_spec(module_name) is None:
        app_logger.error(f"Das Modul '{module_name}' ist nicht installiert.")
        print(f"Fehler: Das Modul '{module_name}' ist nicht installiert.")
        print("Bitte installieren Sie das fehlende Modul und starten Sie das Programm erneut.\n")
        print("Dazu zuerst in das Verzeichnis mit der Datei 'msg_file_renamer.bat' wechseln und dort ein Windows Terminal öffnen.")
        print(r"Dann die zugehörige virtuelle Python-Umgebung durch '.\venv\Scripts\activate.bat' aktivieren.")
        print("Eventuell hat das Verzeichnis für die virtuelle Python-Umgebung statt 'venv' auch einen anderen Namen.")
        print(f"Anschließend kann das fehlende Modul über die folgende Eingabe installiert werden:\n '{install_text}'")
        app_logger.error(f"Das Programm wird beendet.")
        sys.exit(1)
    else:
        app_logger.debug(f"Das Modul '{module_name}' ist installiert.")

if __name__ == '__main__':

    app_logger.info(f"Programm Logdatei: {LOG_FILE_DIRECTORY}")
    debug_level_text = DEBUG_LEVEL_TEXT.get(DEBUG_LEVEL, "UNKNOWN")

    app_logger.info(f"Debug-Modus: {debug_level_text}")
    app_logger.info("Programm gestartet")

    # Überprüfen, ob die erforderlichen Module installiert sind
    check_module_installed('extract_msg', "pip install extract_msg --trusted-host pypi.org --trusted-host files.pythonhosted.org")
    check_module_installed('pandas', "pip install pandas --trusted-host pypi.org --trusted-host files.pythonhosted.org")
    check_module_installed('fpdf', "pip install fpdf2 --trusted-host pypi.org --trusted-host files.pythonhosted.org") # Dient zur Überprüfung, ob fpdf oder fpdf2 installiert ist
    check_module_installed('openpyxl', "pip install openpyxl --trusted-host pypi.org --trusted-host files.pythonhosted.org")
    check_module_installed('win32file', "pip install pywin32 --trusted-host pypi.org --trusted-host files.pythonhosted.org") # Dient der Überprüfung, ob pywin32 installiert ist

    # Argumente des Programmaufrufs über die Kommandozeile auswerten
    parser = argparse.ArgumentParser()
    parser.add_argument("-ntr", "--no_test_run", default=False, action="store_true", help="Testlauf ohne Dateioperationen aktivieren (Default=False)")
    parser.add_argument("-it", "--init_testdata", default=False, action="store_true", help="True/False für Initialisiere Testdaten (Default=False)")
    parser.add_argument("-fd", "--set_filedate", default=False, action="store_true", help="True/False für File Date (Default=False)")
    parser.add_argument("-db", "--debug_mode", default=False, action="store_true", help="True/False für Debug-Mode (Default=False)")
    parser.add_argument("-sd", "--search_directory", type=str, default="", help="Verzeichnispfad als Such-Verzeichnis (Default='').")
    parser.add_argument("-elb", "--excel_log_basename", type=str, default="excel_log_file", help="Dateiname-Anfang für Excel-Log-Aufzeichnung (Default='excel_log_file')")
    parser.add_argument("-elf", "--excel_log_directory", type=str, default="./", help="Verzeichnis für Excel-Log-Aufzeichnung (Default='./')")
    parser.add_argument("-dlf", "--debug_log_directory", type=str, default="./", help="Verzeichnis für Debug-Log-Aufzeichnung (Default='./')")
    parser.add_argument("-spn", "--no_shorten_path_name", default=False, action="store_true", help="True/False für kein Kürzen des Pfades bei Überlänge (Default=False)")
    parser.add_argument("-pdf", "--generate_pdf", default=False, action="store_true", help="True/False für Generieren eines PDF-Files aus MSG-Dateien (Default=False)")
    parser.add_argument("-opdf", "--overwrite_pdf", default=False, action="store_true", help="True/False für Überschreiben eines bereits existierende PDF-Files aus MSG-Dateien (Default=False)")
    parser.add_argument("-rs", "--recursive_search", default=False, action="store_true", help="True/False für die rekursive Suche nach MSG-Dateien (Default=False)"),
    parser.add_argument("-ucf", "--use_knownsender_file", default=False, action="store_true", help="True/False für die Nutzung des Config-Files (Default=False)"),
    parser.add_argument("-cf", "--knownsender_file", type=str, default=ENV_LIST_OF_KNOWN_SENDERS, help="CSV-Datei mit Liste der bekannten Absender")
    parser.add_argument("-mco", "--max_console_output", default=False, action="store_true", help="Maximale Consolen-Ausgabe aktivieren (Default=False)")
    args, unknown = parser.parse_known_args()

    # Unbekannte Parameter ausgeben und Programm beenden
    if unknown:
        app_logger.error(f"Unbekannte Parameter gefunden: {unknown}")
        print(f"Warnung: Unbekannte Parameter gefunden: {unknown}")
        print("Bitte überprüfen Sie die übergebenen Parameter.")
        app_logger.error(f"Das Programm wird beendet.")
        exit()

    # Formatierung der Argumente für die Ausgabe auf der Console oder dem Log-File
    args_formatted = str(args).replace(",", "\n\t\t").replace("Namespace", "").strip("()")

    # Argumente des Programmaufrufs an die Variablen übergeben
    DEBUG_MODE = args.debug_mode
    MAX_CONSOLE_OUTPUT = args.max_console_output
    USE_KNOWNSENDER_FILE = args.use_knownsender_file
    KNOWNSENDER_FILE = args.knownsender_file
    INIT_TESTDATA = args.init_testdata
    TEST_RUN = not args.no_test_run
    RECURSIVE_SEARCH = args.recursive_search
    NO_SHORTEN_PATH_NAME = args.no_shorten_path_name
    SET_FILEDATE = args.set_filedate
    GENERATE_PDF = args.generate_pdf
    OVERWRITE_PDF = args.overwrite_pdf

    app_logger.debug(f"Argumente:")
    # Alles mit Fokus Debug
    app_logger.info(f"DEBUG_MODE = {DEBUG_MODE}")
    app_logger.info(f"MAX_CONSOLE_OUTPUT = {MAX_CONSOLE_OUTPUT}")
    # Known-Sender-File
    app_logger.info(f"USE_KNOWSENDER_FILE = {USE_KNOWNSENDER_FILE}")
    app_logger.info(f"KNOWSENDER_FILE = {KNOWNSENDER_FILE}")
    # Test-Initialisierung
    app_logger.info(f"INIT_TESTDATA = {INIT_TESTDATA}")
    app_logger.info(f"TEST_RUN = {TEST_RUN}")
    # Ablaufsteuerung
    app_logger.info(f"RECURSIVE_SEARCH = {RECURSIVE_SEARCH}")
    app_logger.info(f"NO_SHORTEN_PATH_NAME = {NO_SHORTEN_PATH_NAME}")
    app_logger.info(f"SET_FILEDATE = {SET_FILEDATE}")
    app_logger.info(f"GENERATE_PDF = {GENERATE_PDF}")
    app_logger.info(f"OVERWRITE_PDF = {OVERWRITE_PDF}")

    # Start Ausgabe auf Console
    if MAX_CONSOLE_OUTPUT: print(f"\nTestlauf: {TEST_RUN}\nTestverzeichnis initialisieren: {INIT_TESTDATA}\nZeitstempel der MSG-dateien anpassen: {SET_FILEDATE}\nDebug-Modus: {DEBUG_MODE}")

    # Excel-Log-Datei
    EXCEL_LOG_DIRECTORY = args.excel_log_directory # Verzeichnis für die Excel-Log-Datei
    excel_log_basename = "excel_log_file_" # Basisname für die Excel-Log-Datei
    excel_log_file_path = os.path.join(EXCEL_LOG_DIRECTORY, excel_log_basename) # Pfad und Dateiname für die Excel-Log-Datei
    app_logger.info(f"EXCEL_LOG_DIRECTORY = {EXCEL_LOG_DIRECTORY}")

    # Verzeichnis für die Suche festlegen
    TARGET_DIRECTORY = Path(args.search_directory) # Verzeichnis für die Suche nach MSG-Dateien
    if not TARGET_DIRECTORY:
        TARGET_DIRECTORY = TARGET_DIRECTORY_TEST_DATA
        app_logger.info(f"TARGET_DIRECTORY = TARGET_DIRECTORY_TEST_DATA aus env-Datei übernommen: {TARGET_DIRECTORY_TEST_DATA}")
    else:
        app_logger.info(f"TARGET_DIRECTORY = {TARGET_DIRECTORY}")

    print(f"\n************************************************************************")
    print(f"Verzeichnis für die Suche: {TARGET_DIRECTORY}") # Debugging-Ausgabe: Console
    print(f"************************************************************************")

    # Log-Verzeichnis und Basisname für Excel-Logdateien festlegen
    LOG_TABLE_HEADER = ["Fortlaufende Nummer", "Verzeichnisname", "Original-Filename"]

    # Excel-Logdatei erstellen und den Pfad ausgeben
    excel_log_file_path = create_log_file_neu(excel_log_basename, EXCEL_LOG_DIRECTORY, LOG_TABLE_HEADER, sheet_name="Log")
    if MAX_CONSOLE_OUTPUT: print(f"Excel-Logdatei erstellt: {excel_log_file_path}")
    app_logger.info(f"Excel-Logdatei = {excel_log_file_path}")

    # Ältere Excel-Logdatei löschen
    deleted_excel_file_count = clean_old_excel_files(EXCEL_LOG_DIRECTORY, MAX_EXCEL_LOG_FILE_COUNT, "excel_log_file_")

    if INIT_TESTDATA:
        if MAX_CONSOLE_OUTPUT: print(f"Prüfung ob Zielverzeichnis für Testdaten bereits existiert: {TARGET_DIRECTORY_TEST_DATA}") # Debugging-Ausgabe: Console
        app_logger.debug(f"Prüfung ob Zielverzeichnis für Testdaten bereits existiert: {TARGET_DIRECTORY_TEST_DATA}")  # Debugging-Ausgabe: Log-File

        # Wenn TARGET_DIRECTORY_TEST_DATA noch nicht existiert, dann erstellen
        if not os.path.isdir(TARGET_DIRECTORY_TEST_DATA):
            os.makedirs(TARGET_DIRECTORY_TEST_DATA, exist_ok=True)
            if MAX_CONSOLE_OUTPUT: print(f"Zielverzeichnis neu erstellt: {TARGET_DIRECTORY_TEST_DATA}")
            app_logger.debug(f"Zielverzeichnis neu erstellt: {TARGET_DIRECTORY_TEST_DATA}")
        else:
            if MAX_CONSOLE_OUTPUT: print(f"Zielverzeichnis existiert bereits: {TARGET_DIRECTORY_TEST_DATA}")
            app_logger.debug(f"Zielverzeichnis existiert bereits: {TARGET_DIRECTORY_TEST_DATA}")

        # Bereite das Zielverzeichnis vor und überprüfe den Erfolg
        success = prepare_test_directory(SOURCE_DIRECTORY_TEST_DATA, TARGET_DIRECTORY_TEST_DATA)
        if success:
            if MAX_CONSOLE_OUTPUT: print("Vorbereitung des Testverzeichnisses erfolgreich abgeschlossen.") # Debugging-Ausgabe: Console
            app_logger.debug("\tVorbereitung des Testverzeichnisses erfolgreich abgeschlossen.")  # Debugging-Ausgabe: Log-File
        else:
            if MAX_CONSOLE_OUTPUT: print("Fehler: Die Vorbereitung des Testverzeichnisses ist fehlgeschlagen. Das Programm wird abgebrochen.") # Debugging-Ausgabe: Console
            app_logger.error("Fehler: Die Vorbereitung des Testverzeichnisses ist fehlgeschlagen. Das Programm wird abgebrochen.")  # Debugging-Ausgabe: Log-File
            exit(1)  # Programm abbrechen
    else:
        if MAX_CONSOLE_OUTPUT: print(f"Kein Testverzeichnis vorbereitet.") # Debugging-Ausgabe: Console
        app_logger.debug(f"Kein Testverzeichnis vorbereitet.")  # Debugging-Ausgabe: Log-File

    # Zähler für gefundene sowie umbenannte Dateien und Probleme initialisieren
    msg_file_count = 0
    msg_file_renamed_count = 0
    msg_file_problem_count = 0
    msg_file_shorted_name_count = 0
    msg_file_doublette_count = 0
    msg_file_doublette_deleted_count = 0
    msg_file_doublette_deleted_problem_count = 0
    msg_file_file_creation_date_count = 0
    msg_file_creation_date_problem_count = 0
    msg_file_creation_date_unchanged_count = 0
    msg_file_modification_date_count = 0
    msg_file_modification_date_problem_count = 0
    msg_file_modification_date_unchanged_count = 0
    msg_file_same_name_count = 0
    pdf_file_generated = 0
    pdf_file_skipped = 0

    # Sicherstellen das TARGET_DIRECTORY ein Pfad ist
    TARGET_DIRECTORY = Path(TARGET_DIRECTORY)

    # Add the \\?\ prefix to support long paths on Windows
    TARGET_DIRECTORY = f"\\\\?\\{os.path.abspath(TARGET_DIRECTORY)}"
    app_logger.debug(f"TARGET_DIRECTORY (Windows Long Path Format) = '{TARGET_DIRECTORY}'")  # Debugging-Ausgabe: Log-File

    for pathname, dirs, files in os.walk(TARGET_DIRECTORY):

        # filename = Dateiname
        for filename in files:

            app_logger.debug(f"**************************BEARBEITUNG NÄCHSTE MSG DAIEI************************************")  # Debugging-Ausgabe: Log-File

            # Initialisierung der Variable
            is_msg_file_name_unchanged = False
            is_msg_file_doublette = False
            is_msg_file_doublette_deleted = False
            is_pdf_file_skipped = False
            is_pdf_file_generated = False
            rename_msg_file_result = None

            # Überprüfen, ob die Datei die Endung .msg hat
            if filename.lower().endswith('.msg'):
                print(f"MSG-Datei: '{filename}'")  # Debugging-Ausgabe: Console
                app_logger.debug(f"Aktuelle MSG-Datei zur Bearbeitung: '{filename}'")  # Debugging-Ausgabe: Log-File

                # Absoluter Pfadname der MSG-Datei
                path_and_file_name = os.path.join(pathname, filename)

                # Pfadlänge ermitteln
                path_and_file_name_length = len(path_and_file_name)
                app_logger.debug(f"Pfadlänge aktuelle MSG-Datei: '{path_and_file_name_length}'")  # Debugging-Ausgabe: Log-File

                if MAX_CONSOLE_OUTPUT: print(f"\tAktuelles Verzeichnis: '{os.path.dirname(path_and_file_name)}'")  # Debugging-Ausgabe: Console
                app_logger.debug(f"Aktuelles Verzeichnis: '{os.path.dirname(path_and_file_name)}'")  # Debugging-Ausgabe: Log-File

                msg_file_count += 1 # Zähler erhöhen, MSG-Datei gefunden

                # Überprüfen den Schreib- und Lesezugriff auf die MSG-Datei
                access_result = test_file_access(path_and_file_name)

                if MAX_CONSOLE_OUTPUT: print(f"\tÜberprüfung Zugriff auf aktuelle MSG-Date: {[s.value for s in access_result]}'")  # Debugging-Ausgabe: Console
                app_logger.debug(f"Überprüfung Zugriff auf aktuelle MSG-Date: {[s.value for s in access_result]}'")  # Debugging-Ausgabe: Log-File

                # Nur wenn die MSG-Datei schreibend geöffnet werden kann, ist ein Umbenennen möglich
                if FileAccessStatus.WRITABLE in access_result:
                    if MAX_CONSOLE_OUTPUT: print(f"\tSchreibender Zugriff auf die Datei ist möglich.")  # Debugging-Ausgabe: Console
                    app_logger.debug(f"Schreibender Zugriff auf die Datei ist möglich: {filename}")  # Debugging-Ausgabe: Log-File

                    # Neuen Dateinamen erzeugen
                    app_logger.debug(f"Versuche neuen Dateinamen fzu erzeugen.")  # Debugging-Ausgabe: Log-File
                    new_msg_filename_collection = generate_new_msg_filename(path_and_file_name, use_list_of_known_senders=USE_KNOWNSENDER_FILE, file_list_of_known_senders=KNOWNSENDER_FILE, max_console_output=MAX_CONSOLE_OUTPUT)

                    # Überprüfen, ob new_msg_filename_collection nicht Leer (True) ist
                    if new_msg_filename_collection.new_truncated_msg_filename:

                        # Datei für Anpassung Erstellungs- und Änderungsdatum verfügbar?
                        is_msg_file_for_change_date_available = False

                        # Alter und neuer Name
                        old_path_and_file_name = path_and_file_name

                        # Neuen Filenamen setzen in Abhängigkeit von NO_SHORTEN_PATH_NAME
                        if NO_SHORTEN_PATH_NAME:
                            new_file_name = new_msg_filename_collection.new_msg_filename
                        else:
                            new_file_name = new_msg_filename_collection.new_truncated_msg_filename
                            if new_msg_filename_collection.is_msg_filename_truncated:
                                msg_file_shorted_name_count += 1
                                if MAX_CONSOLE_OUTPUT: print(f"\tNeuer gekürzter Dateiname: '{new_file_name}'")
                                app_logger.debug(f"Neuer gekürzter Dateiname: '{new_file_name}'")  # Debugging-Ausgabe: Log-File

                        # Neuen absoluten Pfad erzeugen
                        new_path_and_file_name = os.path.join(pathname, new_file_name)

                        if MAX_CONSOLE_OUTPUT: print(f"\tNeuer absoluter Pfad: '{new_path_and_file_name}'")
                        app_logger.debug(f"Neuer absoluter Pfad: '{new_path_and_file_name}'")  # Debugging-Ausgabe: Log-File

                        # Pfadlänge ermitteln
                        new_path_and_file_name_length = len(new_path_and_file_name)
                        if MAX_CONSOLE_OUTPUT: print(f"\tPfadlänge neue MSG-Datei: '{new_path_and_file_name_length}'")
                        app_logger.debug(f"Pfadlänge neue MSG-Datei: '{new_path_and_file_name_length}'")  # Debugging-Ausgabe: Log-File

                        # Prüfen, ob Alter und neuer Name gleich sind, dann keine Änderung erforderlich
                        if old_path_and_file_name == new_path_and_file_name:
                            print(f"\tAlter und neuer Dateiname sind gleich.")
                            app_logger.debug(f"Alter und neuer Dateiname sind gleich: '{filename}'")  # Debugging-Ausgabe: Log-File
                            msg_file_same_name_count += 1  # Erfolgszähler erhöhen
                            is_msg_file_name_unchanged = True # Kennzeichnung keine Änderung des Dateinamens erforderlich
                            is_msg_file_for_change_date_available = True # Kennzeichnung für Anpassung Erstellungs- und Änderungsdatum
                        else:
                            # Prüfen, ob die Datei mit neuem Namen bereits existiert, also Doublette
                            if os.path.exists(new_path_and_file_name):
                                print(f"\tDatei ist eine Doublette: '{filename}'")
                                app_logger.debug(f"Datei ist eine Doublette: '{filename}'")  # Debugging-Ausgabe: Log-File
                                msg_file_doublette_count += 1  # Erfolgszähler erhöhen
                                is_msg_file_doublette = True # MSG-Datei mit gleichem neuen Namen existiert bereits - also Doublette
                                is_msg_file_for_change_date_available = False  # Kennzeichnung für Anpassung Erstellungs- und Änderungsdatum

                                # Versuche Doublette zu löschen, wenn nicht Test
                                if not TEST_RUN:
                                    # Versuche, die Datei zu löschen
                                    try:
                                        os.remove(old_path_and_file_name)  # Versuche, die Datei zu löschen
                                        print(f"\tDoublette gelöscht: '{filename}'")
                                        app_logger.debug(f"Doublette gelöscht: '{filename}'")  # Debugging-Ausgabe: Log-File
                                        msg_file_doublette_deleted_count += 1  # Löschzähler erhöhen
                                        is_msg_file_doublette_deleted = True
                                    except Exception as e:
                                        print(f"\tDoublette konnte nicht gelöscht werden: '{filename}'. Fehler: {str(e)}")
                                        app_logger.error(f"Doublette konnte nicht gelöscht werden: '{filename}'. Fehler: {str(e)}")  # Debugging-Ausgabe: Log-File
                                        msg_file_doublette_deleted_problem_count += 1  # Problemzähler erhöhen
                            else:
                                # Wenn kein Testlauf
                                if (not TEST_RUN) and (not is_msg_file_doublette):

                                    if MAX_CONSOLE_OUTPUT: print(f"\t****************************************************")
                                    if MAX_CONSOLE_OUTPUT: print(f"\t* Wenn erforderlich, dann Umbenennung der MSG-Datei.")
                                    if MAX_CONSOLE_OUTPUT: print(f"\t****************************************************")

                                    # Umbenennen der MSG-Datei
                                    rename_msg_file_result = rename_file(old_path_and_file_name, new_path_and_file_name, max_console_output=MAX_CONSOLE_OUTPUT)

                                    # Zähler und Parameter abhängig von erfolgreicher Umbenennung setzen 
                                    if rename_msg_file_result.SUCCESS:
                                        print(f"\tErfolgreiche Umbenennung der Datei in '{new_file_name}'")
                                        app_logger.debug(f"Erfolgreiche Umbenennung der Datei '{filename}' in '{new_file_name}'")  # Debugging-Ausgabe: Log-File
                                        msg_file_renamed_count += 1  # Erfolgszähler erhöhen
                                        is_msg_file_for_change_date_available = True  # Kennzeichnung für Anpassung Erstellungs- und Änderungsdatum
                                    elif rename_msg_file_result.DESTINATION_EXISTS:
                                        print(f"\tDatei ist eine Doublette: '{filename}'")
                                        app_logger.debug(f"Datei ist eine Doublette: '{filename}'")  # Debugging-Ausgabe: Log-File
                                        msg_file_doublette_count += 1  # Erfolgszähler erhöhen
                                        is_msg_file_for_change_date_available = False  # Kennzeichnung für Anpassung Erstellungs- und Änderungsdatum
                                    else:
                                        print(f"\tUmbenennen der Datei fehlgeschlagen: '{rename_msg_file_result}'")
                                        app_logger.debug(f"Umbenennen der Datei '{filename}' fehlgeschlagen: '{rename_msg_file_result}'")  # Debugging-Ausgabe: Log-File
                                        msg_file_problem_count += 1  # Problemzähler erhöhen
                                        is_msg_file_for_change_date_available = False

                        # Wenn die Datei erfolgreich umbenannt wurde oder die Datei bereits mit korrekten Namen existiert und kein Testlauf durchgeführt wird
                        # dann soll das Erstellungsdatum auf das Versanddatum gesetzt werden
                        file_has_new_creation_date = False
                        file_has_new_modification_date = False

                        if is_msg_file_for_change_date_available and (not TEST_RUN) and SET_FILEDATE and (not is_msg_file_doublette):

                            if MAX_CONSOLE_OUTPUT: print(f"\t*************************************************************")
                            if MAX_CONSOLE_OUTPUT: print(f"\t* Wenn erforderlich, dann Zeitstempel der MSG-Datei anpassen.")
                            if MAX_CONSOLE_OUTPUT: print(f"\t*************************************************************")

                            # Setze das Erstelldatum auf das Versanddatum
                            if new_msg_filename_collection.datetime_stamp != "Unbekannt":
                                # Überprüfen, ob datetime_stamp ein datetime-Objekt ist
                                if isinstance(new_msg_filename_collection.datetime_stamp, datetime.datetime):
                                    # Konvertiere das datetime-Objekt in einen String im richtigen Format
                                    datetime_stamp_str = new_msg_filename_collection.datetime_stamp.strftime("%Y-%m-%d %H:%M:%S")
                                else:
                                    datetime_stamp_str = new_msg_filename_collection.datetime_stamp  # Annehmen, dass es bereits ein String ist

                                # Jetzt das Erstellungsdatum auf das Versanddatum setzen
                                set_creation_result = set_file_creation_date(new_path_and_file_name, datetime_stamp_str)

                                if set_creation_result == FileOperationResult.SUCCESS:
                                    msg_file_file_creation_date_count += 1
                                    file_has_new_creation_date = True
                                    if MAX_CONSOLE_OUTPUT: print(f"\tNeues Erstellungsdatum erfolgreich gesetzt.")  # Ausgabe des Ergebnisses
                                    app_logger.debug(f"Neues Erstellungsdatum für '{new_file_name}' erfolgreich gesetzt.")  # Debugging-Ausgabe: Log-File
                                elif set_creation_result == FileOperationResult.TIMESTAMP_MATCH :
                                    msg_file_creation_date_unchanged_count += 1
                                    if MAX_CONSOLE_OUTPUT: print(f"\tSetzen des Erstellungsdatum nicht erforderlich.")  # Ausgabe des Ergebnisses
                                    app_logger.debug(f"Setzen des Erstellungsdatum für '{new_file_name}' nicht erforderlich.")  # Debugging-Ausgabe: Log-File
                                else:
                                    msg_file_creation_date_problem_count += 1
                                    if MAX_CONSOLE_OUTPUT: print(f"\tFehler beim Setzen des Erstellungsdatum: '{set_creation_result}'")  # Ausgabe des Ergebnisses
                                    app_logger.debug(f"Fehler beim Setzen des Erstellungsdatum für '{new_file_name}': '{set_creation_result}'")  # Debugging-Ausgabe: Log-File

                                # Setze das Änderungsdatum auf das Versanddatum
                                set_modification_result = set_file_modification_date(new_path_and_file_name, datetime_stamp_str)
                                if set_modification_result == FileOperationResult.SUCCESS:
                                    msg_file_modification_date_count += 1
                                    file_has_new_modification_date = True
                                    if MAX_CONSOLE_OUTPUT: print(f"\tNeues Änderungsdatum erfolgreich gesetzt.")  # Ausgabe des Ergebnisses
                                    app_logger.debug(f"Neues Änderungsdatum für '{new_file_name}' erfolgreich gesetzt.")  # Debugging-Ausgabe: Log-File
                                elif set_creation_result == FileOperationResult.TIMESTAMP_MATCH :
                                    msg_file_modification_date_unchanged_count += 1
                                    if MAX_CONSOLE_OUTPUT: print(f"\tSetzen des Änderungsdatum nicht erforderlich.")  # Ausgabe des Ergebnisses
                                    app_logger.debug(f"Setzen des Änderungsdatum für '{new_file_name}' nicht erforderlich.")  # Debugging-Ausgabe: Log-File
                                else:
                                    msg_file_modification_date_problem_count += 1
                                    if MAX_CONSOLE_OUTPUT: print(
                                        f"\tFehler beim Setzen des Änderungsdatum: '{set_creation_result}'")  # Ausgabe des Ergebnisses
                                    app_logger.debug(
                                        f"Fehler beim Setzen des Änderungsdatum für '{new_file_name}': '{set_creation_result}'")  # Debugging-Ausgabe: Log-File
                            else:
                                msg_file_creation_date_problem_count += 1
                                msg_file_modification_date_count += 1
                                if MAX_CONSOLE_OUTPUT: print(f"\tKein Versanddatum der MSG-Datei verfügbar.")  # Ausgabe des Ergebnisses
                                app_logger.debug(f"Kein Versanddatum der MSG-Datei verfügbar.")  # Debugging-Ausgabe: Log-File

                        # Wenn GENERATE_PDF True ist, wird eine PDF-Datei aus der MSG-Datei erstellt
                        if GENERATE_PDF and (not is_msg_file_doublette):
                            if MAX_CONSOLE_OUTPUT: print(f"\t***********************************************************")
                            if MAX_CONSOLE_OUTPUT: print(f"\t* Wenn erforderlich, PDF-Datei erzeugen.")
                            if MAX_CONSOLE_OUTPUT: print(f"\t***********************************************************")

                            pdf_path = os.path.splitext(new_path_and_file_name)[0] + ".pdf"

                            # Wenn der -opdf Flag False ist und die PDF-Datei bereits existiert, wird sie nicht überschrieben
                            if not OVERWRITE_PDF and os.path.exists(pdf_path):
                                if MAX_CONSOLE_OUTPUT: print(f"\tPDF-Datei '{pdf_path}' existiert bereits und -opdf ist False. Überspringe Erstellung.")
                                app_logger.info(f"PDF-Datei '{pdf_path}' existiert bereits und -opdf ist False. Überspringe Erstellung.")
                                pdf_file_skipped += 1
                                is_pdf_file_skipped = True
                            else:
                                generate_pdf_from_msg(new_path_and_file_name, 800)
                                if MAX_CONSOLE_OUTPUT: print(f"\tPDF-Datei '{pdf_path}' erzeugt.")
                                app_logger.info(f"PDF-Datei '{pdf_path}' erzeugt.")
                                pdf_file_generated += 1
                                is_pdf_file_generated = True

                elif FileAccessStatus.READABLE in access_result:
                    if MAX_CONSOLE_OUTPUT: print(f"\tNur lesender Zugriff auf die Datei möglich.")
                    app_logger.info(f"Nur lesender Zugriff auf die Datei möglich.")
                    msg_file_problem_count += 1  # Problemzähler erhöhen

                else:
                    if MAX_CONSOLE_OUTPUT: print(f"\tWeder lesender noch schreibender Zugriff auf die Datei möglich: '{[s.value for s in access_result]}'")
                    app_logger.info(f"Weder lesender noch schreibender Zugriff auf die Datei möglich: '{[s.value for s in access_result]}'")
                    msg_file_problem_count += 1  # Problemzähler erhöhen

                # Logeintrag erstellen
                entry = {
                    "Fortlaufende Nummer": msg_file_count,
                    "Verzeichnisname": pathname,
                    "Original-Filename": filename,
                    "Alter absoluter Dateiname": path_and_file_name,
                    "Alte Pfadlänge": path_and_file_name_length,
                    "Versanddatum": new_msg_filename_collection.datetime_stamp,
                    "Formatiertes Versanddatum": new_msg_filename_collection.formatted_timestamp,
                    "Gefundener Absender": new_msg_filename_collection.sender_name,
                    "Gefundener Email-Absender": new_msg_filename_collection.sender_email,
                    "Betreff": new_msg_filename_collection.msg_subject,
                    "Bereinigter Betreff": new_msg_filename_collection.msg_subject_sanitized,
                    "Neuer absoluter Dateiname": new_path_and_file_name,
                    "Neue nicht gekürzte Pfadlänge": new_path_and_file_name_length,
                    "Neuer gekürzter Dateiname": new_msg_filename_collection.new_truncated_msg_filename,
                    "Kürzung Dateiname erforderlich": new_msg_filename_collection.is_msg_filename_truncated,
                    "Alter und neuer Name sind gleich": is_msg_file_name_unchanged,
                    "Neues Erstellungsdatum": file_has_new_creation_date,
                    "Neues Änderungsdatum": file_has_new_modification_date,
                    "Doublette": is_msg_file_doublette,
                    "Doublette gelöscht": is_msg_file_doublette_deleted,
                    "PDF erstellt": is_pdf_file_generated,
                    "PDF übersprungen": is_pdf_file_skipped
                }

                # Eintrag ins Logfile hinzufügen
                log_entry_neu(excel_log_file_path, entry, sheet_name="Log")

        # Wenn keine rekursive Suche gewünscht ist, wird die Schleife beendet
        if not RECURSIVE_SEARCH: break

    # Ausgabe der wichtigsten Konfigurationen
    print(f"\nÜbersicht der Konfigurationen:")
    app_logger.info(f"Übersicht der Konfigurationen:")
    print(f"******************************")
    print(f"Testlauf? {TEST_RUN}")
    app_logger.info(f"Testlauf? {TEST_RUN}")
    print(f"Testverzeichnis initialisieren? {INIT_TESTDATA}")
    app_logger.info(f"Testverzeichnis initialisieren? {INIT_TESTDATA}")
    if INIT_TESTDATA: print(f"Pfad zur Quelle der Testdaten: {SOURCE_DIRECTORY_TEST_DATA}")
    app_logger.info(f"Pfad zur Quelle der Testdaten: {SOURCE_DIRECTORY_TEST_DATA}")
    print(f"Verzeichnis für die Suche nach MSG-Dateien: {TARGET_DIRECTORY}")
    app_logger.info(f"Verzeichnis für die Suche nach MSG-Dateien: {TARGET_DIRECTORY}")
    print(f"Rekursive Suche? {RECURSIVE_SEARCH}")
    app_logger.info(f"Rekursive Suche? {RECURSIVE_SEARCH}")
    print(f"Bei Bedarf in der Tabelle der bekannten Email-Absender? {USE_KNOWNSENDER_FILE}")
    app_logger.info(f"Bei Bedarf in der Tabelle der bekannten Email-Absender? {USE_KNOWNSENDER_FILE}")
    if USE_KNOWNSENDER_FILE:
        print(f"Pfad zur Datei der bekannten Email-Absender: {KNOWNSENDER_FILE}")
        app_logger.info(f"Pfad zur Datei der bekannten Email-Absender: {KNOWNSENDER_FILE}")
    print(f"Zeitstempel der MSG-Dateien anpassen? {SET_FILEDATE}")
    app_logger.info(f"Zeitstempel der MSG-Dateien anpassen? {SET_FILEDATE}")
    if GENERATE_PDF:
        print(f"PDF-Dateien erzeugen? {GENERATE_PDF}")
        app_logger.info(f"PDF-Dateien erzeugen? {GENERATE_PDF}")
        print(f"Existierende PDF-Dateien überschreiben? {OVERWRITE_PDF}")
        app_logger.info(f"Existierende PDF-Dateien überschreiben? {OVERWRITE_PDF}")
    print(f"Debug-Mode? {DEBUG_MODE}")
    app_logger.info(f"Debug-Mode? {DEBUG_MODE}")
    print(f"Debug-Datei: {prog_log_file_path}")
    app_logger.info(f"Debug-Datei: {prog_log_file_path}")
    print(f"Excel-Log-Datei: {excel_log_file_path}")
    app_logger.info(f"Excel-Log-Datei: {excel_log_file_path}")

    # Schreibe Konfiguration
    entry = [
        { "Konfiguration": "Testlauf?", "Wert": TEST_RUN },
        { "Konfiguration": "Testverzeichnis initialisieren?", "Wert": INIT_TESTDATA },
        { "Konfiguration": "Verzeichnis für die Suche nach MSG-Dateien", "Wert": TARGET_DIRECTORY },
        { "Konfiguration": "Bei Bedarf in der Tabelle der bekannten Email-Absender?", "Wert": RECURSIVE_SEARCH },
        { "Konfiguration": "Rekursive Suche?", "Wert": USE_KNOWNSENDER_FILE },
        { "Konfiguration": "Pfad zur Datei der bekannten Email-Absender", "Wert": KNOWNSENDER_FILE },
        { "Konfiguration": "Zeitstempel der MSG-Dateien anpassen?", "Wert": SET_FILEDATE },
        { "Konfiguration": "PDF-Dateien erzeugen?", "Wert": GENERATE_PDF },
        { "Konfiguration": "Existierende PDF-Dateien überschreiben?", "Wert": OVERWRITE_PDF },
        { "Konfiguration": "Debug-Mode?", "Wert": DEBUG_MODE },
        { "Konfiguration": "Debug-Datei", "Wert": prog_log_file_path },
        { "Konfiguration": "Excel-Log-Datei", "Wert": excel_log_file_path }
    ]
    log_entry_neu(excel_log_file_path, entry, sheet_name="Konfiguration")

    # Ausgabe der Ergebnisse
    print(f"\nErgebnisse (auch bei Testlauf):")
    app_logger.info("Ergebnisse (auch bei Testlauf):")
    print(f"*******************************")
    print(f"Anzahl der gefundenen MSG-Dateien: {msg_file_count}")
    app_logger.info(f"Anzahl der gefundenen MSG-Dateien: {msg_file_count}")
    print(f"Anzahl der bereits mit korrekten Namen existierenden MSG-Dateien: {msg_file_same_name_count}")
    app_logger.info(f"Anzahl der bereits mit korrekten Namen existierenden MSG-Dateien: {msg_file_same_name_count}")
    print(f"Anzahl der Dateien mit Problemen: {msg_file_problem_count}")
    app_logger.info(f"Anzahl der Dateien mit Problemen: {msg_file_problem_count}")
    print(f"Anzahl gekürzte Dateinamen: {msg_file_shorted_name_count}")
    app_logger.info(f"Anzahl gekürzte Dateinamen: {msg_file_shorted_name_count}")

    # Schreibe Zusammenfassung Sheet Teil 1
    entry = [
        { "Ergebnis": "Anzahl der gefundenen MSG-Dateien", "Wert": msg_file_count },
        { "Ergebnis": "Anzahl der bereits mit korrekten Namen existierenden MSG-Dateien", "Wert": msg_file_same_name_count },
        { "Ergebnis": "Anzahl der Dateien mit Problemen", "Wert": msg_file_problem_count },
        { "Ergebnis": "Anzahl gekürzte Dateinamen", "Wert": msg_file_shorted_name_count }
    ]
    log_entry_neu(excel_log_file_path, entry, sheet_name="Zusammenfassung")

    if not TEST_RUN:
        print(f"\nErgebnisse der Anpassungen:")
        app_logger.info(f"Ergebnisse der Anpassungen:")
        print(f"***************************")
        print(f"Anzahl der umbenannten Dateien: {msg_file_renamed_count}")
        app_logger.info(f"Anzahl der umbenannten Dateien: {msg_file_renamed_count}")
        print(f"Anzahl gefundener Doubletten: {msg_file_doublette_count}")
        app_logger.info(f"Anzahl gefundener Doubletten: {msg_file_doublette_count}")
        print(f"Anzahl gelöschter Doubletten: {msg_file_doublette_deleted_count}")
        app_logger.info(f"Anzahl gelöschter Doubletten: {msg_file_doublette_deleted_count}")
        print(f"Anzahl nicht gelöschter Doubletten: {msg_file_doublette_deleted_problem_count}")
        app_logger.info(f"Anzahl nicht gelöschter Doubletten: {msg_file_doublette_deleted_problem_count}")

        # Schreibe Zusammenfassung Sheet Teil 2
        entry = [
            { "Ergebnis": "Anzahl der umbenannten Dateien", "Wert": msg_file_renamed_count },
            { "Ergebnis": "Anzahl gefundener Doubletten", "Wert": msg_file_doublette_count },
            { "Ergebnis": "Anzahl gelöschter Doubletten", "Wert": msg_file_doublette_deleted_count },
            { "Ergebnis": "Anzahl nicht gelöschter Doubletten", "Wert": msg_file_doublette_deleted_problem_count }
        ]
        log_entry_neu(excel_log_file_path, entry, sheet_name="Zusammenfassung")

        if SET_FILEDATE:
            print(f"\nErgebnisse bei Anpassung File-Datum:")
            app_logger.info(f"Ergebnisse bei Anpassung File-Datum:")
            print(f"************************************")
            print(f"Anzahl MSG-Dateien mit geändertem Erstellungsdatum: {msg_file_file_creation_date_count}")
            app_logger.info(f"Anzahl MSG-Dateien mit geändertem Erstellungsdatum: {msg_file_file_creation_date_count}")
            print(f"Anzahl MSG-Dateien mit nicht geändertem Erstellungsdatum: {msg_file_creation_date_unchanged_count}")
            app_logger.info(f"Anzahl MSG-Dateien mit nicht geändertem Erstellungsdatum: {msg_file_creation_date_unchanged_count}")
            print(f"Anzahl MSG-Dateien wo Anpassung Erstellungsdatum nicht möglich: {msg_file_creation_date_problem_count}")
            app_logger.info(f"Anzahl MSG-Dateien wo Anpassung Erstellungsdatum nicht möglich: {msg_file_creation_date_problem_count}")
            print(f"Anzahl MSG-Dateien mit geändertem Änderungsdatum: {msg_file_modification_date_count}")
            app_logger.info(f"Anzahl MSG-Dateien mit geändertem Änderungsdatum: {msg_file_modification_date_count}")
            print(f"Anzahl MSG-Dateien mit nicht geändertem Änderungsdatum: {msg_file_modification_date_unchanged_count}")
            app_logger.info(f"Anzahl MSG-Dateien mit nicht geändertem Änderungsdatum: {msg_file_modification_date_unchanged_count}")
            print(f"Anzahl MSG-Dateien wo Anpassung Änderungsdatum nicht möglich: {msg_file_modification_date_problem_count}")
            app_logger.info(f"Anzahl MSG-Dateien wo Anpassung Änderungsdatum nicht möglich: {msg_file_modification_date_problem_count}")

            # Schreibe Zusammenfassung Sheet Teil 3
            entry = [
                { "Ergebnis": "Anzahl MSG-Dateien mit geändertem Erstellungsdatum", "Wert": msg_file_file_creation_date_count },
                { "Ergebnis": "Anzahl MSG-Dateien mit nicht geändertem Erstellungsdatum", "Wert": msg_file_creation_date_unchanged_count },
                { "Ergebnis": "Anzahl MSG-Dateien wo Anpassung Erstellungsdatum nicht möglich", "Wert": msg_file_creation_date_problem_count },
                { "Ergebnis": "Anzahl MSG-Dateien mit geändertem Änderungsdatum", "Wert": msg_file_modification_date_count },
                { "Ergebnis": "Anzahl MSG-Dateien mit nicht geändertem Änderungsdatum", "Wert": msg_file_modification_date_unchanged_count },
                { "Ergebnis": "Anzahl MSG-Dateien wo Anpassung Änderungsdatum nicht möglich", "Wert": msg_file_modification_date_problem_count }
            ]
            log_entry_neu(excel_log_file_path, entry, sheet_name="Zusammenfassung")

        if GENERATE_PDF:
            print(f"\nErgebnis der PDF-Erzeugung:")
            app_logger.info(f"Ergebnis der PDF-Erzeugung:")
            print(f"****************************")
            print(f"Anzahl der erzeugten PDF-Dateien: {pdf_file_generated}")
            app_logger.info(f"Anzahl der erzeugten PDF-Dateien: {pdf_file_generated}")
            print(f"Anzahl der übersprungenen PDF-Dateien: {pdf_file_skipped}")
            app_logger.info(f"Anzahl der übersprungenen PDF-Dateien: {pdf_file_skipped}")

            # Schreibe Zusammenfassung Sheet Teil 4
            entry = [
                { "Ergebnis": "Anzahl der erzeugten PDF-Dateien", "Wert": pdf_file_generated },
                { "Ergebnis": "Anzahl der übersprungenen PDF-Dateien", "Wert": pdf_file_skipped }
            ]
            log_entry_neu(excel_log_file_path, entry, sheet_name="Zusammenfassung")

