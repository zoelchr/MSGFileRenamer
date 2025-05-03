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

Sepzielle Parameter zur Konfiguration
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
--min_console_output
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
import logging
import datetime
import argparse
import sys
import importlib.util
import io
import pandas as pd

from modules.msg_generate_new_filename import generate_new_msg_filename
from utils.file_handling import (rename_file, test_file_access, FileAccessStatus, set_file_creation_date, set_file_modification_date, FileOperationResult)
from modules.msg_handling import create_log_file, log_entry, log_entry_neu, create_log_file_neu
from utils.testset_preparation import prepare_test_directory
from utils.pdf_generation import generate_pdf_from_msg
from pathlib import Path

print("")
os.system('chcp 65001')  # Setzt Konsole auf UTF-8
#original_stdout = sys.stdout
#sys.stdout = io.TextIOWrapper(sys.stdout.detach(), encoding='utf-8')
print("")

# Verzeichnisse für die Tests definieren
# SOURCE_DIRECTORY_TEST_DATA = r'.\data\sample_files\testset-short-longpath'
SOURCE_DIRECTORY_TEST_DATA = r'.\data\sample_files\testset-public'
# SOURCE_DIRECTORY_TEST_DATA = r'.\data\sample_files\testset-short'
TARGET_DIRECTORY_TEST_DATA = r'.\tests\functional\testdir'
TARGET_DIRECTORY = ""
LIST_OF_KNOWN_SENDERS = r'.\config\known_senders.csv' # Liste der bekannten Email-Absender aus einer CSV-Datei

# Maximal zulässige Pfadlänge für Windows 11
MAX_PATH_LENGTH = 260

def setup_logging(file, debug=False):
    log_level=logging.DEBUG if debug else logging.INFO
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(name)s - %(funcName)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
        filename=file
    )

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
        print(f"Fehler: Das Modul '{module_name}' ist nicht installiert.")
        print("Bitte installieren Sie das fehlende Modul und starten Sie das Programm erneut.\n")
        print("Dazu zuerst in das Verzeichnis mit der Datei 'msg_file_renamer.bat' wechseln und dort ein Windows Terminal öffnen.")
        print(r"Dann die zugehörige virtuelle Python-Umgebung durch '.\venv\Scripts\activate.bat' aktivieren.")
        print("Eventuell hat das Verzeichnis für die virtuelle Python-Umgebung statt 'venv' auch einen anderen Namen.")
        print(f"Anschließend kann das fehlende Modul über die folgende Eingabe installiert werden:\n '{install_text}'")
        sys.exit(1)

if __name__ == '__main__':

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
    parser.add_argument("-cf", "--knownsender_file", type=str, default="LIST_OF_KNOWN_SENDERS", help="CSV-Datei mit Liste der bekannten Absender")
    parser.add_argument("-mco", "--max_console_output", default=False, action="store_true", help="Maximale Consolen-Ausgabe aktivieren (Default=False)")
    args, unknown = parser.parse_known_args()

    # Unbekannte Parameter ausgeben und Programm beenden
    if unknown:
        print(f"Warnung: Unbekannte Parameter gefunden: {unknown}")
        print("Bitte überprüfen Sie die übergebenen Parameter.")
        exit()

    args_formatted = str(args).replace(",", "\n\t\t").replace("Namespace", "").strip("()")

    # Argumente des Programmaufrufs an die Variablen übergeben
    INIT_TESTDATA = args.init_testdata
    SET_FILEDATE = args.set_filedate
    DEBUG_MODE = args.debug_mode
    TEST_RUN = not args.no_test_run
    NO_SHORTEN_PATH_NAME = args.no_shorten_path_name
    GENERATE_PDF = args.generate_pdf
    OVERWRITE_PDF = args.overwrite_pdf
    RECURSIVE_SEARCH = args.recursive_search
    USE_KNOWNSENDER_FILE = args.use_knownsender_file
    KNOWNSENDER_FILE = args.knownsender_file
    MAX_CONSOLE_OUTPUT = args.max_console_output

    if MAX_CONSOLE_OUTPUT: print(f"\n****************************************************")
    if MAX_CONSOLE_OUTPUT: print(f"* Informationen zum Programmstart.")
    if MAX_CONSOLE_OUTPUT: print(f"****************************************************")
    if MAX_CONSOLE_OUTPUT: print(f"\nÜbergebene Argumente:\n\t\t {args_formatted}") # Debugging-Ausgabe: Console

    # Start Ausgabe auf Console
    if MAX_CONSOLE_OUTPUT: print(f"\nTestlauf: {TEST_RUN}\nTestverzeichnis initialisieren: {INIT_TESTDATA}\nZeitstempel der MSG-dateien anpassen: {SET_FILEDATE}\nDebug-Modus: {DEBUG_MODE}")

    TARGET_DIRECTORY = Path(args.search_directory) # Verzeichnis für die Suche nach MSG-Dateien

    # Excel-Log-Datei
    excel_log_directory = args.excel_log_directory # Verzeichnis für die Excel-Log-Datei
    excel_log_basename = args.excel_log_basename # Basisname für die Excel-Log-Datei
    excel_log_file_path = os.path.join(excel_log_directory, excel_log_basename) # Pfad und Dateiname für die Excel-Log-Datei

    # Zusätzlicher Sheetname
    excel_log_sheet_name = "Ergebnis"
    
    # Debug-Datei (nutzt den gleichen Namen wie die Excel-Log-Datei, aber mit Endung *.txt)
    debug_log_directory = args.debug_log_directory
    prog_log_file_path = os.path.join(debug_log_directory, "debug_msgfilerenamer" + ".txt") # Pfad und Dateiname für die Debug-Log-Datei

    # Logging für Test und Debugging konfigurieren
    setup_logging(prog_log_file_path, DEBUG_MODE)
    logger = logging.getLogger(prog_log_file_path)
    # Logging für die Programm-Ausführung
    logger.info("Programm gestartet")
    logger.info(f"Debug-Modus: {DEBUG_MODE}")
    logging.info(f"Programm Logdatei: {prog_log_file_path}")

    # Verzeichnis für die Suche festlegen
    if not TARGET_DIRECTORY:
        TARGET_DIRECTORY = TARGET_DIRECTORY_TEST_DATA

    print(f"\n************************************************************************")
    print(f"Verzeichnis für die Suche: {TARGET_DIRECTORY}") # Debugging-Ausgabe: Console
    print(f"************************************************************************")
    logging.debug(f"Verzeichnis für die Suche: {TARGET_DIRECTORY}")  # Debugging-Ausgabe: Log-File

    # Log-Verzeichnis und Basisname für Excel-Logdateien festlegen
    LOG_TABLE_HEADER = ["Fortlaufende Nummer", "Verzeichnisname", "Original-Filename"]

    # Excel-Logdatei erstellen und den Pfad ausgeben
    excel_log_file_path = create_log_file_neu(excel_log_basename, excel_log_directory, LOG_TABLE_HEADER, sheet_name="Log")
    if MAX_CONSOLE_OUTPUT: print(f"Excel-Logdatei erstellt: {excel_log_file_path}")

    # Abhängig von INIT_TESTDATA wird das Testverzeichnis vorbereitet oder nicht
    if MAX_CONSOLE_OUTPUT: print(f"Soll das Testverzeichnis vorbereitet werden? {INIT_TESTDATA}") # Debugging-Ausgabe: Console
    logging.debug(f"Soll das Testverzeichnis vorbereitet werden? {INIT_TESTDATA}")  # Debugging-Ausgabe: Log-File

    if INIT_TESTDATA:
        if MAX_CONSOLE_OUTPUT: print(f"Prüfung ob Zielverzeichnis für Testdaten bereits existiert: {TARGET_DIRECTORY_TEST_DATA}") # Debugging-Ausgabe: Console
        logging.debug(f"Prüfung ob Zielverzeichnis für Testdaten bereits existiert: {TARGET_DIRECTORY_TEST_DATA}")  # Debugging-Ausgabe: Log-File

        # Wenn TARGET_DIRECTORY_TEST_DATA noch nicht existiert, dann erstellen
        if not os.path.isdir(TARGET_DIRECTORY_TEST_DATA):
            os.makedirs(TARGET_DIRECTORY_TEST_DATA, exist_ok=True)
            if MAX_CONSOLE_OUTPUT: print(f"Zielverzeichnis neu erstellt: {TARGET_DIRECTORY_TEST_DATA}")
        else:
            if MAX_CONSOLE_OUTPUT: print(f"Zielverzeichnis existiert bereits: {TARGET_DIRECTORY_TEST_DATA}")

        # Bereite das Zielverzeichnis vor und überprüfe den Erfolg
        success = prepare_test_directory(SOURCE_DIRECTORY_TEST_DATA, TARGET_DIRECTORY_TEST_DATA)
        if success:
            if MAX_CONSOLE_OUTPUT: print("Vorbereitung des Testverzeichnisses erfolgreich abgeschlossen.") # Debugging-Ausgabe: Console
            logging.error("\tVorbereitung des Testverzeichnisses erfolgreich abgeschlossen.")  # Debugging-Ausgabe: Log-File
        else:
            if MAX_CONSOLE_OUTPUT: print("Fehler: Die Vorbereitung des Testverzeichnisses ist fehlgeschlagen. Das Programm wird abgebrochen.") # Debugging-Ausgabe: Console
            logging.error("Fehler: Die Vorbereitung des Testverzeichnisses ist fehlgeschlagen. Das Programm wird abgebrochen.")  # Debugging-Ausgabe: Log-File
            exit(1)  # Programm abbrechen
    else:
        if MAX_CONSOLE_OUTPUT: print(f"Kein Testverzeichnis vorbereitet.") # Debugging-Ausgabe: Console
        logging.debug(f"Kein Testverzeichnis vorbereitet.")  # Debugging-Ausgabe: Log-File

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

    # Durchsuchen des Zielverzeichnisses nach MSG-Dateien
    # pathname = Verzeichnisname, dirs = Unterverzeichnisse, files = List von Dateien
    for pathname, dirs, files in os.walk(TARGET_DIRECTORY):
        # filename = Dateiname
        for filename in files:
            logging.debug(f"\n******************************************************************************************************")  # Debugging-Ausgabe: Log-File

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
                logging.debug(f"\nMSG-Datei: '{filename}'")  # Debugging-Ausgabe: Log-File

                # Absoluter Pfadname der MSG-Datei
                path_and_file_name = os.path.join(pathname, filename)
                # Add the \\?\ prefix to support long paths on Windows
                if os.name == 'nt':  # Check if the OS is Windows
                    path_and_file_name = f"\\\\?\\{os.path.abspath(path_and_file_name)}"
                # Pfadlänge ermitteln
                path_and_file_name_length = len(path_and_file_name)
                logging.debug(f"Pfadlänge aktuelle MSG-Datei: '{path_and_file_name_length}'")  # Debugging-Ausgabe: Log-File

                if MAX_CONSOLE_OUTPUT: print(f"\tAktuelles Verzeichnis: '{os.path.dirname(path_and_file_name)}'")  # Debugging-Ausgabe: Console
                logging.debug(f"\tAktuelles Verzeichnis: '{os.path.dirname(path_and_file_name)}'")  # Debugging-Ausgabe: Log-File

                msg_file_count += 1 # Zähler erhöhen, MSG-Datei gefunden

                # Überprüfen den Schreib- und Lesezugriff auf die MSG-Datei
                access_result = test_file_access(path_and_file_name)

                if MAX_CONSOLE_OUTPUT: print(f"\tÜberprüfung Zugriff auf aktuelle MSG-Date: {[s.value for s in access_result]}'")  # Debugging-Ausgabe: Console
                logging.debug(f"\tÜberprüfung Zugriff auf aktuelle MSG-Date: {[s.value for s in access_result]}'")  # Debugging-Ausgabe: Log-File

                # Nur wenn die MSG-Datei schreibend geöffnet werden kann, ist ein Umbenennen möglich
                if FileAccessStatus.WRITABLE in access_result:
                    if MAX_CONSOLE_OUTPUT: print(f"\tSchreibender Zugriff auf die Datei ist möglich.")  # Debugging-Ausgabe: Console
                    logging.debug(f"\tSchreibender Zugriff auf die Datei ist möglich: {filename}")  # Debugging-Ausgabe: Log-File

                    # Neuen Datenamen erzeugen
                    new_msg_filename_collection = generate_new_msg_filename(path_and_file_name, use_list_of_known_senders=USE_KNOWNSENDER_FILE, file_list_of_known_senders=KNOWNSENDER_FILE, max_console_output=MAX_CONSOLE_OUTPUT)

                    # Wenn Dateiname gekürzt wurde, dann Zähler erhöhen
                    # if new_msg_filename_collection.is_msg_filename_truncated: msg_file_shorted_name_count += 1

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
                                logging.debug(f"Neuer gekürzter Dateiname: '{new_file_name}'")  # Debugging-Ausgabe: Log-File

                        # Neuen absoluten Pfad erzeugen
                        new_path_and_file_name = os.path.join(pathname, new_file_name)
                        # Add the \\?\ prefix to support long paths on Windows
                        if os.name == 'nt':  # Check if the OS is Windows
                            new_path_and_file_name = f"\\\\?\\{os.path.abspath(new_path_and_file_name)}"

                        if MAX_CONSOLE_OUTPUT: print(f"\tNeuer absoluter Pfad: '{new_path_and_file_name}'")
                        logging.debug(f"Neuer absoluter Pfad: '{new_path_and_file_name}'")  # Debugging-Ausgabe: Log-File

                        # Pfadlänge ermitteln
                        new_path_and_file_name_length = len(new_path_and_file_name)
                        if MAX_CONSOLE_OUTPUT: print(f"\tPfadlänge neue MSG-Datei: '{new_path_and_file_name_length}'")
                        logging.debug(f"Pfadlänge neue MSG-Datei: '{new_path_and_file_name_length}'")  # Debugging-Ausgabe: Log-File

                        # Prüfen, ob Alter und neuer Name gleich, dann keine Änderung erforderlich
                        if old_path_and_file_name == new_path_and_file_name:
                            print(f"\tAlter und neuer Dateiname sind gleich.")
                            logging.debug(f"Alter und neuer Dateiname sind gleich: '{filename}'")  # Debugging-Ausgabe: Log-File
                            msg_file_same_name_count += 1  # Erfolgszähler erhöhen
                            is_msg_file_name_unchanged = True # Kennzeichnung keine Änderung des Dateinamens erforderlich
                            is_msg_file_for_change_date_available = True # Kennzeichnung für Anpassung Erstellungs- und Änderungsdatum
                        else:
                            # Prüfen, ob die Datei mit neuem Namen bereits existiert, also Doublette
                            if os.path.exists(new_path_and_file_name):
                                print(f"\tDatei ist eine Doublette: '{filename}'")
                                logging.debug(f"Datei ist eine Doublette: '{filename}'")  # Debugging-Ausgabe: Log-File
                                msg_file_doublette_count += 1  # Erfolgszähler erhöhen
                                is_msg_file_doublette = True # MSG-Datei mit gleichem neuen Namen existiert bereits - also Doublette
                                is_msg_file_for_change_date_available = False  # Kennzeichnung für Anpassung Erstellungs- und Änderungsdatum

                                # Versuche Doublette zu löschen, wenn nicht Test
                                if not TEST_RUN:
                                    # Versuche, die Datei zu löschen
                                    try:
                                        os.remove(old_path_and_file_name)  # Versuche, die Datei zu löschen
                                        print(f"\tDoublette gelöscht: '{filename}'")
                                        logging.debug(f"Doublette gelöscht: '{filename}'")  # Debugging-Ausgabe: Log-File
                                        msg_file_doublette_deleted_count += 1  # Löschzähler erhöhen
                                        is_msg_file_doublette_deleted = True
                                    except Exception as e:
                                        print(f"\tDoublette konnte nicht gelöscht werden: '{filename}'. Fehler: {str(e)}")
                                        logging.error(f"Doublette konnte nicht gelöscht werden: '{filename}'. Fehler: {str(e)}")  # Debugging-Ausgabe: Log-File
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
                                        logging.debug(f"Erfolgreiche Umbenennung der Datei '{filename}' in '{new_file_name}'")  # Debugging-Ausgabe: Log-File
                                        msg_file_renamed_count += 1  # Erfolgszähler erhöhen
                                        is_msg_file_for_change_date_available = True  # Kennzeichnung für Anpassung Erstellungs- und Änderungsdatum
                                    elif rename_msg_file_result.DESTINATION_EXISTS:
                                        print(f"\tDatei ist eine Doublette: '{filename}'")
                                        logging.debug(f"Datei ist eine Doublette: '{filename}'")  # Debugging-Ausgabe: Log-File
                                        msg_file_doublette_count += 1  # Erfolgszähler erhöhen
                                        is_msg_file_for_change_date_available = False  # Kennzeichnung für Anpassung Erstellungs- und Änderungsdatum
                                    else:
                                        print(f"\tUmbenennen der Datei fehlgeschlagen: '{rename_msg_file_result}'")
                                        logging.debug(f"Umbenennen der Datei '{filename}' fehlgeschlagen: '{rename_msg_file_result}'")  # Debugging-Ausgabe: Log-File
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
                                    logging.debug(f"Neues Erstellungsdatum für '{new_file_name}' erfolgreich gesetzt.")  # Debugging-Ausgabe: Log-File
                                elif set_creation_result == FileOperationResult.TIMESTAMP_MATCH :
                                    msg_file_creation_date_unchanged_count += 1
                                    if MAX_CONSOLE_OUTPUT: print(f"\tSetzen des Erstellungsdatum nicht erforderlich.")  # Ausgabe des Ergebnisses
                                    logging.debug(f"Setzen des Erstellungsdatum für '{new_file_name}' nicht erforderlich.")  # Debugging-Ausgabe: Log-File
                                else:
                                    msg_file_creation_date_problem_count += 1
                                    if MAX_CONSOLE_OUTPUT: print(f"\tFehler beim Setzen des Erstellungsdatum: '{set_creation_result}'")  # Ausgabe des Ergebnisses
                                    logging.debug(f"Fehler beim Setzen des Erstellungsdatum für '{new_file_name}': '{set_creation_result}'")  # Debugging-Ausgabe: Log-File

                                # Setze das Änderungsdatum auf das Versanddatum
                                set_modification_result = set_file_modification_date(new_path_and_file_name, datetime_stamp_str)
                                if set_modification_result == FileOperationResult.SUCCESS:
                                    msg_file_modification_date_count += 1
                                    file_has_new_modification_date = True
                                    if MAX_CONSOLE_OUTPUT: print(f"\tNeues Änderungsdatum erfolgreich gesetzt.")  # Ausgabe des Ergebnisses
                                    logging.debug(f"Neues Änderungsdatum für '{new_file_name}' erfolgreich gesetzt.")  # Debugging-Ausgabe: Log-File
                                elif set_creation_result == FileOperationResult.TIMESTAMP_MATCH :
                                    msg_file_modification_date_unchanged_count += 1
                                    if MAX_CONSOLE_OUTPUT: print(f"\tSetzen des Änderungsdatum nicht erforderlich.")  # Ausgabe des Ergebnisses
                                    logging.debug(f"Setzen des Änderungsdatum für '{new_file_name}' nicht erforderlich.")  # Debugging-Ausgabe: Log-File
                                else:
                                    msg_file_modification_date_problem_count += 1
                                    if MAX_CONSOLE_OUTPUT: print(
                                        f"\tFehler beim Setzen des Änderungsdatum: '{set_creation_result}'")  # Ausgabe des Ergebnisses
                                    logging.debug(
                                        f"Fehler beim Setzen des Änderungsdatum für '{new_file_name}': '{set_creation_result}'")  # Debugging-Ausgabe: Log-File
                            else:
                                msg_file_creation_date_problem_count += 1
                                msg_file_modification_date_count += 1
                                if MAX_CONSOLE_OUTPUT: print(f"\tKein Versanddatum der MSG-Datei verfügbar.")  # Ausgabe des Ergebnisses
                                logging.debug(f"Kein Versanddatum der MSG-Datei verfügbar.")  # Debugging-Ausgabe: Log-File

                        # Wenn GENERATE_PDF True ist, wird eine PDF-Datei aus der MSG-Datei erstellt
                        if GENERATE_PDF and (not is_msg_file_doublette):
                            if MAX_CONSOLE_OUTPUT: print(f"\t***********************************************************")
                            if MAX_CONSOLE_OUTPUT: print(f"\t* Wenn erforderlich, PDF-Datei erzeugen.")
                            if MAX_CONSOLE_OUTPUT: print(f"\t***********************************************************")

                            pdf_path = os.path.splitext(new_path_and_file_name)[0] + ".pdf"

                            # Wenn der -opdf Flag False ist und die PDF-Datei bereits existiert, wird sie nicht überschrieben
                            if not OVERWRITE_PDF and os.path.exists(pdf_path):
                                if MAX_CONSOLE_OUTPUT: print(f"\tPDF-Datei '{pdf_path}' existiert bereits und -opdf ist False. Überspringe Erstellung.")
                                logging.info(f"PDF-Datei '{pdf_path}' existiert bereits und -opdf ist False. Überspringe Erstellung.")
                                pdf_file_skipped += 1
                            else:
                                generate_pdf_from_msg(new_path_and_file_name, 800)
                                if MAX_CONSOLE_OUTPUT: print(f"\tPDF-Datei '{pdf_path}' erzeugt.")
                                logging.info(f"PDF-Datei '{pdf_path}' erzeugt.")
                                pdf_file_generated += 1

                elif FileAccessStatus.READABLE in access_result:
                    if MAX_CONSOLE_OUTPUT: print(f"\tNur lesender Zugriff auf die Datei möglich.")
                    logging.info(f"Nur lesender Zugriff auf die Datei möglich.")
                    msg_file_problem_count += 1  # Problemzähler erhöhen

                else:
                    if MAX_CONSOLE_OUTPUT: print(f"\tWeder lesender noch schreibender Zugriff auf die Datei möglich: '{[s.value for s in access_result]}'")
                    logging.info(f"eder lesender noch schreibender Zugriff auf die Datei möglich: '{[s.value for s in access_result]}'")
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
        if (not RECURSIVE_SEARCH): break

    # Ausgabe der wichtigsten Konfigurationen
    print(f"\nÜbersicht der Konfigurationen:")
    print(f"******************************")
    print(f"Testlauf? {TEST_RUN}")
    print(f"Testverzeichnis initialisieren? {INIT_TESTDATA}")
    if INIT_TESTDATA: print(f"Pfad zur Quelle der Testdaten: {SOURCE_DIRECTORY_TEST_DATA}")
    print(f"Verzeichnis für die Suche nach MSG-Dateien: {TARGET_DIRECTORY}")
    print(f"Rekursive Suche? {RECURSIVE_SEARCH}")
    print(f"Bei Bedarf in der Tabelle der bekannten Email-Absender? {USE_KNOWNSENDER_FILE}")
    if USE_KNOWNSENDER_FILE: print(f"Pfad zur Datei der bekannten Email-Absender: {KNOWNSENDER_FILE}")
    print(f"Zeitstempel der MSG-Dateien anpassen? {SET_FILEDATE}")
    print(f"PDF-Dateien erzeugen? {GENERATE_PDF}")
    if GENERATE_PDF: print(f"Existierende PDF-Dateien überschreiben? {OVERWRITE_PDF}")
    print(f"Debug-Mode? {DEBUG_MODE}")
    print(f"Debug-Datei: {prog_log_file_path}")
    print(f"Excel-Log-Datei: {excel_log_file_path}")

    # Schreibe Ergebnis Sheet
    entry = [
        { "Konfiguration": "Testlauf?", "Wert": TEST_RUN },
        { "Konfiguration": "Testverzeichnis initialisieren?", "Wert": INIT_TESTDATA },
        { "Konfiguration": "Verzeichnis für die Suche nach MSG-Dateien", "Wert": TARGET_DIRECTORY },
        { "Konfiguration": "Bei Bedarf in der Tabelle der bekannten Email-Absender?", "Wert": RECURSIVE_SEARCH },
        { "Konfiguration": "Rekursive Suche?", "Wert": USE_KNOWNSENDER_FILE },
        { "Konfiguration": "Pfad zur Datei der bekannten Email-Absender", "Wert": KNOWNSENDER_FILE }
    ]
    log_entry_neu(excel_log_file_path, entry, sheet_name="Konfiguration")

    # Ausgabe der Ergebnisse
    print(f"\nErgebnisse (auch bei Testlauf):")
    print(f"*******************************")
    print(f"Anzahl der gefundenen MSG-Dateien: {msg_file_count}")
    print(f"Anzahl der bereits mit korrekten Namen existierenden MSG-Dateien: {msg_file_same_name_count}")
    print(f"Anzahl der Dateien mit Problemen: {msg_file_problem_count}")
    print(f"Anzahl gekürzte Dateinamen: {msg_file_shorted_name_count}")

    if not TEST_RUN:
        print(f"\nErgebnisse der Anpassungen:")
        print(f"***************************")
        print(f"Anzahl der umbenannten Dateien: {msg_file_renamed_count}")
        print(f"Anzahl gefundener Doubletten: {msg_file_doublette_count}")
        print(f"Anzahl gelöschter Doubletten: {msg_file_doublette_deleted_count}")
        print(f"Anzahl nicht gelöschter Doubletten: {msg_file_doublette_deleted_problem_count}")

        if SET_FILEDATE:
            print(f"\nErgebnisse bei Anpassung File-Datum:")
            print(f"************************************")
            print(f"Anzahl MSG-Dateien mit geändertem Erstellungsdatum: {msg_file_file_creation_date_count}")
            print(f"Anzahl MSG-Dateien mit nicht geändertem Erstellungsdatum: {msg_file_creation_date_unchanged_count}")
            print(f"Anzahl MSG-Dateien wo Anpassung Erstellungsdatum nicht möglich: {msg_file_creation_date_problem_count}")
            print(f"Anzahl MSG-Dateien mit geändertem Änderungsdatum: {msg_file_modification_date_count}")
            print(f"Anzahl MSG-Dateien mit nicht geändertem Änderungssdatum: {msg_file_modification_date_unchanged_count}")
            print(f"Anzahl MSG-Dateien wo Anpassung Änderungsdatum nicht möglich: {msg_file_modification_date_problem_count}")

        if GENERATE_PDF:
            print(f"\nErgebniss der PDF-Erzeugung:")
            print(f"****************************")
            print(f"Anzahl der erzeugten PDF-Dateien: {pdf_file_generated}")
            print(f"Anzahl der übersprungenen PDF-Dateien: {pdf_file_skipped}")

    # Aufräumen
    #sys.stdout.close()
    #sys.stdout = original_stdout  # Optional: Wiederherstellen für spätere Nutzung

    logging.shutdown()
