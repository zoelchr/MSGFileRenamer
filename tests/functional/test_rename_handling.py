"""
test_rename_handling.py

Dieses Modul führt funktionale Tests für das Umbenennen von MSG-Dateien durch. Es überprüft die Fähigkeit, MSG-Dateien basierend auf ihren Metadaten umzubenennen und protokolliert die Ergebnisse. Das Modul umfasst die Initialisierung von Testdaten, das Einrichten von Logdateien und das Verarbeiten von Dateien in einem Testverzeichnis.

Funktionen und Prozesse:
- setup_logging(debug=False): Konfiguriert das Logging für Debugging und Informationsausgaben.
- Hauptprozess: Durchsucht ein Zielverzeichnis nach MSG-Dateien, überprüft den Datei-Zugriff, generiert neue Dateinamen, benennt Dateien um und setzt Erstellungs- und Änderungsdaten.
- Protokollierung: Erstellt Logeinträge für jede verarbeitete Datei und speichert sie in einer Excel-Logdatei.
- Zähler: Hält verschiedene Zähler für gefundene, umbenannte, problematische Dateien und andere relevante Metriken.

Verwendung:
Führen Sie dieses Modul aus, um die Funktionalität des Umbenennens von MSG-Dateien zu testen und die Ergebnisse zu protokollieren. Stellen Sie sicher, dass die Verzeichnisse und Konfigurationen korrekt eingestellt sind, bevor Sie den Testlauf starten.
"""

import os
import logging
from modules.msg_generate_new_filename import generate_new_msg_filename
from utils.file_handling import (rename_file, test_file_access, FileAccessStatus, set_file_creation_date, set_file_modification_date, FileOperationResult)
from modules.msg_handling import create_log_file, log_entry
from utils.testset_preparation import prepare_test_directory
import datetime

# Konfiguration
TEST_RUN = False # Testlauf ohne Dateioperationen aktivieren
INIT_TESTDATA = True # Testdaten initialisieren
SET_FILEDATE = True # Dateidatum setzen

# Log-Verzeichnis und Basisname für Logdateien festlegen
EXCEL_LOG_DIRECTORY = r'D:\Dev\pycharm\MSGFileRenamer\logs'
EXCEL_PROG_LOG_NAME = 'test_rename_prog_log.log'
prog_log_file_path = os.path.join(EXCEL_LOG_DIRECTORY, EXCEL_PROG_LOG_NAME)

# Logging für Test und Debugging konfigurieren
DEBUG_MODE=True # Debugging-Modus aktivieren

def setup_logging(debug=False):
    log_level=logging.DEBUG if debug else logging.INFO
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(name)s - %(funcName)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
        filename=prog_log_file_path
    )

if __name__ == '__main__':
    # Logging für Test und Debugging konfigurieren
    setup_logging(DEBUG_MODE)
    logger = logging.getLogger(prog_log_file_path)
    # Logging für die Programm-Ausführung
    logger.info("Programm gestartet")
    logger.info(f"Debug-Modus: {DEBUG_MODE}")
    logging.info(f"Programm Logdatei: {prog_log_file_path}")

    # Maximal zulässige Pfadlänge für Windows 11
    max_path_length = 260

    # Verzeichnisse für die Tests definieren
    SOURCE_DIRECTORY_TEST_DATA = r'D:\Dev\pycharm\MSGFileRenamer\data\sample_files\testset-short'
    #SOURCE_DIRECTORY_TEST_DATA = r'D:\Dev\pycharm\MSGFileRenamer\data\sample_files\testset-long'
    TARGET_DIRECTORY_TEST_DATA = r'D:\Dev\pycharm\MSGFileRenamer\tests\functional\testdir'

    TARGET_DIRECTORY = ""

    # Verzeichnis für die Suche festlegen
    if not TARGET_DIRECTORY:
        TARGET_DIRECTORY = TARGET_DIRECTORY_TEST_DATA
    print(f"Verzeichnis für die Suche: {TARGET_DIRECTORY}") # Debugging-Ausgabe: Console
    logging.debug(f"Verzeichnis für die Suche: {TARGET_DIRECTORY}")  # Debugging-Ausgabe: Log-File

    # Log-Verzeichnis und Basisname für Excel-Logdateien festlegen
    excel_log_directory = EXCEL_LOG_DIRECTORY
    excel_log_base_name = EXCEL_PROG_LOG_NAME
    LOG_TABLE_HEADER = ["Fortlaufende Nummer", "Verzeichnisname", "Original-Filename"]

    # Excel-Logdatei erstellen und den Pfad ausgeben
    excel_log_file_path = create_log_file(excel_log_base_name, excel_log_directory, LOG_TABLE_HEADER)
    print(f"Excel-Logdatei erstellt: {excel_log_file_path}")

    # Abhängig von INIT_TESTDATA wird das Testverzeichnis vorbereitet oder nicht
    print(f"\nSoll das Testverzeichnis vorbereitet werden? {INIT_TESTDATA}") # Debugging-Ausgabe: Console
    logging.debug(f"Soll das Testverzeichnis vorbereitet werden? {INIT_TESTDATA}")  # Debugging-Ausgabe: Log-File

    if INIT_TESTDATA:
        print(f"\tPrüfung ob Zielverzeichnis für Testdaten bereits existiert: {TARGET_DIRECTORY_TEST_DATA}") # Debugging-Ausgabe: Console
        logging.debug(f"\tPrüfung ob Zielverzeichnis für Testdaten bereits existiert: {TARGET_DIRECTORY_TEST_DATA}")  # Debugging-Ausgabe: Log-File

        # Wenn TARGET_DIRECTORY_TEST_DATA noch nicht existiert, dann erstellen
        if not os.path.isdir(TARGET_DIRECTORY_TEST_DATA):
            os.makedirs(TARGET_DIRECTORY_TEST_DATA, exist_ok=True)
            print(f"\tZielverzeichnis neu erstellt: {TARGET_DIRECTORY_TEST_DATA}")
        else:
            print(f"\tZielverzeichnis existiert bereits: {TARGET_DIRECTORY_TEST_DATA}")

        # Bereite das Zielverzeichnis vor und überprüfe den Erfolg
        success = prepare_test_directory(SOURCE_DIRECTORY_TEST_DATA, TARGET_DIRECTORY_TEST_DATA)
        if success:
            print("\tVorbereitung des Testverzeichnisses erfolgreich abgeschlossen.") # Debugging-Ausgabe: Console
            logging.error("\tVorbereitung des Testverzeichnisses erfolgreich abgeschlossen.")  # Debugging-Ausgabe: Log-File
        else:
            print("\tFehler: Die Vorbereitung des Testverzeichnisses ist fehlgeschlagen. Das Programm wird abgebrochen.") # Debugging-Ausgabe: Console
            logging.error("\tFehler: Die Vorbereitung des Testverzeichnisses ist fehlgeschlagen. Das Programm wird abgebrochen.")  # Debugging-Ausgabe: Log-File
            exit(1)  # Programm abbrechen
    else:
        print(f"\tKein Testverzeichnis vorbereitet.") # Debugging-Ausgabe: Console
        logging.debug(f"\tKein Testverzeichnis vorbereitet.")  # Debugging-Ausgabe: Log-File

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
    msg_file_modification_date_count = 0
    msg_file_modification_date_problem_count = 0
    msg_file_same_name_count = 0

    # Durchsuchen des Zielverzeichnisses nach MSG-Dateien
    # pathname = Verzeichnisname, dirs = Unterverzeichnisse, files = List von Dateien
    for pathname, dirs, files in os.walk(TARGET_DIRECTORY_TEST_DATA):
        # filename = Dateiname
        for filename in files:
            logging.debug(f"\n******************************************************'")  # Debugging-Ausgabe: Log-File

            # Überprüfen, ob die Datei die Endung .msg hat
            if filename.lower().endswith('.msg'):
                print(f"\nAktuelle MSG-Datei: '{filename}'")  # Debugging-Ausgabe: Console
                logging.debug(f"\nAktuelle MSG-Datei: '{filename}'")  # Debugging-Ausgabe: Log-File

                # Absoluter Pfadname der MSG-Datei
                path_and_file_name = os.path.join(pathname, filename)

                print(f"\tVollständiger Pfad der MSG-Date: '{path_and_file_name}'")  # Debugging-Ausgabe: Console
                logging.debug(f"\tVollständiger Pfad der MSG-Datei: '{path_and_file_name}'")  # Debugging-Ausgabe: Log-File

                msg_file_count += 1 # Zähler erhöhen, MSG-Datei gefunden

                # Überprüfen den Schreib- und Lesezugriff auf die MSG-Datei
                access_result = test_file_access(path_and_file_name)

                print(f"\tÜberprüfung Zugriff auf aktuelle MSG-Date: {[s.value for s in access_result]}'")  # Debugging-Ausgabe: Console
                logging.debug(f"\tÜberprüfung Zugriff auf aktuelle MSG-Date: {[s.value for s in access_result]}'")  # Debugging-Ausgabe: Log-File

                # Nur wenn die MSG-Datei schreibend geöffnet werden kann, ist ein Umbenennen möglich
                if FileAccessStatus.WRITABLE in access_result:
                    print(f"\tSchreibender Zugriff auf die Datei möglich: {filename}")  # Debugging-Ausgabe: Console
                    logging.debug(f"\tSchreibender Zugriff auf die Datei möglich: {filename}")  # Debugging-Ausgabe: Log-File

                    # Neuen Datenamen erzeugen
                    new_msg_filename_collection = generate_new_msg_filename(path_and_file_name)

                    # Wenn Dateiname gekürzt wurde, dann Zähler erhöhen
                    if new_msg_filename_collection.is_msg_filename_truncated: msg_file_shorted_name_count += 1

                    # Überprüfen, ob new_msg_filename_collection nicht Leer (True) ist
                    if new_msg_filename_collection.new_truncated_msg_filename:

                        # Datei für Anpassung Erstellungs- und Änderungsdatum verfügbar?
                        is_msg_file_for_change_date_available = False

                        # Alter und neuer Name
                        old_path_and_file_name = path_and_file_name
                        new_file_name = new_msg_filename_collection.new_truncated_msg_filename
                        new_path_and_file_name = os.path.join(pathname, new_msg_filename_collection.new_truncated_msg_filename)

                        # Prüfen, ob Alter und neuer Name gleich, dann keine Änderung erforderlich
                        if old_path_and_file_name == new_path_and_file_name:
                            print(f"\tAlter und neuer Dateiname sind gleich: '{filename}'")
                            logging.debug(f"Alter und neuer Dateiname sind gleich: '{filename}'")  # Debugging-Ausgabe: Log-File
                            msg_file_same_name_count += 1  # Erfolgszähler erhöhen
                            is_msg_file_for_change_date_available = True # Kennzeichnung für Anpassung Erstellungs- und Änderungsdatum
                        else:
                            # Prüfen, ob die Datei mit neuem Namen bereits existiert, also Doublette
                            if os.path.exists(new_path_and_file_name):
                                print(f"Datei ist eine Doublette: '{filename}'")
                                logging.debug(f"Datei ist eine Doublette: '{filename}'")  # Debugging-Ausgabe: Log-File
                                msg_file_doublette_count += 1  # Erfolgszähler erhöhen
                                is_msg_file_for_change_date_available = True  # Kennzeichnung für Anpassung Erstellungs- und Änderungsdatum

                                # Versuche Doublette zu löschen, wenn nicht Test
                                if not TEST_RUN:
                                    # Versuche, die Datei zu löschen
                                    try:
                                        os.remove(old_path_and_file_name)  # Versuche, die Datei zu löschen
                                        print(f"Doublette gelöscht: '{filename}'")
                                        logging.debug(f"Doublette gelöscht: '{filename}'")  # Debugging-Ausgabe: Log-File
                                        msg_file_doublette_deleted_count += 1  # Löschzähler erhöhen
                                    except Exception as e:
                                        print(f"Doublette konnte nicht gelöscht werden: '{filename}'. Fehler: {str(e)}")
                                        logging.error(f"Doublette konnte nicht gelöscht werden: '{filename}'. Fehler: {str(e)}")  # Debugging-Ausgabe: Log-File
                                        msg_file_doublette_deleted_problem_count += 1  # Problemzähler erhöhen
                            else:
                                # Wenn kein Testlauf
                                if (not TEST_RUN):

                                    # Umbenennen der MSG-Datei
                                    rename_msg_file_result = rename_file(old_path_and_file_name, new_path_and_file_name)

                                    # rename_msg_file_result gleich "Datei erfolgreich umbenannt"
                                    if rename_msg_file_result.SUCCESS:
                                        print(f"\tErfolgreiche Umbenennung der Datei '{filename}' in '{new_file_name}'")
                                        logging.debug(f"Erfolgreiche Umbenennung der Datei '{filename}' in '{new_file_name}'")  # Debugging-Ausgabe: Log-File
                                        msg_file_renamed_count += 1  # Erfolgszähler erhöhen
                                        is_msg_file_for_change_date_available = True  # Kennzeichnung für Anpassung Erstellungs- und Änderungsdatum
                                    elif rename_msg_file_result.DESTINATION_EXISTS:
                                        print(f"\tDatei ist eine Doublette: '{filename}'")
                                        logging.debug(f"Datei ist eine Doublette: '{filename}'")  # Debugging-Ausgabe: Log-File
                                        msg_file_doublette_count += 1  # Erfolgszähler erhöhen
                                        is_msg_file_for_change_date_available = True  # Kennzeichnung für Anpassung Erstellungs- und Änderungsdatum
                                    else:
                                        print(f"\tUmbenennen der Datei '{filename}' fehlgeschlagen: '{rename_msg_file_result}'")
                                        logging.debug(f"Umbenennen der Datei '{filename}' fehlgeschlagen: '{rename_msg_file_result}'")  # Debugging-Ausgabe: Log-File
                                        msg_file_problem_count += 1  # Problemzähler erhöhen

                        # Wenn die Datei erfolgreich umbenannt wurde oder die Datei bereits mit korrekten Namen existiert und kein Testlauf durchgeführt wird
                        # dann soll das Erstellungsdatum auf das Versanddatum gesetzt werden
                        if is_msg_file_for_change_date_available and (not TEST_RUN) and SET_FILEDATE:
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
                                    print(f"\tNeues Erstellungsdatum für '{new_file_name}' erfolgreich gesetzt.")  # Ausgabe des Ergebnisses
                                    logging.debug(f"Neues Erstellungsdatum für '{new_file_name}' erfolgreich gesetzt.")  # Debugging-Ausgabe: Log-File
                                else:
                                    msg_file_creation_date_problem_count += 1
                                    print(
                                        f"\tFehler beim Setzen des Erstellungsdatum für '{new_file_name}': '{set_creation_result}'")  # Ausgabe des Ergebnisses
                                    logging.debug(
                                        f"Fehler beim Setzen des Erstellungsdatum für '{new_file_name}': '{set_creation_result}'")  # Debugging-Ausgabe: Log-File

                                # Setze das Änderungsdatum auf das Versanddatum
                                set_modification_result = set_file_modification_date(new_path_and_file_name, datetime_stamp_str)
                                if set_modification_result == FileOperationResult.SUCCESS:
                                    msg_file_modification_date_count += 1
                                    print(f"\tNeues Änderungsdatum für '{new_file_name}' erfolgreich gesetzt.")  # Ausgabe des Ergebnisses
                                    logging.debug(f"Neues Änderungsdatum für '{new_file_name}' erfolgreich gesetzt.")  # Debugging-Ausgabe: Log-File
                                else:
                                    msg_file_modification_date_problem_count += 1
                                    print(
                                        f"\tFehler beim Setzen des Änderungsdatum für '{new_file_name}': '{set_creation_result}'")  # Ausgabe des Ergebnisses
                                    logging.debug(
                                        f"Fehler beim Setzen des Änderungsdatum für '{new_file_name}': '{set_creation_result}'")  # Debugging-Ausgabe: Log-File

                elif FileAccessStatus.READABLE in access_result:
                    print(f"\tNur lesender Zugriff auf die Datei möglich.")
                    msg_file_problem_count += 1  # Problemzähler erhöhen

                else:
                    print(f"\tWeder lesender noch schreibender Zugriff auf die Datei möglich: '{[s.value for s in access_result]}'")
                    msg_file_problem_count += 1  # Problemzähler erhöhen

                # Logeintrag erstellen
                entry = {
                    "Fortlaufende Nummer": msg_file_count,
                    "Verzeichnisname": pathname,
                    "Original-Filename": filename,
                    "Versanddatum": new_msg_filename_collection.datetime_stamp,
                    "Formatiertes Versanddatum": new_msg_filename_collection.formatted_timestamp,
                    "Gefundener Absender": new_msg_filename_collection.sender_name,
                    "Gefundener Email-Absender": new_msg_filename_collection.sender_email,
                    "Betreff": new_msg_filename_collection.msg_subject,
                    "Bereinigter Betreff": new_msg_filename_collection.msg_subject_sanitized,
                    "Neuer Dateiname": new_msg_filename_collection.new_msg_filename,
                    "Neuer gekürzter Dateiname": new_msg_filename_collection.new_truncated_msg_filename,
                    "Kürzung Dateiname erforderlich": new_msg_filename_collection.is_msg_filename_truncated
                }
            else:
                # Verkürzter Logeintrag erstellen, generate_new_msg_filename kein Ergebnis liefert
                entry = {
                    "Fortlaufende Nummer": msg_file_count,
                    "Verzeichnisname": pathname,
                    "Original-Filename": filename
                }

            # Eintrag ins Logfile hinzufügen
            log_entry(excel_log_file_path, entry)

    # Ausgabe der Ergebnisse
    print(f"\nTestlauf: {TEST_RUN}")
    print(f"Testverzeichnis initialisieren: {INIT_TESTDATA}")
    print(f"Zeitstempel der MSG-dateien anpassen: {SET_FILEDATE}")
    print(f"\nAnzahl der gefundenen MSG-Dateien: {msg_file_count}")
    print(f"Anzahl der umbenannten Dateien: {msg_file_renamed_count}")
    print(f"Anzahl der Dateien mit Problemen: {msg_file_problem_count}")
    print(f"Anzahl gekürzte Dateinamen: {msg_file_shorted_name_count}")
    print(f"Anzahl gefundener Doubletten: {msg_file_doublette_count}")
    print(f"Anzahl gelöschter Doubletten: {msg_file_doublette_deleted_count}")
    print(f"Anzahl nicht gelöschter Doubletten: {msg_file_doublette_deleted_problem_count}")
    print(f"Anzahl MSG-Dateien mit geändertem Erstellungsdatum: {msg_file_file_creation_date_count}")
    print(f"Anzahl MSG-Dateien mit nicht geändertem Erstellungsdatum: {msg_file_creation_date_problem_count}")
    print(f"Anzahl MSG-Dateien mit geändertem Änderungsdatum: {msg_file_modification_date_count}")
    print(f"Anzahl MSG-Dateien mit nicht geändertem Änderungsdatum: {msg_file_modification_date_problem_count}")
    print(f"Alter und neuer Dateiname sind gleich: {msg_file_same_name_count}")
