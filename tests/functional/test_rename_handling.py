"""
test_rename_handling.py
"""
import os
import logging
from modules.msg_generate_new_filename import generate_new_msg_filename
from utils.file_handling import (rename_file2, test_file_access2, FileAccessStatus, set_file_creation_date, set_file_date)
from modules.msg_handling import create_log_file, log_entry
from utils.testset_preparation import prepare_test_directory
import datetime

# Konfiguration
TEST_RUN = True # Testlauf ohne Dateioperationen aktivieren
INIT_TESTDATA = True # Testdaten initialisieren
SET_FILEDATE = True # Dateidatum setzen

# Log-Verzeichnis und Basisname für Logdateien festlegen
LOG_DIRECTORY = r'D:\Dev\pycharm\MSGFileRenamer\logs'
PROG_LOG_NAME = 'test_rename_prog_log.log'
prog_log_file_path = os.path.join(LOG_DIRECTORY, PROG_LOG_NAME)

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
    #SOURCE_DIRECTORY_TEST_DATA = r'D:\Dev\pycharm\MSGFileRenamer\data\sample_files\testset-short'
    SOURCE_DIRECTORY_TEST_DATA = r'D:\Dev\pycharm\MSGFileRenamer\data\sample_files\testset-long'
    TARGET_DIRECTORY_TEST_DATA = r'D:\Dev\pycharm\MSGFileRenamer\tests\functional\testdir'
    DEFAULT_EXCEL_LOG_DIRECTORY = (r"D:\Dev\pycharm\MSGFileRenamer\logs")
    DEFAULT_EXCEL_LOG_BASE_NAME = 'msg_log'

    # Log-Verzeichnis und Basisname für Logdateien festlegen
    excel_log_directory = DEFAULT_EXCEL_LOG_DIRECTORY
    excel_log_base_name = DEFAULT_EXCEL_LOG_BASE_NAME
    LOG_TABLE_HEADER = ["Fortlaufende Nummer", "Verzeichnisname", "Filename"]

    # Logdatei erstellen und den Pfad ausgeben
    excel_log_file_path = create_log_file(excel_log_base_name, excel_log_directory, LOG_TABLE_HEADER)
    print(f"Logdatei erstellt: {excel_log_file_path}")

    # Wenn TARGET_DIRECTORY_TEST_DATA noch nicht existiert, dann erstellen
    if not os.path.isdir(TARGET_DIRECTORY_TEST_DATA):
        os.makedirs(TARGET_DIRECTORY_TEST_DATA, exist_ok=True)

    # Bereite das Zielverzeichnis vor und überprüfe den Erfolg
    success = prepare_test_directory(SOURCE_DIRECTORY_TEST_DATA, TARGET_DIRECTORY_TEST_DATA)
    if not success:
        print("Fehler: Die Vorbereitung des Testverzeichnisses ist fehlgeschlagen. Das Programm wird abgebrochen.")
        logging.error("Fehler: Die Vorbereitung des Testverzeichnisses ist fehlgeschlagen. Das Programm wird abgebrochen.")  # Debugging-Ausgabe: Log-File
        exit(1)  # Programm abbrechen

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
                access_result = test_file_access2(path_and_file_name)

                print(f"\tÜberprüfung Zugriff auf aktuelle MSG-Date: {[s.value for s in access_result]}'")  # Debugging-Ausgabe: Console
                logging.debug(f"\tÜberprüfung Zugriff auf aktuelle MSG-Date: {[s.value for s in access_result]}'")  # Debugging-Ausgabe: Log-File

                # Nur wenn die MSG-Datei schreibend geöffnet werden kann, ist ein Umbenennen möglich
                if FileAccessStatus.WRITABLE in access_result:
                    print(f"\tSchreibender Zugriff auf die Datei möglich: {filename}")  # Debugging-Ausgabe: Console
                    logging.debug(f"\tSchreibender Zugriff auf die Datei möglich: {filename}")  # Debugging-Ausgabe: Log-File

                    # Neuen Datenamen erzeugen
                    new_msg_filename_collection = generate_new_msg_filename(path_and_file_name)
                    print(f"\tErgebnis von 'new_msg_filename_collection': '{new_msg_filename_collection}'")

                    # Wenn Dateiname gekürzt wurde, dann Zähler erhöhen
                    if new_msg_filename_collection.is_msg_filename_truncated: msg_file_shorted_name_count += 1

                    # Überprüfen, ob new_msg_filename_collection nicht Leer (True) ist
                    if new_msg_filename_collection.new_truncated_msg_filename:

                        # Alter und neuer Name
                        old_path_and_file_name = path_and_file_name
                        new_path_and_file_name = os.path.join(pathname, new_msg_filename_collection.new_truncated_msg_filename)

                        # Prüfen, ob Alter und neuer Name gleich, dann keine Änderung erforderlich
                        if old_path_and_file_name == new_path_and_file_name:
                            print(f"Alter und neuer Dateiname sind gleich: '{filename}'")
                            logging.debug(f"Alter und neuer Dateiname sind gleich: '{filename}'")  # Debugging-Ausgabe: Log-File
                            msg_file_same_name_count += 1  # Erfolgszähler erhöhen
                        else:
                            # Prüfen, ob die Datei mit neuem Namen bereits existiert, also Doublette
                            if os.path.exists(new_path_and_file_name):
                                print(f"Datei ist eine Doublette: '{filename}'")
                                logging.debug(f"Datei ist eine Doublette: '{filename}'")  # Debugging-Ausgabe: Log-File
                                msg_file_doublette_count += 1  # Erfolgszähler erhöhen

                                # Versuche Doublette zu löschen
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
                                # Umbenennen der MSG-Datei
                                rename_msg_file_result = rename_file2(old_path_and_file_name, new_path_and_file_name)

                                # rename_msg_file_result gleich "Datei erfolgreich umbenannt"
                                if rename_msg_file_result.SUCCESS:
                                    print(f"Erfolgreiche Umbenennung der Datei: '{filename}'")
                                    logging.debug(f"Erfolgreiche Umbenennung der Datei: '{filename}'")  # Debugging-Ausgabe: Log-File
                                    msg_file_renamed_count += 1  # Erfolgszähler erhöhen

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
                                        if not set_creation_result.startswith("Fehler"):
                                            msg_file_file_creation_date_count += 1
                                            print(f"Erfolgreiche Änderung des Erstellungsdatum der Datei: '{filename}'")
                                            logging.debug(f"Erfolgreiche Änderung des Erstellungsdatum der Datei: '{filename}'")  # Debugging-Ausgabe: Log-File
                                        else:
                                            msg_file_creation_date_problem_count += 1
                                            print(f"Fehler beim Ändern des Erstellungsdatum der Datei: '{filename}'")
                                            logging.error(f"Fehler beim Ändern des Erstellungsdatum der Datei: '{filename}' wegen {set_creation_result}")  # Debugging-Ausgabe: Log-File

                                        # Jetzt das Änderungsdatum auf das gleiche Datum
                                        set_modification_result = set_file_date(new_path_and_file_name, datetime_stamp_str)
                                        if not set_creation_result.startswith("Fehler"):
                                            msg_file_modification_date_count += 1
                                            print(f"Erfolgreiche Änderung des Änderungsdatums der Datei: '{filename}'")
                                            logging.debug(f"Erfolgreiche Änderung des Änderungsdatums der Datei: '{filename}'")  # Debugging-Ausgabe: Log-File
                                        else:
                                            msg_file_modification_date_problem_count += 1
                                            print(f"Fehler beim Ändern des Änderungsdatums der Datei: '{filename}'")
                                            logging.error(f"Fehler beim Ändern des Änderungsdatums der Datei: '{filename}' wegen {set_creation_result}")  # Debugging-Ausgabe: Log-File

                                elif rename_msg_file_result.DESTINATION_EXISTS:
                                    print(f"Datei ist eine Doublette: '{filename}'")
                                    logging.debug(f"Datei ist eine Doublette: '{filename}'")  # Debugging-Ausgabe: Log-File
                                    msg_file_doublette_count += 1  # Erfolgszähler erhöhen
                                else:
                                    print(f"Umbenennen der Datei '{filename}' fehlgeschlagen: '{rename_msg_file_result}'")
                                    logging.debug(f"Umbenennen der Datei '{filename}' fehlgeschlagen: '{rename_msg_file_result}'")  # Debugging-Ausgabe: Log-File
                                    msg_file_problem_count += 1  # Problemzähler erhöhen

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
    print(f"\nAnzahl der gefundenen MSG-Dateien: {msg_file_count}")
    print(f"Anzahl der umbenannten Dateien: {msg_file_renamed_count}")
    print(f"Anzahl der Dateien mit Problemen: {msg_file_problem_count}")
    print(f"Anzahl gekürzte Dateinamen: {msg_file_shorted_name_count}")
    print(f"Anzahl gefundener Doubletten: {msg_file_doublette_count}")
    print(f"Anzahl gelöschter Doubletten: {msg_file_doublette_deleted_count}")
    print(f"Anzahl nicht gelöschter Doubletten: {msg_file_doublette_deleted_problem_count}")
    print(f"Anzahl geänderter Erstellungsdaten: {msg_file_file_creation_date_count}")
    print(f"Anzahl nicht geänderter Erstellungsdaten: {msg_file_creation_date_problem_count}")
    print(f"Anzahl geänderter Änderungsdaten: {msg_file_modification_date_count}")
    print(f"Anzahl nicht geänderter Änderungsdaten: {msg_file_modification_date_problem_count}")
    print(f"Alter und neuer Dateiname sind gleich: {msg_file_same_name_count}")
